# -*- coding: utf-8 -*-
"""Atualização Descrição por Keynotes - sincroniza Keynotes do modelo com o Excel.

Base única de Keynotes entre o Revit e a planilha do escritório
(OCA_XXX_Lista_de_Materiais - aba MATERIAIS).

Onde cada dado mora:
  * Keynote ........ parâmetro nativo KEYNOTE dos TIPOS usados no modelo e dos
                     MATERIAIS usados pelos elementos (opcional).
  * Descrição PT ... parâmetro nativo "Description" (tipo / material).
  * Descrição IN ... parâmetro compartilhado "Descrição IN" (tipo / material),
                     associado pelo GUID 70307f6e-3b31-4049-bed4-62dfc3570fec
                     (OCA_Parametros_Template.txt). Quando o modo grava no
                     modelo e falta em alguma categoria, oferece vincular ao
                     projeto como parâmetro de TIPO.
  * Excel .......... aba MATERIAIS: colunas KEYNOTE, DESCRIÇÃO, DESCRIÇÃO IN.

Verificação nos DOIS sentidos (tabela única, agrupada pela estrutura de
agrupamentos de Keynotes: CATEGORIA > GRUPO/prefixo, abas DADOS e KEYNOTE BÁSICO):
  * Excel -> Modelo: o que está no Excel e difere do tipo/material no modelo.
  * Modelo -> Excel: o que está no modelo e falta/difere no Excel (keynote
    novo entra no FIM da tabela com PREFIXO/SEQUENCIA + fórmulas do template).
  * Ambos: as duas listas juntas; o usuário marca o que vai para cada lado.
    Um mesmo campo do mesmo keynote não pode ir para os dois lados.
  Padrões: Excel -> Modelo marcado; Modelo -> Excel marcado só quando preenche
  vazio (sobrescrever o Excel começa desmarcado). Valor vazio nunca apaga o
  outro lado. Keynote do Excel não usado no modelo nunca é apagado.

Destaque temporário (vista ativa, desfeito ao fechar a tabela):
  Keynote + Description ............ nada
  Keynote sem Description .......... VERMELHO
  Description sem Keynote .......... VERMELHO
  Sem Keynote e sem Description .... LARANJA

TXT de Keynotes do Revit (opcional, confirma no final): gerado do Excel com a
mesma estrutura do arquivo do escritório - CATEGORIA/prefixos, PREFIXO/GRUPO/
CATEGORIA e KEYNOTE / "DESCRIÇÃO | DESCRIÇÃO IN" / PREFIXO - e carregado no
projeto.

Modos: Somente analisar (nada é gravado) · Modelo -> Excel · Excel -> Modelo ·
Ambos (simultâneo).

O Excel é lido e gravado direto no XML do .xlsx (sem Excel/COM). Antes de
gravar é feito um backup com data/hora ao lado do arquivo. O arquivo precisa
estar FECHADO no Excel.

Ao final pergunta se gera o relatório HTML (sempre gerado em "Somente
analisar"): verificação nos dois sentidos, tabela final do Excel (filtro por
agrupamento, só alterados, download), destaques e inconsistências.
"""

__title__ = "Atualização\nDescrição\npor Keynotes"
__doc__ = ("Compara os Keynotes do modelo (tipos e materiais) com a planilha "
           "Excel (aba MATERIAIS), atualiza o Excel a partir do modelo ou as "
           "descrições do modelo a partir do Excel e gera um relatório de "
           "inconsistências agrupado por categoria.")

import os
import re
import json
import codecs
from datetime import datetime
from collections import OrderedDict, defaultdict

from pyrevit import revit, DB, script, forms
from oca_ui import build_xaml, alert_title, brand_report, output_header

import clr
clr.AddReference("System.Xml")
for _asm in ("System.IO.Compression",
             "System.IO.Compression.FileSystem",   # .NET Framework (Revit <= 2024)
             "System.IO.Compression.ZipFile"):     # .NET 8 (Revit 2025+)
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
from System import Guid
from System.IO import (StreamReader, StreamWriter, File, FileMode, Path,
                       FileAccess, FileShare)
from System.IO.Compression import ZipFile, ZipArchiveMode
from System.Text import UTF8Encoding
from System.Xml import XmlDocument, XmlNamespaceManager

doc = revit.doc
BIP = DB.BuiltInParameter
INVALID_ID = DB.ElementId.InvalidElementId
HERE = os.path.dirname(__file__)

output = script.get_output()
output.set_title("Atualização Descrição por Keynotes")
ALERT_TITLE = alert_title(__title__)


# ------------------------------------------------------------------
# Configuração do escritório
# ------------------------------------------------------------------
SHEET_NAME = u"MATERIAIS"          # aba com os keynotes
DATA_SHEET_NAME = u"DADOS"         # aba PREFIXO -> GRUPO -> CATEGORIA
KN_SHEET_NAME = u"KEYNOTE BÁSICO"  # aba com CATEGORIA -> lista de prefixos (topo do TXT)
NO_GROUP = u"SEM AGRUPAMENTO"
# Parâmetro compartilhado da descrição em inglês (OCA_Parametros_Template.txt).
# A associação é feita pelo GUID - um parâmetro de mesmo nome com outro GUID
# NÃO é considerado (não aparece nas tabelas/etiquetas do template).
PARAM_IN_NAME = u"Descrição IN"
PARAM_IN_GUID_STR = "70307f6e-3b31-4049-bed4-62dfc3570fec"
PARAM_IN_GROUP = u"Informações do Modelo"        # grupo no arquivo compartilhado
PARAM_IN_GUID = Guid(PARAM_IN_GUID_STR)
MATERIALS_LABEL = u"Materiais"     # categoria usada no relatório p/ materiais

MODES = OrderedDict([
    ("analyze", u"Somente analisar"),
    ("to_excel", u"Modelo → Excel"),
    ("to_model", u"Excel → Modelo"),
    ("both", u"Ambos (simultâneo)"),
])
# sentidos de verificação incluídos em cada modo
MODE_DIRS = {"analyze": ("model", "excel"), "to_excel": ("excel",),
             "to_model": ("model",), "both": ("model", "excel")}


# ------------------------------------------------------------------
# Helpers (mesmo padrão das outras ferramentas OCA)
# ------------------------------------------------------------------
def to_unicode(value):
    """Coerce any Revit / .NET / byte value into clean unicode text."""
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    if isinstance(value, str):
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return value.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return value.decode("ascii", "replace")
    return unicode(value)


def eid_int(eid):
    """ElementId -> int (Revit 2024+ usa .Value, versões antigas .IntegerValue)."""
    raw = eid.Value if hasattr(eid, "Value") else eid.IntegerValue
    return int(raw)


def clean(text):
    """Texto de parâmetro/célula sem espaços nas pontas. O conteúdo é mantido
    exatamente (todos os dígitos/símbolos do keynote contam)."""
    return to_unicode(text).strip()


_ACCENTS = {u"Á": u"A", u"À": u"A", u"Â": u"A", u"Ã": u"A", u"Ä": u"A",
            u"É": u"E", u"È": u"E", u"Ê": u"E", u"Ë": u"E",
            u"Í": u"I", u"Ì": u"I", u"Î": u"I", u"Ï": u"I",
            u"Ó": u"O", u"Ò": u"O", u"Ô": u"O", u"Õ": u"O", u"Ö": u"O",
            u"Ú": u"U", u"Ù": u"U", u"Û": u"U", u"Ü": u"U", u"Ç": u"C"}


def norm_header(text):
    """'Descrição  IN' -> 'DESCRICAO IN' (maiúsculas, sem acento, 1 espaço)."""
    t = to_unicode(text).upper().replace(u"_", u" ").replace(u"-", u" ")
    t = u"".join(_ACCENTS.get(ch, ch) for ch in t)
    return u" ".join(t.split())


def elem_name(elem):
    try:
        return to_unicode(DB.Element.Name.__get__(elem))
    except Exception:
        try:
            return to_unicode(elem.Name)
        except Exception:
            return u""


def family_name(elem):
    try:
        return to_unicode(elem.FamilyName)
    except Exception:
        return u""


def open_in_browser(path):
    try:
        from System.Diagnostics import Process, ProcessStartInfo
        psi = ProcessStartInfo(path)
        psi.UseShellExecute = True
        Process.Start(psi)
        return True
    except Exception:
        pass
    try:
        from System.Diagnostics import Process
        Process.Start("explorer.exe", '"{}"'.format(path))
        return True
    except Exception:
        pass
    try:
        import webbrowser
        return webbrowser.open("file:///" + path.replace("\\", "/"))
    except Exception:
        return False


# ------------------------------------------------------------------
# 1. Janela de opções (padrão visual OCA - lib/oca_ui)
# ------------------------------------------------------------------
KN_BODY = u"""
  <StackPanel>
    <TextBlock Style="{StaticResource oca.Hint}" Margin="0,0,0,14"
               Text="Descrição PT = parâmetro Description; Descrição IN = parâmetro compartilhado 'Descrição IN' (GUID 70307f6e). Antes de gravar abre a tabela de verificação para marcar o que muda no modelo e o que muda no Excel."/>

    <HeaderedContentControl Header="PLANILHA EXCEL" Style="{StaticResource oca.Section}">
      <StackPanel>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBox x:Name="xlsx"/>
          <Button x:Name="browse" Grid.Column="1" Content="Procurar..." Margin="8,0,0,0"/>
        </Grid>
        <TextBlock Text=".xlsx - feche o arquivo no Excel antes de gravar." Style="{StaticResource oca.Hint}" Margin="0,4,0,0"/>
      </StackPanel>
    </HeaderedContentControl>

    <HeaderedContentControl Header="VERIFICAÇÃO / ATUALIZAÇÃO" Style="{StaticResource oca.Section}">
      <StackPanel>
        <RadioButton x:Name="m_analyze" GroupName="mode">
          <TextBlock TextWrapping="Wrap"><Run Text="Somente analisar"/><Run Text="  ·  tabela de verificação + relatório - nada é gravado" FontSize="11" Foreground="{StaticResource oca.Ink2}"/></TextBlock>
        </RadioButton>
        <RadioButton x:Name="m_to_excel" GroupName="mode">
          <TextBlock TextWrapping="Wrap"><Run Text="Modelo → Excel"/><Run Text="  ·  dados do modelo comparados com o Excel" FontSize="11" Foreground="{StaticResource oca.Ink2}"/></TextBlock>
        </RadioButton>
        <RadioButton x:Name="m_to_model" GroupName="mode">
          <TextBlock TextWrapping="Wrap"><Run Text="Excel → Modelo"/><Run Text="  ·  dados do Excel comparados com o modelo" FontSize="11" Foreground="{StaticResource oca.Ink2}"/></TextBlock>
        </RadioButton>
        <RadioButton x:Name="m_both" GroupName="mode">
          <TextBlock TextWrapping="Wrap"><Run Text="Ambos, simultâneo"/><Run Text="  ·  escolher na tabela o que vai para cada lado" FontSize="11" Foreground="{StaticResource oca.Ink2}"/></TextBlock>
        </RadioButton>
      </StackPanel>
    </HeaderedContentControl>

    <HeaderedContentControl Header="ESCOPO" Style="{StaticResource oca.Section}">
      <StackPanel>
        <RadioButton x:Name="s_model" GroupName="scope" Content="Modelo inteiro"/>
        <RadioButton x:Name="s_view" GroupName="scope" Content="Somente elementos visíveis na vista ativa"/>
        <CheckBox x:Name="materials" Content="Incluir Keynotes de materiais" Margin="0,6,0,6"/>
        <CheckBox x:Name="highlight" Content="Destacar no modelo (vista ativa) durante a verificação: vermelho / laranja"/>
      </StackPanel>
    </HeaderedContentControl>

    <HeaderedContentControl Header="TXT DE KEYNOTES DO REVIT" Style="{StaticResource oca.Section}">
      <StackPanel>
        <CheckBox x:Name="txt_on" Margin="0,0,0,8"
                  Content="Gerar o TXT a partir do Excel e carregar no projeto (confirma no final)"/>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBox x:Name="txt_path"/>
          <Button x:Name="txt_browse" Grid.Column="1" Content="Salvar como..." Margin="8,0,0,0"
                 />
        </Grid>
      </StackPanel>
    </HeaderedContentControl>
  </StackPanel>"""

KN_FOOTER = u"""
  <Button x:Name="ok" Content="Executar" Style="{StaticResource oca.Primary}"/>
  <Button x:Name="cancel" Content="Cancelar" Margin="8,0,0,0"/>"""

KN_XAML = build_xaml(title=__title__, subtitle=__doc__, body=KN_BODY,
                     footer_right=KN_FOOTER, size="M")


def default_txt_path(xlsx_path):
    """Mesmo nome/pasta da planilha (OCA_XXX_Lista_de_Materiais_0X.txt)."""
    return os.path.splitext(xlsx_path)[0] + u".txt" if xlsx_path else u""


class KeynoteWindow(forms.WPFWindow):
    def __init__(self, xaml, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.result = None
        self.xlsx.Text = cfg["xlsx"]
        {"to_excel": self.m_to_excel,
         "to_model": self.m_to_model,
         "both": self.m_both}.get(cfg["mode"], self.m_analyze).IsChecked = True
        (self.s_view if cfg["scope"] == "view" else self.s_model).IsChecked = True
        self.materials.IsChecked = bool(cfg["materials"])
        self.highlight.IsChecked = bool(cfg["highlight"])
        self.txt_on.IsChecked = bool(cfg["txt_on"])
        self.txt_path.Text = cfg["txt_path"] or default_txt_path(cfg["xlsx"])
        self.browse.Click += self._browse
        self.txt_browse.Click += self._txt_browse
        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    def _browse(self, sender, args):
        picked = forms.pick_file(file_ext="xlsx",
                                 title="Selecione a planilha de Keynotes")
        if picked:
            self.xlsx.Text = picked
            if not to_unicode(self.txt_path.Text).strip():
                self.txt_path.Text = default_txt_path(picked)

    def _txt_browse(self, sender, args):
        cur = to_unicode(self.txt_path.Text).strip() or default_txt_path(
            to_unicode(self.xlsx.Text).strip())
        kw = {"file_ext": "txt",
              "default_name": os.path.basename(cur) if cur else u"",
              "init_dir": os.path.dirname(cur) if cur else u""}
        try:
            picked = forms.save_file(title="TXT de Keynotes do Revit", **kw)
        except TypeError:                       # pyRevit antigo sem 'title'
            picked = forms.save_file(**kw)
        if picked:
            self.txt_path.Text = picked

    def _ok(self, sender, args):
        path = to_unicode(self.xlsx.Text).strip().strip('"')
        if not path or not os.path.isfile(path) or not path.lower().endswith(".xlsx"):
            forms.alert(u"Selecione um arquivo .xlsx válido.",
                        sub_msg=u"Use Procurar... para escolher a planilha de Keynotes.", title=ALERT_TITLE)
            return
        txt = to_unicode(self.txt_path.Text).strip().strip('"') or default_txt_path(path)
        if self.txt_on.IsChecked and not os.path.isdir(os.path.dirname(txt) or u"."):
            forms.alert(u"A pasta do TXT de Keynotes não existe.", sub_msg=txt, title=ALERT_TITLE)
            return
        if self.m_to_excel.IsChecked:
            mode = "to_excel"
        elif self.m_to_model.IsChecked:
            mode = "to_model"
        elif self.m_both.IsChecked:
            mode = "both"
        else:
            mode = "analyze"
        self.result = {
            "xlsx": path,
            "mode": mode,
            "scope": "view" if self.s_view.IsChecked else "model",
            "materials": bool(self.materials.IsChecked),
            "highlight": bool(self.highlight.IsChecked),
            "txt_on": bool(self.txt_on.IsChecked),
            "txt_path": txt,
        }
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.confirmed = False
        self.Close()


config = script.get_config()
saved_cfg = {"xlsx": to_unicode(config.get_option("xlsx_path", u"")),
             "mode": config.get_option("mode", "analyze"),
             "scope": config.get_option("scope", "model"),
             "materials": config.get_option("materials", True),
             "highlight": config.get_option("highlight", True),
             "txt_on": config.get_option("txt_on", False),
             "txt_path": to_unicode(config.get_option("txt_path", u""))}

win = KeynoteWindow(KN_XAML, saved_cfg)
win.ShowDialog()
if not win.confirmed:
    script.exit()
opts = win.result

config.xlsx_path = opts["xlsx"]
config.mode = opts["mode"]
config.scope = opts["scope"]
config.materials = opts["materials"]
config.highlight = opts["highlight"]
config.txt_on = opts["txt_on"]
config.txt_path = opts["txt_path"]
script.save_config()

XLSX_PATH = opts["xlsx"]
MODE = opts["mode"]

active_view = doc.ActiveView
use_view = opts["scope"] == "view"
if use_view:
    bad = set()
    for _n in ("Schedule", "DrawingSheet", "Report", "ColumnSchedule",
               "PanelSchedule", "Internal", "Undefined", "ProjectBrowser",
               "SystemBrowser"):
        _vt = getattr(DB.ViewType, _n, None)
        if _vt is not None:
            bad.add(_vt)
    if active_view is None or getattr(active_view, "IsTemplate", False) \
            or active_view.ViewType in bad:
        forms.alert(u"A vista ativa não é uma vista gráfica (tabela, folha...).",
                    sub_msg=u"Abra uma planta, corte, elevação ou 3D, ou use o escopo "
                            u"'Modelo inteiro'.", title=ALERT_TITLE, exitscript=True)


# ------------------------------------------------------------------
# 2. Leitura do Excel (XML dentro do .xlsx, sem Excel instalado)
# ------------------------------------------------------------------
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def ns_manager(xdoc):
    nsm = XmlNamespaceManager(xdoc.NameTable)
    nsm.AddNamespace("a", NS_MAIN)
    nsm.AddNamespace("r", NS_REL)
    nsm.AddNamespace("p", NS_PKG)
    return nsm


def load_xml(text):
    d = XmlDocument()
    d.PreserveWhitespace = True
    d.LoadXml(text)
    return d


def col_letters(idx):
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    idx += 1
    s = u""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = unichr(65 + rem) + s
    return s


def split_ref(ref):
    """'E12' -> (4, 12)."""
    col, num = u"", u""
    for ch in ref:
        if ch.isalpha():
            col += ch
        else:
            num += ch
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1, (int(num) if num else 0)


def zip_read(archive, name):
    entry = archive.GetEntry(name)
    if entry is None:
        return None
    reader = StreamReader(entry.Open(), UTF8Encoding(False))
    try:
        return reader.ReadToEnd()
    finally:
        reader.Dispose()


def resolve_sheets(archive):
    """Return OrderedDict norm(sheet name) -> (name, part path, position)."""
    wb = load_xml(zip_read(archive, "xl/workbook.xml"))
    rels = load_xml(zip_read(archive, "xl/_rels/workbook.xml.rels"))
    nsm_wb, nsm_rel = ns_manager(wb), ns_manager(rels)
    targets = {}
    for rel in rels.SelectNodes("//p:Relationship", nsm_rel):
        tgt = to_unicode(rel.GetAttribute("Target")).replace("\\", "/")
        if tgt.startswith("/"):
            tgt = tgt[1:]
        elif not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        targets[to_unicode(rel.GetAttribute("Id"))] = tgt
    out = OrderedDict()
    for pos, sh in enumerate(wb.SelectNodes("//a:sheets/a:sheet", nsm_wb)):
        name = to_unicode(sh.GetAttribute("name"))
        rid = to_unicode(sh.GetAttribute("id", NS_REL))
        out[norm_header(name)] = (name, targets.get(rid, u""), pos)
    return out


def read_shared_strings(archive):
    text = zip_read(archive, "xl/sharedStrings.xml")
    if not text:
        return []
    d = load_xml(text)
    nsm = ns_manager(d)
    return [u"".join(to_unicode(t.InnerText) for t in si.SelectNodes(".//a:t", nsm))
            for si in d.SelectNodes("//a:si", nsm)]


def parse_sheet(xml_text, shared):
    """-> (grid {row: {col: text}}, rows_with_formula set)."""
    d = load_xml(xml_text)
    nsm = ns_manager(d)
    grid, formula_rows = {}, set()
    for c in d.SelectNodes("//a:sheetData/a:row/a:c", nsm):
        ref = to_unicode(c.GetAttribute("r"))
        if not ref:
            continue
        ci, rn = split_ref(ref)
        if c.SelectSingleNode("a:f", nsm) is not None:
            formula_rows.add(rn)
        ctype = to_unicode(c.GetAttribute("t"))
        val = u""
        if ctype == u"inlineStr":
            val = u"".join(to_unicode(t.InnerText) for t in c.SelectNodes(".//a:t", nsm))
        else:
            v = c.SelectSingleNode("a:v", nsm)
            if v is not None:
                raw = to_unicode(v.InnerText)
                if ctype == u"s":
                    try:
                        val = shared[int(raw)]
                    except Exception:
                        val = u""
                else:
                    val = raw
        if val != u"":
            grid.setdefault(rn, {})[ci] = val
    return grid, formula_rows


HEADER_ALIASES = {
    "key": [u"KEYNOTE", u"KEYNOTES", u"NOTA CHAVE", u"KEY"],
    "pt": [u"DESCRICAO", u"DESCRICAO PT", u"DESCRICAO PORTUGUES", u"DESCRIPTION",
           u"DESCRIPTION PT"],
    "en": [u"DESCRICAO IN", u"DESCRICAO EM IN", u"DESCRICAO EN", u"DESCRICAO INGLES",
           u"DESCRIPTION IN", u"DESCRIPTION EN"],
    "cat": [u"CATEGORIA"],
    "grp": [u"GRUPO"],
    "prefix": [u"PREFIXO"],
    "seq": [u"SEQUENCIA"],
    "obs": [u"OBSERVACOES", u"OBSERVACAO"],
}


def read_excel(path):
    archive = ZipFile.OpenRead(path)
    try:
        sheets = resolve_sheets(archive)
        shared = read_shared_strings(archive)

        # aba MATERIAIS; se não existir, primeira aba que tenha coluna KEYNOTE
        candidates = []
        if norm_header(SHEET_NAME) in sheets:
            candidates.append(sheets[norm_header(SHEET_NAME)])
        candidates += [s for k, s in sheets.items() if k != norm_header(SHEET_NAME)]

        chosen = None
        for name, part, pos in candidates:
            text = zip_read(archive, part) if part else None
            if not text:
                continue
            grid, frows = parse_sheet(text, shared)
            for rn in sorted(grid.keys())[:15]:
                heads = dict((ci, norm_header(v)) for ci, v in grid[rn].items())
                if u"KEYNOTE" in heads.values() or u"NOTA CHAVE" in heads.values():
                    chosen = (name, part, pos, grid, frows, rn, heads)
                    break
            if chosen:
                break
        if chosen is None:
            raise Exception(u"Nenhuma aba com a coluna KEYNOTE foi encontrada "
                            u"(procurada primeiro a aba '{}').".format(SHEET_NAME))

        # estrutura de agrupamentos (mesma do TXT de Keynotes do escritório):
        #   DADOS          A=PREFIXO  B=GRUPO  C=CATEGORIA   (linhas "AC  ACESSÓRIOS  ACESSÓRIOS")
        #   KEYNOTE BÁSICO A=CATEGORIA B=prefixos          (linhas "ACESSÓRIOS  AC, DS, ...")
        # valores BRUTOS guardados para o TXT sair idêntico (inclui espaços finais)
        prefix_map, dados_rows, kn_headers = {}, [], []
        dados = sheets.get(norm_header(DATA_SHEET_NAME))
        if dados and dados[1]:
            dtext = zip_read(archive, dados[1])
            if dtext:
                dgrid, _ = parse_sheet(dtext, shared)
                for rn in sorted(dgrid.keys()):
                    cells = dgrid[rn]
                    pf = clean(cells.get(0, u""))
                    if not pf:
                        continue
                    dados_rows.append((to_unicode(cells.get(0, u"")),
                                       to_unicode(cells.get(1, u"")),
                                       to_unicode(cells.get(2, u""))))
                    if pf not in prefix_map:
                        prefix_map[pf] = (clean(cells.get(1, u"")), clean(cells.get(2, u"")))
        kb = sheets.get(norm_header(KN_SHEET_NAME))
        if kb and kb[1]:
            ktext = zip_read(archive, kb[1])
            if ktext:
                kgrid, kfrows = parse_sheet(ktext, shared)
                for rn in sorted(kgrid.keys()):
                    if rn in kfrows:            # a partir daqui são fórmulas =DADOS!...
                        break
                    a_, b_ = kgrid[rn].get(0, u""), kgrid[rn].get(1, u"")
                    if clean(a_) and clean(b_):
                        kn_headers.append((to_unicode(a_), to_unicode(b_)))
        if not kn_headers:                      # sem a aba: monta a partir de DADOS
            order = OrderedDict()
            for pf, grp, cat in dados_rows:
                order.setdefault(cat, []).append(clean(pf))
            kn_headers = [(cat, u", ".join(pfs)) for cat, pfs in order.items()]
    finally:
        archive.Dispose()

    name, part, pos, grid, frows, header_row, heads = chosen
    cols = {}
    for key, aliases in HEADER_ALIASES.items():
        for ci in sorted(heads.keys()):
            if heads[ci] in aliases:
                cols[key] = ci
                break
    missing = [k for k in ("key", "pt", "en") if k not in cols]
    if missing:
        raise Exception(u"Colunas não encontradas na aba '{}': {}".format(
            name, u", ".join({"key": u"KEYNOTE", "pt": u"DESCRIÇÃO",
                              "en": u"DESCRIÇÃO IN"}[m] for m in missing)))

    def cell(cells, key):
        ci = cols.get(key)
        return clean(cells.get(ci, u"")) if ci is not None else u""

    rows, last_row = [], header_row
    for rn in sorted(grid.keys()):
        if rn <= header_row:
            continue
        cells = grid[rn]
        rec = dict((k, cell(cells, k)) for k in HEADER_ALIASES.keys())
        rec["row"] = rn
        # KEYNOTE é fórmula CONCAT(PREFIXO:SEQUENCIA); sem valor em cache
        # (arquivo nunca recalculado) monta a partir das duas colunas.
        if not rec["key"] and rec["prefix"] and rec["seq"]:
            rec["key"] = rec["prefix"] + rec["seq"]
        if any(rec[k] for k in ("key", "pt", "en", "prefix", "seq")):
            last_row = max(last_row, rn)
        if rec["key"]:
            rows.append(rec)

    by_key, dups = OrderedDict(), defaultdict(list)
    for rec in rows:
        if rec["key"] in by_key:
            dups[rec["key"]].append(rec["row"])
        else:
            by_key[rec["key"]] = rec

    return {
        "path": path, "sheet": name, "part": part, "pos": pos,
        "header_row": header_row, "cols": cols, "rows": rows,
        "by_key": by_key, "dups": dict(dups), "last_row": last_row,
        "grid": grid, "formula_rows": frows, "prefix_map": prefix_map,
        "dados_rows": dados_rows, "kn_headers": kn_headers,
        "cat_order": [clean(c) for c, _ in kn_headers],
    }


try:
    XL = read_excel(XLSX_PATH)
except Exception as ex:
    forms.alert(u"Não foi possível ler a planilha.", sub_msg=to_unicode(ex),
                title=ALERT_TITLE, exitscript=True)


# ------------------------------------------------------------------
# 3. Keynotes do modelo (tipos usados + materiais usados)
# ------------------------------------------------------------------
def param_text(p):
    if p is None or not p.HasValue:
        return u""
    try:
        return clean(p.AsString())
    except Exception:
        return u""


def keynote_of(elem):
    try:
        return param_text(elem.get_Parameter(BIP.KEYNOTE_PARAM))
    except Exception:
        return u""


def desc_param(elem):
    """Parâmetro nativo Description (independe do idioma do Revit)."""
    try:
        p = elem.get_Parameter(BIP.ALL_MODEL_DESCRIPTION)
        if p is not None:
            return p
    except Exception:
        pass
    for nm in (u"Description", u"Descrição"):
        try:
            p = elem.LookupParameter(nm)
            if p is not None and p.StorageType == DB.StorageType.String:
                return p
        except Exception:
            pass
    return None


def in_param(elem):
    """Parâmetro compartilhado 'Descrição IN', localizado pelo GUID
    (None se o parâmetro não estiver vinculado a este tipo / material)."""
    try:
        p = elem.get_Parameter(PARAM_IN_GUID)
        if p is not None and p.StorageType == DB.StorageType.String:
            return p
    except Exception:
        pass
    return None


def in_param_clash(elem):
    """True se existe um parâmetro chamado 'Descrição IN' que NÃO é o
    compartilhado do escritório (GUID diferente / parâmetro de projeto)."""
    try:
        p = elem.LookupParameter(PARAM_IN_NAME)
        if p is None:
            return False
        return not (p.IsShared and p.GUID == PARAM_IN_GUID)
    except Exception:
        return False


def is_model_category(cat):
    try:
        return cat.CategoryType == DB.CategoryType.Model
    except Exception:
        return False


def collect_model(view_only, include_materials):
    col = (DB.FilteredElementCollector(doc, active_view.Id) if view_only
           else DB.FilteredElementCollector(doc))
    instances = col.WhereElementIsNotElementType().ToElements()

    type_info, mat_info = {}, {}      # id int -> [ElementId, count, set(categorias)]
    for e in instances:
        cat = e.Category
        if cat is None:
            continue
        cname = to_unicode(cat.Name)
        try:
            tid = e.GetTypeId()
        except Exception:
            tid = None
        if tid is not None and tid != INVALID_ID:
            info = type_info.setdefault(eid_int(tid), [tid, 0, set()])
            info[1] += 1
            info[2].add(cname)
        if include_materials and is_model_category(cat):
            seen = set()
            for paint in (False, True):
                try:
                    mids = list(e.GetMaterialIds(paint))
                except Exception:
                    continue
                for mid in mids:
                    k = eid_int(mid)
                    if k in seen:
                        continue
                    seen.add(k)
                    info = mat_info.setdefault(k, [mid, 0, set()])
                    info[1] += 1
                    info[2].add(cname)

    holders = []

    def add(kind, elem, category, count, used_in):
        key = keynote_of(elem)
        if not key:
            return
        pp, ep = desc_param(elem), in_param(elem)
        holders.append({
            "kind": kind, "elem": elem, "id": eid_int(elem.Id),
            "category": category, "family": family_name(elem) if kind == "type" else u"",
            "name": elem_name(elem), "count": count, "usedIn": sorted(used_in),
            "key": key, "pt_p": pp, "en_p": ep,
            "en_clash": ep is None and in_param_clash(elem),
            "pt": param_text(pp), "en": param_text(ep) if ep is not None else None,
        })

    for tid, count, cats in type_info.values():
        t = doc.GetElement(tid)
        if t is None:
            continue
        cname = to_unicode(t.Category.Name) if t.Category is not None else sorted(cats)[0]
        add("type", t, cname, count, cats)
    for mid, count, cats in mat_info.values():
        m = doc.GetElement(mid)
        if m is None:
            continue
        add("material", m, MATERIALS_LABEL, count, cats)

    by_key = OrderedDict()
    for h in sorted(holders, key=lambda x: (x["key"], x["category"], x["name"])):
        by_key.setdefault(h["key"], []).append(h)
    return by_key


KN_PARENT = {}        # chave -> chave pai no arquivo de keynotes carregado no Revit


def keynote_file_texts():
    """Texto (e pai) de cada chave do .txt de keynotes carregado no projeto."""
    out = {}
    try:
        table = DB.KeynoteTable.GetKeynoteTable(doc)
        for entry in table.GetKeyBasedTreeEntries():
            try:
                k = clean(entry.Key)
                out[k] = clean(entry.KeynoteText)
                KN_PARENT[k] = clean(entry.ParentKey)
            except Exception:
                continue
    except Exception:
        pass
    return out


MODEL = collect_model(use_view, opts["materials"])
FILE_TEXT = keynote_file_texts()


# ------------------------------------------------------------------
# 3b. Vincular o parâmetro compartilhado 'Descrição IN' (quando grava no modelo)
# ------------------------------------------------------------------
def in_param_missing_categories():
    """Categorias (Category) de tipos/materiais com keynote sem o parâmetro."""
    cats = OrderedDict()
    for holders in MODEL.values():
        for h in holders:
            if h["en_p"] is not None:
                continue
            try:
                if h["kind"] == "material":
                    cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Materials)
                else:
                    cat = h["elem"].Category
            except Exception:
                cat = None
            if cat is not None and getattr(cat, "AllowsBoundParameters", True):
                cats.setdefault(eid_int(cat.Id), cat)
    return list(cats.values())


def _insert_binding(defn, binding, reinsert):
    """Insert/ReInsert no grupo 'Dados de identidade' (API 2022+ e anterior)."""
    pb = doc.ParameterBindings
    fn = pb.ReInsert if reinsert else pb.Insert
    group = getattr(DB, "GroupTypeId", None)
    if group is not None:
        try:
            return fn(defn, binding, group.IdentityData)
        except Exception:
            pass
    bipg = getattr(DB, "BuiltInParameterGroup", None)
    if bipg is not None:
        try:
            return fn(defn, binding, bipg.PG_IDENTITY_DATA)
        except Exception:
            pass
    return fn(defn, binding)


SHARED_PARAM_TXT = u"\r\n".join([
    u"# This is a Revit shared parameter file.",
    u"# Do not edit manually.",
    u"*META\tVERSION\tMINVERSION",
    u"META\t2\t1",
    u"*GROUP\tID\tNAME",
    u"GROUP\t1\t" + PARAM_IN_GROUP,
    u"*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION"
    u"\tUSERMODIFIABLE\tHIDEWHENNOVALUE",
    u"PARAM\t{}\t{}\tTEXT\t\t1\t1\t\t1\t0".format(PARAM_IN_GUID_STR, PARAM_IN_NAME),
    u""])


def bind_in_param(cats):
    """Vincula 'Descrição IN' como parâmetro de TIPO às categorias informadas.
    Se o parâmetro ainda não existe no projeto, a definição é lida de um
    arquivo compartilhado temporário com o MESMO GUID do template OCA (o
    arquivo configurado no Revit é restaurado em seguida).
    -> (ok, mensagem)"""
    app = doc.Application
    catset = app.Create.NewCategorySet()
    for c in cats:
        catset.Insert(c)

    spe = None
    try:
        spe = DB.SharedParameterElement.Lookup(doc, PARAM_IN_GUID)
    except Exception:
        spe = None

    t = DB.Transaction(doc, u"Vincular parâmetro compartilhado Descrição IN")
    t.Start()
    try:
        if spe is not None:
            idef = spe.GetDefinition()
            binding = doc.ParameterBindings.get_Item(idef)
            if binding is None:
                ok = _insert_binding(idef, app.Create.NewTypeBinding(catset), False)
            elif isinstance(binding, DB.InstanceBinding):
                t.RollBack()
                return False, (u"'{}' já está no projeto como parâmetro de INSTÂNCIA. "
                               u"Altere para TIPO em Gerenciar > Parâmetros do projeto "
                               u"e rode de novo.".format(PARAM_IN_NAME))
            else:
                for c in cats:
                    binding.Categories.Insert(c)
                ok = _insert_binding(idef, binding, True)
        else:
            tmp = os.path.join(Path.GetTempPath(), u"OCA_Descricao_IN_shared.txt")
            with codecs.open(tmp, "w", encoding="utf-16") as f:
                f.write(SHARED_PARAM_TXT)
            original = app.SharedParametersFilename
            try:
                app.SharedParametersFilename = tmp
                dfile = app.OpenSharedParameterFile()
                ext = None
                for g in dfile.Groups:
                    ext = g.Definitions.get_Item(PARAM_IN_NAME)
                    if ext is not None:
                        break
                if ext is None:
                    raise Exception(u"definição não encontrada no arquivo temporário")
                ok = _insert_binding(ext, app.Create.NewTypeBinding(catset), False)
            finally:
                try:
                    app.SharedParametersFilename = original or u""
                except Exception:
                    pass
                try:
                    os.remove(tmp)
                except Exception:
                    pass
        if not ok:
            t.RollBack()
            return False, (u"o Revit recusou o vínculo (existe outro parâmetro chamado "
                           u"'{}' no projeto?)".format(PARAM_IN_NAME))
        t.Commit()
        return True, u"vinculado como parâmetro de tipo em: {}".format(
            u", ".join(to_unicode(c.Name) for c in cats))
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        return False, to_unicode(ex)


if "model" in MODE_DIRS[MODE] and MODE != "analyze":
    missing_cats = in_param_missing_categories()
    clash = sorted(set(h["category"] for hs in MODEL.values() for h in hs if h["en_clash"]))
    if missing_cats:
        msg = (u"O parâmetro compartilhado '{}' (GUID {}) não está disponível em:\n\n"
               u"{}\n\nSem ele, a descrição em inglês não é gravada nesses itens.").format(
                   PARAM_IN_NAME, PARAM_IN_GUID_STR,
                   u"\n".join(u"  • " + to_unicode(c.Name) for c in missing_cats))
        if clash:
            msg += (u"\n\nAtenção: em {} já existe um parâmetro '{}' que NÃO é o "
                    u"compartilhado do escritório (GUID diferente). Ele será ignorado."
                    ).format(u", ".join(clash), PARAM_IN_NAME)
        if forms.alert(u"Vincular o parâmetro '{}' ao projeto como parâmetro de TIPO nessas "
                       u"categorias?".format(PARAM_IN_NAME), sub_msg=msg, title=ALERT_TITLE,
                       yes=True, no=True):
            ok, info = bind_in_param(missing_cats)
            output.print_md(u"- Parâmetro **{}**: {}{}".format(
                PARAM_IN_NAME, u"" if ok else u":warning: não vinculado - ", info))
            if ok:
                MODEL = collect_model(use_view, opts["materials"])   # relê os parâmetros

if not MODEL and MODE == "to_model":
    forms.alert(u"Nenhum tipo ou material com Keynote preenchido foi encontrado "
                u"no escopo escolhido.", title=ALERT_TITLE, exitscript=True)


# ------------------------------------------------------------------
# 4. Comparação
# ------------------------------------------------------------------
def weighted_values(holders, field):
    """Valores distintos não vazios, do mais usado para o menos usado."""
    w = defaultdict(int)
    for h in holders:
        v = h.get(field)
        if v:
            w[v] += max(1, h["count"])
    return [v for v, _ in sorted(w.items(), key=lambda kv: (-kv[1], kv[0]))]


def split_keynote(key):
    """'RE01' -> ('RE', '01'); 'PCF-01/02' -> ('PCF', '-01/02'). A fórmula
    CONCAT(PREFIXO:SEQUENCIA) do template devolve o keynote original inteiro."""
    i = 0
    while i < len(key) and key[i].isalpha():
        i += 1
    return key[:i], key[i:]


changes = []     # {"target","action","key","field","old","new","where","id"}
skipped = []     # {"key","where","id","eid","reason"}
xl_written = {"backup": u"", "appended": 0, "filled": 0}


# ------------------------------------------------------------------
# 5a. Modelo -> Excel
# ------------------------------------------------------------------
def plan_to_excel_rows():
    """Sentido MODELO -> EXCEL: uma linha por keynote do modelo que falta no
    Excel ou cuja descrição difere. Valor vazio no modelo nunca apaga o Excel.
    Linhas que SOBRESCREVEM célula preenchida começam desmarcadas."""
    rows = []
    for key, holders in MODEL.items():
        pts, ens = weighted_values(holders, "pt"), weighted_values(holders, "en")
        pt = pts[0] if pts else u""
        en = ens[0] if ens else u""
        # tipo/material "representante": o mais usado entre os que têm o valor proposto
        cands = [h for h in holders if (pt and h["pt"] == pt)] or holders
        rep = sorted(cands, key=lambda h: -h["count"])[0]
        multi = len(pts) > 1 or len(ens) > 1
        rec = XL["by_key"].get(key)
        if rec is None:
            prefix, seq = split_keynote(key)
            grp, cat = XL["prefix_map"].get(prefix, (u"", u""))
            note = u"Incluído via Revit em {}".format(datetime.now().strftime("%d/%m/%Y"))
            if prefix not in XL["prefix_map"]:
                note += u" - prefixo não cadastrado em DADOS"
            if multi:
                note += u" - modelo com descrições diferentes, revisar"
            rows.append({"dir": "excel", "key": key, "rep": rep, "nh": len(holders),
                         "multi": multi, "rec": None,
                         "append": {"key": key, "prefix": prefix, "seq": seq, "grp": grp,
                                    "cat": cat, "pt": pt, "en": en, "obs": note},
                         "pt_old": None, "en_old": None, "pt_new": pt, "en_new": en,
                         "over": set(), "blocked": {}, "enabled": True})
            continue
        pt_new = pt if (pt and pt != rec["pt"]) else None
        en_new = en if (en and en != rec["en"]) else None
        if pt_new is None and en_new is None:
            continue
        over = set(f for f, v in (("pt", pt_new), ("en", en_new)) if v is not None and rec[f])
        rows.append({"dir": "excel", "key": key, "rep": rep, "nh": len(holders),
                     "multi": multi, "rec": rec, "append": None,
                     "pt_old": rec["pt"], "en_old": rec["en"],
                     "pt_new": pt_new, "en_new": en_new,
                     "over": over, "blocked": {}, "enabled": True})
    return rows


def backup_path(path):
    """<arquivo>_backup_AAAAMMDD_HHMMSS.<ext>, sem colidir com um backup já existente."""
    base, ext = os.path.splitext(path)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cand, n = u"{}_backup_{}{}".format(base, stamp, ext), 1
    while os.path.exists(cand):
        n += 1
        cand = u"{}_backup_{}_{}{}".format(base, stamp, n, ext)
    return cand


def ensure_file_closed(path):
    try:
        fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite,
                       getattr(FileShare, "None"))
        fs.Dispose()
        return True
    except Exception:
        return False


def write_excel(fills, appends):
    """Grava direto no XML da aba: preenche F/G, acrescenta linhas no fim da
    tabela (PREFIXO/SEQUENCIA + fórmulas do template com valor em cache) e
    ajusta AutoFiltro/dimensão. O resto do arquivo não é tocado."""
    backup = backup_path(XLSX_PATH)
    File.Copy(XLSX_PATH, backup, False)

    cols = XL["cols"]
    archive = ZipFile.Open(XLSX_PATH, ZipArchiveMode.Update)
    try:
        sd = load_xml(zip_read(archive, XL["part"]))
        nsm = ns_manager(sd)
        sheet_data = sd.SelectSingleNode("//a:sheetData", nsm)
        row_nodes = {}
        for r in sheet_data.SelectNodes("a:row", nsm):
            try:
                row_nodes[int(r.GetAttribute("r"))] = r
            except Exception:
                continue

        def get_row(rn):
            node = row_nodes.get(rn)
            if node is not None:
                return node
            node = sd.CreateElement("row", NS_MAIN)
            node.SetAttribute("r", str(rn))
            after = [k for k in row_nodes if k > rn]
            if after:
                sheet_data.InsertBefore(node, row_nodes[min(after)])
            else:
                sheet_data.AppendChild(node)
            row_nodes[rn] = node
            return node

        def set_cell(rn, ci, text=None, formula=None):
            row = get_row(rn)
            ref = col_letters(ci) + str(rn)
            target, before = None, None
            for c in row.SelectNodes("a:c", nsm):
                cref = to_unicode(c.GetAttribute("r"))
                if cref == ref:
                    target = c
                    break
                if cref and split_ref(cref)[0] > ci:
                    before = c
                    break
            if target is None:
                target = sd.CreateElement("c", NS_MAIN)
                target.SetAttribute("r", ref)
                if before is not None:
                    row.InsertBefore(target, before)
                else:
                    row.AppendChild(target)
            while target.HasChildNodes:              # mantém só o estilo (s)
                target.RemoveChild(target.FirstChild)
            if target.HasAttribute("t"):
                target.RemoveAttribute("t")
            text = to_unicode(text or u"")
            if formula:
                target.SetAttribute("t", "str")
                f = sd.CreateElement("f", NS_MAIN)
                f.InnerText = formula
                target.AppendChild(f)
                v = sd.CreateElement("v", NS_MAIN)
                v.InnerText = text
                target.AppendChild(v)
            elif text:
                target.SetAttribute("t", "inlineStr")
                is_el = sd.CreateElement("is", NS_MAIN)
                t_el = sd.CreateElement("t", NS_MAIN)
                t_el.InnerText = text
                is_el.AppendChild(t_el)
                target.AppendChild(is_el)

        # preencher descrições vazias
        for rec, field, value in fills:
            set_cell(rec["row"], cols[field], value)
            rec[field] = value

        # acrescentar no fim da tabela: próxima linha sem conteúdo e sem fórmula
        last = XL["last_row"]
        rn = last
        for item in appends:
            rn += 1
            while rn in XL["grid"] or rn in XL["formula_rows"]:
                rn += 1
            if "prefix" in cols and "seq" in cols:
                c_pf, c_sq = col_letters(cols["prefix"]), col_letters(cols["seq"])
                set_cell(rn, cols["prefix"], item["prefix"])
                set_cell(rn, cols["seq"], item["seq"])
                set_cell(rn, cols["key"], item["key"],
                         formula=u"_xlfn.CONCAT({0}{2}:{1}{2})".format(c_pf, c_sq, rn))
                if "grp" in cols:
                    set_cell(rn, cols["grp"], item["grp"], formula=(
                        u'IFERROR(VLOOKUP({0}{1},{2}!A:B,2,FALSE),"")'.format(
                            c_pf, rn, DATA_SHEET_NAME)))
                    if "cat" in cols:
                        set_cell(rn, cols["cat"], item["cat"], formula=(
                            u'IFERROR(VLOOKUP({0}{1},{2}!B:C,2,FALSE),"")'.format(
                                col_letters(cols["grp"]), rn, DATA_SHEET_NAME)))
            else:
                set_cell(rn, cols["key"], item["key"])
            set_cell(rn, cols["pt"], item["pt"])
            set_cell(rn, cols["en"], item["en"])
            if "obs" in cols:
                set_cell(rn, cols["obs"], item["obs"])
            last = rn
            rec = {"row": rn, "key": item["key"], "pt": item["pt"], "en": item["en"],
                   "cat": item["cat"], "grp": item["grp"], "prefix": item["prefix"],
                   "seq": item["seq"], "obs": item["obs"]}
            XL["rows"].append(rec)
            XL["by_key"][item["key"]] = rec

        # AutoFiltro e dimensão acompanham a nova última linha
        def grow_ref(ref):
            m = re.match(r"^(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)$", ref)
            if m and int(m.group(2)) < last:
                return m.group(1) + str(last)
            return ref

        af = sd.SelectSingleNode("//a:autoFilter", nsm)
        if af is not None and af.GetAttribute("ref"):
            af.SetAttribute("ref", grow_ref(to_unicode(af.GetAttribute("ref"))))
        dim = sd.SelectSingleNode("//a:dimension", nsm)
        if dim is not None and dim.GetAttribute("ref"):
            dim.SetAttribute("ref", grow_ref(to_unicode(dim.GetAttribute("ref"))))

        # workbook: intervalo do filtro + recalcular fórmulas ao abrir
        wb = load_xml(zip_read(archive, "xl/workbook.xml"))
        wnsm = ns_manager(wb)
        for dn in wb.SelectNodes("//a:definedNames/a:definedName", wnsm):
            if to_unicode(dn.GetAttribute("name")) == u"_xlnm._FilterDatabase" and \
                    to_unicode(dn.GetAttribute("localSheetId")) == str(XL["pos"]):
                txt = to_unicode(dn.InnerText)
                m = re.match(r"^(.*!\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)$", txt)
                if m and int(m.group(2)) < last:
                    dn.InnerText = m.group(1) + str(last)
        calc = wb.SelectSingleNode("//a:calcPr", wnsm)
        if calc is not None:
            calc.SetAttribute("fullCalcOnLoad", "1")

        for part, xdoc in ((XL["part"], sd), ("xl/workbook.xml", wb)):
            old = archive.GetEntry(part)
            if old is not None:
                old.Delete()
            entry = archive.CreateEntry(part)
            writer = StreamWriter(entry.Open(), UTF8Encoding(False))
            try:
                writer.Write(xdoc.OuterXml)
            finally:
                writer.Dispose()
    finally:
        archive.Dispose()
    return backup


# ------------------------------------------------------------------
# 5b. Excel -> Modelo
# ------------------------------------------------------------------
def owned_by_other(elem):
    if not doc.IsWorkshared:
        return False
    try:
        return (DB.WorksharingUtils.GetCheckoutStatus(doc, elem.Id)
                == DB.CheckoutStatus.OwnedByOtherUser)
    except Exception:
        return False


def where_label(h):
    if h["kind"] == "material":
        return u"Material: {}".format(h["name"])
    fam = h["family"]
    return u"{}: {}{}".format(h["category"], (fam + u" - ") if fam else u"", h["name"])


def plan_to_model():
    """-> (plan, bloqueados). plan: alterações possíveis; bloqueados: diferenças
    que não podem ser gravadas (parâmetro ausente, somente leitura, workshared)."""
    plan, blocked = [], []
    for key, holders in MODEL.items():
        rec = XL["by_key"].get(key)
        if rec is None:
            continue
        for h in holders:
            for field, label, pkey in (("pt", u"Description", "pt_p"),
                                       ("en", PARAM_IN_NAME, "en_p")):
                new = rec[field]
                if not new:
                    continue                      # Excel vazio nunca apaga o modelo
                p = h[pkey]
                if p is not None:
                    label = to_unicode(p.Definition.Name)
                old = h[field] or u""
                if old == new:
                    continue
                reason = None
                if p is None and field == "en" and h["en_clash"]:
                    reason = (u"'{}' existente não é o parâmetro compartilhado "
                              u"(GUID diferente)".format(label))
                elif p is None:
                    reason = (u"parâmetro '{}' não existe neste {}".format(
                        label, u"material" if h["kind"] == "material" else u"tipo"))
                elif p.IsReadOnly:
                    reason = u"parâmetro '{}' somente leitura".format(label)
                elif owned_by_other(h["elem"]):
                    reason = u"elemento emprestado por outro usuário (workshared)"
                item = {"h": h, "field": field, "label": label, "p": p,
                        "old": old, "new": new, "reason": reason}
                (blocked if reason else plan).append(item)
    return plan, blocked


# ------------------------------------------------------------------
# 5c. Agrupamento de Keynotes (estrutura do TXT: CATEGORIA > GRUPO/prefixo)
# ------------------------------------------------------------------
def agrup_of(key):
    """-> (categoria, rótulo do grupo). Ordem de busca: aba DADOS do Excel
    (prefixo), colunas CATEGORIA/GRUPO da linha do keynote, hierarquia do
    arquivo de keynotes carregado no Revit."""
    prefix = split_keynote(key)[0]
    if prefix in XL["prefix_map"]:
        grp, cat = XL["prefix_map"][prefix]
        return (cat or NO_GROUP), (u"{} · {}".format(prefix, grp) if grp else prefix)
    rec = XL["by_key"].get(key)
    if rec is not None and (rec.get("cat") or rec.get("grp")):
        return (rec.get("cat") or NO_GROUP), (
            u"{} · {}".format(prefix, rec["grp"]) if rec.get("grp") else prefix)
    chain, k, seen = [], key, set()
    while KN_PARENT.get(k) and k not in seen:
        seen.add(k)
        k = KN_PARENT[k]
        chain.append(k)
    if chain:
        grp = chain[0]
        label = u"{} · {}".format(grp, FILE_TEXT.get(grp, u"")) if len(chain) > 1 else grp
        return chain[-1], label
    return NO_GROUP, (prefix or u"-")


def cat_rank(cat):
    order = XL["cat_order"]
    if cat in order:
        return order.index(cat)
    return len(order) + (1 if cat == NO_GROUP else 0)


# ------------------------------------------------------------------
# 5d. Destaque temporário no modelo (vista ativa)
#     Caso 1  keynote + description ......... nada
#     Caso 2  keynote + description vazia ... VERMELHO
#     Caso 3  keynote vazio + description ... VERMELHO
#     Caso 4  keynote vazio + description vazia LARANJA
#   Avaliado nos TIPOS (onde ficam Keynote e Description). As cores são
#   sobreposições de vista aplicadas numa transação que é DESFEITA ao fechar a
#   tabela de verificação - nada fica gravado no modelo.
# ------------------------------------------------------------------
HL_CASES = {
    2: (u"red", u"Keynote preenchido, Description vazia"),
    3: (u"red", u"Keynote vazio, Description preenchida"),
    4: (u"orange", u"Keynote e Description vazios"),
}
HL_RGB = {u"red": (232, 74, 74), u"orange": (255, 140, 0)}
hl_note = u""


def view_allows_overrides(view):
    if view is None or getattr(view, "IsTemplate", False):
        return False
    try:
        return bool(view.AreGraphicsOverridesAllowed())
    except Exception:
        pass
    bad = set()
    for _n in ("Schedule", "DrawingSheet", "Report", "ColumnSchedule", "PanelSchedule",
               "Internal", "Undefined", "ProjectBrowser", "SystemBrowser"):
        _vt = getattr(DB.ViewType, _n, None)
        if _vt is not None:
            bad.add(_vt)
    return view.ViewType not in bad


def collect_highlight():
    """Tipos (com parâmetro Keynote) dos elementos VISÍVEIS na vista ativa
    que caem nos casos 2, 3 ou 4."""
    global hl_note
    if not view_allows_overrides(active_view):
        hl_note = u"vista ativa não aceita sobreposição de cores (tabela, folha...)"
        return []
    by_type = {}
    try:
        elems = (DB.FilteredElementCollector(doc, active_view.Id)
                 .WhereElementIsNotElementType().ToElements())
    except Exception:
        elems = []
    for e in elems:
        cat = e.Category
        if cat is None or not is_model_category(cat):
            continue
        try:
            tid = e.GetTypeId()
        except Exception:
            continue
        if tid is None or tid == INVALID_ID:
            continue
        info = by_type.setdefault(eid_int(tid), [tid, [], to_unicode(cat.Name)])
        info[1].append(e.Id)
    out = []
    for tid, ids, cname in by_type.values():
        t = doc.GetElement(tid)
        if t is None:
            continue
        try:
            kp = t.get_Parameter(BIP.KEYNOTE_PARAM)
        except Exception:
            kp = None
        if kp is None:
            continue                       # categoria sem parâmetro Keynote
        key, desc = param_text(kp), param_text(desc_param(t))
        if key and desc:
            continue                       # caso 1
        case = 2 if key else (3 if desc else 4)
        color, label = HL_CASES[case]
        out.append({"case": case, "color": color, "label": label, "category": cname,
                    "family": family_name(t), "type": elem_name(t), "typeId": eid_int(tid),
                    "key": key, "desc": desc, "count": len(ids), "ids": ids})
    out.sort(key=lambda x: (x["case"], x["category"], x["family"], x["type"]))
    return out


def solid_fill_id():
    try:
        for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
            try:
                if fp.GetFillPattern().IsSolidFill:
                    return fp.Id
            except Exception:
                continue
    except Exception:
        pass
    return INVALID_ID


def apply_highlight(hl):
    """Aplica as cores numa transação ABERTA e redesenha a vista. Devolve a
    transação para ser desfeita (RollBack) quando a tabela fechar."""
    if not hl:
        return None
    fill = solid_fill_id()
    ogs = {}
    for color, (r_, g_, b_) in HL_RGB.items():
        c = DB.Color(r_, g_, b_)
        o = DB.OverrideGraphicSettings()
        o.SetProjectionLineColor(c)
        o.SetCutLineColor(c)
        if fill != INVALID_ID:
            o.SetSurfaceForegroundPatternId(fill)
            o.SetSurfaceForegroundPatternColor(c)
            o.SetCutForegroundPatternId(fill)
            o.SetCutForegroundPatternColor(c)
        ogs[color] = o
    t = DB.Transaction(doc, u"Destaque temporário - Keynotes")
    t.Start()
    for item in hl:
        for eid in item["ids"]:
            try:
                active_view.SetElementOverrides(eid, ogs[item["color"]])
            except Exception:
                pass
    try:
        doc.Regenerate()
        revit.uidoc.RefreshActiveView()
    except Exception:
        pass
    return t


def remove_highlight(t):
    if t is None:
        return
    try:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        revit.uidoc.RefreshActiveView()
    except Exception:
        pass


# ------------------------------------------------------------------
# 5e. Linhas de verificação nos dois sentidos
#     dir "model": EXCEL -> MODELO (uma linha por tipo/material)
#     dir "excel": MODELO -> EXCEL (uma linha por keynote)
# ------------------------------------------------------------------
def model_rows():
    plan, blocked = plan_to_model()
    by_h = OrderedDict()
    for it in plan + blocked:
        by_h.setdefault(it["h"]["id"], []).append(it)
    rows = []
    for its in by_h.values():
        h = its[0]["h"]
        r = {"dir": "model", "key": h["key"], "h": h, "items": its,
             "pt_old": h["pt"], "en_old": h["en"], "pt_new": None, "en_new": None,
             "over": set(), "blocked": {}}
        for it in its:
            if it["reason"]:
                r["blocked"][it["field"]] = it["reason"]
            else:
                r[it["field"] + "_new"] = it["new"]
                if it["old"]:
                    r["over"].add(it["field"])
        r["enabled"] = any(not it["reason"] for it in its)
        rows.append(r)
    return rows, plan, blocked


def fam_type(h):
    if h["kind"] == "material":
        return u"Material", h["name"]
    return (h["family"] or h["category"]), h["name"]


def row_view(r):
    """Texto das colunas da tabela (janela e HTML)."""
    cat, grp = agrup_of(r["key"])
    if r["dir"] == "model":
        fam, typ = fam_type(r["h"])
    else:
        fam, typ = fam_type(r["rep"])
        if r["nh"] > 1:
            typ = u"{}  (+{} tipo/material)".format(typ, r["nh"] - 1)

    def cur(v, no_param=False):
        if no_param:
            return u"(sem parâmetro)"
        if v is None:
            return u"(não está no Excel)"
        return v or u"(vazio)"

    def new(f):
        v = r[f + "_new"]
        if v is not None:
            return v or u"(vazio)"
        if f in r["blocked"]:
            return u"(bloqueado)"
        return u"(mantém)"

    if r["dir"] == "model":
        pt_cur, en_cur = cur(r["pt_old"]), cur(r["en_old"], r["h"]["en"] is None)
        dest = u"→ Modelo"
    else:
        pt_cur, en_cur = cur(r["pt_old"]), cur(r["en_old"])
        dest = u"→ Excel · nova linha" if r["append"] else u"→ Excel"
    notes = []
    if r["over"]:
        notes.append(u"sobrescreve")
    if r["dir"] == "excel" and r["multi"]:
        notes.append(u"modelo com valores diferentes")
    for f, why in sorted(r["blocked"].items()):
        notes.append(u"{} bloqueado: {}".format(u"PT" if f == "pt" else u"IN", why))
    if notes:
        dest += u" · " + u" · ".join(notes)
    if not r["enabled"]:
        kind = u"blocked"
    elif r["over"]:
        kind = u"over"
    else:
        kind = u"fill"
    return {"cat": cat, "grp": grp, "family": fam, "type": typ, "key": r["key"],
            "ptCur": pt_cur, "ptNew": new("pt"), "enCur": en_cur, "enNew": new("en"),
            "dest": dest, "dir": r["dir"], "kind": kind}


def writes_of(r):
    return set(f for f in ("pt", "en") if r[f + "_new"] is not None)


def conflicts(a, b):
    """Mesmo keynote, sentidos opostos e o mesmo campo sendo gravado."""
    return (a["dir"] != b["dir"] and a["key"] == b["key"]
            and bool(writes_of(a) & writes_of(b)))


# ------------------------------------------------------------------
# 5f. Janela de verificação: tabela única agrupada por Agrupamento
# ------------------------------------------------------------------
for _asm in ("PresentationFramework", "PresentationCore", "WindowsBase", "System.Data"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
import System
from System.Data import DataTable

VERIFY_BODY = u"""
  <Grid>
    <Grid.Resources>
      <Style x:Key="wrap" TargetType="TextBlock">
        <Setter Property="TextWrapping" Value="Wrap"/>
        <Setter Property="Padding" Value="4,2,4,2"/>
      </Style>
      <Style x:Key="mono" TargetType="TextBlock">
        <Setter Property="FontFamily" Value="Consolas"/>
        <Setter Property="Padding" Value="4,2,4,2"/>
      </Style>
    </Grid.Resources>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row="0">
      <Border Style="{StaticResource oca.Msg.Info}" Margin="0"
              Visibility="{Binding Visibility, ElementName=info}">
        <TextBlock x:Name="info" Style="{StaticResource oca.MsgText}"/>
      </Border>
      <Border Style="{StaticResource oca.Msg.Err}" Margin="0,6,0,0"
              Visibility="{Binding Visibility, ElementName=hlinfo}">
        <TextBlock x:Name="hlinfo" Style="{StaticResource oca.MsgText}"/>
      </Border>
      <TextBlock Style="{StaticResource oca.Hint}" Margin="0,8,0,0"
                 Text="→ Modelo = valor do Excel gravado no tipo/material  ·  → Excel = valor do modelo gravado na planilha  ·  Verde = preenche vazio  ·  Âmbar = sobrescreve valor existente  ·  Cinza = não pode ser gravado. Marcar a linha do agrupamento marca o grupo inteiro (linhas visíveis). Um mesmo campo não pode ir para os dois lados ao mesmo tempo: marcar um lado desmarca o outro."/>
    </StackPanel>

    <WrapPanel Grid.Row="1" Margin="0,12,0,8">
      <TextBlock Text="Sentido" Style="{StaticResource oca.Label}"/>
      <ComboBox x:Name="f_dir" Width="170" Margin="0,0,16,0"/>
      <TextBlock Text="Agrupamento" Style="{StaticResource oca.Label}"/>
      <ComboBox x:Name="f_cat" Width="210" Margin="0,0,16,0"/>
      <TextBlock Text="Buscar" Style="{StaticResource oca.Label}"/>
      <TextBox x:Name="f_text" Width="200" Margin="0,0,16,0"/>
      <Button x:Name="b_all" Content="Marcar visíveis" Margin="0,0,8,0"/>
      <Button x:Name="b_none" Content="Desmarcar visíveis" Margin="0,0,8,0"/>
      <Button x:Name="b_fill" Content="Somente preencher vazios"/>
    </WrapPanel>

    <DataGrid x:Name="grid" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False"
              CanUserDeleteRows="False" CanUserSortColumns="False" CanUserReorderColumns="False"
              CanUserResizeRows="False" SelectionMode="Single" SelectionUnit="FullRow"
              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
              EnableRowVirtualization="True" FrozenColumnCount="1">
      <DataGrid.RowStyle>
        <Style TargetType="DataGridRow" BasedOn="{StaticResource {x:Type DataGridRow}}">
          <Style.Triggers>
            <DataTrigger Binding="{Binding Kind}" Value="group">
              <Setter Property="Background" Value="{StaticResource oca.AccentSoft}"/>
              <Setter Property="Foreground" Value="{StaticResource oca.Accent}"/>
              <Setter Property="FontWeight" Value="Bold"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Kind}" Value="fill">
              <Setter Property="Foreground" Value="{StaticResource oca.Ok}"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Kind}" Value="over">
              <Setter Property="Foreground" Value="{StaticResource oca.Warn}"/>
            </DataTrigger>
            <DataTrigger Binding="{Binding Kind}" Value="blocked">
              <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
            </DataTrigger>
          </Style.Triggers>
        </Style>
      </DataGrid.RowStyle>
      <DataGrid.Columns>
        <DataGridTemplateColumn Header="✓" Width="36">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0"
                        IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}"
                        IsEnabled="{Binding Enabled}"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
        <DataGridTextColumn Header="Agrupamento de Keynotes" Binding="{Binding Grupo}" Width="190" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
        <DataGridTextColumn Header="Família" Binding="{Binding Familia}" Width="150" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
        <DataGridTextColumn Header="Tipo" Binding="{Binding Tipo}" Width="160" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
        <DataGridTextColumn Header="Keynote" Binding="{Binding Keynote}" Width="80" IsReadOnly="True" ElementStyle="{StaticResource mono}"/>
        <DataGridTextColumn Header="Descrição atual" Binding="{Binding PtAtual}" Width="200" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
        <DataGridTextColumn Header="Descrição nova" Binding="{Binding PtNova}" Width="200" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
        <DataGridTextColumn Header="Descrição em IN atual" Binding="{Binding EnAtual}" Width="180" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
        <DataGridTextColumn Header="Descrição em IN nova" Binding="{Binding EnNova}" Width="180" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
        <DataGridTextColumn Header="Destino" Binding="{Binding Destino}" Width="200" IsReadOnly="True" ElementStyle="{StaticResource wrap}"/>
      </DataGrid.Columns>
    </DataGrid>
  </Grid>"""

VERIFY_FOOTER_LEFT = u"""
  <TextBlock x:Name="counter" VerticalAlignment="Center" Foreground="{StaticResource oca.Ink}"/>"""

VERIFY_FOOTER = u"""
  <Button x:Name="b_apply" Content="Gravar marcados" Style="{StaticResource oca.Primary}"/>
  <Button x:Name="b_cancel" Content="Cancelar" Margin="8,0,0,0"/>"""

VERIFY_XAML = build_xaml(title=u"Verificação · Modelo × Excel",
                         subtitle=u"Atualização Descrição por Keynotes · marque o que muda no modelo "
                                  u"e o que muda na planilha.",
                         body=VERIFY_BODY, footer_right=VERIFY_FOOTER,
                         footer_left=VERIFY_FOOTER_LEFT, size="L", height=780, width=1400)


DIR_LABEL = {"all": u"Todos", "model": u"Excel → Modelo", "excel": u"Modelo → Excel"}
COLS = ("Grupo", "Familia", "Tipo", "Keynote", "PtAtual", "PtNova", "EnAtual",
        "EnNova", "Destino")


def _flt(text):
    """Texto seguro para DataView.RowFilter."""
    t = to_unicode(text).replace(u"'", u"''")
    for ch in u"[]*%":
        t = t.replace(ch, u"")
    return t


class VerifyWindow(forms.WPFWindow):
    def __init__(self, rows, readonly, info_text, hl_text):
        forms.WPFWindow.__init__(self, VERIFY_XAML, literal_string=True)
        self.confirmed, self.readonly, self._busy = False, readonly, False
        self.rows = rows
        self.by_rid = dict((r["rid"], r) for r in rows)
        self.info.Text = info_text
        self.hlinfo.Text = hl_text

        tb = DataTable("verificacao")
        for name, typ in (("RowId", System.Int32), ("Sel", System.Boolean),
                          ("Enabled", System.Boolean), ("Kind", System.String),
                          ("Dir", System.String), ("Cat", System.String)) + \
                tuple((c, System.String) for c in COLS):
            tb.Columns.Add(name, clr.GetClrType(typ))
        self.table, self.drow, self.cats = tb, {}, []

        gid, last_cat = 0, None
        for r in rows:
            v = r["view"]
            if v["cat"] != last_cat:
                gid -= 1
                last_cat = v["cat"]
                self.cats.append(v["cat"])
                g = tb.NewRow()
                g["RowId"], g["Kind"], g["Dir"], g["Cat"] = gid, u"group", u"", v["cat"]
                g["Sel"], g["Enabled"] = False, True
                g["Grupo"] = v["cat"]
                for c in COLS[1:]:
                    g[c] = u""
                tb.Rows.Add(g)
                self.drow[gid] = g
            d = tb.NewRow()
            d["RowId"], d["Kind"], d["Dir"], d["Cat"] = r["rid"], v["kind"], r["dir"], v["cat"]
            d["Sel"], d["Enabled"] = bool(r["sel"]), bool(r["enabled"]) and not readonly
            for c, k in zip(COLS, ("grp", "family", "type", "key", "ptCur", "ptNew",
                                   "enCur", "enNew", "dest")):
                d[c] = v[k]
            tb.Rows.Add(d)
            self.drow[r["rid"]] = d

        self.grid.ItemsSource = tb.DefaultView
        dirs = sorted(set(r["dir"] for r in rows))
        self.f_dir.ItemsSource = [DIR_LABEL["all"]] + [DIR_LABEL[d] for d in dirs]
        self.f_dir.SelectedIndex = 0
        self.f_dir.IsEnabled = len(dirs) > 1
        self.f_cat.ItemsSource = [u"Todos"] + self.cats
        self.f_cat.SelectedIndex = 0
        self.f_dir.SelectionChanged += self._filter
        self.f_cat.SelectionChanged += self._filter
        self.f_text.TextChanged += self._filter
        tb.ColumnChanged += self._on_changed
        self.b_all.Click += lambda s, a: self._set_visible(lambda r: True)
        self.b_none.Click += lambda s, a: self._set_visible(lambda r: False)
        self.b_fill.Click += lambda s, a: self._set_visible(lambda r: not r["over"])
        self.b_apply.Click += self._apply
        self.b_cancel.Click += self._cancel
        if readonly:
            from System.Windows import Visibility
            self.grid.Columns[0].Visibility = Visibility.Collapsed
            for b in (self.b_all, self.b_none, self.b_fill, self.b_apply):
                b.Visibility = Visibility.Collapsed
            self.b_cancel.Content = u"Fechar"
        self._sync_groups()
        self._update_counter()

    # ---- filtros ----
    def _filter(self, sender, args):
        parts = []
        d = self.f_dir.SelectedItem
        for k, lbl in DIR_LABEL.items():
            if k != "all" and d == lbl:
                parts.append(u"(Kind = 'group' OR Dir = '{}')".format(k))
        c = self.f_cat.SelectedItem
        if c and c != u"Todos":
            parts.append(u"Cat = '{}'".format(_flt(c)))
        t = _flt(self.f_text.Text).strip()
        if t:
            like = u" OR ".join(u"{} LIKE '%{}%'".format(col, t) for col in COLS[1:])
            parts.append(u"(Kind = 'group' OR {})".format(like))
        try:
            self.table.DefaultView.RowFilter = u" AND ".join(parts)
        except Exception:
            self.table.DefaultView.RowFilter = u""
        self._sync_groups()
        self._update_counter()

    def _visible_rids(self):
        out = []
        for drv in self.table.DefaultView:
            rid = int(drv.Row["RowId"])
            if rid >= 0:
                out.append(rid)
        return out

    # ---- seleção ----
    def _check(self, rid, value):
        """Marca/desmarca respeitando bloqueio e conflito de sentido."""
        r = self.by_rid[rid]
        if not r["enabled"] or self.readonly:
            value = False
        if value:
            for o in self.rows:
                if o is not r and o["sel"] and conflicts(o, r):
                    o["sel"] = False
                    self.drow[o["rid"]]["Sel"] = False
        r["sel"] = bool(value)
        self.drow[rid]["Sel"] = bool(value)

    def _on_changed(self, sender, e):
        if self._busy or e.Column.ColumnName != "Sel":
            return
        self._busy = True
        try:
            rid = int(e.Row["RowId"])
            val = e.Row["Sel"] == True
            if rid < 0:
                cat = e.Row["Cat"]
                for vr in self._visible_rids():
                    if self.by_rid[vr]["view"]["cat"] == cat and self.by_rid[vr]["enabled"]:
                        self._check(vr, val)
            else:
                self._check(rid, val)
            self._sync_groups()
            self._update_counter()
        finally:
            self._busy = False

    def _set_visible(self, rule):
        self._busy = True
        try:
            vis = self._visible_rids()
            # sentido Excel -> Modelo primeiro: em conflito, o Excel (base revisada) vence
            vis.sort(key=lambda rid: 0 if self.by_rid[rid]["dir"] == "model" else 1)
            for rid in vis:
                self._check(rid, False)
            for rid in vis:
                r = self.by_rid[rid]
                if r["enabled"] and rule(r) and not any(
                        o["sel"] and conflicts(o, r) for o in self.rows):
                    self._check(rid, True)
            self._sync_groups()
            self._update_counter()
        finally:
            self._busy = False

    def _sync_groups(self):
        was = self._busy
        self._busy = True
        try:
            vis = set(self._visible_rids())
            for gid, g in self.drow.items():
                if gid >= 0:
                    continue
                kids = [r for r in self.rows if r["view"]["cat"] == g["Cat"]
                        and r["rid"] in vis and r["enabled"]]
                g["Enabled"] = bool(kids) and not self.readonly
                g["Sel"] = bool(kids) and all(r["sel"] for r in kids)
                n = len([r for r in self.rows if r["view"]["cat"] == g["Cat"] and r["rid"] in vis])
                g["Destino"] = u"{} linha(s)".format(n)
        finally:
            self._busy = was

    def _update_counter(self):
        sel = self.selected()
        tot = len([r for r in self.rows if r["enabled"]])
        to_m = len([r for r in sel if r["dir"] == "model"])
        to_x = len([r for r in sel if r["dir"] == "excel"])
        over = len([r for r in sel if r["over"]])
        if self.readonly:
            self.counter.Text = u"{} diferença(s) encontradas · somente análise, nada será gravado".format(
                len(self.rows))
        else:
            self.counter.Text = (u"{} de {} marcadas  ·  {} → Modelo  ·  {} → Excel  ·  "
                                 u"{} sobrescrevem valor existente".format(
                                     len(sel), tot, to_m, to_x, over))

    def selected(self):
        return [r for r in self.rows if r["sel"] and r["enabled"]]

    def _apply(self, sender, args):
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# ------------------------------------------------------------------
# 5g. TXT de Keynotes do Revit a partir do Excel
#     CATEGORIA <tab> lista de prefixos <tab>          (aba KEYNOTE BÁSICO)
#     PREFIXO   <tab> GRUPO <tab> CATEGORIA            (aba DADOS)
#     KEYNOTE   <tab> DESCRIÇÃO | DESCRIÇÃO IN <tab> PREFIXO   (aba MATERIAIS)
# ------------------------------------------------------------------
def _txt_field(text):
    return u" ".join(to_unicode(text).replace(u"\t", u" ").split())


def build_keynote_txt():
    """-> (texto, avisos). Estrutura idêntica ao TXT do escritório."""
    lines, warns = [], []
    for cat, plist in XL["kn_headers"]:
        lines.append(u"{}\t{}\t".format(cat, plist))
    known = set()
    for pf, grp, cat in XL["dados_rows"]:
        lines.append(u"{}\t{}\t{}".format(pf, grp, cat))
        known.add(clean(pf))
    seen = set(known) | set(clean(c) for c, _ in XL["kn_headers"])
    for rec in XL["rows"]:
        key = rec["key"]
        if key in seen:
            if key not in known:
                warns.append(u"{}: repetido no Excel (só a 1ª linha entra no TXT)".format(key))
            continue
        seen.add(key)
        text = u" | ".join(_txt_field(v) for v in (rec["pt"], rec["en"]) if _txt_field(v))
        prefix = split_keynote(key)[0]
        parent = prefix if prefix in known else u""
        if not parent:
            warns.append(u"{}: prefixo '{}' não está em DADOS - entra sem agrupamento".format(
                key, prefix))
        lines.append(u"{}\t{}\t{}".format(_txt_field(key), text, parent))
    return u"\r\n".join(lines) + u"\r\n", warns


def write_keynote_txt(path, text):
    """Grava em ANSI (cp1252, como o arquivo atual do escritório); se houver
    caractere fora do cp1252, grava em Unicode (UTF-16), que o Revit também lê.
    Se o arquivo já existe, guarda um backup com data/hora."""
    backup = u""
    if os.path.isfile(path):
        backup = backup_path(path)
        File.Copy(path, backup, False)
    try:
        data, enc = text.encode("cp1252"), u"ANSI (cp1252)"
    except UnicodeEncodeError:
        data, enc = codecs.BOM_UTF16_LE + text.encode("utf-16-le"), u"Unicode (UTF-16)"
    with open(path, "wb") as f:
        f.write(data)
    return backup, enc


def load_keynote_txt(path):
    """Carrega o TXT como tabela de keynotes do projeto. -> (ok, mensagem)"""
    t = DB.Transaction(doc, u"Carregar TXT de Keynotes")
    t.Start()
    try:
        mp = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        ref = DB.ExternalResourceReference.CreateLocalResource(
            doc, DB.ExternalResourceTypes.BuiltInExternalResourceTypes.KeynoteTable,
            mp, DB.PathType.Absolute)
        res = DB.KeyBasedTreeEntriesLoadResults()
        DB.KeynoteTable.GetKeynoteTable(doc).LoadFrom(ref, res)
        t.Commit()
        errs = []
        try:
            errs += [to_unicode(x) for x in res.GetFileSyntaxErrors()]
        except Exception:
            pass
        try:
            errs += [to_unicode(x) for x in res.GetKeyBasedTreeEntryErrors()]
        except Exception:
            pass
        return True, (u"carregado no projeto" + (
            u" com {} aviso(s) do Revit".format(len(errs)) if errs else u""))
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        return False, to_unicode(ex)


# ------------------------------------------------------------------
# 5h. Executa: tabela de verificação -> grava o que foi marcado
# ------------------------------------------------------------------
# (holder id, campo) -> motivo de NÃO ter sido gravado (usado no resultado)
not_written = {}
DIRS = MODE_DIRS[MODE]
READONLY = MODE == "analyze"

m_rows, m_plan, m_blocked = model_rows() if "model" in DIRS else ([], [], [])
for it in m_blocked:
    h = it["h"]
    not_written[(h["id"], it["field"])] = it["reason"]
    skipped.append({"key": h["key"], "where": where_label(h), "id": h["id"],
                    "eid": h["elem"].Id, "reason": it["reason"]})
x_rows = plan_to_excel_rows() if "excel" in DIRS else []

VROWS = m_rows + x_rows
for r in VROWS:
    r["view"] = row_view(r)
    # padrão: Excel -> Modelo marcado; Modelo -> Excel marcado só se não sobrescreve
    r["sel"] = r["enabled"] and not READONLY and (r["dir"] == "model" or not r["over"])
VROWS.sort(key=lambda r: (cat_rank(r["view"]["cat"]), r["view"]["cat"], r["view"]["grp"],
                          r["key"], 0 if r["dir"] == "model" else 1,
                          r["view"]["family"], r["view"]["type"]))
for i, r in enumerate(VROWS):
    r["rid"] = i
# garante que os padrões não tragam conflito (Excel -> Modelo vence)
for r in VROWS:
    if r["sel"] and r["dir"] == "excel" and any(
            o["sel"] and conflicts(o, r) for o in VROWS if o["dir"] == "model"):
        r["sel"] = False

HL = collect_highlight() if opts["highlight"] else []
hl_red = sum(x["count"] for x in HL if x["color"] == u"red")
hl_orange = sum(x["count"] for x in HL if x["color"] == u"orange")

keys_with_info = [k for k in MODEL if XL["by_key"].get(k) and
                  (XL["by_key"][k]["pt"] or XL["by_key"][k]["en"])]
info_text = (u"{} keynote(s) no modelo · {} com descrição no Excel · {} fora do Excel  ·  "
             u"{} linha(s) Excel → Modelo · {} linha(s) Modelo → Excel  ·  modo: {}".format(
                 len(MODEL), len(keys_with_info),
                 len([k for k in MODEL if not XL["by_key"].get(k)]),
                 len(m_rows), len(x_rows), MODES[MODE]))
if not opts["highlight"]:
    hl_text = u"Destaque no modelo desligado."
elif hl_note:
    hl_text = u"Destaque no modelo indisponível: " + hl_note
else:
    hl_text = (u"Destacados na vista ativa (temporário, desfeito ao fechar): {} elemento(s) em "
               u"VERMELHO (keynote sem description / description sem keynote) · {} em LARANJA "
               u"(sem keynote e sem description).".format(hl_red, hl_orange))

chosen_rows = []
if VROWS or HL:
    hl_tx = apply_highlight(HL)
    try:
        vw = VerifyWindow(VROWS, READONLY, info_text, hl_text)
        vw.ShowDialog()
        if vw.confirmed and not READONLY:
            chosen_rows = vw.selected()
        elif not READONLY:
            output.print_md(u"**Cancelado - nada foi gravado.**")
    finally:
        remove_highlight(hl_tx)
else:
    output.print_md(u"**Verificação:** nenhuma diferença entre o modelo e o Excel.")

chosen_ids = set(r["rid"] for r in chosen_rows)
for r in VROWS:
    r["status"] = u"Somente análise" if READONLY else (
        u"Bloqueado" if not r["enabled"] else
        (u"Marcado" if r["rid"] in chosen_ids else u"Não marcado"))

# ---- Excel -> Modelo: grava os tipos/materiais marcados ----
chosen_items = [it for r in chosen_rows if r["dir"] == "model"
                for it in r["items"] if not it["reason"]]
chosen_item_ids = set(id(it) for it in chosen_items)
for it in m_plan:
    if id(it) not in chosen_item_ids:
        not_written[(it["h"]["id"], it["field"])] = (
            u"não selecionado pelo usuário" +
            (u" (modelo diferente do Excel)" if it["old"] else u""))

failed_rows = set()
if chosen_items:
    t = DB.Transaction(doc, u"Atualizar descrições por Keynote")
    t.Start()
    try:
        for r in chosen_rows:
            if r["dir"] != "model":
                continue
            for it in r["items"]:
                if it["reason"]:
                    continue
                h, field, old, new = it["h"], it["field"], it["old"], it["new"]
                try:
                    it["p"].Set(new)
                    h[field] = new
                    changes.append({"target": "model",
                                    "action": "overwrite" if old else "fill",
                                    "key": h["key"], "field": it["label"],
                                    "fieldKey": field, "category": h["category"],
                                    "old": old, "new": new, "where": where_label(h),
                                    "id": h["id"]})
                except Exception as ex:
                    failed_rows.add(r["rid"])
                    not_written[(h["id"], field)] = to_unicode(ex)
                    skipped.append({"key": h["key"], "where": where_label(h),
                                    "id": h["id"], "eid": h["elem"].Id,
                                    "reason": to_unicode(ex)})
        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert(u"Atualização do modelo desfeita (rollback). Nada foi gravado no modelo.",
                    sub_msg=to_unicode(ex), title=ALERT_TITLE, exitscript=True)

# ---- Modelo -> Excel: grava as linhas marcadas ----
x_chosen = [r for r in chosen_rows if r["dir"] == "excel"]
if x_chosen:
    fills, appends = [], []
    for r in x_chosen:
        if r["append"]:
            appends.append(r["append"])
            continue
        for f in ("pt", "en"):
            if r[f + "_new"] is not None:
                fills.append((r["rec"], f, r[f + "_new"]))
    xl_ok = True
    while not ensure_file_closed(XLSX_PATH):
        if not forms.alert(u"A planilha está aberta (ou bloqueada). Tentar de novo?",
                           sub_msg=u"{}\n\nFeche o arquivo no Excel e clique em Sim para tentar de novo.\n"
                                   u"Não = não gravar no Excel.".format(XLSX_PATH),
                           title=ALERT_TITLE, yes=True, no=True):
            xl_ok = False
            break
    if xl_ok:
        snapshot = [(rec, f, rec[f]) for rec, f, _ in fills]
        try:
            xl_written["backup"] = write_excel(fills, appends)
            xl_written["appended"] = len(appends)
            xl_written["filled"] = len(fills)
            for a in appends:
                changes.append({"target": "excel", "action": "append", "key": a["key"],
                                "field": u"Linha nova", "old": u"",
                                "new": u"{} | {}".format(a["pt"] or u"-", a["en"] or u"-"),
                                "where": XL["sheet"], "id": None})
            for rec, f, old in snapshot:
                changes.append({"target": "excel", "action": "overwrite" if old else "fill",
                                "key": rec["key"], "fieldKey": f,
                                "field": u"DESCRIÇÃO" if f == "pt" else u"DESCRIÇÃO IN",
                                "old": old, "new": rec[f],
                                "where": u"{} · linha {}".format(XL["sheet"], rec["row"]),
                                "id": None})
        except Exception as ex:
            xl_ok = False
            forms.alert(u"Falha ao gravar a planilha.",
                        sub_msg=u"Se o arquivo não abrir, restaure o backup criado ao lado dele "
                                u"(*_backup_AAAAMMDD_HHMMSS.xlsx).\n\n{}".format(to_unicode(ex)),
                        title=ALERT_TITLE)
    if not xl_ok:
        for r in x_chosen:
            failed_rows.add(r["rid"])

for r in VROWS:
    if r["rid"] in failed_rows:
        r["status"] = u"Não gravado"
    elif r["status"] == u"Marcado":
        r["status"] = u"Gravado"

# ---- TXT de Keynotes (opcional, confirma antes) ----
txt_info = {"on": opts["txt_on"], "path": opts["txt_path"], "written": False,
            "loaded": False, "msg": u"", "backup": u"", "encoding": u"", "warnings": [],
            "lines": 0}
if opts["txt_on"]:
    txt_text, txt_warns = build_keynote_txt()
    txt_info["warnings"] = txt_warns
    txt_info["lines"] = len(txt_text.splitlines())
    n_keys = txt_info["lines"] - len(XL["kn_headers"]) - len(XL["dados_rows"])
    if forms.alert(u"Gravar o TXT de Keynotes do Revit a partir do Excel e carregar no projeto?",
                   sub_msg=u"{}\n\n{} agrupamento(s) · {} grupo(s) · {} keynote(s)\n"
                   u"Texto do keynote = DESCRIÇÃO | DESCRIÇÃO IN{}".format(
                       opts["txt_path"], len(XL["kn_headers"]), len(XL["dados_rows"]), n_keys,
                       u"\n\n{} aviso(s) - ver relatório.".format(len(txt_warns))
                       if txt_warns else u""), title=ALERT_TITLE, yes=True, no=True):
        try:
            txt_info["backup"], txt_info["encoding"] = write_keynote_txt(
                opts["txt_path"], txt_text)
            txt_info["written"] = True
            ok, msg = load_keynote_txt(opts["txt_path"])
            txt_info["loaded"], txt_info["msg"] = ok, msg
            if ok:
                FILE_TEXT.clear()
                FILE_TEXT.update(keynote_file_texts())
        except Exception as ex:
            txt_info["msg"] = to_unicode(ex)
    else:
        txt_info["msg"] = u"não gravado (cancelado pelo usuário)"


# ------------------------------------------------------------------
# 6. Inconsistências (estado APÓS a operação)
# ------------------------------------------------------------------
ISSUES = OrderedDict([
    ("model_only", {"label": u"No modelo, fora do Excel", "color": "var(--amber)",
                    "hint": u"Rode Modelo → Excel para incluir."}),
    ("excel_only", {"label": u"No Excel, não usado no modelo", "color": "var(--dim)",
                    "hint": u"Só informativo - nunca é apagado."}),
    ("diff", {"label": u"Descrição diferente Excel × modelo", "color": "var(--mag)",
              "hint": u"Rode Excel → Modelo para alinhar (o Excel vale)."}),
    ("model_conflict", {"label": u"Mesmo keynote, descrições diferentes no modelo",
                        "color": "var(--red)",
                        "hint": u"Tipos/materiais com o mesmo keynote e textos distintos."}),
    ("no_desc", {"label": u"Sem descrição", "color": "var(--red)",
                 "hint": u"Nem PT nem IN, no Excel ou no modelo."}),
    ("pt_no_en", {"label": u"PT sem IN", "color": "var(--violet)",
                  "hint": u"Tem descrição em português e falta a em inglês."}),
    ("en_no_pt", {"label": u"IN sem PT", "color": "var(--violet)",
                  "hint": u"Tem descrição em inglês e falta a em português."}),
    ("excel_dup", {"label": u"Keynote repetido no Excel", "color": "var(--amber)",
                   "hint": u"Só a primeira linha é usada na sincronização."}),
    ("no_in_param", {"label": u"Sem parâmetro compartilhado 'Descrição IN'", "color": "var(--amber)",
                     "hint": u"Vincule o parâmetro compartilhado a essas categorias."}),
])

keynotes_out = []
issue_count = OrderedDict((k, 0) for k in ISSUES)
all_keys = list(MODEL.keys()) + [k for k in XL["by_key"] if k not in MODEL]

for key in all_keys:
    holders = MODEL.get(key, [])
    rec = XL["by_key"].get(key)
    in_model, in_excel = bool(holders), rec is not None
    xl_pt = rec["pt"] if rec else u""
    xl_en = rec["en"] if rec else u""
    m_pt = weighted_values(holders, "pt")
    with_in = [h for h in holders if h["en"] is not None]
    m_en = weighted_values(with_in, "en")

    flags = []
    if in_model and not in_excel:
        flags.append("model_only")
    if in_excel and not in_model:
        flags.append("excel_only")
    if in_model and in_excel:
        if set(h["pt"] for h in holders) != set([xl_pt]) or \
                (with_in and set(h["en"] for h in with_in) != set([xl_en])):
            flags.append("diff")
    if len(m_pt) > 1 or len(m_en) > 1:
        flags.append("model_conflict")
    eff_pt = xl_pt or (m_pt[0] if m_pt else u"")
    eff_en = xl_en or (m_en[0] if m_en else u"")
    if not eff_pt and not eff_en:
        flags.append("no_desc")
    elif eff_pt and not eff_en:
        flags.append("pt_no_en")
    elif eff_en and not eff_pt:
        flags.append("en_no_pt")
    if key in XL["dups"]:
        flags.append("excel_dup")
    if holders and len(with_in) < len(holders):
        flags.append("no_in_param")
    for f in flags:
        issue_count[f] += 1

    if holders:
        cats = sorted(set(h["category"] for h in holders))
    else:
        cats = [u"Excel · {}".format(rec["cat"] or rec["grp"] or u"sem categoria")]

    agcat, aggrp = agrup_of(key)
    keynotes_out.append({
        "key": key, "inModel": in_model, "inExcel": in_excel,
        "agCat": agcat, "agGrp": aggrp, "agRank": cat_rank(agcat),
        "categories": cats, "flags": flags,
        "excelRow": rec["row"] if rec else None,
        "excelCategory": rec["cat"] if rec else u"",
        "excelGroup": rec["grp"] if rec else u"",
        "excelDupRows": XL["dups"].get(key, []),
        "xlPt": xl_pt, "xlEn": xl_en, "modelPt": m_pt, "modelEn": m_en,
        "fileText": FILE_TEXT.get(key, u""),
        "holders": [{
            "kind": h["kind"], "id": h["id"], "category": h["category"],
            "family": h["family"], "name": h["name"], "count": h["count"],
            "usedIn": h["usedIn"], "pt": h["pt"], "en": h["en"],
            "hasIn": h["en"] is not None,
        } for h in holders],
    })

missing_in_by_cat = defaultdict(int)
for holders in MODEL.values():
    for h in holders:
        if h["en"] is None:
            missing_in_by_cat[h["category"]] += 1

cat_summary = defaultdict(lambda: {"keynotes": 0, "issues": 0})
for k in keynotes_out:
    for c in k["categories"]:
        cat_summary[c]["keynotes"] += 1
        if k["inModel"] and any(f != "excel_only" for f in k["flags"]):
            cat_summary[c]["issues"] += 1


# ------------------------------------------------------------------
# 7. Resultado da operação: preenchidos / alterados / faltando,
#    agrupado por categoria do modelo e depois por keynote
# ------------------------------------------------------------------
FIELD_LABEL = {"pt": u"Description", "en": PARAM_IN_NAME}
XL_FIELD_LABEL = {"pt": u"DESCRIÇÃO", "en": u"DESCRIÇÃO IN"}


def _ag(entry):
    entry["agCat"], entry["agGrp"] = agrup_of(entry["key"])
    entry["agRank"] = cat_rank(entry["agCat"])
    return entry


def build_result():
    """Preenchidos / alterados / faltando, nos sentidos incluídos no modo.
    'target' = model | excel | both."""
    res = {"target": None, "filled": [], "changed": [], "missing": []}
    if READONLY:
        return res
    do_model, do_excel = "model" in DIRS, "excel" in DIRS
    res["target"] = "both" if (do_model and do_excel) else ("model" if do_model else "excel")
    if do_model:
        written = set()
        for c in changes:
            if c["target"] != "model":
                continue
            written.add((c["id"], c["fieldKey"]))
            entry = {"dest": u"Modelo", "category": c["category"], "key": c["key"],
                     "where": c["where"], "id": c["id"], "field": c["field"],
                     "old": c["old"], "new": c["new"]}
            res["changed" if c["action"] == "overwrite" else "filled"].append(_ag(entry))
        for key, holders in MODEL.items():
            rec = XL["by_key"].get(key)
            for h in holders:
                for field, pkey in (("pt", "pt_p"), ("en", "en_p")):
                    if (h["id"], field) in written:
                        continue
                    label = to_unicode(h[pkey].Definition.Name) if h[pkey] is not None \
                        else FIELD_LABEL[field]
                    value = h[field] or u""
                    reason = not_written.get((h["id"], field))
                    if not reason and value:
                        continue                   # preenchido e sem pendência
                    if not reason:
                        if h[pkey] is None and h.get("en_clash"):
                            reason = (u"'{}' existente não é o parâmetro "
                                      u"compartilhado (GUID diferente)".format(label))
                        elif h[pkey] is None:
                            reason = u"parâmetro '{}' não existe".format(label)
                        elif rec is None:
                            reason = u"keynote não está no Excel"
                        elif not rec[field]:
                            reason = u"sem descrição no Excel"
                        else:
                            reason = u"vazio"
                    res["missing"].append(_ag({
                        "dest": u"Modelo", "category": h["category"], "key": key,
                        "where": where_label(h), "id": h["id"], "field": label,
                        "old": value, "new": u"", "reason": reason}))
    if do_excel:
        for c in changes:
            if c["target"] != "excel":
                continue
            cats = sorted(set(h["category"] for h in MODEL.get(c["key"], []))) or [u"-"]
            entry = {"dest": u"Excel", "category": u", ".join(cats), "key": c["key"],
                     "where": c["where"], "id": None, "field": c["field"],
                     "old": c["old"], "new": c["new"]}
            res["changed" if c["action"] == "overwrite" else "filled"].append(_ag(entry))
        for key, holders in MODEL.items():
            rec = XL["by_key"].get(key)
            cats = u", ".join(sorted(set(h["category"] for h in holders)))
            if rec is None:
                res["missing"].append(_ag({
                    "dest": u"Excel", "category": cats, "key": key, "where": XL["sheet"],
                    "id": None, "field": u"Linha", "old": u"", "new": u"",
                    "reason": u"keynote não gravado na planilha"}))
                continue
            for field in ("pt", "en"):
                if not rec[field]:
                    res["missing"].append(_ag({
                        "dest": u"Excel", "category": cats, "key": key,
                        "where": u"{} · linha {}".format(XL["sheet"], rec["row"]),
                        "id": None, "field": XL_FIELD_LABEL[field], "old": u"",
                        "new": u"", "reason": u"vazio no Excel e no modelo"}))
    for k in ("filled", "changed", "missing"):
        res[k].sort(key=lambda e: (e["agRank"], e["agCat"], e["key"], e["dest"],
                                   e["where"], e["field"]))
    return res


RESULT = build_result()


def excel_table():
    """Aba MATERIAIS como ficou no fim (estado gravado), com as células
    alteradas nesta execução marcadas (antes -> depois)."""
    old_cells, appended = {}, set()
    for c in changes:
        if c["target"] != "excel":
            continue
        if c["action"] == "append":
            appended.add(c["key"])
        else:
            old_cells[(c["key"], c["fieldKey"])] = (c["action"], c["old"])
    out = []
    for rec in sorted(XL["rows"], key=lambda x: x["row"]):
        cat, grp = agrup_of(rec["key"])
        ch = {}
        for f in ("pt", "en"):
            if (rec["key"], f) in old_cells:
                act, old = old_cells[(rec["key"], f)]
                ch[f] = {"action": act, "old": old}
        out.append({"row": rec["row"], "agCat": cat, "agGrp": grp, "agRank": cat_rank(cat),
                    "categoria": rec.get("cat") or cat, "grupo": rec.get("grp") or u"",
                    "prefixo": rec.get("prefix") or u"", "seq": rec.get("seq") or u"",
                    "key": rec["key"], "pt": rec["pt"], "en": rec["en"],
                    "obs": rec.get("obs") or u"", "isNew": rec["key"] in appended,
                    "changed": ch, "inModel": rec["key"] in MODEL})
    return out


# destaque recalculado depois das gravações (para o relatório)
HL_FINAL = collect_highlight() if opts["highlight"] else []
hl_red = sum(x["count"] for x in HL_FINAL if x["color"] == u"red")
hl_orange = sum(x["count"] for x in HL_FINAL if x["color"] == u"orange")


# ------------------------------------------------------------------
# 8. Janela de saída + relatório HTML (opcional nos modos que gravam)
# ------------------------------------------------------------------
output_header(output, __title__, u"Keynotes × Excel - {}".format(MODES[MODE]))
output.print_md(u"## Keynotes × Excel - {}".format(MODES[MODE]))
output.print_md(u"- **{}** keynote(s) no modelo ({} tipo(s)/material(is)) · **{}** no Excel "
                u"(aba *{}*) · **{}** em ambos".format(
                    len(MODEL), sum(len(v) for v in MODEL.values()), len(XL["by_key"]),
                    XL["sheet"],
                    sum(1 for k in keynotes_out if k["inModel"] and k["inExcel"])))
if RESULT["target"]:
    output.print_md(u"- Resultado no **{}**: **{}** preenchido(s) · **{}** alterado(s) · "
                    u"**{}** faltando / pendente(s)".format(
                        {"model": u"modelo", "excel": u"Excel",
                         "both": u"modelo e no Excel"}[RESULT["target"]],
                        len(RESULT["filled"]), len(RESULT["changed"]),
                        len(RESULT["missing"])))
if opts["highlight"]:
    output.print_md(u"- Destaque temporário (vista ativa): **{}** elemento(s) em vermelho · "
                    u"**{}** em laranja{}".format(hl_red, hl_orange,
                                                 u" - " + hl_note if hl_note else u""))
if txt_info["on"]:
    output.print_md(u"- TXT de Keynotes: {}{}".format(
        (u"gravado em `{}` ({}) e {}".format(txt_info["path"], txt_info["encoding"],
                                             txt_info["msg"]) if txt_info["written"]
         else txt_info["msg"]),
        u" · :warning: {} aviso(s)".format(len(txt_info["warnings"]))
        if txt_info["warnings"] else u""))
for c, v in ISSUES.items():
    if issue_count[c]:
        output.print_md(u"- {}: **{}**".format(v["label"], issue_count[c]))
if xl_written["backup"]:
    output.print_md(u"- Excel gravado: **{}** adicionado(s), **{}** preenchido(s). "
                    u"Backup: `{}`".format(xl_written["appended"], xl_written["filled"],
                                           xl_written["backup"]))
if skipped:
    output.print_md(u"### :warning: {} item(ns) não gravado(s)".format(len(skipped)))
    output.print_table(
        table_data=[[s["key"], output.linkify(s["eid"]),
                     s["where"], s["reason"]] for s in skipped],
        columns=[u"Keynote", u"Elemento", u"Tipo / Material", u"Motivo"])
if missing_in_by_cat:
    output.print_md(u"- Parâmetro **{}** ausente em: {}".format(
        PARAM_IN_NAME, u", ".join(u"{} ({})".format(c, n)
                                      for c, n in sorted(missing_in_by_cat.items()))))

if MODE != "analyze":
    make_html = forms.alert(
        u"Gerar o relatório HTML?",
        sub_msg=u"{} preenchido(s) · {} alterado(s) · {} faltando / pendente(s)\n"
        u"(agrupados por agrupamento de Keynotes)".format(
            len(RESULT["filled"]), len(RESULT["changed"]), len(RESULT["missing"])),
        title=ALERT_TITLE, yes=True, no=True)
    if not make_html:
        script.exit()

data = {
    "project": to_unicode(doc.Title),
    "generated": datetime.now().strftime("%d/%m/%Y %H:%M"),
    "mode": MODE, "modeLabel": MODES[MODE],
    "scope": (u"Vista ativa: " + to_unicode(active_view.Name)) if use_view else u"Modelo inteiro",
    "materials": opts["materials"],
    "inParam": PARAM_IN_NAME,
    "excel": {"path": XLSX_PATH, "name": os.path.basename(XLSX_PATH), "sheet": XL["sheet"],
              "rows": len(XL["rows"]), "backup": xl_written["backup"]},
    "summary": {"model": len(MODEL), "excel": len(XL["by_key"]),
                "both": sum(1 for k in keynotes_out if k["inModel"] and k["inExcel"]),
                "holders": sum(len(v) for v in MODEL.values())},
    "issues": [{"code": c, "label": v["label"], "color": v["color"], "hint": v["hint"],
                "count": issue_count[c]} for c, v in ISSUES.items()],
    "categories": [{"name": c, "keynotes": v["keynotes"], "issues": v["issues"]}
                   for c, v in sorted(cat_summary.items())],
    "missingIn": [{"category": c, "count": n} for c, n in sorted(missing_in_by_cat.items())],
    "keynotes": keynotes_out,
    "catOrder": XL["cat_order"] + [NO_GROUP],
    "verify": [dict(r["view"], status=r["status"]) for r in VROWS],
    "excelTable": excel_table(),
    "highlight": [dict((k, v) for k, v in x.items() if k != "ids") for x in HL_FINAL],
    "highlightNote": hl_note if opts["highlight"] else u"desligado",
    "txt": txt_info,
    "result": RESULT,
    "changes": changes,
    "skipped": [dict((k, v) for k, v in s.items() if k != "eid") for s in skipped],
}

# modelo do relatório: keynotes.html (aceita script.html por compatibilidade)
template_path = None
for _name in ("keynotes.html", "script.html"):
    if os.path.isfile(os.path.join(HERE, _name)):
        template_path = os.path.join(HERE, _name)
        break
if template_path is None:
    forms.alert(u"Modelo do relatório não encontrado (keynotes.html).",
                sub_msg=u"Pasta do botão:\n{}".format(HERE), title=ALERT_TITLE, exitscript=True)
with codecs.open(template_path, "r", encoding="utf-8") as f:
    html = f.read()
html = brand_report(html.replace("__DATA__", json.dumps(data, ensure_ascii=False,
                                                        default=lambda o: to_unicode(o))))
out_path = script.get_document_data_file("keynotes_report", "html")
with codecs.open(out_path, "w", encoding="utf-8") as f:
    f.write(html)

if not open_in_browser(out_path):
    forms.alert(u"Relatório gerado, mas não abriu automaticamente.",
                sub_msg=u"Abra o arquivo manualmente:\n{}".format(out_path), title=ALERT_TITLE)
