# -*- coding: utf-8 -*-
"""ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE.

Lê a planilha de Keynotes do escritório (OCA_XXX_Lista_de_Materiais - aba
MATERIAIS) e atualiza as descrições dos TIPOS do modelo pelo código de Keynote.
Fluxo único: EXCEL -> MODELO. Escopo: modelo inteiro. Materiais não são lidos.

Onde cada dado mora:
  * Keynote ........ parâmetro nativo KEYNOTE dos TIPOS usados no modelo.
  * Descrição PT ... parâmetro nativo "Description" do tipo.
  * Descrição IN ... parâmetro compartilhado "Descrição IN" do tipo, associado
                     pelo GUID 70307f6e-3b31-4049-bed4-62dfc3570fec
                     (OCA_Parametros_Template.txt). Se faltar em alguma
                     categoria, oferece vincular ao projeto como parâmetro de TIPO.
  * Excel .......... aba MATERIAIS: colunas KEYNOTE, DESCRIÇÃO, DESCRIÇÃO IN
                     (somente leitura - o Excel não é alterado).

Aba 1 - configurações: planilha, destaque na vista ativa e TXT de Keynotes.

Aba 2 - tabela de verificação (uma janela, duas tabelas, busca e agrupamento
por Keynote ou por Categoria):
  1. Excel -> Modelo: o que o Excel tem e difere do modelo. Tipos com o mesmo
     Keynote, mesma Descrição, mesma Descrição IN e mesma Categoria viram UMA
     linha. Valor vazio no Excel nunca apaga o modelo.
  2. Elementos do modelo SEM Keynote ou SEM Descrição: editável - digite
     Keynote / Descrição / Descrição IN, "Verificar" confere com o Excel e
     completa os campos vazios a partir dele.
  "Gravar marcados" grava no modelo só as linhas marcadas das duas tabelas.

Destaque temporário (vista ativa, desfeito ao fechar a tabela): LARANJA nos
elementos cujo tipo não tem Keynote, Descrição ou Descrição IN.

TXT de Keynotes do Revit (opcional, confirma no final): gerado do Excel com a
mesma estrutura do arquivo do escritório - CATEGORIA/prefixos, PREFIXO/GRUPO/
CATEGORIA e KEYNOTE / "DESCRIÇÃO | DESCRIÇÃO IN" / PREFIXO - e carregado no
projeto.

Ao final pergunta se gera o relatório HTML.
"""

__title__ = "Atualização\ndas Descrições\npor Keynote"
__doc__ = ("Lê a planilha de Keynotes (aba MATERIAIS) e atualiza a Descrição, "
           "Descrição IN e Texto do Keynote do modelo pelo código de Keynote.\n"
           "Fluxo: Excel > Modelo.")

import os
import re
import json
import codecs
from datetime import datetime
from collections import OrderedDict, defaultdict

from pyrevit import revit, DB, script, forms

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

TOOL_TITLE = u"ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE"
# fluxo único: Excel -> Modelo (o Excel é só lido)
MODES = OrderedDict([("to_model", u"Excel → Modelo")])
MODE_DIRS = {"to_model": ("model",)}


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
# 1. Janela de opções
# ------------------------------------------------------------------
KN_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Atualização das Descrições do Modelo por Keynote" Height="Auto" Width="580"
        SizeToContent="Height" WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize" Background="#0E1526">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,12,0,4"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#D9E8F5"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,8,0,0"/>
    </Style>
  </Window.Resources>
  <StackPanel Margin="18">
    <TextBlock Text="ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE" FontSize="15"
               FontWeight="SemiBold" Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,0,0,4"
               Text="Lê a aba MATERIAIS da planilha e atualiza a Description e a Descrição IN dos tipos do modelo pelo Keynote. Fluxo: Excel → Modelo (o Excel não é alterado). Escopo: modelo inteiro, sem materiais. Na próxima etapa você confere a tabela e grava só o que marcar."/>

    <TextBlock Text="1. Planilha Excel de Keynotes (.xlsx)"/>
    <DockPanel LastChildFill="True">
      <Button x:Name="browse" DockPanel.Dock="Right" Content="Procurar..." Width="90"
              Height="26" Margin="8,0,0,0"/>
      <TextBox x:Name="xlsx" Height="26" Padding="4,3,4,2"/>
    </DockPanel>

    <TextBlock Text="2. Fluxo e escopo"/>
    <TextBlock Margin="0,0,0,0" FontSize="12" Foreground="#D9E8F5"
               Text="Excel → Modelo  ·  modelo inteiro  ·  somente tipos (materiais não são lidos)"/>

    <TextBlock Text="3. Verificação visual"/>
    <CheckBox x:Name="highlight" Margin="0,2,0,0"
              Content="Destacar em LARANJA na vista ativa os elementos sem Keynote, Descrição ou Descrição IN"/>

    <TextBlock Text="4. TXT de Keynotes do Revit"/>
    <CheckBox x:Name="txt_on" Margin="0,2,0,6"
              Content="Gerar o TXT a partir do Excel e carregar no projeto (confirma no final)"/>
    <DockPanel LastChildFill="True">
      <Button x:Name="txt_browse" DockPanel.Dock="Right" Content="Salvar como..." Width="90"
              Height="26" Margin="8,0,0,0"/>
      <TextBox x:Name="txt_path" Height="26" Padding="4,3,4,2"/>
    </DockPanel>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
      <Button x:Name="cancel" Content="Cancelar" Width="90" Height="28" Margin="0,0,10,0"/>
      <Button x:Name="ok" Content="Executar  →" Width="120" Height="28"/>
    </StackPanel>
  </StackPanel>
</Window>
"""


def default_txt_path(xlsx_path):
    """Mesmo nome/pasta da planilha (OCA_XXX_Lista_de_Materiais_0X.txt)."""
    return os.path.splitext(xlsx_path)[0] + u".txt" if xlsx_path else u""


class KeynoteWindow(forms.WPFWindow):
    def __init__(self, xaml, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.result = None
        self.xlsx.Text = cfg["xlsx"]
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
            forms.alert(u"Selecione um arquivo .xlsx válido.", title=__title__.replace("\n", " "))
            return
        txt = to_unicode(self.txt_path.Text).strip().strip('"') or default_txt_path(path)
        if self.txt_on.IsChecked and not os.path.isdir(os.path.dirname(txt) or u"."):
            forms.alert(u"Pasta do TXT de Keynotes não existe:\n{}".format(txt))
            return
        self.result = {
            "xlsx": path,
            "mode": "to_model",            # fluxo único
            "scope": "model",              # sempre o modelo inteiro
            "materials": False,            # materiais não são lidos
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
             "highlight": config.get_option("highlight", True),
             "txt_on": config.get_option("txt_on", False),
             "txt_path": to_unicode(config.get_option("txt_path", u""))}

win = KeynoteWindow(KN_XAML, saved_cfg)
win.ShowDialog()
if not win.confirmed:
    script.exit()
opts = win.result

config.xlsx_path = opts["xlsx"]
config.highlight = opts["highlight"]
config.txt_on = opts["txt_on"]
config.txt_path = opts["txt_path"]
script.save_config()

XLSX_PATH = opts["xlsx"]
MODE = "to_model"                  # Excel -> Modelo
active_view = doc.ActiveView
use_view = False                   # escopo: modelo inteiro


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
    forms.alert(u"Não foi possível ler a planilha:\n{}".format(to_unicode(ex)),
                exitscript=True)


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
    """Categorias (Category) sem o parâmetro: tipos com keynote (tabela 1) e
    tipos com parâmetro Keynote ainda vazio (tabela 2)."""
    cats = OrderedDict()
    for info in scan_types():
        try:
            if in_param(info["t"]) is None and info["t"].Category is not None:
                cats.setdefault(eid_int(info["t"].Category.Id), info["t"].Category)
        except Exception:
            continue
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
# 5a. Backup de arquivos gravados (TXT de Keynotes)
# ------------------------------------------------------------------
def backup_path(path):
    """<arquivo>_backup_AAAAMMDD_HHMMSS.<ext>, sem colidir com um backup já existente."""
    base, ext = os.path.splitext(path)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cand, n = u"{}_backup_{}{}".format(base, stamp, ext), 1
    while os.path.exists(cand):
        n += 1
        cand = u"{}_backup_{}_{}{}".format(base, stamp, n, ext)
    return cand


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
# 5d. Situação dos TIPOS no modelo + destaque temporário (vista ativa)
#   LARANJA: tipo sem Keynote, sem Descrição (Description) ou sem
#   Descrição IN (parâmetro ausente ou vazio). As cores são sobreposições de
#   vista aplicadas numa transação DESFEITA ao fechar a tabela - nada fica
#   gravado no modelo.
# ------------------------------------------------------------------
HL_RGB = (255, 140, 0)
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


def scan_types(view=None):
    """Tipos COM parâmetro Keynote usados por elementos de categorias de modelo
    (no modelo inteiro ou só os visíveis na vista). -> lista de dicts."""
    try:
        col = (DB.FilteredElementCollector(doc, view.Id) if view is not None
               else DB.FilteredElementCollector(doc))
        elems = col.WhereElementIsNotElementType().ToElements()
    except Exception:
        elems = []
    by_type = OrderedDict()
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
        info = by_type.setdefault(eid_int(tid), {"tid": tid, "ids": [],
                                                 "category": to_unicode(cat.Name)})
        info["ids"].append(e.Id)
    out = []
    for info in by_type.values():
        t = doc.GetElement(info["tid"])
        if t is None:
            continue
        try:
            if t.get_Parameter(BIP.KEYNOTE_PARAM) is None:
                continue                   # categoria sem parâmetro Keynote
        except Exception:
            continue
        info["t"] = t
        out.append(info)
    return out


def type_status(t):
    """-> (keynote, descrição, descrição IN, tem parâmetro IN, campos faltando)"""
    key = keynote_of(t)
    desc = param_text(desc_param(t))
    ip = in_param(t)
    en = param_text(ip) if ip is not None else u""
    missing = [lbl for lbl, v in ((u"Keynote", key), (u"Descrição", desc),
                                  (PARAM_IN_NAME, en)) if not v]
    return key, desc, en, ip is not None, missing


def collect_highlight():
    """Tipos visíveis na vista ativa sem Keynote, Descrição ou Descrição IN."""
    global hl_note
    if not view_allows_overrides(active_view):
        hl_note = u"vista ativa não aceita sobreposição de cores (tabela, folha...)"
        return []
    out = []
    for info in scan_types(active_view):
        t = info["t"]
        key, desc, en, has_in, missing = type_status(t)
        if not missing:
            continue
        out.append({"color": u"orange", "missing": missing, "category": info["category"],
                    "family": family_name(t), "type": elem_name(t), "typeId": eid_int(t.Id),
                    "key": key, "desc": desc, "en": en, "hasIn": has_in,
                    "count": len(info["ids"]), "ids": info["ids"]})
    out.sort(key=lambda x: (x["category"], x["family"], x["type"]))
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
    """Aplica o LARANJA numa transação ABERTA e redesenha a vista. Devolve a
    transação para ser desfeita (RollBack) quando a tabela fechar."""
    if not hl:
        return None
    fill = solid_fill_id()
    c = DB.Color(*HL_RGB)
    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(c)
    ogs.SetCutLineColor(c)
    if fill != INVALID_ID:
        ogs.SetSurfaceForegroundPatternId(fill)
        ogs.SetSurfaceForegroundPatternColor(c)
        ogs.SetCutForegroundPatternId(fill)
        ogs.SetCutForegroundPatternColor(c)
    t = DB.Transaction(doc, u"Destaque temporário - Keynotes")
    t.Start()
    for item in hl:
        for eid in item["ids"]:
            try:
                active_view.SetElementOverrides(eid, ogs)
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
# 5e. Tabela 1 - Excel -> Modelo
#     Tipos com o MESMO Keynote, mesma Descrição, mesma Descrição IN e mesma
#     Categoria (e a mesma alteração proposta) viram UMA linha.
# ------------------------------------------------------------------
def model_rows():
    plan, blocked = plan_to_model()
    by_h = OrderedDict()
    for it in plan + blocked:
        by_h.setdefault(it["h"]["id"], []).append(it)
    merged = OrderedDict()
    for its in by_h.values():
        h = its[0]["h"]
        new = {"pt": None, "en": None}
        blk = {}
        for it in its:
            if it["reason"]:
                blk[it["field"]] = it["reason"]
            else:
                new[it["field"]] = it["new"]
        gkey = (h["key"], h["category"], h["pt"], h["en"], new["pt"], new["en"],
                tuple(sorted(blk.items())))
        r = merged.get(gkey)
        if r is None:
            r = merged[gkey] = {
                "dir": "model", "key": h["key"], "h": h, "hs": [], "items": [],
                "category": h["category"], "pt_old": h["pt"], "en_old": h["en"],
                "pt_new": new["pt"], "en_new": new["en"], "over": set(), "blocked": blk}
        r["hs"].append(h)
        r["items"].extend(its)
        for it in its:
            if not it["reason"] and it["old"]:
                r["over"].add(it["field"])
    rows = list(merged.values())
    for r in rows:
        r["enabled"] = any(not it["reason"] for it in r["items"])
    return rows, plan, blocked


def _names(values, n=2):
    vals = []
    for v in values:
        if v and v not in vals:
            vals.append(v)
    if len(vals) <= n:
        return u", ".join(vals)
    return u"{} (+{})".format(u", ".join(vals[:n]), len(vals) - n)


def row_view(r):
    """Texto das colunas da tabela 1 (janela e HTML)."""
    cat, grp = agrup_of(r["key"])
    hs = r["hs"]
    fam = _names([h["family"] or h["category"] for h in hs])
    if len(hs) == 1:
        typ = hs[0]["name"]
    else:
        typ = u"{} tipos: {}".format(len(hs), _names([h["name"] for h in hs], 3))

    def cur(v, no_param=False):
        if no_param:
            return u"(sem parâmetro)"
        return v or u"(vazio)"

    def new(f):
        v = r[f + "_new"]
        if v is not None:
            return v or u"(vazio)"
        if f in r["blocked"]:
            return u"(bloqueado)"
        return u"(mantém)"

    dest = u"→ Modelo"
    notes = []
    if r["over"]:
        notes.append(u"sobrescreve")
    for f, why in sorted(r["blocked"].items()):
        notes.append(u"{} bloqueado: {}".format(u"PT" if f == "pt" else u"IN", why))
    if notes:
        dest += u" · " + u" · ".join(notes)
    kind = u"blocked" if not r["enabled"] else (u"over" if r["over"] else u"fill")
    return {"cat": cat, "grp": grp, "category": r["category"], "family": fam,
            "type": typ, "nTypes": len(hs), "count": sum(h["count"] for h in hs),
            "key": r["key"], "ptCur": cur(r["pt_old"]), "ptNew": new("pt"),
            "enCur": cur(r["en_old"], r["h"]["en"] is None), "enNew": new("en"),
            "dest": dest, "dir": r["dir"], "kind": kind}


# ------------------------------------------------------------------
# 5e'. Tabela 2 - tipos do modelo SEM Keynote ou SEM Descrição (editável)
# ------------------------------------------------------------------
def collect_incomplete():
    rows = []
    for info in scan_types():
        t = info["t"]
        key, desc, en, has_in, missing = type_status(t)
        if key and desc:
            continue
        rows.append({"t": t, "id": eid_int(t.Id), "category": info["category"],
                     "family": family_name(t), "name": elem_name(t),
                     "count": len(info["ids"]), "hasIn": has_in,
                     "key0": key, "pt0": desc, "en0": en,
                     "key": key, "pt": desc, "en": en,
                     "sel": False, "check": u"", "status": u""})
    rows.sort(key=lambda r: (r["category"], r["family"], r["name"]))
    return rows


def inc_changed(r):
    """Algum campo digitado difere do valor atual? (vazio nunca apaga)"""
    return any(clean(r[f]) and clean(r[f]) != r[f + "0"] for f in ("key", "pt", "en"))


def verify_incomplete(r, autofill):
    """Confere a linha com o Excel. autofill: completa Descrição / Descrição IN
    vazias com os valores do Excel quando o Keynote existe na planilha."""
    for f in ("key", "pt", "en"):
        r[f] = clean(r[f])
    rec = XL["by_key"].get(r["key"]) if r["key"] else None
    filled = []
    if autofill and rec is not None:
        if not r["pt"] and rec["pt"]:
            r["pt"] = rec["pt"]
            filled.append(u"Descrição")
        if not r["en"] and rec["en"] and r["hasIn"]:
            r["en"] = rec["en"]
            filled.append(PARAM_IN_NAME)
    errs, warns = [], []
    if not r["key"]:
        errs.append(u"falta Keynote")
    elif rec is None:
        warns.append(u"Keynote não está no Excel")
    if not r["pt"]:
        errs.append(u"falta Descrição")
    elif rec is not None and rec["pt"] and r["pt"] != rec["pt"]:
        warns.append(u"Descrição difere do Excel")
    if not r["hasIn"]:
        warns.append(u"sem parâmetro {}".format(PARAM_IN_NAME))
    elif not r["en"]:
        warns.append(u"{} vazia".format(PARAM_IN_NAME))
    elif rec is not None and rec["en"] and r["en"] != rec["en"]:
        warns.append(u"{} difere do Excel".format(PARAM_IN_NAME))
    parts = []
    if errs:
        parts.append(u"Erro: " + u", ".join(errs))
    if warns:
        parts.append(u"Aviso: " + u", ".join(warns))
    if filled:
        parts.append(u"Completado do Excel: " + u", ".join(filled))
    r["check"] = u" · ".join(parts) if parts else u"OK"
    return r["check"]


# ------------------------------------------------------------------
# 5f. Janela de verificação (Aba 2): duas tabelas, busca e agrupamento
# ------------------------------------------------------------------
for _asm in ("PresentationFramework", "PresentationCore", "WindowsBase", "System.Data"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
import System
from System.Data import DataTable

_GRID_STYLE = u"""
              AutoGenerateColumns="False" CanUserAddRows="False"
              CanUserDeleteRows="False" CanUserSortColumns="False" CanUserReorderColumns="False"
              CanUserResizeRows="False" HeadersVisibility="Column" SelectionMode="Single"
              SelectionUnit="FullRow" GridLinesVisibility="Horizontal"
              HorizontalGridLinesBrush="#1C2B44" Background="#0A1120" RowBackground="#0E1526"
              AlternatingRowBackground="#111C31" Foreground="#D9E8F5" BorderBrush="#2A4A66"
              BorderThickness="1" FontFamily="Segoe UI" FontSize="12"
              VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
              EnableRowVirtualization="True" FrozenColumnCount="1"
              VirtualizingPanel.IsVirtualizingWhenGrouping="True"
              ColumnHeaderStyle="{StaticResource head}" RowStyle="{StaticResource row}" """

# cabeçalho de grupo em linha única, na largura toda da tabela
_GROUP_STYLE = u"""
      <DataGrid.GroupStyle>
        <GroupStyle>
          <GroupStyle.HeaderTemplate>
            <DataTemplate>
              <Border Background="#1B2A47" BorderBrush="#2A4A66" BorderThickness="0,1,0,1"
                      Padding="8,4,8,4">
                <DockPanel LastChildFill="True">
                  <Button DockPanel.Dock="Right" Content="Desmarcar grupo" Tag="{Binding Name}"
                          Height="22" Padding="8,0,8,0" Margin="6,0,0,0" FontSize="11"/>
                  <Button DockPanel.Dock="Right" Content="Marcar grupo" Tag="{Binding Name}"
                          Height="22" Padding="8,0,8,0" Margin="6,0,0,0" FontSize="11"/>
                  <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="{Binding Name}" Foreground="#65E3FF" FontWeight="Bold"
                               FontFamily="Segoe UI" TextTrimming="CharacterEllipsis"/>
                    <TextBlock Text="{Binding ItemCount, StringFormat='   ·   {0} linha(s)'}"
                               Foreground="#7A8FA9" FontFamily="Segoe UI"/>
                  </StackPanel>
                </DockPanel>
              </Border>
            </DataTemplate>
          </GroupStyle.HeaderTemplate>
        </GroupStyle>
      </DataGrid.GroupStyle>"""

_CHECK_COL = u"""
        <DataGridTemplateColumn Header="✓" Width="36">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                        IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}"
                        IsEnabled="{Binding Enabled}"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>"""


def _col(header, binding, width, style=u"wrap", readonly=True):
    return (u'<DataGridTextColumn Header="{}" Binding="{{Binding {}}}" Width="{}" '
            u'IsReadOnly="{}" ElementStyle="{{StaticResource {}}}"/>'.format(
                header, binding, width, u"True" if readonly else u"False", style))


VERIFY_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Atualização das Descrições do Modelo por Keynote · Verificação"
        Width="1450" Height="900" MinWidth="900" MinHeight="560"
        WindowStartupLocation="CenterScreen" Background="#0E1526">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Height" Value="28"/>
      <Setter Property="Padding" Value="12,0,12,0"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style x:Key="wrap" TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
    </Style>
    <Style x:Key="mono" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
    </Style>
    <Style x:Key="edit" TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
      <Setter Property="Background" Value="#16233B"/>
    </Style>
    <Style x:Key="lbl" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#65E3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,16,0"/>
    </Style>
    <Style x:Key="head" TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#13203A"/>
      <Setter Property="Foreground" Value="#65E3FF"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="6,6,6,6"/>
      <Setter Property="BorderBrush" Value="#2A4A66"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style x:Key="row" TargetType="DataGridRow">
      <Style.Triggers>
        <DataTrigger Binding="{Binding Kind}" Value="fill">
          <Setter Property="Foreground" Value="#7BE3A0"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Kind}" Value="over">
          <Setter Property="Foreground" Value="#FFB454"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Kind}" Value="blocked">
          <Setter Property="Foreground" Value="#7A8FA9"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Kind}" Value="changed">
          <Setter Property="Foreground" Value="#7BE3A0"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="3*"/>
      <RowDefinition Height="6"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="2*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0">
      <TextBlock Text="ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE" FontSize="15"
                 FontWeight="SemiBold" Foreground="#65E3FF" FontFamily="Segoe UI"/>
      <TextBlock x:Name="info" TextWrapping="Wrap" FontSize="12" Foreground="#CFE3FF"
                 FontFamily="Segoe UI" Margin="0,6,0,0"/>
      <TextBlock x:Name="hlinfo" TextWrapping="Wrap" FontSize="12" Foreground="#FFB454"
                 FontFamily="Segoe UI" Margin="0,4,0,0"/>
    </StackPanel>
    <WrapPanel Grid.Row="1" Margin="0,10,0,6">
      <TextBlock Text="Agrupar por " Foreground="#CFE3FF" VerticalAlignment="Center" FontFamily="Segoe UI"/>
      <ComboBox x:Name="f_mode" Width="130" Height="26" Margin="0,0,14,0"/>
      <TextBlock Text="Grupo " Foreground="#CFE3FF" VerticalAlignment="Center" FontFamily="Segoe UI"/>
      <ComboBox x:Name="f_grp" Width="230" Height="26" Margin="0,0,14,0"/>
      <TextBlock Text="Buscar " Foreground="#CFE3FF" VerticalAlignment="Center" FontFamily="Segoe UI"/>
      <TextBox x:Name="f_text" Width="260" Height="26" Margin="0,0,14,0" Padding="4,3,4,2"/>
    </WrapPanel>
    <DockPanel Grid.Row="2" Margin="0,6,0,6" LastChildFill="False">
      <TextBlock Style="{StaticResource lbl}"
                 Text="1 · EXCEL → MODELO  (tipos iguais agrupados numa linha)"/>
      <Button x:Name="b_all" Content="Marcar visíveis"/>
      <Button x:Name="b_none" Content="Desmarcar visíveis"/>
      <Button x:Name="b_fill" Content="Somente preencher vazios"/>
      <TextBlock DockPanel.Dock="Right" VerticalAlignment="Center" FontSize="11"
                 Foreground="#7A8FA9" FontFamily="Segoe UI"
                 Text="Verde = preenche vazio · Âmbar = sobrescreve · Cinza = não pode ser gravado"/>
    </DockPanel>
    <DataGrid x:Name="grid" Grid.Row="3" """ + _GRID_STYLE + u""">
      <DataGrid.Columns>""" + _CHECK_COL + u"""
        """ + _col(u"Agrupamento", u"Grupo", 150) + u"""
        """ + _col(u"Categoria", u"Categoria", 110) + u"""
        """ + _col(u"Família", u"Familia", 130) + u"""
        """ + _col(u"Tipo", u"Tipo", 170) + u"""
        """ + _col(u"Keynote", u"Keynote", 75, u"mono") + u"""
        """ + _col(u"Descrição atual", u"PtAtual", 180) + u"""
        """ + _col(u"Descrição nova", u"PtNova", 180) + u"""
        """ + _col(u"Descrição em IN atual", u"EnAtual", 160) + u"""
        """ + _col(u"Descrição em IN nova", u"EnNova", 160) + u"""
        """ + _col(u"Destino", u"Destino", 150) + u"""
      </DataGrid.Columns>""" + _GROUP_STYLE + u"""
    </DataGrid>
    <GridSplitter Grid.Row="4" Height="6" HorizontalAlignment="Stretch" Background="#1C2B44"/>
    <DockPanel Grid.Row="5" Margin="0,8,0,6" LastChildFill="False">
      <TextBlock Style="{StaticResource lbl}"
                 Text="2 · ELEMENTOS DO MODELO SEM KEYNOTE OU SEM DESCRIÇÃO  (clique duas vezes para digitar)"/>
      <Button x:Name="b2_all" Content="Marcar visíveis"/>
      <Button x:Name="b2_none" Content="Desmarcar visíveis"/>
      <Button x:Name="b2_check" Content="Verificar"/>
    </DockPanel>
    <DataGrid x:Name="grid2" Grid.Row="6" """ + _GRID_STYLE + u""">
      <DataGrid.Columns>""" + _CHECK_COL + u"""
        """ + _col(u"Categoria", u"Categoria", 110) + u"""
        """ + _col(u"Família", u"Familia", 140) + u"""
        """ + _col(u"Tipo", u"Tipo", 180) + u"""
        """ + _col(u"Qtd", u"Qtd", 45, u"mono") + u"""
        """ + _col(u"Keynote ✎", u"Keynote", 95, u"edit", False) + u"""
        """ + _col(u"Descrição ✎", u"Descricao", 230, u"edit", False) + u"""
        """ + _col(u"Descrição em IN ✎", u"DescricaoIN", 210, u"edit", False) + u"""
        """ + _col(u"Verificação", u"Verificacao", 300) + u"""
      </DataGrid.Columns>""" + _GROUP_STYLE + u"""
    </DataGrid>
    <DockPanel Grid.Row="7" Margin="0,12,0,0" LastChildFill="False">
      <TextBlock x:Name="counter" DockPanel.Dock="Left" VerticalAlignment="Center"
                 Foreground="#CFE3FF" FontFamily="Segoe UI" TextWrapping="Wrap" MaxWidth="900"/>
      <Button x:Name="b_apply" DockPanel.Dock="Right" Content="Gravar marcados"
              Width="160" Margin="0"/>
      <Button x:Name="b_cancel" DockPanel.Dock="Right" Content="Cancelar" Width="100"
              Margin="0,0,10,0"/>
    </DockPanel>
  </Grid>
</Window>
"""

GROUP_MODES = [u"Keynote", u"Categoria"]
COLS1 = ("Grupo", "Categoria", "Familia", "Tipo", "Keynote", "PtAtual", "PtNova",
         "EnAtual", "EnNova", "Destino")
COLS2 = ("Categoria", "Familia", "Tipo", "Qtd", "Keynote", "Descricao", "DescricaoIN",
         "Verificacao")
EDIT2 = {"Keynote": "key", "Descricao": "pt", "DescricaoIN": "en"}
NO_KEY = u"(sem keynote)"


def _flt(text):
    """Texto seguro para DataView.RowFilter."""
    t = to_unicode(text).replace(u"'", u"''")
    for ch in u"[]*%":
        t = t.replace(ch, u"")
    return t


def _key_label(key):
    if not key:
        return NO_KEY
    cat, grp = agrup_of(key)
    return u"{}   ·   {} ({})".format(key, grp, cat)


class VerifyWindow(forms.WPFWindow):
    def __init__(self, rows, inc, info_text, hl_text):
        forms.WPFWindow.__init__(self, VERIFY_XAML, literal_string=True)
        self.confirmed, self._busy = False, False
        self.rows, self.inc = rows, inc
        self.by_rid = dict((r["rid"], r) for r in rows)
        self.by_rid2 = dict((r["rid"], r) for r in inc)
        self.info.Text = info_text
        self.hlinfo.Text = hl_text

        self.t1 = self._new_table(COLS1)
        self.t2 = self._new_table(COLS2)
        self.d1, self.d2 = {}, {}
        self.grid.ItemsSource = self.t1.DefaultView
        self.grid2.ItemsSource = self.t2.DefaultView
        from System.Windows.Data import PropertyGroupDescription
        from System.Windows import RoutedEventHandler
        from System.Windows.Controls.Primitives import ButtonBase
        for g in (self.grid, self.grid2):
            g.Items.GroupDescriptions.Add(PropertyGroupDescription("Grp"))
        # botões "Marcar grupo / Desmarcar grupo" dos cabeçalhos de grupo
        self.grid.AddHandler(ButtonBase.ClickEvent, RoutedEventHandler(self._grp_click1))
        self.grid2.AddHandler(ButtonBase.ClickEvent, RoutedEventHandler(self._grp_click2))

        self.f_mode.ItemsSource = GROUP_MODES
        self.f_mode.SelectedIndex = 0
        self._rebuild()

        self.f_mode.SelectionChanged += self._on_mode
        self.f_grp.SelectionChanged += self._filter
        self.f_text.TextChanged += self._filter
        self.t1.ColumnChanged += self._on_changed1
        self.t2.ColumnChanged += self._on_changed2
        self.b_all.Click += lambda s, a: self._set_visible1(lambda r: True)
        self.b_none.Click += lambda s, a: self._set_visible1(lambda r: False)
        self.b_fill.Click += lambda s, a: self._set_visible1(lambda r: not r["over"])
        self.b2_all.Click += lambda s, a: self._set_visible2(True)
        self.b2_none.Click += lambda s, a: self._set_visible2(False)
        self.b2_check.Click += self._verify_all
        self.b_apply.Click += self._apply
        self.b_cancel.Click += self._cancel

    # ---- montagem das tabelas (agrupadas pela coluna Grp) ----
    @staticmethod
    def _new_table(cols):
        tb = DataTable("t")
        for name, typ in (("RowId", System.Int32), ("Sel", System.Boolean),
                          ("Enabled", System.Boolean), ("Kind", System.String),
                          ("Grp", System.String)) + tuple((c, System.String) for c in cols):
            tb.Columns.Add(name, clr.GetClrType(typ))
        return tb

    def _mode(self):
        return self.f_mode.SelectedItem or GROUP_MODES[0]

    def _grp1(self, r):
        return _key_label(r["key"]) if self._mode() == u"Keynote" else r["category"]

    def _grp2(self, r):
        return _key_label(clean(r["key"])) if self._mode() == u"Keynote" else r["category"]

    def _fill(self, tb, store, items, grp_of, values_of, sort_extra):
        tb.Rows.Clear()
        store.clear()
        ordered = sorted(items, key=lambda r: (grp_of(r) == NO_KEY, grp_of(r)) + sort_extra(r))
        for r in ordered:
            g = grp_of(r)
            row = tb.NewRow()
            row["RowId"], row["Grp"] = r["rid"], g
            vals = values_of(r)
            row["Kind"] = vals.pop("_kind")
            row["Sel"], row["Enabled"] = bool(r["sel"]), bool(r.get("enabled", True))
            for c, v in vals.items():
                row[c] = v
            tb.Rows.Add(row)
            store[r["rid"]] = row

    def _vals1(self, r):
        v = r["view"]
        return dict(zip(COLS1, (v["grp"], v["category"], v["family"], v["type"], v["key"],
                                v["ptCur"], v["ptNew"], v["enCur"], v["enNew"], v["dest"])),
                    _kind=v["kind"])

    def _vals2(self, r):
        return dict(zip(COLS2, (r["category"], r["family"], r["name"], to_unicode(r["count"]),
                                r["key"], r["pt"], r["en"], r["check"])),
                    _kind=u"changed" if inc_changed(r) else u"inc")

    def _rebuild(self):
        self._busy = True
        try:
            self._fill(self.t1, self.d1, self.rows, self._grp1, self._vals1,
                       lambda r: (r["key"], r["view"]["category"], r["view"]["family"],
                                  r["view"]["type"]))
            self._fill(self.t2, self.d2, self.inc, self._grp2, self._vals2,
                       lambda r: (r["category"], r["family"], r["name"]))
            groups = []
            for tb in (self.t1, self.t2):
                for row in tb.Rows:
                    if row["Grp"] not in groups:
                        groups.append(row["Grp"])
            self.f_grp.ItemsSource = [u"Todos"] + sorted(groups)
            self.f_grp.SelectedIndex = 0
        finally:
            self._busy = False
        self._filter(None, None)

    def _on_mode(self, sender, args):
        self._commit()
        self._rebuild()

    # ---- filtros (busca e grupo valem para as duas tabelas) ----
    def _filter(self, sender, args):
        if self._busy:
            return
        g = self.f_grp.SelectedItem
        t = _flt(self.f_text.Text).strip()
        for tb, cols in ((self.t1, COLS1), (self.t2, COLS2)):
            parts = []
            if g and g != u"Todos":
                parts.append(u"Grp = '{}'".format(_flt(g)))
            if t:
                like = u" OR ".join(u"{} LIKE '%{}%'".format(c, t) for c in cols)
                parts.append(u"({})".format(like))
            try:
                tb.DefaultView.RowFilter = u" AND ".join(parts)
            except Exception:
                tb.DefaultView.RowFilter = u""
        self._update_counter()

    @staticmethod
    def _visible(tb):
        return [int(drv.Row["RowId"]) for drv in tb.DefaultView]

    @staticmethod
    def _group_button(args):
        """-> (grupo, marcar?) se o clique veio de um botão de cabeçalho de grupo."""
        from System.Windows.Controls import Button
        src = args.OriginalSource
        if not isinstance(src, Button) or src.Tag is None:
            return None, None
        return to_unicode(src.Tag), to_unicode(src.Content).startswith(u"Marcar")

    def _grp_click1(self, sender, args):
        grp, val = self._group_button(args)
        if grp is None:
            return
        self._busy = True
        try:
            for rid in self._visible(self.t1):
                if self.d1[rid]["Grp"] == grp:
                    self._check1(rid, val)
            self._update_counter()
        finally:
            self._busy = False

    def _grp_click2(self, sender, args):
        grp, val = self._group_button(args)
        if grp is None:
            return
        self._busy = True
        try:
            for rid in self._visible(self.t2):
                if self.d2[rid]["Grp"] == grp:
                    self._check2(rid, val)
            self._update_counter()
        finally:
            self._busy = False

    def _commit(self):
        """Confirma uma célula em edição na tabela 2 antes de ler os valores."""
        try:
            from System.Windows.Controls import DataGridEditingUnit
            self.grid2.CommitEdit(DataGridEditingUnit.Cell, True)
            self.grid2.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            pass

    # ---- tabela 1 ----
    def _check1(self, rid, value):
        r = self.by_rid[rid]
        r["sel"] = bool(value) and r["enabled"]
        self.d1[rid]["Sel"] = r["sel"]

    def _on_changed1(self, sender, e):
        if self._busy or e.Column.ColumnName != "Sel":
            return
        self._busy = True
        try:
            self._check1(int(e.Row["RowId"]), e.Row["Sel"] == True)
            self._update_counter()
        finally:
            self._busy = False

    def _set_visible1(self, rule):
        self._busy = True
        try:
            for rid in self._visible(self.t1):
                self._check1(rid, rule(self.by_rid[rid]))
            self._update_counter()
        finally:
            self._busy = False

    # ---- tabela 2 (editável) ----
    def _check2(self, rid, value):
        r = self.by_rid2[rid]
        r["sel"] = bool(value)
        self.d2[rid]["Sel"] = r["sel"]

    def _on_changed2(self, sender, e):
        name = e.Column.ColumnName
        if self._busy or name not in ("Sel",) + tuple(EDIT2.keys()):
            return
        self._busy = True
        try:
            rid = int(e.Row["RowId"])
            if name == "Sel":
                self._check2(rid, e.Row["Sel"] == True)
            else:
                r = self.by_rid2[rid]
                r[EDIT2[name]] = to_unicode(e.Row[name] or u"")
                verify_incomplete(r, autofill=False)
                e.Row["Verificacao"] = r["check"]
                e.Row["Kind"] = u"changed" if inc_changed(r) else u"inc"
                if inc_changed(r):
                    self._check2(rid, True)
            self._update_counter()
        finally:
            self._busy = False

    def _set_visible2(self, value):
        self._busy = True
        try:
            for rid in self._visible(self.t2):
                self._check2(rid, value)
            self._update_counter()
        finally:
            self._busy = False

    def _verify_all(self, sender, args):
        """Confere todas as linhas com o Excel e completa os campos vazios."""
        self._commit()
        self._busy = True
        try:
            for r in self.inc:
                verify_incomplete(r, autofill=True)
                row = self.d2[r["rid"]]
                for col, f in EDIT2.items():
                    row[col] = r[f]
                row["Verificacao"] = r["check"]
                row["Kind"] = u"changed" if inc_changed(r) else u"inc"
                if inc_changed(r):
                    self._check2(r["rid"], True)
            self._update_counter()
        finally:
            self._busy = False

    # ---- comum ----
    def _update_counter(self):
        s1, s2 = self.selected(), self.selected_inc()
        tot1 = len([r for r in self.rows if r["enabled"]])
        n_over = len([r for r in s1 if r["over"]])
        n_types = sum(len(r["hs"]) for r in s1)
        n_edit = len([r for r in s2 if inc_changed(r)])
        self.counter.Text = (
            u"Tabela 1: {} de {} linha(s) marcada(s) ({} tipo(s), {} sobrescreve(m))   ·   "
            u"Tabela 2: {} de {} marcada(s), {} com valor digitado/completado".format(
                len(s1), tot1, n_types, n_over, len(s2), len(self.inc), n_edit))

    def selected(self):
        return [r for r in self.rows if r["sel"] and r["enabled"]]

    def selected_inc(self):
        return [r for r in self.inc if r["sel"]]

    def _apply(self, sender, args):
        self._commit()
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
# 5h. Executa: tabela de verificação -> grava no modelo o que foi marcado
# ------------------------------------------------------------------
# (holder id, campo) -> motivo de NÃO ter sido gravado (usado no resultado)
not_written = {}
DIRS = MODE_DIRS[MODE]

if "model" in MODE_DIRS[MODE] and MODE != "analyze":
    missing_cats = in_param_missing_categories()
    clash = sorted(set(h["category"] for hs in MODEL.values() for h in hs if h["en_clash"]))
    if missing_cats:
        msg = (u"O parâmetro compartilhado '{}' (GUID {}) não está disponível em:\n\n"
               u"{}\n\nVincular ao projeto como parâmetro de TIPO nessas categorias?\n"
               u"(sem ele, a descrição em inglês não é gravada nesses itens)").format(
                   PARAM_IN_NAME, PARAM_IN_GUID_STR,
                   u"\n".join(u"  • " + to_unicode(c.Name) for c in missing_cats))
        if clash:
            msg += (u"\n\nAtenção: em {} já existe um parâmetro '{}' que NÃO é o "
                    u"compartilhado do escritório (GUID diferente). Ele será ignorado."
                    ).format(u", ".join(clash), PARAM_IN_NAME)
        if forms.alert(msg, yes=True, no=True):
            ok, info = bind_in_param(missing_cats)
            output.print_md(u"- Parâmetro **{}**: {}{}".format(
                PARAM_IN_NAME, u"" if ok else u":warning: não vinculado - ", info))
            if ok:
                MODEL = collect_model(use_view, opts["materials"])   # relê os parâmetros



m_rows, m_plan, m_blocked = model_rows()
for it in m_blocked:
    h = it["h"]
    not_written[(h["id"], it["field"])] = it["reason"]
    skipped.append({"key": h["key"], "where": where_label(h), "id": h["id"],
                    "eid": h["elem"].Id, "reason": it["reason"]})

VROWS = m_rows
for r in VROWS:
    r["view"] = row_view(r)
    r["sel"] = r["enabled"]           # Excel -> Modelo começa marcado
VROWS.sort(key=lambda r: (r["key"], r["category"], r["view"]["family"], r["view"]["type"]))
for i, r in enumerate(VROWS):
    r["rid"] = i

INC = collect_incomplete()
for i, r in enumerate(INC):
    r["rid"] = i
    verify_incomplete(r, autofill=False)

HL = collect_highlight() if opts["highlight"] else []
hl_orange = sum(x["count"] for x in HL)

keys_with_info = [k for k in MODEL if XL["by_key"].get(k) and
                  (XL["by_key"][k]["pt"] or XL["by_key"][k]["en"])]
info_text = (u"{} keynote(s) no modelo · {} com descrição no Excel · {} fora do Excel  ·  "
             u"Tabela 1: {} linha(s) ({} tipo(s))  ·  Tabela 2: {} tipo(s) sem Keynote ou "
             u"sem Descrição".format(
                 len(MODEL), len(keys_with_info),
                 len([k for k in MODEL if not XL["by_key"].get(k)]),
                 len(VROWS), sum(len(r["hs"]) for r in VROWS), len(INC)))
if not opts["highlight"]:
    hl_text = u"Destaque no modelo desligado."
elif hl_note:
    hl_text = u"Destaque no modelo indisponível: " + hl_note
else:
    hl_text = (u"Destaque LARANJA na vista ativa (temporário, desfeito ao fechar): {} elemento(s) "
               u"de {} tipo(s) sem Keynote, Descrição ou {}.".format(
                   hl_orange, len(HL), PARAM_IN_NAME))

chosen_rows, chosen_inc = [], []
if VROWS or INC or HL:
    hl_tx = apply_highlight(HL)
    try:
        vw = VerifyWindow(VROWS, INC, info_text, hl_text)
        vw.ShowDialog()
        if vw.confirmed:
            chosen_rows = vw.selected()
            chosen_inc = vw.selected_inc()
        else:
            output.print_md(u"**Cancelado - nada foi gravado.**")
    finally:
        remove_highlight(hl_tx)
else:
    output.print_md(u"**Verificação:** nada a atualizar - o modelo já coincide com o Excel "
                    u"e todos os tipos têm Keynote e Descrição.")

chosen_ids = set(r["rid"] for r in chosen_rows)
for r in VROWS:
    r["status"] = (u"Bloqueado" if not r["enabled"] else
                   (u"Marcado" if r["rid"] in chosen_ids else u"Não marcado"))

# ---- Tabela 1 (Excel -> Modelo): grava os tipos das linhas marcadas ----
chosen_items = [it for r in chosen_rows for it in r["items"] if not it["reason"]]
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
        forms.alert(u"Atualização do modelo desfeita (rollback):\n{}".format(to_unicode(ex)),
                    exitscript=True)

for r in VROWS:
    if r["rid"] in failed_rows:
        r["status"] = u"Não gravado"
    elif r["status"] == u"Marcado":
        r["status"] = u"Gravado"

# ---- Tabela 2 (tipos sem Keynote / sem Descrição): grava o que foi digitado ----
INC_FIELDS = (("key", u"Keynote", lambda t: t.get_Parameter(BIP.KEYNOTE_PARAM)),
              ("pt", u"Description", desc_param),
              ("en", PARAM_IN_NAME, in_param))
chosen_inc_ids = set(r["rid"] for r in chosen_inc)
for r in INC:
    r["status"] = u"Não marcado"
    r["written"] = []
to_write = [r for r in chosen_inc if inc_changed(r)]
for r in chosen_inc:
    if not inc_changed(r):
        r["status"] = u"Sem alteração"
if to_write:
    t = DB.Transaction(doc, u"Completar Keynote / Descrição dos tipos")
    t.Start()
    try:
        for r in to_write:
            tp = r["t"]
            where = u"{}: {}{}".format(r["category"],
                                       (r["family"] + u" - ") if r["family"] else u"", r["name"])
            errs = []
            if owned_by_other(tp):
                errs.append(u"tipo emprestado por outro usuário (workshared)")
            else:
                for f, label, getp in INC_FIELDS:
                    new, old = clean(r[f]), r[f + "0"]
                    if not new or new == old:
                        continue
                    try:
                        p = getp(tp)
                    except Exception:
                        p = None
                    if p is None:
                        errs.append(u"parâmetro '{}' não existe".format(label))
                        continue
                    if param_text(p) == new:
                        r["written"].append(f)      # já gravado pela tabela 1
                        continue
                    if p.IsReadOnly:
                        errs.append(u"'{}' somente leitura".format(label))
                        continue
                    try:
                        p.Set(new)
                        r["written"].append(f)
                        changes.append({"target": "model",
                                        "action": "overwrite" if old else "fill",
                                        "key": clean(r["key"]) or NO_KEY, "field": label,
                                        "fieldKey": "inc_" + f, "category": r["category"],
                                        "old": old, "new": new, "where": where,
                                        "id": r["id"]})
                    except Exception as ex:
                        errs.append(to_unicode(ex))
            if errs:
                r["status"] = (u"Gravado em parte: " if r["written"] else u"Não gravado: ") + \
                    u"; ".join(errs)
                skipped.append({"key": clean(r["key"]) or NO_KEY, "where": where,
                                "id": r["id"], "eid": tp.Id, "reason": u"; ".join(errs)})
            else:
                r["status"] = u"Gravado"
        t.Commit()
    except Exception as ex:
        t.RollBack()
        forms.alert(u"Gravação da tabela 2 desfeita (rollback):\n{}".format(to_unicode(ex)),
                    exitscript=True)

# relê o modelo para o relatório refletir o que ficou gravado
if changes:
    MODEL = collect_model(use_view, opts["materials"])

# ---- TXT de Keynotes (opcional, confirma antes) ----
txt_info = {"on": opts["txt_on"], "path": opts["txt_path"], "written": False,
            "loaded": False, "msg": u"", "backup": u"", "encoding": u"", "warnings": [],
            "lines": 0}
if opts["txt_on"]:
    txt_text, txt_warns = build_keynote_txt()
    txt_info["warnings"] = txt_warns
    txt_info["lines"] = len(txt_text.splitlines())
    n_keys = txt_info["lines"] - len(XL["kn_headers"]) - len(XL["dados_rows"])
    if forms.alert(u"Gravar o TXT de Keynotes do Revit a partir do Excel e carregar no "
                   u"projeto?\n\n{}\n\n{} agrupamento(s) · {} grupo(s) · {} keynote(s)\n"
                   u"Texto do keynote = DESCRIÇÃO | DESCRIÇÃO IN{}".format(
                       opts["txt_path"], len(XL["kn_headers"]), len(XL["dados_rows"]), n_keys,
                       u"\n\n{} aviso(s) - ver relatório.".format(len(txt_warns))
                       if txt_warns else u""), yes=True, no=True):
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
hl_orange = sum(x["count"] for x in HL_FINAL)


# ------------------------------------------------------------------
# 8. Janela de saída + relatório HTML (opcional nos modos que gravam)
# ------------------------------------------------------------------
output.print_md(u"## {}".format(TOOL_TITLE))
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
    output.print_md(u"- Destaque temporário (vista ativa): **{}** elemento(s) em laranja "
                    u"(sem Keynote, Descrição ou {}){}".format(
                        hl_orange, PARAM_IN_NAME, u" - " + hl_note if hl_note else u""))
inc_written = [r for r in INC if r["status"].startswith(u"Gravado")]
if INC:
    output.print_md(u"- Tabela 2 (tipos sem Keynote / Descrição): **{}** de {} tipo(s) "
                    u"completado(s)".format(len(inc_written), len(INC)))
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

if not forms.alert(
        u"Gerar o relatório HTML?\n\n"
        u"{} preenchido(s) · {} alterado(s) · {} faltando / pendente(s)".format(
            len(RESULT["filled"]), len(RESULT["changed"]), len(RESULT["missing"])),
        yes=True, no=True):
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
    "title": TOOL_TITLE,
    "verify": [dict(r["view"], status=r["status"]) for r in VROWS],
    "incomplete": [{"category": r["category"], "family": r["family"], "type": r["name"],
                    "typeId": r["id"], "count": r["count"], "hasIn": r["hasIn"],
                    "key0": r["key0"], "pt0": r["pt0"], "en0": r["en0"],
                    "key": clean(r["key"]), "pt": clean(r["pt"]), "en": clean(r["en"]),
                    "written": r.get("written", []), "check": r["check"],
                    "status": r["status"]} for r in INC],
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
    forms.alert(u"Modelo do relatório não encontrado (keynotes.html) na pasta do botão:\n"
                u"{}".format(HERE), exitscript=True)
with codecs.open(template_path, "r", encoding="utf-8") as f:
    html = f.read()
html = html.replace("__DATA__", json.dumps(data, ensure_ascii=False,
                                           default=lambda o: to_unicode(o)))
out_path = script.get_document_data_file("keynotes_report", "html")
with codecs.open(out_path, "w", encoding="utf-8") as f:
    f.write(html)

if not open_in_browser(out_path):
    forms.alert(u"Relatório gerado, mas não abriu automaticamente:\n\n{}".format(out_path))
