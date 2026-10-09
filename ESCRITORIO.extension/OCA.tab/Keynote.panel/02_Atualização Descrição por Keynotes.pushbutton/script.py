# -*- coding: utf-8 -*-
# ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE  (pyRevit - Revit 2025/2026+)
#
# Lê a planilha de Keynotes do escritório (OCA_XXX_Lista_de_Materiais - aba
# MATERIAIS) e atualiza as descrições dos TIPOS do modelo pelo código de Keynote.
# Fluxo único: EXCEL > MODELO. Escopo: modelo inteiro. Materiais não são lidos.
# (Comentário, não docstring: o tooltip do botão vem SÓ de __doc__ abaixo.)
#
# Onde cada dado mora:
#   * Keynote ........ parâmetro nativo KEYNOTE dos TIPOS usados no modelo.
#   * Descrição PT ... parâmetro nativo "Description" do tipo.
#   * Descrição IN ... parâmetro compartilhado "Descrição IN" do tipo, associado
#                      pelo GUID 70307f6e-3b31-4049-bed4-62dfc3570fec
#                      (OCA_Parametros_Template.txt). Se faltar em alguma
#                      categoria, oferece vincular ao projeto como parâmetro de TIPO.
#   * Excel .......... aba MATERIAIS: colunas KEYNOTE, DESCRIÇÃO, DESCRIÇÃO IN
#                      (somente leitura - o Excel não é alterado).
#
# Aba 1 - resumo, fluxo (Excel > Modelo), planilha e caminho do TXT de Keynote
#         (sugestão: OCA_XXX_Lista de Materiais_XX.txt, editável).
#
# Aba 2 - janela NÃO modal (dá para navegar no modelo). Três tabelas
# recolhíveis, uma linha por TIPO (família + tipo), colunas ajustáveis, busca
# (botões Buscar / Limpar busca) e "Agrupar por" Sem agrupamento / Keynote /
# Categoria em cada tabela:
#   1. Excel > Modelo: o que o Excel tem e difere do modelo. Valor vazio no
#      Excel nunca apaga o modelo.
#   2. Tipos SEM Keynote ou SEM Descrição: só o Keynote é digitado. A linha
#      fica destacada quando a Keynote existe no Excel (correspondência - não
#      quer dizer gravado). "Verificar no Excel" traz Descrição / Descrição IN
#      da planilha para os campos vazios.
#   3. Tipos com Keynote que não existe no Excel (só consulta).
#   Clicar numa linha seleciona no Revit todas as instâncias do tipo e
#   enquadra na vista ativa (sem transação, a janela continua aberta).
#
# Botões: Voltar (1ª aba) · Gerar TXT de Keynote (do Excel, carrega no projeto,
# pede confirmação) · Dashboard HTML (qualidade dos dados, offline) · Gravar
# marcados (pede confirmação, grava e atualiza as tabelas) · Finalizar (só
# fecha - não grava nem gera nada).
#
# TXT de Keynote do Revit: mesma estrutura do arquivo do escritório -
# CATEGORIA/prefixos, PREFIXO/GRUPO/CATEGORIA e KEYNOTE / "DESCRIÇÃO |
# DESCRIÇÃO IN" / PREFIXO.

__title__ = "Atualização\ndas Descrições\npor Keynote"
__persistentengine__ = True        # janela não modal + ExternalEvent
__doc__ = ("Confere as Keynotes dos tipos do modelo com a planilha de materiais "
           "(aba MATERIAIS) e atualiza Description e Descrição IN pelo código de "
           "Keynote.\n\n"
           "Fluxo: Excel > Modelo. Nada é gravado sem confirmação.\n"
           "Inclui conferência em tabelas, seleção no modelo, TXT de Keynote e "
           "dashboard de qualidade.")

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
  </Window.Resources>
  <StackPanel Margin="18">
    <TextBlock Text="ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE" FontSize="15"
               FontWeight="SemiBold" Foreground="#65E3FF" Margin="0,0,0,4"/>

    <TextBlock Text="Resumo" FontWeight="SemiBold"/>
    <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#9FB3CC" Margin="0,0,0,0"
               Text="Lê a aba MATERIAIS da planilha de Keynote e atualiza a Descrição e a Descrição IN dos tipos do modelo pelo código de Keynote. Na próxima etapa você confere as tabelas e grava só o que marcar (com confirmação). O TXT de Keynote é gerado pelo botão 'Gerar TXT de Keynote'."/>

    <TextBlock Text="Fluxo" FontWeight="SemiBold"/>
    <TextBlock Margin="0,0,0,0" FontSize="13" Foreground="#D9E8F5"
               Text="Excel  &gt;  Modelo"/>

    <TextBlock Text="Planilha Excel de Keynote (.xlsx)" FontWeight="SemiBold"/>
    <DockPanel LastChildFill="True">
      <Button x:Name="browse" DockPanel.Dock="Right" Content="Procurar..." Width="90"
              Height="26" Margin="8,0,0,0"/>
      <TextBox x:Name="xlsx" Height="26" Padding="4,3,4,2"/>
    </DockPanel>

    <TextBlock Text="TXT de Keynote do Revit" FontWeight="SemiBold"/>
    <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#9FB3CC" Margin="0,0,0,6"
               Text="Gerado a partir do Excel e carregado no projeto pelo botão 'Gerar TXT de Keynote' da próxima etapa. Troque XXX (projeto) e XX (revisão) no nome. Se o arquivo já existir, o comando pede confirmação antes de substituir (e guarda um backup)."/>
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


TXT_DEFAULT_NAME = u"OCA_XXX_Lista de Materiais_XX.txt"   # XXX = projeto · XX = revisão


def default_txt_path(xlsx_path):
    """Sugestão: OCA_XXX_Lista de Materiais_XX.txt na pasta da planilha
    (o usuário troca XXX / XX no campo antes de executar)."""
    folder = os.path.dirname(xlsx_path) if xlsx_path else u""
    return os.path.join(folder, TXT_DEFAULT_NAME) if folder else TXT_DEFAULT_NAME


def _old_default_txt(xlsx_path):
    """Sugestão antiga (nome da planilha + .txt): não é mais reaproveitada."""
    return os.path.splitext(xlsx_path)[0] + u".txt" if xlsx_path else u""


class KeynoteWindow(forms.WPFWindow):
    def __init__(self, xaml, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.result = None
        self.xlsx.Text = cfg["xlsx"]
        saved = cfg["txt_path"]
        if not saved or os.path.normcase(saved) == os.path.normcase(_old_default_txt(cfg["xlsx"])):
            saved = default_txt_path(cfg["xlsx"])
        self.txt_path.Text = saved
        self.browse.Click += self._browse
        self.txt_browse.Click += self._txt_browse
        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    def _browse(self, sender, args):
        picked = forms.pick_file(file_ext="xlsx",
                                 title="Selecione a planilha de Keynotes")
        if picked:
            cur = to_unicode(self.txt_path.Text).strip()
            self.xlsx.Text = picked
            if not cur or os.path.basename(cur) == TXT_DEFAULT_NAME:
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
        if not txt.lower().endswith(".txt"):
            txt += u".txt"
        if not os.path.isdir(os.path.dirname(txt) or u"."):
            forms.alert(u"Pasta do TXT de Keynote não existe:\n{}".format(txt))
            return
        self.result = {
            "xlsx": path,
            "mode": "to_model",            # fluxo único
            "scope": "model",              # sempre o modelo inteiro
            "materials": False,            # materiais não são lidos
            "txt_on": True,                # TXT gerado pelo botão 'Gerar TXT de Keynote'
            "txt_path": txt,
        }
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.confirmed = False
        self.Close()


config = script.get_config()


def ask_options(cfg=None):
    """Mostra a 1ª aba (modal). -> opções escolhidas, ou None se cancelou.
    cfg: valores iniciais; sem cfg usa os da última execução."""
    if cfg is None:
        cfg = {"xlsx": to_unicode(config.get_option("xlsx_path", u"")),
               "txt_path": to_unicode(config.get_option("txt_path", u""))}
    w = KeynoteWindow(KN_XAML, cfg)
    w.ShowDialog()
    if not w.confirmed:
        return None
    config.xlsx_path = w.result["xlsx"]
    config.txt_path = w.result["txt_path"]
    script.save_config()
    return w.result


opts = None                        # opções da 1ª aba (definidas em start_session)
XLSX_PATH = u""
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


XL = None                          # planilha lida (start_session)


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


MODEL = OrderedDict()               # lido em start_session
FILE_TEXT = {}


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
# 5d. Situação dos TIPOS no modelo (sem keynote / descrição / descrição IN)
#   O comando NÃO colore mais o modelo: a conferência é feita selecionando a
#   linha da tabela, que seleciona e mostra os elementos na vista ativa.
#   As cores deixadas pelas versões anteriores são localizadas e removidas
#   (seção 9).
# ------------------------------------------------------------------
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


# ---- cores deixadas no modelo pelas versões anteriores deste comando ----
# Cores usadas antes (laranja, vermelho, lilás, rosa claro, roxo, rosa, azul),
# puras e no tom claro (45% de branco). Os corretos ficavam em meio-tom + 70%.
_OLD_RGB = ((255, 140, 0), (232, 74, 74), (200, 162, 200), (255, 182, 193),
            (142, 68, 173), (255, 105, 180), (52, 152, 219))
OLD_COLORS = set(_OLD_RGB) | set(
    tuple(int(round(c + (255 - c) * 0.45)) for c in rgb) for rgb in _OLD_RGB)
_OGS_COLORS = ("ProjectionLineColor", "CutLineColor", "SurfaceForegroundPatternColor",
               "CutForegroundPatternColor", "SurfaceBackgroundPatternColor",
               "CutBackgroundPatternColor", "ProjectionFillColor", "CutFillColor")


def _rgb(c):
    try:
        if c is not None and c.IsValid:
            return (int(c.Red), int(c.Green), int(c.Blue))
    except Exception:
        pass
    return None


def is_old_highlight(o):
    """Sobreposição igual à que o comando aplicava (2+ campos com a mesma cor
    do comando, ou meio-tom + 70% de transparência)."""
    try:
        if o.Halftone and int(o.Transparency) == 70:
            return True
    except Exception:
        pass
    hits = defaultdict(int)
    for name in _OGS_COLORS:
        try:
            rgb = _rgb(getattr(o, name, None))
        except Exception:
            rgb = None
        if rgb in OLD_COLORS:
            hits[rgb] += 1
    return any(n >= 2 for n in hits.values())


def _colors_log_path(name="keynotes_cores"):
    try:
        return script.get_document_data_file(name, "json")
    except Exception:
        return None


def graphic_views():
    out = []
    try:
        for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
            if view_allows_overrides(v):
                out.append(v)
    except Exception:
        pass
    return out


def find_old_highlight(views, progress=None):
    """-> [(vista, [ElementId com a cor antiga])]"""
    found = []
    for n, v in enumerate(views):
        if progress is not None:
            try:
                progress.update_progress(n + 1, len(views))
            except Exception:
                pass
        ids = []
        try:
            for e in DB.FilteredElementCollector(doc, v.Id).WhereElementIsNotElementType():
                cat = e.Category
                if cat is None or not is_model_category(cat):
                    continue
                try:
                    if is_old_highlight(v.GetElementOverrides(e.Id)):
                        ids.append(e.Id)
                except Exception:
                    continue
        except Exception:
            continue
        if ids:
            found.append((v, ids))
    return found


def clear_old_highlight(found):
    """Volta a sobreposição desses elementos ao padrão. -> (limpos, falhas)"""
    ok, bad = 0, 0
    t = DB.Transaction(doc, u"Remover cores antigas - Keynotes")
    t.Start()
    try:
        for v, ids in found:
            for eid in ids:
                try:
                    v.SetElementOverrides(eid, DB.OverrideGraphicSettings())
                    ok += 1
                except Exception:
                    bad += 1
        t.Commit()
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        return 0, ok + bad
    return ok, bad


# ---- ações no Revit com a janela aberta (janela não modal) ----
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from System.Collections.Generic import List


class _RevitAction(IExternalEventHandler):
    """Fila de funções executadas no contexto da API do Revit."""

    def __init__(self):
        self.queue = []

    def Execute(self, uiapp):
        while self.queue:
            fn = self.queue.pop(0)
            try:
                fn()
            except SystemExit:
                pass
            except Exception as ex:
                forms.alert(u"Erro: {}".format(to_unicode(ex)))

    def GetName(self):
        return u"Atualização das Descrições por Keynote"


REVIT_ACTION = _RevitAction()
REVIT_EVENT = ExternalEvent.Create(REVIT_ACTION)


def run_in_revit(fn):
    REVIT_ACTION.queue.append(fn)
    REVIT_EVENT.Raise()


def show_in_view(ids):
    """Seleciona os elementos e aproxima na vista ativa. -> mensagem"""
    if not ids:
        return u""
    uidoc = revit.uidoc
    view = doc.ActiveView

    def select():
        try:
            uidoc.Selection.SetElementIds(List[DB.ElementId](ids))
        except Exception:
            pass
    select()
    try:
        visible = set(eid_int(i) for i in DB.FilteredElementCollector(doc, view.Id)
                      .WhereElementIsNotElementType().ToElementIds())
    except Exception:
        visible = set()
    vis = [i for i in ids if eid_int(i) in visible]
    if not vis:
        return (u"{} elemento(s) selecionado(s), mas nenhum está visível na vista ativa "
                u"({}). Abra uma vista 3D e clique em 'Mostrar na vista ativa'.".format(
                    len(ids), to_unicode(view.Name)))
    try:
        uidoc.ShowElements(List[DB.ElementId](vis))
    except Exception:
        pass
    select()                              # continua selecionado depois do zoom
    try:
        uidoc.RefreshActiveView()
    except Exception:
        pass
    return u"{} de {} elemento(s) selecionado(s) e mostrado(s) na vista '{}'.".format(
        len(vis), len(ids), to_unicode(view.Name))


# ------------------------------------------------------------------
# 5e. Tabela 1 - Excel -> Modelo
#     Uma linha por Keynote + Categoria + Família + Tipo: tipos diferentes com
#     o mesmo Keynote NÃO são juntados (as instâncias do tipo ficam na Qtd).
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
        gkey = (h["key"], h["category"], h["family"], h["name"], h["id"])
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
def collect_incomplete(t1_keys):
    """-> (tabela 2, tabela 3)
    Tabela 2: tipos sem Keynote, ou sem Descrição cujo Keynote está no Excel
              mas NÃO aparece na tabela 1 (o que está na tabela 1 já será
              atualizado por ela).
    Tabela 3: tipos com Keynote preenchido que NÃO existe no Excel (só
              informativo - o código precisa ser incluído na planilha)."""
    inc, notx = [], []
    for info in scan_types():
        t = info["t"]
        key, desc, en, has_in, missing = type_status(t)
        row = {"t": t, "id": eid_int(t.Id), "category": info["category"],
               "family": family_name(t), "name": elem_name(t),
               "count": len(info["ids"]), "hasIn": has_in,
               "key0": key, "pt0": desc, "en0": en,
               "key": key, "pt": desc, "en": en,
               "sel": False, "check": u"", "status": u""}
        if key and key not in XL["by_key"]:
            row["check"] = u"Keynote {} não está no Excel".format(key)
            notx.append(row)
            continue
        if key and desc:
            continue
        if key and key in t1_keys:
            continue                         # a tabela 1 já atualiza este keynote
        inc.append(row)
    for lst in (inc, notx):
        lst.sort(key=lambda r: (r["category"], r["family"], r["name"]))
    return inc, notx


def inc_changed(r):
    """Algum campo digitado difere do valor atual? (vazio nunca apaga)"""
    return any(clean(r[f]) and clean(r[f]) != r[f + "0"] for f in ("key", "pt", "en"))


def excel_rec(key):
    """Linha do Excel (aba MATERIAIS) do keynote, ou None (vazio / inexistente)."""
    key = clean(key)
    return XL["by_key"].get(key) if key else None


def verify_incomplete(r, autofill):
    """Confere a linha com o Excel. autofill: completa Descrição / Descrição IN
    vazias com os valores do Excel quando o Keynote existe na planilha.
    Descrição do modelo DIFERENTE da do Excel não é substituída aqui: depois de
    gravar o keynote, a diferença aparece na tabela 1 (Excel -> Modelo)."""
    for f in ("key", "pt", "en"):
        r[f] = clean(r[f])
    rec = excel_rec(r["key"])
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
        errs.append(u"falta Descrição" + (u" (vazia também no Excel)"
                                          if rec is not None and not rec["pt"] else u""))
    elif rec is not None and rec["pt"] and r["pt"] != rec["pt"]:
        warns.append(u"Descrição atual difere do Excel (mantida; após gravar o keynote "
                     u"a troca aparece na tabela 1)")
    if not r["hasIn"]:
        warns.append(u"sem parâmetro {}".format(PARAM_IN_NAME))
    elif not r["en"]:
        warns.append(u"{} vazia".format(PARAM_IN_NAME))
    elif rec is not None and rec["en"] and r["en"] != rec["en"]:
        warns.append(u"{} difere do Excel".format(PARAM_IN_NAME))
    parts = []
    if rec is not None:
        parts.append(u"Keynote encontrada no Excel (linha {})".format(rec["row"]))
        pending = [lbl for f, lbl, ok in (("pt", u"Descrição", True),
                                          ("en", PARAM_IN_NAME, r["hasIn"]))
                   if ok and not r[f] and rec[f]]
        if pending and not autofill:
            parts.append(u"clique em 'Verificar no Excel' para usar: " + u", ".join(pending))
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

# tabela com altura do conteúdo: quem rola é a janela (barra à direita);
# sem virtualização, para o agrupamento sempre redesenhar todas as linhas
_GRID_STYLE = u"""
              AutoGenerateColumns="False" CanUserAddRows="False"
              CanUserDeleteRows="False" CanUserSortColumns="False" CanUserReorderColumns="False"
              CanUserResizeColumns="True"
              CanUserResizeRows="False" HeadersVisibility="Column" SelectionMode="Single"
              SelectionUnit="FullRow" GridLinesVisibility="Horizontal"
              HorizontalGridLinesBrush="#1C2B44" Background="#0A1120" AlternationCount="2"
              Foreground="#D9E8F5" BorderBrush="#2A4A66"
              BorderThickness="1" FontFamily="Segoe UI" FontSize="12" Margin="0,4,0,0"
              VerticalScrollBarVisibility="Disabled" HorizontalScrollBarVisibility="Auto"
              EnableRowVirtualization="False" EnableColumnVirtualization="False"
              VirtualizingPanel.IsVirtualizing="False"
              VirtualizingPanel.IsVirtualizingWhenGrouping="False"
              ScrollViewer.CanContentScroll="False"
              ColumnHeaderStyle="{StaticResource head}" RowStyle="{StaticResource row}" """

# cor da CÉLULA que vai mudar (coluna "<campo>Mark" da linha):
#   fill = preenche vazio · over = sobrescreve · edit = digitado · blocked = não grava
#   excel = vem do Excel e será gravado · diff = Excel difere do modelo (não grava)
MARK_COLORS = (("fill", "#1E5A3A", "#E8FFF0"), ("over", "#6B4A12", "#FFF3DD"),
               ("edit", "#1F4E8C", "#EAF3FF"), ("excel", "#1E5A3A", "#E8FFF0"),
               ("diff", "#6B4A12", "#FFF3DD"))


def _mark_style(key, mark_col):
    trig = u"".join(
        u'<DataTrigger Binding="{{Binding {0}}}" Value="{1}">'
        u'<Setter Property="Background" Value="{2}"/><Setter Property="Foreground" Value="{3}"/>'
        u'<Setter Property="FontWeight" Value="SemiBold"/></DataTrigger>'.format(mark_col, v, bg, fg)
        for v, bg, fg in MARK_COLORS)
    trig += (u'<DataTrigger Binding="{{Binding {0}}}" Value="blocked">'
             u'<Setter Property="Foreground" Value="#7A8FA9"/></DataTrigger>'.format(mark_col))
    return (u'<Style x:Key="{0}" TargetType="DataGridCell" '
            u'BasedOn="{{StaticResource {{x:Type DataGridCell}}}}">'
            u'<Style.Triggers>{1}</Style.Triggers></Style>'.format(key, trig))


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


def _col(header, binding, width, style=u"wrap", readonly=True, cell=None):
    """Colunas de texto >= 100 começam proporcionais (ocupam a largura da janela);
    as estreitas (Qtd, Keynote) começam com largura fixa. O usuário pode arrastar
    a divisa do cabeçalho para alargar ou estreitar qualquer coluna."""
    size = (u'Width="{}*" MinWidth="{}"'.format(width, int(width * 0.55)) if width >= 100
            else u'Width="{}"'.format(width))
    return (u'<DataGridTextColumn Header="{}" Binding="{{Binding {}}}" {} '
            u'IsReadOnly="{}" ElementStyle="{{StaticResource {}}}"{}/>'.format(
                header, binding, size, u"True" if readonly else u"False", style,
                u' CellStyle="{{StaticResource {}}}"'.format(cell) if cell else u""))


VERIFY_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Atualização das Descrições do Modelo por Keynote · Verificação"
        Width="1450" Height="980" MinWidth="900" MinHeight="620"
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
      <Setter Property="Background" Value="#1E3150"/>
      <Setter Property="FontFamily" Value="Consolas"/>
    </Style>
    <Style x:Key="lbl" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#65E3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,16,0"/>
    </Style>
    <Style x:Key="small" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
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
      <Setter Property="Background" Value="#0E1526"/>
      <Style.Triggers>
        <Trigger Property="AlternationIndex" Value="1">
          <Setter Property="Background" Value="#111C31"/>
        </Trigger>
        <DataTrigger Binding="{Binding Kind}" Value="blocked">
          <Setter Property="Foreground" Value="#7A8FA9"/>
        </DataTrigger>
        <!-- tabela 2: keynote encontrada no Excel (correspondência, não "gravado") -->
        <DataTrigger Binding="{Binding Kind}" Value="match">
          <Setter Property="Background" Value="#123B47"/>
          <Setter Property="BorderBrush" Value="#3FB6C9"/>
          <Setter Property="BorderThickness" Value="3,0,0,0"/>
        </DataTrigger>
        <!-- linha de título do grupo (Agrupar por), na largura toda da tabela -->
        <DataTrigger Binding="{Binding Kind}" Value="group">
          <Setter Property="Focusable" Value="False"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="DataGridRow">
                <Border Background="#1B2A47" BorderBrush="#2A4A66" BorderThickness="0,1,0,1"
                        Padding="8,5,8,5">
                  <TextBlock Text="{Binding Grp}" Foreground="#65E3FF" FontWeight="Bold"
                             FontFamily="Segoe UI"/>
                </Border>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </DataTrigger>
      </Style.Triggers>
    </Style>
    """ + _mark_style(u"mPt", u"PtMark") + u"""
    """ + _mark_style(u"mEn", u"EnMark") + u"""
    """ + _mark_style(u"mKey", u"KeyMark") + u"""
  </Window.Resources>
  <DockPanel Margin="16,12,4,12">
    <StackPanel DockPanel.Dock="Top" Margin="0,0,12,0">
      <TextBlock Text="ATUALIZAÇÃO DAS DESCRIÇÕES DO MODELO POR KEYNOTE   ·   Excel &gt; Modelo"
                 FontSize="15" FontWeight="SemiBold" Foreground="#65E3FF" FontFamily="Segoe UI"/>
      <TextBlock x:Name="info" TextWrapping="Wrap" FontSize="12" Foreground="#CFE3FF"
                 FontFamily="Segoe UI" Margin="0,6,0,0"/>
      <WrapPanel Margin="0,10,0,0">
        <TextBlock Text="Buscar " Style="{StaticResource small}"/>
        <TextBox x:Name="f_text" Width="300" Height="26" Margin="0,0,8,0" Padding="4,3,4,2"/>
        <Button x:Name="b_search" Content="Buscar" Height="26"/>
        <Button x:Name="b_clear" Content="Limpar busca" Height="26"/>
        <TextBlock x:Name="searchinfo" Style="{StaticResource small}" Foreground="#7A8FA9"
                   Margin="8,0,0,0"/>
      </WrapPanel>
      <TextBlock x:Name="viewinfo" TextWrapping="Wrap" FontSize="12" Foreground="#7BE3A0"
                 FontFamily="Segoe UI" Margin="0,4,0,0"/>
    </StackPanel>
    <Grid DockPanel.Dock="Bottom" Margin="0,10,12,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Button x:Name="b_back" Grid.Column="0" Content="←  Voltar" MinWidth="100"
              Margin="0,0,14,0" ToolTip="Volta para a primeira aba (planilha e TXT)"/>
      <TextBlock x:Name="counter" Grid.Column="1" VerticalAlignment="Center" Margin="0,0,12,0"
                 Foreground="#CFE3FF" FontFamily="Segoe UI" TextWrapping="Wrap"/>
      <Button x:Name="b_txt" Grid.Column="2" Content="Gerar TXT de Keynote" MinWidth="150"
              Margin="0,0,10,0"
              ToolTip="Gera o TXT de Keynote a partir do Excel e carrega no projeto (pede confirmação)"/>
      <Button x:Name="b_dash" Grid.Column="3" Content="Dashboard HTML" MinWidth="130"
              Margin="0,0,10,0"
              ToolTip="Abre no navegador o painel de qualidade dos dados da última análise (não altera o modelo)"/>
      <Button x:Name="b_apply" Grid.Column="4" Content="Gravar marcados" MinWidth="150"
              Margin="0,0,10,0" ToolTip="Grava no modelo as linhas marcadas (pede confirmação)"/>
      <Button x:Name="b_finish" Grid.Column="5" Content="Finalizar" MinWidth="110"
              Margin="0" FontWeight="SemiBold" ToolTip="Fecha o comando sem gravar nada"/>
    </Grid>
    <ScrollViewer x:Name="scroll" VerticalScrollBarVisibility="Visible"
                  HorizontalScrollBarVisibility="Disabled" Margin="0,8,0,0">
      <StackPanel Margin="0,0,12,0">
        <Border Background="#13203A" BorderBrush="#2A4A66" BorderThickness="1"
                Padding="6,4,8,4" Margin="0,4,0,0">
          <DockPanel LastChildFill="False">
            <Button x:Name="tg1" Content="▼" Width="28" Height="24" Padding="0" Margin="0,0,8,0"
                    ToolTip="Recolher / expandir a tabela"/>
            <TextBlock x:Name="tt1" Style="{StaticResource lbl}" Margin="0,0,8,0" Cursor="Hand"
                       Text="1 · EXCEL &gt; MODELO"/>
            <TextBlock x:Name="h1" Style="{StaticResource small}" Margin="0,0,18,0"
                       Foreground="#7A8FA9"/>
            <TextBlock Text="Agrupar por " Style="{StaticResource small}"/>
            <ComboBox x:Name="f_mode1" Width="150" Height="26" Margin="0,0,14,0"/>
          </DockPanel>
        </Border>
        <StackPanel x:Name="body1" Margin="0,6,0,0">
            <TextBlock Style="{StaticResource small}" FontSize="11" Foreground="#7A8FA9"
                       TextWrapping="Wrap"
                       Text="Uma linha por tipo (família + tipo): tipos diferentes com a mesma keynote ficam em linhas separadas. Clique numa linha para selecionar o tipo no Revit.   Célula verde = preenche vazio  ·  Âmbar = substitui o valor atual  ·  Cinza = não pode ser gravado"/>
            <DataGrid x:Name="grid" """ + _GRID_STYLE + u""">
              <DataGrid.Columns>""" + _CHECK_COL + u"""
                """ + _col(u"Agrupamento", u"Grupo", 150) + u"""
                """ + _col(u"Categoria", u"Categoria", 110) + u"""
                """ + _col(u"Família", u"Familia", 130) + u"""
                """ + _col(u"Tipo", u"Tipo", 170) + u"""
                """ + _col(u"Keynote", u"Keynote", 75, u"mono") + u"""
                """ + _col(u"Descrição atual", u"PtAtual", 180) + u"""
                """ + _col(u"Descrição nova", u"PtNova", 180, cell=u"mPt") + u"""
                """ + _col(u"Descrição em IN atual", u"EnAtual", 160) + u"""
                """ + _col(u"Descrição em IN nova", u"EnNova", 160, cell=u"mEn") + u"""
                """ + _col(u"Destino", u"Destino", 130) + u"""
              </DataGrid.Columns>
            </DataGrid>
        </StackPanel>
        <Border Background="#13203A" BorderBrush="#2A4A66" BorderThickness="1"
                Padding="6,4,8,4" Margin="0,16,0,0">
          <DockPanel LastChildFill="False">
            <Button x:Name="tg2" Content="▼" Width="28" Height="24" Padding="0" Margin="0,0,8,0"
                    ToolTip="Recolher / expandir a tabela"/>
            <TextBlock x:Name="tt2" Style="{StaticResource lbl}" Margin="0,0,8,0" Cursor="Hand"
                       Text="2 · ELEMENTOS SEM KEYNOTE OU SEM DESCRIÇÃO"/>
            <TextBlock x:Name="h2" Style="{StaticResource small}" Margin="0,0,18,0"
                       Foreground="#7A8FA9"/>
            <TextBlock Text="Agrupar por " Style="{StaticResource small}"/>
            <ComboBox x:Name="f_mode2" Width="150" Height="26" Margin="0,0,14,0"/>
            <Button x:Name="b2_check" Content="Verificar no Excel" Height="26"/>
          </DockPanel>
        </Border>
        <StackPanel x:Name="body2" Margin="0,6,0,0">
            <TextBlock Style="{StaticResource small}" FontSize="11" Foreground="#7A8FA9"
                       TextWrapping="Wrap"
                       Text="Clique duas vezes no Keynote para digitar. Linha azul-esverdeada = keynote encontrada no Excel (só indica correspondência; nada é gravado até 'Gravar marcados'). 'Verificar no Excel' traz a Descrição e a Descrição em IN da planilha para os campos vazios.   Célula azul = keynote digitado  ·  Verde = vem do Excel e será gravado  ·  Âmbar = Excel difere do modelo (mantido)"/>
            <DataGrid x:Name="grid2" """ + _GRID_STYLE + u""">
              <DataGrid.Columns>""" + _CHECK_COL + u"""
                """ + _col(u"Categoria", u"Categoria", 105) + u"""
                """ + _col(u"Família", u"Familia", 125) + u"""
                """ + _col(u"Tipo", u"Tipo", 160) + u"""
                """ + _col(u"Qtd", u"Qtd", 45, u"mono") + u"""
                """ + _col(u"Keynote ✎", u"Keynote", 95, u"edit", False, u"mKey") + u"""
                """ + _col(u"Descrição atual", u"PtAtual", 160) + u"""
                """ + _col(u"Descrição Excel", u"PtExcel", 170, cell=u"mPt") + u"""
                """ + _col(u"Descrição em IN atual", u"EnAtual", 140) + u"""
                """ + _col(u"Descrição em IN Excel", u"EnExcel", 150, cell=u"mEn") + u"""
                """ + _col(u"Verificação", u"Verificacao", 230) + u"""
              </DataGrid.Columns>
            </DataGrid>
        </StackPanel>
        <Border Background="#13203A" BorderBrush="#2A4A66" BorderThickness="1"
                Padding="6,4,8,4" Margin="0,16,0,0">
          <DockPanel LastChildFill="False">
            <Button x:Name="tg3" Content="▼" Width="28" Height="24" Padding="0" Margin="0,0,8,0"
                    ToolTip="Recolher / expandir a tabela"/>
            <TextBlock x:Name="tt3" Style="{StaticResource lbl}" Margin="0,0,8,0" Cursor="Hand"
                       Text="3 · ELEMENTOS COM KEYNOTE QUE NÃO ESTÁ NO EXCEL"/>
            <TextBlock x:Name="h3" Style="{StaticResource small}" Margin="0,0,18,0"
                       Foreground="#7A8FA9"/>
            <TextBlock Text="Agrupar por " Style="{StaticResource small}"/>
            <ComboBox x:Name="f_mode3" Width="150" Height="26" Margin="0,0,14,0"/>
          </DockPanel>
        </Border>
        <StackPanel x:Name="body3" Margin="0,6,0,8">
            <TextBlock Style="{StaticResource small}" FontSize="11" Foreground="#7A8FA9"
                       TextWrapping="Wrap"
                       Text="Só consulta: inclua o código na planilha do Excel e rode o comando de novo."/>
            <DataGrid x:Name="grid3" """ + _GRID_STYLE + u""" IsReadOnly="True">
              <DataGrid.Columns>
                """ + _col(u"Categoria", u"Categoria", 120) + u"""
                """ + _col(u"Família", u"Familia", 150) + u"""
                """ + _col(u"Tipo", u"Tipo", 220) + u"""
                """ + _col(u"Qtd", u"Qtd", 45, u"mono") + u"""
                """ + _col(u"Keynote", u"Keynote", 95, u"mono") + u"""
                """ + _col(u"Descrição atual", u"Descricao", 240) + u"""
                """ + _col(u"Descrição em IN atual", u"DescricaoIN", 220) + u"""
                """ + _col(u"Situação", u"Verificacao", 240) + u"""
              </DataGrid.Columns>
            </DataGrid>
        </StackPanel>
      </StackPanel>
    </ScrollViewer>
  </DockPanel>
</Window>
"""

GROUP_MODES = [u"Sem agrupamento", u"Keynote", u"Categoria"]
COLS1 = ("Grupo", "Categoria", "Familia", "Tipo", "Keynote", "PtAtual", "PtNova",
         "EnAtual", "EnNova", "Destino")
COLS2 = ("Categoria", "Familia", "Tipo", "Qtd", "Keynote", "PtAtual", "PtExcel",
         "EnAtual", "EnExcel", "Verificacao")
COLS3 = ("Categoria", "Familia", "Tipo", "Qtd", "Keynote", "Descricao", "DescricaoIN",
         "Verificacao")
MARKS = ("PtMark", "EnMark", "KeyMark")
NO_KEY = u"(sem keynote)"
TYPE_INST = {}        # id do tipo -> [ElementId das instâncias] (modelo inteiro)


def _fold(text):
    """Texto para busca: minúsculo e sem acento ('Descrição' acha 'descricao')."""
    return norm_header(text).lower()


def _key_label(key):
    if not key:
        return NO_KEY
    cat, grp = agrup_of(key)
    if cat == NO_GROUP:                   # prefixo sem agrupamento na aba DADOS
        return key
    return u"{}   ·   {} ({})".format(key, grp, cat)


def _ask(owner, text, title=u"Atualização das Descrições por Keynote"):
    """Pergunta Sim/Não (janela WPF; pode ser chamada com a tabela aberta)."""
    from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, MessageBoxResult
    try:
        res = MessageBox.Show(owner, text, title, MessageBoxButton.YesNo,
                              MessageBoxImage.Question)
    except Exception:
        res = MessageBox.Show(text, title, MessageBoxButton.YesNo, MessageBoxImage.Question)
    return res == MessageBoxResult.Yes


def _tell(owner, text, title=u"Atualização das Descrições por Keynote", warn=False):
    from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
    img = MessageBoxImage.Warning if warn else MessageBoxImage.Information
    try:
        MessageBox.Show(owner, text, title, MessageBoxButton.OK, img)
    except Exception:
        MessageBox.Show(text, title, MessageBoxButton.OK, img)


class VerifyWindow(forms.WPFWindow):
    """Janela NÃO modal: o Revit continua utilizável. Tudo que usa a API do
    Revit (selecionar/mostrar, gravar, carregar o TXT) passa por run_in_revit
    (ExternalEvent). Busca e agrupamento são feitos aqui, em Python: cada
    tabela é remontada com as linhas que atendem à busca e, se houver
    agrupamento, com uma linha de título por grupo (Kind = "group")."""

    def __init__(self, rows, inc, notx, info_text):
        forms.WPFWindow.__init__(self, VERIFY_XAML, literal_string=True)
        self.closing_for_back = False
        self._busy = False
        self.rows, self.inc, self.notx = rows, inc, notx
        self.by_rid = dict((r["rid"], r) for r in rows)
        self.by_rid2 = dict((r["rid"], r) for r in inc)
        self.by_rid3 = dict((r["rid"], r) for r in notx)
        self.info.Text = info_text
        self.viewinfo.Text = u""
        self.searchinfo.Text = u""
        self._query = u""                 # termo aplicado pelo botão Buscar
        self._last_ids = []
        self.t1 = self.t2 = self.t3 = None
        self.d1, self.d2, self.d3 = {}, {}, {}
        self._vis = {1: [], 2: [], 3: []}  # RowId das linhas visíveis (sem títulos)
        for combo, idx in ((self.f_mode1, 1), (self.f_mode2, 2), (self.f_mode3, 1)):
            combo.ItemsSource = GROUP_MODES
            combo.SelectedIndex = idx      # tab. 1 e 3: Keynote · tab. 2: Categoria
        self._rebuild()

        # eventos (ligados UMA vez; as tabelas novas ligam só o ColumnChanged)
        self.f_mode1.SelectionChanged += self._on_mode
        self.f_mode2.SelectionChanged += self._on_mode
        self.f_mode3.SelectionChanged += self._on_mode
        self.b_search.Click += self._search
        self.b_clear.Click += self._clear_search
        self.f_text.KeyDown += self._on_key
        self.grid.SelectionChanged += self._on_select1
        self.grid2.SelectionChanged += self._on_select2
        self.grid3.SelectionChanged += self._on_select3
        # recolher / expandir: botão ▼ ou clique no título da tabela
        for n in (1, 2, 3):
            getattr(self, "tg%d" % n).Click += (lambda k: lambda s, a: self._toggle(k))(n)
            getattr(self, "tt%d" % n).MouseLeftButtonUp += (lambda k: lambda s, a: self._toggle(k))(n)
        # a roda do mouse sobre uma tabela rola a janela (as tabelas não rolam)
        for g in (self.grid, self.grid2, self.grid3):
            g.PreviewMouseWheel += self._wheel
        self.b2_check.Click += self._verify_all
        self.b_back.Click += self._back
        self.b_txt.Click += self._txt
        self.b_dash.Click += self._dashboard
        self.b_apply.Click += self._apply
        self.b_finish.Click += self._finish

    # ---- montagem das tabelas ----
    @staticmethod
    def _new_table(cols):
        tb = DataTable("t")
        for name, typ in (("RowId", System.Int32), ("Sel", System.Boolean),
                          ("Enabled", System.Boolean), ("Kind", System.String),
                          ("Grp", System.String)) + tuple((c, System.String)
                                                         for c in cols + MARKS):
            tb.Columns.Add(name, clr.GetClrType(typ))
        return tb

    @staticmethod
    def _mode(combo):
        return GROUP_MODES[min(max(combo.SelectedIndex, 0), len(GROUP_MODES) - 1)]

    @staticmethod
    def _grp_of(mode, key, category):
        if mode == u"Keynote":
            return _key_label(clean(key))
        if mode == u"Categoria":
            return category or u"(sem categoria)"
        return u""

    def _matches(self, vals, grp):
        if not self._query:
            return True
        hay = u" ".join(to_unicode(v) for k, v in vals.items() if not k.startswith(u"_"))
        return self._query in _fold(hay + u" " + grp)

    def _fill(self, n, tb, store, items, mode, key_of, values_of, sort_key):
        """Monta a tabela n: filtra pela busca, ordena e, se houver
        agrupamento, insere a linha de título de cada grupo. Cada registro
        aparece uma única vez."""
        store.clear()
        entries = []
        for r in items:
            g = self._grp_of(mode, key_of(r), r["category"])
            vals = values_of(r)
            if self._matches(vals, g):
                entries.append((g, r, vals))
        if mode == GROUP_MODES[0]:
            entries.sort(key=lambda e: sort_key(e[1]))
        else:
            entries.sort(key=lambda e: (e[0] == NO_KEY, e[0]) + sort_key(e[1]))
        counts = defaultdict(int)
        for g, _r, _v in entries:
            counts[g] += 1
        last, gi = None, 0
        vis = []
        for g, r, vals in entries:
            if mode != GROUP_MODES[0] and g != last:
                last, gi = g, gi + 1
                hrow = tb.NewRow()
                hrow["RowId"], hrow["Kind"] = -gi, u"group"
                hrow["Sel"], hrow["Enabled"] = False, False
                hrow["Grp"] = u"{}      ·  {} linha(s)".format(g, counts[g])
                tb.Rows.Add(hrow)
            row = tb.NewRow()
            row["RowId"], row["Grp"] = r["rid"], g
            vals = dict(vals)
            row["Kind"] = vals.pop("_kind")
            row["Sel"], row["Enabled"] = bool(r["sel"]), bool(r.get("enabled", True))
            for c, v in vals.items():
                row[c] = v
            tb.Rows.Add(row)
            store[r["rid"]] = row
            vis.append(r["rid"])
        self._vis[n] = vis

    @staticmethod
    def _mark1(r, f):
        if r[f + "_new"] is not None:
            return u"over" if f in r["over"] else u"fill"
        return u"blocked" if f in r["blocked"] else u""

    def _vals1(self, r):
        v = r["view"]
        d = dict(zip(COLS1, (v["grp"], v["category"], v["family"], v["type"], v["key"],
                             v["ptCur"], v["ptNew"], v["enCur"], v["enNew"], v["dest"])),
                 _kind=v["kind"])
        d.update(PtMark=self._mark1(r, "pt"), EnMark=self._mark1(r, "en"), KeyMark=u"")
        return d

    @staticmethod
    def _cmp2(r, f, rec, has_param=True):
        """-> (texto da coluna 'Excel', marca da célula) para o campo f."""
        if not has_param:
            return u"(sem parâmetro no tipo)", u""
        if rec is None:
            return (u"(keynote não está no Excel)" if clean(r["key"]) else u""), u""
        xl = rec[f]
        if r[f] and r[f] != r[f + "0"]:
            return r[f], u"excel"              # vem do Excel e será gravado
        if not xl:
            return u"(vazio no Excel)", u""
        if r[f + "0"] and xl != r[f + "0"]:
            return xl, u"diff"                 # Excel difere do modelo (mantido)
        return xl, u""

    def _vals2(self, r):
        rec = excel_rec(r["key"])
        pt_x, pt_m = self._cmp2(r, "pt", rec)
        en_x, en_m = self._cmp2(r, "en", rec, r["hasIn"])
        d = dict(zip(COLS2, (r["category"], r["family"], r["name"], to_unicode(r["count"]),
                             r["key"], r["pt0"] or u"(vazio)", pt_x,
                             (r["en0"] or u"(vazio)") if r["hasIn"] else u"(sem parâmetro)",
                             en_x, r["check"])),
                 _kind=u"match" if rec is not None else
                 (u"changed" if inc_changed(r) else u"inc"))
        d.update(KeyMark=u"edit" if r["key"] and r["key"] != r["key0"] else u"",
                 PtMark=pt_m, EnMark=en_m)
        return d

    def _vals3(self, r):
        d = dict(zip(COLS3, (r["category"], r["family"], r["name"], to_unicode(r["count"]),
                             r["key"], r["pt"] or u"(vazio)",
                             (r["en"] or u"(vazio)") if r["hasIn"] else u"(sem parâmetro)",
                             r["check"])), _kind=u"info")
        d.update(PtMark=u"", EnMark=u"", KeyMark=u"")
        return d

    @staticmethod
    def _bind(grid, tb):
        """Troca a fonte da grade por uma tabela NOVA (nunca apaga e refaz as
        linhas da tabela ligada, o que deixava a grade vazia)."""
        grid.ItemsSource = None
        try:
            grid.Items.GroupDescriptions.Clear()
        except Exception:
            pass
        grid.ItemsSource = tb.DefaultView

    def _rebuild(self):
        """Remonta as 3 tabelas com o agrupamento e a busca atuais."""
        self._commit()
        keep = dict((n, self._selected_rid(g)) for n, g in
                    ((1, self.grid), (2, self.grid2), (3, self.grid3)))
        self._busy = True
        try:
            self.t1 = self._new_table(COLS1)
            self.t2 = self._new_table(COLS2)
            self.t3 = self._new_table(COLS3)
            self._fill(1, self.t1, self.d1, self.rows, self._mode(self.f_mode1),
                       lambda r: r["key"], self._vals1,
                       lambda r: (r["key"], r["view"]["category"], r["view"]["family"],
                                  r["view"]["type"]))
            self._fill(2, self.t2, self.d2, self.inc, self._mode(self.f_mode2),
                       lambda r: r["key"], self._vals2,
                       lambda r: (r["category"], r["family"], r["name"]))
            self._fill(3, self.t3, self.d3, self.notx, self._mode(self.f_mode3),
                       lambda r: r["key"], self._vals3,
                       lambda r: (r["key"], r["category"], r["family"], r["name"]))
            for n, grid, tb in ((1, self.grid, self.t1), (2, self.grid2, self.t2),
                                (3, self.grid3, self.t3)):
                self._bind(grid, tb)
                self._reselect(grid, keep[n])
            self.t1.ColumnChanged += self._on_changed1
            self.t2.ColumnChanged += self._on_changed2
        except Exception as ex:
            self.viewinfo.Text = u"Erro ao montar as tabelas: {}".format(to_unicode(ex))
        finally:
            self._busy = False
        self._update_counter()

    @staticmethod
    def _reselect(grid, rid):
        """Mantém destacada a linha que estava selecionada (sem reenviar ao Revit)."""
        if rid is None:
            return
        try:
            for drv in grid.ItemsSource:
                if int(drv.Row["RowId"]) == rid:
                    grid.SelectedItem = drv
                    return
        except Exception:
            pass

    def _on_mode(self, sender, args):
        self._rebuild()

    # ---- busca: só ao clicar em Buscar (ou Enter); Limpar busca restaura ----
    def _search(self, sender, args):
        self._query = _fold(to_unicode(self.f_text.Text).strip())
        self._rebuild()
        if self._query:
            self.searchinfo.Text = u"Busca '{}': {} linha(s) encontrada(s)".format(
                to_unicode(self.f_text.Text).strip(),
                sum(len(v) for v in self._vis.values()))
        else:
            self.searchinfo.Text = u""

    def _clear_search(self, sender, args):
        self.f_text.Text = u""
        self._query = u""
        self.searchinfo.Text = u""
        self._rebuild()

    def _on_key(self, sender, args):
        try:
            from System.Windows.Input import Key
            if args.Key in (Key.Enter, Key.Return):
                args.Handled = True
                self._search(None, None)
        except Exception:
            pass

    # ---- recolher / expandir e rolagem da janela ----
    def _toggle(self, n):
        from System.Windows import Visibility
        body, btn = getattr(self, "body%d" % n), getattr(self, "tg%d" % n)
        show = body.Visibility != Visibility.Visible
        body.Visibility = Visibility.Visible if show else Visibility.Collapsed
        btn.Content = u"▼" if show else u"▶"

    def _wheel(self, sender, args):
        args.Handled = True
        self.scroll.ScrollToVerticalOffset(self.scroll.VerticalOffset - args.Delta * 0.4)

    def _commit(self):
        """Confirma células/linhas em edição (tabelas 1 e 2) antes de ler os
        valores ou remontar as tabelas."""
        try:
            from System.Windows.Controls import DataGridEditingUnit
        except Exception:
            return
        for g in (self.grid, self.grid2):
            try:
                g.CommitEdit(DataGridEditingUnit.Cell, True)
                g.CommitEdit(DataGridEditingUnit.Row, True)
            except Exception:
                try:
                    g.CancelEdit()
                except Exception:
                    pass

    # ---- Revit: selecionar e mostrar na vista ativa (via ExternalEvent) ----
    @staticmethod
    def _ids_of_types(type_ids):
        out = []
        for tid in type_ids:
            out.extend(TYPE_INST.get(tid, []))
        return out

    @staticmethod
    def _selected_rid(grid):
        """RowId da linha selecionada; None para nada / linha de título."""
        item = grid.SelectedItem
        if item is None:
            return None
        try:
            rid = int(item["RowId"])
        except Exception:
            return None
        return rid if rid >= 0 else None

    def _on_select1(self, sender, args):
        rid = None if self._busy else self._selected_rid(self.grid)
        if rid is not None and rid in self.by_rid:
            self._show(self._ids_of_types([h["id"] for h in self.by_rid[rid]["hs"]]))

    def _on_select2(self, sender, args):
        rid = None if self._busy else self._selected_rid(self.grid2)
        if rid is not None and rid in self.by_rid2:
            self._show(self._ids_of_types([self.by_rid2[rid]["id"]]))

    def _on_select3(self, sender, args):
        rid = None if self._busy else self._selected_rid(self.grid3)
        if rid is not None and rid in self.by_rid3:
            self._show(self._ids_of_types([self.by_rid3[rid]["id"]]))

    def _show(self, ids):
        """Seleciona no Revit TODAS as instâncias do tipo da linha (substitui a
        seleção anterior) e enquadra na vista ativa. Sem transação."""
        self._last_ids = list(ids)
        if not ids:
            self.viewinfo.Text = u"Nenhuma instância deste tipo no modelo."
            return
        win = self

        def act():
            win.viewinfo.Text = show_in_view(win._last_ids)
        run_in_revit(act)

    # ---- tabela 1 ----
    def _check1(self, rid, value):
        r = self.by_rid[rid]
        r["sel"] = bool(value) and r["enabled"]
        if rid in self.d1:
            self.d1[rid]["Sel"] = r["sel"]

    def _on_changed1(self, sender, e):
        if self._busy or e.Column.ColumnName != "Sel":
            return
        rid = int(e.Row["RowId"])
        if rid < 0:
            return
        self._busy = True
        try:
            self._check1(rid, e.Row["Sel"] == True)
            self._update_counter()
        finally:
            self._busy = False

    # ---- tabela 2 (só o Keynote é digitado; Descrições vêm do Excel) ----
    def _check2(self, rid, value):
        r = self.by_rid2[rid]
        r["sel"] = bool(value)
        if rid in self.d2:
            self.d2[rid]["Sel"] = r["sel"]

    def _refresh2(self, r):
        row = self.d2.get(r["rid"])
        if row is None:                   # fora da busca atual
            return
        vals = self._vals2(r)
        row["Kind"] = vals.pop("_kind")
        for k, v in vals.items():
            row[k] = v

    def _on_changed2(self, sender, e):
        name = e.Column.ColumnName
        if self._busy or name not in ("Sel", "Keynote"):
            return
        rid = int(e.Row["RowId"])
        if rid < 0:
            return
        self._busy = True
        try:
            if name == "Sel":
                self._check2(rid, e.Row["Sel"] == True)
            else:
                r = self.by_rid2[rid]
                r["key"] = clean(e.Row["Keynote"] or u"")
                # novo keynote: confere no Excel (destaque + valores do Excel);
                # descrições só são usadas depois de 'Verificar no Excel'
                r["pt"], r["en"] = r["pt0"], r["en0"]
                verify_incomplete(r, autofill=False)
                self._refresh2(r)
                self._check2(rid, inc_changed(r))
            self._update_counter()
        finally:
            self._busy = False

    def _verify_all(self, sender, args):
        """Confere cada Keynote no Excel e traz Descrição / Descrição IN de lá
        para os campos VAZIOS. Só marca as linhas: gravar é outro passo."""
        self._commit()
        self._busy = True
        n_match = 0
        try:
            for r in self.inc:
                verify_incomplete(r, autofill=True)
                if excel_rec(r["key"]) is not None:
                    n_match += 1
                self._refresh2(r)
                if inc_changed(r):
                    self._check2(r["rid"], True)
            self._update_counter()
        finally:
            self._busy = False
        self.viewinfo.Text = (u"Verificado no Excel: {} de {} linha(s) com keynote encontrada. "
                              u"Nada foi gravado - confira e use 'Gravar marcados'.".format(
                                  n_match, len(self.inc)))

    # ---- comum ----
    def _update_counter(self):
        s1, s2 = self.selected(), self.selected_inc()
        tot1 = len([r for r in self.rows if r["enabled"]])
        n_over = len([r for r in s1 if r["over"]])
        n_edit = len([r for r in s2 if inc_changed(r)])
        self.counter.Text = (
            u"Tabela 1: {} de {} linha(s) marcada(s) ({} sobrescreve(m))   ·   "
            u"Tabela 2: {} de {} marcada(s), {} com alteração   ·   "
            u"Tabela 3: {} tipo(s) com Keynote fora do Excel".format(
                len(s1), tot1, n_over, len(s2), len(self.inc), n_edit, len(self.notx)))
        tots = (len(self.rows), len(self.inc), len(self.notx))
        for n, lbl, extra in ((1, self.h1, u"{} marcada(s)".format(len(s1))),
                              (2, self.h2, u"{} marcada(s)".format(len(s2))),
                              (3, self.h3, u"")):
            vis, tot = len(self._vis[n]), tots[n - 1]
            txt = u"{} linha(s)".format(tot) if vis == tot else \
                u"{} de {} linha(s) na busca".format(vis, tot)
            lbl.Text = u"·  " + txt + (u"  ·  " + extra if extra else u"")

    def selected(self):
        return [r for r in self.rows if r["sel"] and r["enabled"]]

    def selected_inc(self):
        return [r for r in self.inc if r["sel"]]

    def reload(self, rows, inc, notx, info_text, msg):
        """Depois de gravar: novas linhas (o que ainda falta) nas 3 tabelas.
        Busca e agrupamento escolhidos continuam valendo."""
        self.rows, self.inc, self.notx = rows, inc, notx
        self.by_rid = dict((r["rid"], r) for r in rows)
        self.by_rid2 = dict((r["rid"], r) for r in inc)
        self.by_rid3 = dict((r["rid"], r) for r in notx)
        self.info.Text = info_text
        self.viewinfo.Text = msg
        self._rebuild()
        self.b_apply.IsEnabled = True

    def _apply(self, sender, args):
        """Grava as linhas marcadas - SÓ depois da confirmação - e atualiza as
        tabelas (a janela continua aberta)."""
        self._commit()
        sel1, sel2 = self.selected(), [r for r in self.selected_inc() if inc_changed(r)]
        if not sel1 and not sel2:
            self.viewinfo.Text = u"Nada marcado para gravar."
            return
        n_over = len([r for r in sel1 if r["over"]])
        if not _ask(self, u"Gravar no modelo?\n\n"
                          u"Tabela 1 (Excel > Modelo): {} tipo(s){}\n"
                          u"Tabela 2 (Keynote / descrições): {} tipo(s)\n\n"
                          u"Ctrl+Z no Revit desfaz a gravação.".format(
                              len(sel1), u" - {} substitui(em) valor atual".format(n_over)
                              if n_over else u"", len(sel2))):
            self.viewinfo.Text = u"Gravação cancelada - nada foi alterado."
            return
        self.b_apply.IsEnabled = False
        self.viewinfo.Text = u"Gravando no modelo..."
        run_in_revit(lambda: save_run(sel1, sel2))

    def _txt(self, sender, args):
        """Gera o TXT de Keynote a partir do Excel e carrega no projeto (pede
        confirmação; avisa se vai substituir um arquivo existente)."""
        path = opts["txt_path"]
        try:
            text, warns = build_keynote_txt()
        except Exception as ex:
            _tell(self, u"Não foi possível montar o TXT:\n{}".format(to_unicode(ex)), warn=True)
            return
        n_lines = len(text.splitlines())
        msg = (u"Gerar o TXT de Keynote a partir do Excel e carregar no projeto?\n\n{}\n\n"
               u"{} linha(s) · texto do keynote = DESCRIÇÃO | DESCRIÇÃO IN".format(path, n_lines))
        if os.path.isfile(path):
            msg += (u"\n\nATENÇÃO: o arquivo já existe e será SUBSTITUÍDO "
                    u"(um backup com data/hora é guardado na mesma pasta).")
        if warns:
            msg += u"\n\n{} aviso(s): {}".format(len(warns), u"; ".join(warns[:3]))
        if not _ask(self, msg):
            self.viewinfo.Text = u"TXT de Keynote não gerado (cancelado)."
            return
        win = self

        def act():
            try:
                backup, enc = write_keynote_txt(path, text)
            except Exception as ex:
                win.viewinfo.Text = u"TXT não gravado: {}".format(to_unicode(ex))
                _tell(win, win.viewinfo.Text, warn=True)
                return
            ok, info = load_keynote_txt(path)
            if ok:
                FILE_TEXT.clear()
                FILE_TEXT.update(keynote_file_texts())
            win.viewinfo.Text = u"TXT gravado ({}){} e {}.".format(
                enc, u", backup: " + os.path.basename(backup) if backup else u"",
                info if ok else u"NÃO carregado no projeto: " + info)
            if not ok:
                _tell(win, win.viewinfo.Text, warn=True)
        run_in_revit(act)

    def _dashboard(self, sender, args):
        """Dashboard HTML com os dados da última análise (não lê nem grava o
        modelo de novo)."""
        try:
            path = write_dashboard()
        except Exception as ex:
            _tell(self, u"Não foi possível gerar o dashboard:\n{}".format(to_unicode(ex)),
                  warn=True)
            return
        if open_in_browser(path):
            self.viewinfo.Text = u"Dashboard aberto no navegador: {}".format(path)
        else:
            self.viewinfo.Text = u"Dashboard gerado (abra no navegador): {}".format(path)

    def _back(self, sender, args):
        """Volta para a 1ª aba. O que já foi gravado continua no modelo;
        marcações não gravadas são descartadas."""
        self._commit()
        pend = len(self.selected()) + len([r for r in self.selected_inc() if inc_changed(r)])
        if pend and not _ask(self, u"Voltar para a primeira aba?\n\n{} linha(s) marcada(s) "
                                   u"ainda não foram gravadas e serão descartadas.".format(pend)):
            return
        self.closing_for_back = True
        self.Close()
        run_in_revit(go_back)

    def _finish(self, sender, args):
        """Encerra o comando: só fecha a interface. Não grava, não relê o
        modelo, não gera TXT nem relatório."""
        self.Close()
        try:
            output.close()
        except Exception:
            pass


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

def offer_in_binding():
    """Oferece vincular o parâmetro compartilhado 'Descrição IN' onde falta."""
    global MODEL
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


VERIFY_LOG = []          # linhas da tabela 1 já gravadas (todas as rodadas)
INC_LOG = []             # linhas da tabela 2 já gravadas (todas as rodadas)


def analyze_model():
    """(Re)analisa Excel x modelo e monta as 3 tabelas. Roda no início e
    depois de cada 'Gravar marcados'."""
    global m_rows, m_plan, m_blocked, VROWS, INC, NOTX
    m_rows, m_plan, m_blocked = model_rows()
    VROWS = m_rows
    for r in VROWS:
        r["view"] = row_view(r)
        r["sel"] = r["enabled"]           # Excel -> Modelo começa marcado
    VROWS.sort(key=lambda r: (r["key"], r["category"], r["view"]["family"],
                              r["view"]["type"]))
    for i, r in enumerate(VROWS):
        r["rid"] = i
    INC, NOTX = collect_incomplete(set(r["key"] for r in VROWS))
    for i, r in enumerate(INC):
        r["rid"] = i
        verify_incomplete(r, autofill=False)
    for i, r in enumerate(NOTX):
        r["rid"] = i
    TYPE_INST.clear()
    recs = []
    for info in scan_types():             # instâncias de cada tipo (mostrar na vista)
        TYPE_INST[eid_int(info["t"].Id)] = list(info["ids"])
        try:
            recs.append(classify_type(info["t"], info))
        except Exception:
            pass                           # tipo ilegível: fica fora (não conta como correto)
    recs.sort(key=lambda x: (x["category"], x["family"], x["type"]))
    DASH["records"] = recs
    DASH["analyzedAt"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def make_info_text():
    keys_with_info = [k for k in MODEL if XL["by_key"].get(k) and
                      (XL["by_key"][k]["pt"] or XL["by_key"][k]["en"])]
    return (u"{} keynote(s) no modelo · {} com descrição no Excel · {} fora do Excel  ·  "
            u"Tabela 1: {} linha(s) ({} tipo(s))  ·  Tabela 2: {} tipo(s) sem Keynote ou "
            u"sem Descrição  ·  Tabela 3: {} tipo(s) com Keynote fora do Excel".format(
                len(MODEL), len(keys_with_info),
                len([k for k in MODEL if not XL["by_key"].get(k)]),
                len(VROWS), sum(len(r["hs"]) for r in VROWS), len(INC), len(NOTX)))


info_text = u""


# ------------------------------------------------------------------
# 5i. "Gravar marcados" (depois da confirmação): grava, reanalisa e ATUALIZA
#     as tabelas (a janela continua aberta). Roda via ExternalEvent.
#     "Finalizar" só fecha a interface (nada é gravado nem gerado).
# ------------------------------------------------------------------
VW = None


def inc_view(r):
    return {"category": r["category"], "family": r["family"], "type": r["name"],
            "typeId": r["id"], "count": r["count"], "hasIn": r["hasIn"],
            "key0": r["key0"], "pt0": r["pt0"], "en0": r["en0"],
            "key": clean(r["key"]), "pt": clean(r["pt"]), "en": clean(r["en"]),
            "written": list(r.get("written", [])), "check": r["check"],
            "status": r["status"]}


def save_run(chosen_rows, chosen_inc):
    global INC_FIELDS, MODEL, chosen_ids, chosen_inc_ids, chosen_item_ids, chosen_items
    global errs, f, failed_rows, field, getp, h, it, label, n_before, n_saved
    global new, old, p, r, t, to_write, tp, where
    n_before = len(changes)
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
            del changes[n_before:]            # nada desta transação ficou no modelo
            failed_rows.update(r["rid"] for r in chosen_rows)
            forms.alert(u"Gravação da tabela 1 desfeita (rollback) - nada foi alterado:\n{}"
                        .format(to_unicode(ex)))

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
        n_before2 = len(changes)
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
            del changes[n_before2:]
            for r in to_write:
                r["status"], r["written"] = u"Não gravado (rollback)", []
            forms.alert(u"Gravação da tabela 2 desfeita (rollback) - nada foi alterado:\n{}"
                        .format(to_unicode(ex)))

    # guarda o que foi gravado nesta rodada
    for r in chosen_rows:
        VERIFY_LOG.append(dict(r["view"], status=r["status"]))
    for r in chosen_inc:
        INC_LOG.append(inc_view(r))
    n_saved = len(changes) - n_before

    # relê o modelo e refaz as tabelas
    MODEL = collect_model(use_view, opts["materials"])
    analyze_model()
    if VW is not None:
        VW.reload(VROWS, INC, NOTX, make_info_text(),
                  u"{} parâmetro(s) gravado(s) no modelo. Tabelas atualizadas - "
                  u"grave mais ou clique em Finalizar.".format(n_saved))


# ------------------------------------------------------------------
# 5j. Dashboard HTML de qualidade dos dados (botão "Dashboard HTML")
#   Usa o retrato gravado por analyze_model() (DASH): NÃO relê nem grava o
#   modelo. Registro = TIPO (família + tipo) usado por elementos do modelo
#   em categoria com parâmetro Keynote (mesmo universo das tabelas).
#   Critérios por registro:
#     falta ........ Keynote vazio · Description vazia · Descrição IN vazia
#     divergência .. Keynote não está no Excel · Description diferente do
#                    Excel · Descrição IN diferente do Excel (só compara quando
#                    os DOIS lados têm valor)
#     não verificado parâmetro Descrição IN inexistente no tipo, ou Excel sem
#                    descrição para comparar com a do modelo
#   Situação (uma por registro, nesta ordem): Informação faltante >
#   Divergência > Não verificado > Correto.
#   Saúde = corretos ÷ (total analisado − não verificados).
# ------------------------------------------------------------------
DASH = {"records": [], "analyzedAt": u""}


def classify_type(t, info):
    """Situação de UM tipo para o dashboard (só leitura)."""
    key, desc, en, has_in, _missing = type_status(t)
    rec = XL["by_key"].get(key) if key else None
    missing, diverg, unver = [], [], []
    if not key:
        missing.append(u"Keynote vazio")
    if not desc:
        missing.append(u"Description vazia")
    if not has_in:
        unver.append(u"Parâmetro {} não existe no tipo".format(PARAM_IN_NAME))
    elif not en:
        missing.append(u"{} vazia".format(PARAM_IN_NAME))
    if key and rec is None:
        diverg.append(u"Keynote não está no Excel")
    if rec is not None:
        if desc and rec["pt"] and desc != rec["pt"]:
            diverg.append(u"Description diferente do Excel")
        elif desc and not rec["pt"]:
            unver.append(u"Excel sem Description para comparar")
        if has_in and en and rec["en"] and en != rec["en"]:
            diverg.append(u"{} diferente do Excel".format(PARAM_IN_NAME))
        elif has_in and en and not rec["en"]:
            unver.append(u"Excel sem {} para comparar".format(PARAM_IN_NAME))
    if missing:
        status = "missing"
    elif diverg:
        status = "diverg"
    elif unver:
        status = "unverified"
    else:
        status = "ok"
    return {"id": eid_int(t.Id), "category": info["category"], "family": family_name(t),
            "type": elem_name(t), "count": len(info["ids"]), "key": key,
            "pt": desc, "en": en, "hasIn": has_in, "inExcel": rec is not None,
            "xlRow": rec["row"] if rec is not None else None,
            "xlPt": rec["pt"] if rec is not None else u"",
            "xlEn": rec["en"] if rec is not None else u"",
            "status": status, "missing": missing, "diverg": diverg, "unverified": unver}


def build_dashboard_data():
    recs = DASH["records"]
    n = len(recs)
    by = defaultdict(int)
    for r in recs:
        by[r["status"]] += 1
    flag_missing = len([r for r in recs if r["missing"]])
    flag_diverg = len([r for r in recs if r["diverg"]])
    evaluated = n - by["unverified"]
    occ = OrderedDict()                   # ocorrência -> [grupo, quantidade]
    for r in recs:
        for grp in ("missing", "diverg", "unverified"):
            for o in r[grp]:
                occ.setdefault(o, [grp, 0])[1] += 1
    return {
        "project": to_unicode(doc.Title),
        "excel": os.path.basename(XLSX_PATH), "excelPath": XLSX_PATH,
        "sheet": XL["sheet"], "inParam": PARAM_IN_NAME,
        "analyzedAt": DASH["analyzedAt"],
        "generatedAt": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "savedThisRun": len(changes),
        "totals": {"records": n, "ok": by["ok"], "missing": by["missing"],
                   "diverg": by["diverg"], "unverified": by["unverified"],
                   "evaluated": evaluated, "flagMissing": flag_missing,
                   "flagDiverg": flag_diverg,
                   "instances": sum(r["count"] for r in recs)},
        "occurrences": [{"label": k, "group": v[0], "count": v[1]} for k, v in occ.items()],
        "records": recs,
    }


def write_dashboard():
    """Grava o HTML (modelo keynotes.html + dados) e devolve o caminho."""
    template = os.path.join(HERE, "keynotes.html")
    if not os.path.isfile(template):
        raise IOError(u"Modelo do dashboard não encontrado: {}".format(template))
    with codecs.open(template, "r", encoding="utf-8") as f:
        html = f.read()
    data = json.dumps(build_dashboard_data(), ensure_ascii=False,
                      default=lambda o: to_unicode(o))
    html = html.replace("__DATA__", data.replace(u"</", u"<\\/"))
    out_path = script.get_document_data_file("keynotes_dashboard", "html")
    with codecs.open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


# ------------------------------------------------------------------
# 9. Remove as cores que as versões anteriores deixaram no modelo e abre a
#    tabela de verificação (NÃO modal)
# ------------------------------------------------------------------
def cleanup_old_colors():
    """Na 1ª execução desta versão (por documento) procura em TODAS as vistas;
    depois, só na vista ativa e nas vistas anotadas pela versão anterior."""
    done_mark = _colors_log_path("keynotes_cores_limpeza")
    old_log = _colors_log_path()
    logged = set()
    if old_log and os.path.isfile(old_log):
        try:
            with codecs.open(old_log, "r", encoding="utf-8") as f:
                logged = set(int(v) for v in json.load(f).get("views", {}).keys())
        except Exception:
            logged = set()
    first = not (done_mark and os.path.isfile(done_mark))
    if first:
        views = graphic_views()
    else:
        views = [v for v in graphic_views()
                 if eid_int(v.Id) in logged or eid_int(v.Id) == eid_int(active_view.Id)]
    if not views:
        found = []
    elif len(views) > 3:
        with forms.ProgressBar(title=u"Procurando cores antigas do comando nas vistas "
                                     u"({value} de {max_value})") as pb:
            found = find_old_highlight(views, pb)
    else:
        found = find_old_highlight(views)

    def mark_done():
        for path in (old_log,):
            try:
                if path and os.path.isfile(path):
                    os.remove(path)
            except Exception:
                pass
        try:
            if done_mark:
                with codecs.open(done_mark, "w", encoding="utf-8") as f:
                    json.dump({"done": datetime.now().isoformat()}, f)
        except Exception:
            pass

    if not found:
        mark_done()
        return
    n = sum(len(ids) for _v, ids in found)
    lines = [u"  • {}: {} elemento(s)".format(to_unicode(v.Name), len(ids))
             for v, ids in found[:12]]
    if len(found) > 12:
        lines.append(u"  • ... e mais {} vista(s)".format(len(found) - 12))
    if forms.alert(u"As versões anteriores deste comando deixaram cores no modelo "
                   u"({} elemento(s) em {} vista(s)):\n\n{}\n\nRemover agora? A "
                   u"sobreposição gráfica desses elementos volta ao padrão da vista."
                   .format(n, len(found), u"\n".join(lines)), yes=True, no=True):
        ok, bad = clear_old_highlight(found)
        output.print_md(u"- Cores antigas do comando removidas: **{}** elemento(s){}.".format(
            ok, u" · {} não puderam ser limpos (vista emprestada por outro usuário?)"
            .format(bad) if bad else u""))
        if not bad:
            mark_done()
    else:
        output.print_md(u"- :warning: {} elemento(s) ainda com as cores antigas do comando "
                        u"(pergunta de novo na próxima execução).".format(n))


def start_session(new_opts):
    """Lê o Excel e o modelo com as opções da 1ª aba, analisa e abre a 2ª aba.
    Roda no início do comando e de novo pelo botão Voltar (via ExternalEvent).
    -> True se a janela abriu."""
    global opts, XLSX_PATH, XL, MODEL, FILE_TEXT, active_view, info_text, VW
    global not_written
    opts = new_opts
    XLSX_PATH = opts["xlsx"]
    active_view = doc.ActiveView
    try:
        XL = read_excel(XLSX_PATH)
    except Exception as ex:
        forms.alert(u"Não foi possível ler a planilha:\n{}\n\n{}".format(
            XLSX_PATH, to_unicode(ex)))
        return False
    MODEL = collect_model(use_view, opts["materials"])
    FILE_TEXT = keynote_file_texts()
    offer_in_binding()
    not_written = {}
    del VERIFY_LOG[:]
    del INC_LOG[:]
    analyze_model()
    info_text = make_info_text()
    if not VROWS and not INC:
        info_text += u"  ·  Nada a atualizar: o modelo já coincide com o Excel."
    VW = VerifyWindow(VROWS, INC, NOTX, info_text)
    VW.Show()
    return True


def go_back():
    """Botão Voltar: reabre a 1ª aba com as opções atuais e, se confirmada,
    analisa de novo. Cancelar na 1ª aba encerra o comando."""
    new_opts = ask_options({"xlsx": opts["xlsx"], "txt_path": opts["txt_path"]})
    if new_opts is not None:
        start_session(new_opts)


_first_opts = ask_options()
if _first_opts is None:
    script.exit()
try:
    cleanup_old_colors()
except Exception as _ex:
    output.print_md(u"- :warning: Não foi possível verificar as cores antigas: {}".format(
        to_unicode(_ex)))
start_session(_first_opts)
