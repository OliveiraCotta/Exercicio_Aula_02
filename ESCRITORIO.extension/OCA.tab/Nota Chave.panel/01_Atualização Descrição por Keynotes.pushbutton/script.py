# -*- coding: utf-8 -*-
"""Atualização Descrição por Keynotes - sincroniza Keynotes do modelo com o Excel.

Base única de Keynotes entre o Revit e a planilha do escritório
(OCA_XXX_Lista_de_Materiais - aba MATERIAIS).

Onde cada dado mora:
  * Keynote ........ parâmetro nativo KEYNOTE dos TIPOS usados no modelo e dos
                     MATERIAIS usados pelos elementos (opcional).
  * Descrição PT ... parâmetro nativo "Description" (tipo / material).
  * Descrição IN ... parâmetro compartilhado "Descrição em IN" (tipo / material).
  * Excel .......... aba MATERIAIS: colunas KEYNOTE, DESCRIÇÃO, DESCRIÇÃO IN.

Fonte da verdade (evita que uma direção desfaça a outra):
  * EXISTÊNCIA do keynote -> modelo. Keynote novo no modelo é ADICIONADO ao
    fim da tabela do Excel (PREFIXO/SEQUENCIA + fórmulas do template).
  * DESCRIÇÕES -> Excel (base revisada). Modelo -> Excel só PREENCHE células
    vazias; divergências vão para o relatório e nunca são sobrescritas.
  * Keynote do Excel não usado no modelo -> só relatório, nunca apagado.
  * Célula vazia no Excel NUNCA apaga valor no modelo.

Modos:
  0. Somente analisar  - relatório de inconsistências, nada é gravado.
  1. Modelo -> Excel   - adiciona keynotes e preenche descrições vazias.
  2. Excel  -> Modelo  - grava Description / Descrição em IN nos tipos e
                         materiais. Antes de gravar abre uma janela de
                         verificação em árvore (Categoria > Keynote > alteração)
                         para o usuário marcar o que deve ser gravado.

O Excel é lido e gravado direto no XML do .xlsx (sem Excel/COM). Antes de
gravar é feito um backup com data/hora ao lado do arquivo. O arquivo precisa
estar FECHADO no Excel.

Ao final pergunta se gera o relatório HTML (sempre gerado em "Somente
analisar"): preenchidos, alterados e faltando/pendentes por categoria do
modelo > keynote, mais as inconsistências Excel × modelo.
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

import clr
clr.AddReference("System.Xml")
for _asm in ("System.IO.Compression",
             "System.IO.Compression.FileSystem",   # .NET Framework (Revit <= 2024)
             "System.IO.Compression.ZipFile"):     # .NET 8 (Revit 2025+)
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
from System.IO import (StreamReader, StreamWriter, File, FileMode,
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
# nome do parâmetro compartilhado da descrição em inglês (primeiro que existir)
PARAM_IN_NAMES = [u"Descrição em IN", u"Descrição IN", u"Descricao em IN",
                  u"Descricao IN", u"Description IN"]
MATERIALS_LABEL = u"Materiais"     # categoria usada no relatório p/ materiais

MODES = OrderedDict([
    ("analyze", u"Somente analisar"),
    ("to_excel", u"Modelo → Excel"),
    ("to_model", u"Excel → Modelo"),
])


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
        Title="Atualização Descrição por Keynotes" Height="Auto" Width="540"
        SizeToContent="Height" WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize" Background="#0E1526">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,12,0,4"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="#D9E8F5"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#D9E8F5"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,8,0,0"/>
    </Style>
  </Window.Resources>
  <StackPanel Margin="18">
    <TextBlock Text="KEYNOTES  ×  EXCEL" FontSize="15" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,0,0,4"
               Text="Compara os Keynotes dos tipos e materiais usados no modelo com a aba MATERIAIS da planilha. Descrição PT = parâmetro Description; Descrição IN = parâmetro compartilhado 'Descrição em IN'. Ao final abre um relatório por categoria."/>

    <TextBlock Text="Planilha Excel (.xlsx - feche o arquivo no Excel antes de gravar)"/>
    <DockPanel LastChildFill="True">
      <Button x:Name="browse" DockPanel.Dock="Right" Content="Procurar..." Width="90"
              Height="26" Margin="8,0,0,0"/>
      <TextBox x:Name="xlsx" Height="26" Padding="4,3,4,2"/>
    </DockPanel>

    <TextBlock Text="Operação"/>
    <RadioButton x:Name="m_analyze" GroupName="mode"
                 Content="Somente analisar  (relatório - nada é gravado)"/>
    <RadioButton x:Name="m_to_excel" GroupName="mode"
                 Content="Modelo → Excel  (adiciona keynotes novos e preenche descrições vazias)"/>
    <RadioButton x:Name="m_to_model" GroupName="mode"
                 Content="Excel → Modelo  (atualiza Description e Descrição em IN no Revit)"/>

    <TextBlock Text="Escopo"/>
    <RadioButton x:Name="s_model" GroupName="scope" Content="Modelo inteiro"/>
    <RadioButton x:Name="s_view" GroupName="scope" Content="Somente elementos visíveis na vista ativa"/>
    <CheckBox x:Name="materials" Content="Incluir Keynotes de materiais"/>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
      <Button x:Name="cancel" Content="Cancelar" Width="90" Height="28" Margin="0,0,10,0"/>
      <Button x:Name="ok" Content="Executar" Width="120" Height="28"/>
    </StackPanel>
  </StackPanel>
</Window>
"""


class KeynoteWindow(forms.WPFWindow):
    def __init__(self, xaml, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.result = None
        self.xlsx.Text = cfg["xlsx"]
        {"to_excel": self.m_to_excel,
         "to_model": self.m_to_model}.get(cfg["mode"], self.m_analyze).IsChecked = True
        (self.s_view if cfg["scope"] == "view" else self.s_model).IsChecked = True
        self.materials.IsChecked = bool(cfg["materials"])
        self.browse.Click += self._browse
        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    def _browse(self, sender, args):
        picked = forms.pick_file(file_ext="xlsx",
                                 title="Selecione a planilha de Keynotes")
        if picked:
            self.xlsx.Text = picked

    def _ok(self, sender, args):
        path = to_unicode(self.xlsx.Text).strip().strip('"')
        if not path or not os.path.isfile(path) or not path.lower().endswith(".xlsx"):
            forms.alert(u"Selecione um arquivo .xlsx válido.", title=__title__.replace("\n", " "))
            return
        if self.m_to_excel.IsChecked:
            mode = "to_excel"
        elif self.m_to_model.IsChecked:
            mode = "to_model"
        else:
            mode = "analyze"
        self.result = {
            "xlsx": path,
            "mode": mode,
            "scope": "view" if self.s_view.IsChecked else "model",
            "materials": bool(self.materials.IsChecked),
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
             "materials": config.get_option("materials", True)}

win = KeynoteWindow(KN_XAML, saved_cfg)
win.ShowDialog()
if not win.confirmed:
    script.exit()
opts = win.result

config.xlsx_path = opts["xlsx"]
config.mode = opts["mode"]
config.scope = opts["scope"]
config.materials = opts["materials"]
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
        forms.alert(u"A vista ativa não é uma vista gráfica (tabela, folha...).\n"
                    u"Abra uma planta, corte, elevação ou 3D, ou use o escopo "
                    u"'Modelo inteiro'.", exitscript=True)


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

        prefix_map = {}
        dados = sheets.get(norm_header(DATA_SHEET_NAME))
        if dados and dados[1]:
            dtext = zip_read(archive, dados[1])
            if dtext:
                dgrid, _ = parse_sheet(dtext, shared)
                for rn, cells in dgrid.items():
                    pf = clean(cells.get(0, u""))
                    if pf and pf not in prefix_map:
                        prefix_map[pf] = (clean(cells.get(1, u"")), clean(cells.get(2, u"")))
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
    """Parâmetro compartilhado da descrição em inglês (None se não existir)."""
    for nm in PARAM_IN_NAMES:
        try:
            p = elem.LookupParameter(nm)
            if p is not None and p.StorageType == DB.StorageType.String:
                return p
        except Exception:
            continue
    return None


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


def keynote_file_texts():
    """Texto do arquivo .txt de keynotes carregado no projeto (só informativo)."""
    out = {}
    try:
        table = DB.KeynoteTable.GetKeynoteTable(doc)
        for entry in table.GetKeyBasedTreeEntries():
            try:
                out[clean(entry.Key)] = clean(entry.KeynoteText)
            except Exception:
                continue
    except Exception:
        pass
    return out


MODEL = collect_model(use_view, opts["materials"])
FILE_TEXT = keynote_file_texts()

if not MODEL and MODE == "to_model":
    forms.alert(u"Nenhum tipo ou material com Keynote preenchido foi encontrado "
                u"no escopo escolhido.", exitscript=True)


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
def plan_to_excel():
    fills, appends = [], []
    for key, holders in MODEL.items():
        pts, ens = weighted_values(holders, "pt"), weighted_values(holders, "en")
        pt = pts[0] if pts else u""
        en = ens[0] if ens else u""
        rec = XL["by_key"].get(key)
        if rec is None:
            prefix, seq = split_keynote(key)
            grp, cat = XL["prefix_map"].get(prefix, (u"", u""))
            note = u"Incluído via Revit em {}".format(datetime.now().strftime("%d/%m/%Y"))
            if prefix not in XL["prefix_map"]:
                note += u" - prefixo não cadastrado em DADOS"
            if len(pts) > 1 or len(ens) > 1:
                note += u" - modelo com descrições diferentes, revisar"
            appends.append({"key": key, "prefix": prefix, "seq": seq, "grp": grp,
                            "cat": cat, "pt": pt, "en": en, "obs": note})
            continue
        if not rec["pt"] and pt:
            fills.append((rec, "pt", pt))
        if not rec["en"] and en:
            fills.append((rec, "en", en))
    return fills, appends


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
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(XLSX_PATH)
    backup = u"{}_backup_{}{}".format(base, stamp, ext)
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


if MODE == "to_excel":
    fills, appends = plan_to_excel()
    if not fills and not appends:
        output.print_md(u"**Modelo → Excel:** nada a gravar - todos os Keynotes do "
                        u"modelo já estão na planilha e nenhuma descrição vazia pode "
                        u"ser preenchida pelo modelo.")
    else:
        output.print_md(u"## Modelo → Excel - prévia")
        if appends:
            output.print_md(u"**{}** Keynote(s) serão ADICIONADOS ao fim da aba "
                            u"*{}*:".format(len(appends), XL["sheet"]))
            output.print_table(
                table_data=[[a["key"], a["prefix"], a["seq"], a["pt"] or u"-",
                             a["en"] or u"-"] for a in appends],
                columns=[u"Keynote", u"Prefixo", u"Sequência", u"Descrição", u"Descrição IN"])
        if fills:
            output.print_md(u"**{}** célula(s) vazia(s) serão PREENCHIDAS:".format(len(fills)))
            output.print_table(
                table_data=[[r["key"], r["row"], u"DESCRIÇÃO" if f == "pt" else u"DESCRIÇÃO IN", v]
                            for r, f, v in fills],
                columns=[u"Keynote", u"Linha", u"Coluna", u"Valor"])
        output.print_md(u"_Descrições já preenchidas no Excel NÃO são alteradas._")

        while not ensure_file_closed(XLSX_PATH):
            if not forms.alert(u"A planilha está aberta (ou bloqueada).\n\n{}\n\n"
                               u"Feche o arquivo no Excel e clique em Sim para "
                               u"tentar de novo.".format(XLSX_PATH), yes=True, no=True):
                script.exit()

        if forms.alert(u"{} Keynote(s) adicionados e {} descrição(ões) preenchidas no "
                       u"Excel.\nUm backup será criado antes.\n\nGravar?".format(
                           len(appends), len(fills)), yes=True, no=True):
            snapshot = [(r, f, r[f]) for r, f, _ in fills]
            try:
                xl_written["backup"] = write_excel(fills, appends)
                xl_written["appended"] = len(appends)
                xl_written["filled"] = len(fills)
                for a in appends:
                    changes.append({"target": "excel", "action": "append", "key": a["key"],
                                    "field": u"Linha", "old": u"",
                                    "new": u"{} | {}".format(a["pt"] or u"-", a["en"] or u"-"),
                                    "where": XL["sheet"], "id": None})
                for r, f, old in snapshot:
                    changes.append({"target": "excel", "action": "fill", "key": r["key"],
                                    "field": u"DESCRIÇÃO" if f == "pt" else u"DESCRIÇÃO IN",
                                    "old": u"", "new": r[f],
                                    "where": u"{} · linha {}".format(XL["sheet"], r["row"]),
                                    "id": None})
            except Exception as ex:
                forms.alert(u"Falha ao gravar a planilha. Se o arquivo não abrir, "
                            u"restaure o backup criado ao lado dele "
                            u"(*_backup_AAAAMMDD_HHMMSS.xlsx).\n\n{}".format(
                                to_unicode(ex)), exitscript=True)
        else:
            output.print_md(u"**Cancelado - planilha não alterada.**")


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
                                       ("en", PARAM_IN_NAMES[0], "en_p")):
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
                if p is None:
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


# ---- janela de verificação: Categoria > Keynote > alteração (com checkbox) ----
for _asm in ("PresentationFramework", "PresentationCore", "WindowsBase"):
    try:
        clr.AddReference(_asm)
    except Exception:
        pass

SELECT_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Excel → Modelo · Verificar alterações" Width="1000" Height="700"
        MinWidth="700" MinHeight="420" WindowStartupLocation="CenterScreen"
        Background="#0E1526">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Height" Value="28"/>
      <Setter Property="Padding" Value="12,0,12,0"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
  </Window.Resources>
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Grid.Row="0">
      <TextBlock Text="EXCEL → MODELO  ·  MARQUE O QUE DEVE SER GRAVADO" FontSize="15"
                 FontWeight="SemiBold" Foreground="#65E3FF" FontFamily="Segoe UI"/>
      <TextBlock x:Name="info" TextWrapping="Wrap" FontSize="12" Foreground="#CFE3FF"
                 FontFamily="Segoe UI" Margin="0,6,0,0"/>
      <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" FontFamily="Segoe UI"
                 Margin="0,4,0,0"
                 Text="Verde = preenche parâmetro vazio  ·  Âmbar = SOBRESCREVE valor existente  ·  Cinza = não pode ser gravado (motivo ao lado). Parâmetros de tipo valem para todas as instâncias do tipo."/>
    </StackPanel>
    <WrapPanel Grid.Row="1" Margin="0,12,0,8">
      <Button x:Name="b_all" Content="Marcar tudo"/>
      <Button x:Name="b_none" Content="Desmarcar tudo"/>
      <Button x:Name="b_fill" Content="Somente preencher vazios"/>
      <Button x:Name="b_expand" Content="Expandir"/>
      <Button x:Name="b_collapse" Content="Recolher"/>
    </WrapPanel>
    <Border Grid.Row="2" BorderBrush="#2A4A66" BorderThickness="1">
      <TreeView x:Name="tree" Background="#0A1120" BorderThickness="0" Padding="4"/>
    </Border>
    <DockPanel Grid.Row="3" Margin="0,12,0,0" LastChildFill="False">
      <TextBlock x:Name="counter" DockPanel.Dock="Left" VerticalAlignment="Center"
                 Foreground="#CFE3FF" FontFamily="Segoe UI"/>
      <Button x:Name="b_apply" DockPanel.Dock="Right" Content="Gravar selecionados"
              Width="160" Margin="0"/>
      <Button x:Name="b_cancel" DockPanel.Dock="Right" Content="Cancelar" Width="100"
              Margin="0,0,10,0"/>
    </DockPanel>
  </Grid>
</Window>
"""

C_TEXT, C_KEY, C_FILL, C_OVER, C_OFF = "#D9E8F5", "#65E3FF", "#50D682", "#FFB454", "#7A8FA9"


def short(text, n=80):
    text = to_unicode(text).replace(u"\n", u" ")
    return text if len(text) <= n else text[:n - 1] + u"…"


def holder_label(h):
    if h["kind"] == "material":
        return u"Material: {}".format(h["name"])
    return u"Tipo: {}{}".format((h["family"] + u" - ") if h["family"] else u"", h["name"])


class TNode(object):
    """Um nível da árvore: CheckBox dentro de um TreeViewItem."""
    def __init__(self, text, color, item=None, enabled=True, bold=False):
        from System.Windows.Controls import CheckBox, TextBlock, TreeViewItem
        from System.Windows import FontWeights, VerticalAlignment
        from System.Windows.Media import SolidColorBrush, ColorConverter
        self.children, self.parent, self.item, self.enabled = [], None, item, enabled
        tb = TextBlock()
        tb.Text = text
        tb.Foreground = SolidColorBrush(ColorConverter.ConvertFromString(color))
        if bold:
            tb.FontWeight = FontWeights.SemiBold
        self.cb = CheckBox()
        self.cb.Content = tb
        self.cb.IsEnabled = enabled
        self.cb.IsChecked = bool(enabled)
        self.cb.VerticalContentAlignment = VerticalAlignment.Center
        self.tvi = TreeViewItem()
        self.tvi.Header = self.cb

    def add(self, child):
        child.parent = self
        self.children.append(child)
        self.tvi.Items.Add(child.tvi)
        return child

    def leaves(self):
        if not self.children:
            return [self]
        out = []
        for c in self.children:
            out.extend(c.leaves())
        return out


class SelectWindow(forms.WPFWindow):
    def __init__(self, plan, blocked, info_text):
        forms.WPFWindow.__init__(self, SELECT_XAML, literal_string=True)
        self.confirmed = False
        self.roots = []
        self.info.Text = info_text

        groups = OrderedDict()               # categoria -> keynote -> [itens]
        for it in sorted(plan + blocked, key=lambda i: (
                i["h"]["category"], i["h"]["key"], i["h"]["name"], i["field"])):
            groups.setdefault(it["h"]["category"], OrderedDict()) \
                  .setdefault(it["h"]["key"], []).append(it)

        for cat, keys in groups.items():
            n_items = sum(1 for its in keys.values() for i in its if not i["reason"])
            cnode = TNode(u"{}    ·  {} keynote(s)  ·  {} alteração(ões)".format(
                cat, len(keys), n_items), C_TEXT, bold=True)
            for key, its in keys.items():
                rec = XL["by_key"].get(key) or {}
                n_ok = sum(1 for i in its if not i["reason"])
                n_over = sum(1 for i in its if not i["reason"] and i["old"])
                ktxt = u"{}    PT: {}    |    IN: {}    ({} alteração(ões){})".format(
                    key, short(rec.get("pt") or u"-", 60), short(rec.get("en") or u"-", 60),
                    n_ok, u", {} sobrescreve(m)".format(n_over) if n_over else u"")
                knode = cnode.add(TNode(ktxt, C_KEY, bold=True))
                for it in its:
                    h = it["h"]
                    if it["reason"]:
                        txt = u"{}  —  {}: não pode ser gravado ({})".format(
                            holder_label(h), it["label"], it["reason"])
                        knode.add(TNode(txt, C_OFF, item=it, enabled=False))
                        continue
                    txt = u"{}  —  {}:  “{}”  →  “{}”".format(
                        holder_label(h), it["label"],
                        short(it["old"], 50) if it["old"] else u"(vazio)", short(it["new"], 60))
                    knode.add(TNode(txt, C_OVER if it["old"] else C_FILL, item=it))
                knode.tvi.IsExpanded = False
            cnode.tvi.IsExpanded = True
            self.roots.append(cnode)
            self.tree.Items.Add(cnode.tvi)

        for node in self._all_nodes():
            node.cb.Click += self._make_click(node)
        self.b_all.Click += lambda s, a: self._set_leaves(lambda it: True)
        self.b_none.Click += lambda s, a: self._set_leaves(lambda it: False)
        self.b_fill.Click += lambda s, a: self._set_leaves(lambda it: not it["old"])
        self.b_expand.Click += lambda s, a: self._expand(True)
        self.b_collapse.Click += lambda s, a: self._expand(False)
        self.b_apply.Click += self._apply
        self.b_cancel.Click += self._cancel
        self._refresh()

    def _all_nodes(self):
        stack, out = list(self.roots), []
        while stack:
            n = stack.pop()
            out.append(n)
            stack.extend(n.children)
        return out

    def _make_click(self, node):
        def handler(sender, args):
            state = node.cb.IsChecked == True
            for leaf in node.leaves():
                if leaf.enabled:
                    leaf.cb.IsChecked = state
            self._refresh()
        return handler

    def _set_leaves(self, rule):
        for root in self.roots:
            for leaf in root.leaves():
                if leaf.enabled:
                    leaf.cb.IsChecked = bool(rule(leaf.item))
        self._refresh()

    def _expand(self, state):
        """Categorias ficam sempre abertas; abre/fecha o nível dos keynotes."""
        for root in self.roots:
            root.tvi.IsExpanded = True
            for k in root.children:
                k.tvi.IsExpanded = state

    def _sync(self, node):
        """Pais: marcado / desmarcado / parcial (None) conforme as folhas."""
        if not node.children:
            return
        for c in node.children:
            self._sync(c)
        states = [l.cb.IsChecked == True for l in node.leaves() if l.enabled]
        if not states:
            node.cb.IsChecked, node.cb.IsEnabled = False, False
        elif all(states):
            node.cb.IsChecked = True
        elif not any(states):
            node.cb.IsChecked = False
        else:
            node.cb.IsChecked = None

    def _refresh(self):
        for root in self.roots:
            self._sync(root)
        sel = self.selected_items()
        total = sum(1 for r in self.roots for l in r.leaves() if l.enabled)
        keys = set((i["h"]["category"], i["h"]["key"]) for i in sel)
        over = sum(1 for i in sel if i["old"])
        self.counter.Text = (u"{} de {} alteração(ões) selecionada(s)  ·  {} keynote(s)  ·  "
                             u"{} sobrescreve(m) valor existente".format(
                                 len(sel), total, len(keys), over))

    def selected_items(self):
        return [l.item for r in self.roots for l in r.leaves()
                if l.enabled and l.cb.IsChecked == True]

    def _apply(self, sender, args):
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.confirmed = False
        self.Close()


# (holder id, campo) -> motivo de NÃO ter sido gravado (usado no resultado)
not_written = {}

if MODE == "to_model":
    plan, blocked = plan_to_model()
    for it in blocked:
        h = it["h"]
        not_written[(h["id"], it["field"])] = it["reason"]
        skipped.append({"key": h["key"], "where": where_label(h), "id": h["id"],
                        "eid": h["elem"].Id, "reason": it["reason"]})

    keys_with_info = [k for k in MODEL if XL["by_key"].get(k) and
                      (XL["by_key"][k]["pt"] or XL["by_key"][k]["en"])]
    keys_to_change = set(it["h"]["key"] for it in plan)
    info_text = (u"{} keynote(s) do modelo têm descrição no Excel  ·  {} com alteração a "
                 u"gravar  ·  {} já iguais ao Excel  ·  {} sem descrição no Excel  ·  "
                 u"{} fora do Excel.".format(
                     len(keys_with_info), len(keys_to_change),
                     len([k for k in keys_with_info if k not in keys_to_change
                          and not any(it["h"]["key"] == k for it in blocked)]),
                     len([k for k in MODEL if XL["by_key"].get(k) and k not in keys_with_info]),
                     len([k for k in MODEL if not XL["by_key"].get(k)])))

    if not plan:
        output.print_md(u"**Excel → Modelo:** nada a gravar - as descrições do modelo "
                        u"já coincidem com a planilha (ou não há permissão para gravar).")
    else:
        sw = SelectWindow(plan, blocked, info_text)
        sw.ShowDialog()
        if not sw.confirmed:
            chosen = []
            output.print_md(u"**Cancelado - modelo não alterado.**")
        else:
            chosen = sw.selected_items()
        chosen_ids = set(id(it) for it in chosen)
        for it in plan:
            if id(it) not in chosen_ids:
                not_written[(it["h"]["id"], it["field"])] = (
                    u"não selecionado pelo usuário" +
                    (u" (modelo diferente do Excel)" if it["old"] else u""))

        if chosen:
            t = DB.Transaction(doc, u"Atualizar descrições por Keynote")
            t.Start()
            try:
                for it in chosen:
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
                        not_written[(h["id"], field)] = to_unicode(ex)
                        skipped.append({"key": h["key"], "where": where_label(h),
                                        "id": h["id"], "eid": h["elem"].Id,
                                        "reason": to_unicode(ex)})
                t.Commit()
            except Exception as ex:
                t.RollBack()
                forms.alert(u"Atualização desfeita (rollback):\n{}".format(to_unicode(ex)),
                            exitscript=True)
        elif sw.confirmed:
            output.print_md(u"**Nenhuma alteração selecionada - modelo não alterado.**")


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
    ("no_in_param", {"label": u"Sem parâmetro 'Descrição em IN'", "color": "var(--amber)",
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

    keynotes_out.append({
        "key": key, "inModel": in_model, "inExcel": in_excel,
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
FIELD_LABEL = {"pt": u"Description", "en": PARAM_IN_NAMES[0]}
XL_FIELD_LABEL = {"pt": u"DESCRIÇÃO", "en": u"DESCRIÇÃO IN"}


def build_result():
    res = {"target": None, "filled": [], "changed": [], "missing": []}
    if MODE == "to_model":
        res["target"] = "model"
        written = set()
        for c in changes:
            if c["target"] != "model":
                continue
            written.add((c["id"], c["fieldKey"]))
            entry = {"category": c["category"], "key": c["key"], "where": c["where"],
                     "id": c["id"], "field": c["field"], "old": c["old"], "new": c["new"]}
            res["changed" if c["action"] == "overwrite" else "filled"].append(entry)
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
                        if h[pkey] is None:
                            reason = u"parâmetro '{}' não existe".format(label)
                        elif rec is None:
                            reason = u"keynote não está no Excel"
                        elif not rec[field]:
                            reason = u"sem descrição no Excel"
                        else:
                            reason = u"vazio"
                    res["missing"].append({
                        "category": h["category"], "key": key, "where": where_label(h),
                        "id": h["id"], "field": label, "old": value, "new": u"",
                        "reason": reason})
    elif MODE == "to_excel":
        res["target"] = "excel"
        for c in changes:
            if c["target"] != "excel":
                continue
            for cat in sorted(set(h["category"] for h in MODEL.get(c["key"], []))) or [u"-"]:
                res["filled"].append({
                    "category": cat, "key": c["key"], "where": c["where"], "id": None,
                    "field": c["field"] if c["action"] != "append" else u"Linha nova",
                    "old": u"", "new": c["new"]})
        for key, holders in MODEL.items():
            rec = XL["by_key"].get(key)
            for cat in sorted(set(h["category"] for h in holders)):
                if rec is None:
                    res["missing"].append({
                        "category": cat, "key": key, "where": XL["sheet"], "id": None,
                        "field": u"Linha", "old": u"", "new": u"",
                        "reason": u"keynote não gravado na planilha"})
                    continue
                for field in ("pt", "en"):
                    if not rec[field]:
                        res["missing"].append({
                            "category": cat, "key": key,
                            "where": u"{} · linha {}".format(XL["sheet"], rec["row"]),
                            "id": None, "field": XL_FIELD_LABEL[field], "old": u"",
                            "new": u"", "reason": u"vazio no Excel e no modelo"})
    for k in ("filled", "changed", "missing"):
        res[k].sort(key=lambda e: (e["category"], e["key"], e["where"], e["field"]))
    return res


RESULT = build_result()


# ------------------------------------------------------------------
# 8. Janela de saída + relatório HTML (opcional nos modos que gravam)
# ------------------------------------------------------------------
output.print_md(u"## Keynotes × Excel - {}".format(MODES[MODE]))
output.print_md(u"- **{}** keynote(s) no modelo ({} tipo(s)/material(is)) · **{}** no Excel "
                u"(aba *{}*) · **{}** em ambos".format(
                    len(MODEL), sum(len(v) for v in MODEL.values()), len(XL["by_key"]),
                    XL["sheet"],
                    sum(1 for k in keynotes_out if k["inModel"] and k["inExcel"])))
if RESULT["target"]:
    output.print_md(u"- Resultado no **{}**: **{}** preenchido(s) · **{}** alterado(s) · "
                    u"**{}** faltando / pendente(s)".format(
                        u"modelo" if RESULT["target"] == "model" else u"Excel",
                        len(RESULT["filled"]), len(RESULT["changed"]),
                        len(RESULT["missing"])))
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
        PARAM_IN_NAMES[0], u", ".join(u"{} ({})".format(c, n)
                                      for c, n in sorted(missing_in_by_cat.items()))))

if MODE != "analyze":
    make_html = forms.alert(
        u"Gerar o relatório HTML?\n\n"
        u"{} preenchido(s) · {} alterado(s) · {} faltando / pendente(s)\n"
        u"(agrupados por categoria do modelo e por keynote)".format(
            len(RESULT["filled"]), len(RESULT["changed"]), len(RESULT["missing"])),
        yes=True, no=True)
    if not make_html:
        script.exit()

data = {
    "project": to_unicode(doc.Title),
    "generated": datetime.now().strftime("%d/%m/%Y %H:%M"),
    "mode": MODE, "modeLabel": MODES[MODE],
    "scope": (u"Vista ativa: " + to_unicode(active_view.Name)) if use_view else u"Modelo inteiro",
    "materials": opts["materials"],
    "inParam": PARAM_IN_NAMES[0],
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
    "result": RESULT,
    "changes": changes,
    "skipped": [dict((k, v) for k, v in s.items() if k != "eid") for s in skipped],
}

# o template pode se chamar script.html (repositório) ou keynotes.html
template_path = None
for _name in ("script.html", "keynotes.html"):
    if os.path.isfile(os.path.join(HERE, _name)):
        template_path = os.path.join(HERE, _name)
        break
if template_path is None:
    forms.alert(u"Modelo do relatório não encontrado (script.html) na pasta do botão:\n"
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
