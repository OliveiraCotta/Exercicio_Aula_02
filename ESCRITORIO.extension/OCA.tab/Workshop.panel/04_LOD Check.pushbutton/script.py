# -*- coding: utf-8 -*-
"""LOD Check - grade the modelled level of development of building elements.

Works on the ACTIVE VIEW. For every wall, floor, roof, ceiling and curtain
panel visible in the view it works out an achieved LOD from rules read from a
reference file (lod_rules.csv / .xlsx) kept next to this script, so the rules
stay traceable to the BIM Forum LOD spec and easy to edit. It then colours the
view (green = meets your target LOD, red = below) using VIEW-ONLY overrides -
the model is never changed - and opens an HTML dashboard.

Achieved LOD (default logic, thresholds come from the reference file):
  * 200 - the element is modelled (a generic placeholder).
  * 300 - a real assembly: the element's type has a compound structure with
          MinLayers+ layers, each carrying a material.
  * 350 - 300 plus a Fire Rating value.
"""

__title__ = "LOD\nCheck"
__doc__ = "Check walls/floors/roofs/ceilings/curtain panels against an LOD target, colour the view and open a dashboard."

import os
import json
import codecs
from datetime import datetime

from pyrevit import revit, DB, script, forms

doc = revit.doc
BIP = DB.BuiltInParameter

output = script.get_output()


# ------------------------------------------------------------------
# small helpers (mirrors the other PA Academy tools)
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


def get_id(elem):
    eid = elem.Id
    raw = eid.Value if hasattr(eid, "Value") else eid.IntegerValue
    return int(raw)


def param_display_value(p):
    """Best-effort readable string for any parameter, else empty string."""
    if p is None or not p.HasValue:
        return u""
    st = p.StorageType
    try:
        if st == DB.StorageType.String:
            return to_unicode(p.AsString() or u"")
        vs = p.AsValueString()
        if vs:
            return to_unicode(vs)
        if st == DB.StorageType.Integer:
            return to_unicode(p.AsInteger())
        if st == DB.StorageType.ElementId:
            return to_unicode(p.AsElementId().IntegerValue)
        return u""
    except Exception:
        return u""


_LEVEL_BIP_NAMES = ("LEVEL_PARAM", "SCHEDULE_LEVEL_PARAM",
                    "FAMILY_LEVEL_PARAM", "ROOM_LEVEL_ID")


def element_level_name(elem):
    """Best-effort level name for any element, else 'N/A'."""
    try:
        lid = elem.LevelId
        if lid is not None and lid.IntegerValue != -1:
            lvl = doc.GetElement(lid)
            if lvl is not None:
                return to_unicode(lvl.Name)
    except Exception:
        pass
    try:
        lvl = getattr(elem, "Level", None)
        if lvl is not None:
            return to_unicode(lvl.Name)
    except Exception:
        pass
    for bname in _LEVEL_BIP_NAMES:
        bip = getattr(BIP, bname, None)
        if bip is None:
            continue
        try:
            p = elem.get_Parameter(bip)
            if p is not None and p.HasValue:
                vs = param_display_value(p)
                if vs:
                    return vs
        except Exception:
            continue
    return u"N/A"


# ------------------------------------------------------------------
# 0. Guard: we need a graphical view we can override.
# ------------------------------------------------------------------
view = doc.ActiveView
_BAD_VIEWTYPES = set()
for _n in ("Schedule", "DrawingSheet", "Legend", "Report", "ColumnSchedule",
           "PanelSchedule", "SystemsAnalysisReport", "Internal", "Undefined"):
    _vt = getattr(DB.ViewType, _n, None)
    if _vt is not None:
        _BAD_VIEWTYPES.add(_vt)

if view is None or getattr(view, "IsTemplate", False) or view.ViewType in _BAD_VIEWTYPES:
    forms.alert(
        "Open a graphical view (plan, section, elevation, 3D or detail) and run "
        "LOD Check again.\n\nThe active view can't take element colour overrides "
        "(it's a schedule, sheet, legend or a view template).",
        title="LOD Check", exitscript=True)


# ------------------------------------------------------------------
# 1. Ask which LOD to check against.
# ------------------------------------------------------------------
LOD_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="LOD Check - Target" Height="Auto" Width="440"
        SizeToContent="Height" WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize" Background="#0E1526">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,10,0,3"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="26"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
    </Style>
  </Window.Resources>
  <StackPanel Margin="18">
    <TextBlock Text="CHECK AGAINST LOD" FontSize="15" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock Text="Pick the target level of development. Every wall, floor, roof, ceiling and curtain panel in the active view is graded against it - green if it meets the target, red if it's below. Colours are view-only overrides; your model is not changed."
               TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,0,0,6"/>

    <TextBlock Text="Target LOD"/>
    <ComboBox x:Name="lod"/>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
      <Button x:Name="cancel" Content="Cancel" Width="80" Height="28" Margin="0,0,10,0"/>
      <Button x:Name="ok" Content="Check" Width="110" Height="28"/>
    </StackPanel>
  </StackPanel>
</Window>
"""

LOD_ITEMS = [u"LOD 200 - modelled (generic)",
             u"LOD 300 - real assembly (layers + materials)",
             u"LOD 350 - assembly + fire rating"]
LOD_VALUES = {LOD_ITEMS[0]: 200, LOD_ITEMS[1]: 300, LOD_ITEMS[2]: 350}


class LODWindow(forms.WPFWindow):
    def __init__(self, xaml, default_label):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.lod.ItemsSource = list(LOD_ITEMS)
        self.lod.SelectedItem = default_label
        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    def _ok(self, sender, args):
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.confirmed = False
        self.Close()

    @property
    def selected(self):
        return self.lod.SelectedItem


win = LODWindow(LOD_XAML, LOD_ITEMS[1])   # default LOD 300
win.ShowDialog()
if not win.confirmed:
    script.exit()
target_lod = LOD_VALUES.get(win.selected or LOD_ITEMS[1], 300)


# ------------------------------------------------------------------
# 2. Read the LOD rules from the reference file (CSV or XLSX).
#    Rules are traceable to the spec (descriptions) AND drive detection
#    (MinLayers, RequireMaterial, FireRatingParam). Anything not listed
#    falls back to the sensible defaults below.
# ------------------------------------------------------------------
HERE = os.path.dirname(__file__)

# Categories we grade, in report order.
CATS = [
    (u"Walls", DB.BuiltInCategory.OST_Walls),
    (u"Floors", DB.BuiltInCategory.OST_Floors),
    (u"Roofs", DB.BuiltInCategory.OST_Roofs),
    (u"Ceilings", DB.BuiltInCategory.OST_Ceilings),
    (u"Curtain Panels", DB.BuiltInCategory.OST_CurtainWallPanels),
]

_GENERIC_DEFAULT = {
    "lod200": u"Modelled as a generic element of approximate size and location.",
    "lod300": u"Type has a compound structure of 2+ layers, each with a material.",
    "lod350": u"As LOD 300, plus a Fire Rating value.",
    "minlayers": 2, "requirematerial": True, "fireparam": u"Fire Rating",
}

DEFAULT_RULES = {}
for _cn, _ in CATS:
    r = dict(_GENERIC_DEFAULT)
    DEFAULT_RULES[_cn.lower()] = r


def _truthy(v):
    return to_unicode(v).strip().lower() in (u"yes", u"y", u"true", u"1", u"t")


def _int_or(v, default):
    try:
        return int(float(to_unicode(v).strip()))
    except Exception:
        return default


def _norm_header(h):
    return to_unicode(h).strip().lower().replace(u" ", u"").replace(u"_", u"")


# map many possible header spellings to our canonical keys
_HEADER_ALIASES = {
    u"category": "category", u"cat": "category",
    u"lod200": "lod200", u"200": "lod200",
    u"lod300": "lod300", u"300": "lod300",
    u"lod350": "lod350", u"350": "lod350",
    u"minlayers": "minlayers", u"minlayer": "minlayers", u"layers": "minlayers",
    u"requirematerial": "requirematerial",
    u"requirematerialperlayer": "requirematerial",
    u"material": "requirematerial", u"materials": "requirematerial",
    u"fireratingparam": "fireparam", u"firerating": "fireparam",
    u"fireparam": "fireparam", u"fire": "fireparam",
}


def parse_csv(text):
    """Minimal RFC-4180 CSV parser -> list of rows (each a list of unicode)."""
    rows, field, row = [], [], []
    i, n, in_q = 0, len(text), False
    cur = []
    while i < n:
        ch = text[i]
        if in_q:
            if ch == u'"':
                if i + 1 < n and text[i + 1] == u'"':
                    cur.append(u'"')
                    i += 1
                else:
                    in_q = False
            else:
                cur.append(ch)
        else:
            if ch == u'"':
                in_q = True
            elif ch == u',':
                row.append(u"".join(cur))
                cur = []
            elif ch in (u'\n', u'\r'):
                if ch == u'\r' and i + 1 < n and text[i + 1] == u'\n':
                    i += 1
                row.append(u"".join(cur))
                cur = []
                rows.append(row)
                row = []
            else:
                cur.append(ch)
        i += 1
    if cur or row:
        row.append(u"".join(cur))
        rows.append(row)
    # drop fully-empty rows
    return [r for r in rows if any(c.strip() for c in r)]


def read_csv_rows(path):
    with codecs.open(path, "r", encoding="utf-8-sig") as f:
        text = f.read()
    grid = parse_csv(text)
    if not grid:
        return []
    headers = [_norm_header(h) for h in grid[0]]
    out = []
    for r in grid[1:]:
        d = {}
        for idx, h in enumerate(headers):
            key = _HEADER_ALIASES.get(h)
            if key and idx < len(r):
                d[key] = r[idx]
        if d.get("category"):
            out.append(d)
    return out


def _split_ref(ref):
    col, num = u"", u""
    for ch in ref:
        if ch.isalpha():
            col += ch
        else:
            num += ch
    return col.upper(), (int(num) if num else 0)


def _col_index(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def read_xlsx_rows(path):
    """Read the first worksheet of an .xlsx with no external libraries
    (System.IO.Compression + System.Xml). Returns list of dict rows."""
    import clr
    clr.AddReference("System.IO.Compression")
    clr.AddReference("System.IO.Compression.FileSystem")
    clr.AddReference("System.Xml")
    from System.IO.Compression import ZipFile
    from System.IO import StreamReader
    from System.Xml import XmlDocument, XmlNamespaceManager

    NS = u"http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    za = ZipFile.OpenRead(path)

    def read_entry(name):
        e = za.GetEntry(name)
        if e is None:
            return None
        sr = StreamReader(e.Open())
        try:
            return sr.ReadToEnd()
        finally:
            sr.Close()

    try:
        shared_xml = read_entry("xl/sharedStrings.xml")
        # first worksheet is conventionally sheet1.xml
        sheet_xml = read_entry("xl/worksheets/sheet1.xml")
    finally:
        za.Dispose()

    if sheet_xml is None:
        return []

    shared = []
    if shared_xml:
        sd = XmlDocument()
        sd.LoadXml(shared_xml)
        sns = XmlNamespaceManager(sd.NameTable)
        sns.AddNamespace("a", NS)
        for si in sd.SelectNodes("//a:si", sns):
            parts = si.SelectNodes(".//a:t", sns)
            shared.append(u"".join(to_unicode(t.InnerText) for t in parts))

    wd = XmlDocument()
    wd.LoadXml(sheet_xml)
    wns = XmlNamespaceManager(wd.NameTable)
    wns.AddNamespace("a", NS)

    grid = {}     # rownum -> {colindex: value}
    for row in wd.SelectNodes("//a:row", wns):
        for c in row.SelectNodes("a:c", wns):
            ref = to_unicode(c.GetAttribute("r"))
            col, rnum = _split_ref(ref)
            if not col:
                continue
            t = to_unicode(c.GetAttribute("t"))
            val = u""
            if t == u"s":
                vnode = c.SelectSingleNode("a:v", wns)
                if vnode is not None:
                    try:
                        val = shared[int(vnode.InnerText)]
                    except Exception:
                        val = u""
            elif t == u"inlineStr":
                isn = c.SelectSingleNode("a:is", wns)
                if isn is not None:
                    parts = isn.SelectNodes(".//a:t", wns)
                    val = u"".join(to_unicode(p.InnerText) for p in parts)
            else:
                vnode = c.SelectSingleNode("a:v", wns)
                if vnode is not None:
                    val = to_unicode(vnode.InnerText)
            grid.setdefault(rnum, {})[_col_index(col)] = val

    if not grid:
        return []
    row_nums = sorted(grid.keys())
    header_row = grid[row_nums[0]]
    headers = {}      # colindex -> canonical key
    for ci, raw in header_row.items():
        key = _HEADER_ALIASES.get(_norm_header(raw))
        if key:
            headers[ci] = key
    out = []
    for rn in row_nums[1:]:
        cells = grid[rn]
        d = {}
        for ci, key in headers.items():
            if ci in cells:
                d[key] = cells[ci]
        if d.get("category"):
            out.append(d)
    return out


def find_rules_file(folder):
    """Pick the reference file: prefer names containing 'lod', then most
    recently modified, among .csv/.xlsx in the button folder."""
    cands = []
    try:
        for fn in os.listdir(folder):
            low = fn.lower()
            if low.endswith(".csv") or (low.endswith(".xlsx") and not low.startswith("~$")):
                full = os.path.join(folder, fn)
                try:
                    mtime = os.path.getmtime(full)
                except Exception:
                    mtime = 0
                pref = 1 if "lod" in low else 0
                cands.append((pref, mtime, full))
    except Exception:
        return None
    if not cands:
        return None
    cands.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return cands[0][2]


rules_path = find_rules_file(HERE)
rules_kind = u"default"
file_rows = []
rules_error = u""
if rules_path:
    try:
        if rules_path.lower().endswith(".xlsx"):
            file_rows = read_xlsx_rows(rules_path)
            rules_kind = u"xlsx"
        else:
            file_rows = read_csv_rows(rules_path)
            rules_kind = u"csv"
    except Exception as ex:
        rules_error = to_unicode(str(ex))
        file_rows = []
        rules_kind = u"default"

# index file rows by lowercased category name
file_by_cat = {}
for r in file_rows:
    cat = to_unicode(r.get("category")).strip()
    if cat:
        file_by_cat[cat.lower()] = r


def resolve_rule(cat_name):
    """Merge file row (if any) over the category default; report the source."""
    base = dict(DEFAULT_RULES.get(cat_name.lower(), _GENERIC_DEFAULT))
    row = file_by_cat.get(cat_name.lower())
    src = u"default"
    if row:
        src = u"file"
        if to_unicode(row.get("lod200")).strip():
            base["lod200"] = to_unicode(row["lod200"]).strip()
        if to_unicode(row.get("lod300")).strip():
            base["lod300"] = to_unicode(row["lod300"]).strip()
        if to_unicode(row.get("lod350")).strip():
            base["lod350"] = to_unicode(row["lod350"]).strip()
        if "minlayers" in row and to_unicode(row["minlayers"]).strip():
            base["minlayers"] = _int_or(row["minlayers"], base["minlayers"])
        if "requirematerial" in row and to_unicode(row["requirematerial"]).strip():
            base["requirematerial"] = _truthy(row["requirematerial"])
        if to_unicode(row.get("fireparam")).strip():
            base["fireparam"] = to_unicode(row["fireparam"]).strip()
    base["source"] = src
    base["category"] = cat_name
    return base


RULES = {cn: resolve_rule(cn) for cn, _ in CATS}


# ------------------------------------------------------------------
# 3. LOD detection helpers.
# ------------------------------------------------------------------
def get_type(elem):
    try:
        tid = elem.GetTypeId()
        if tid is not None and tid.IntegerValue > 0:
            return doc.GetElement(tid)
    except Exception:
        pass
    return None


def analyze_compound(etype):
    """Return (layer_count, layers_with_material) for a type's compound
    structure, or (0, 0) if it has none."""
    if etype is None:
        return 0, 0
    getcs = getattr(etype, "GetCompoundStructure", None)
    if getcs is None:
        return 0, 0
    try:
        cs = getcs()
    except Exception:
        cs = None
    if cs is None:
        return 0, 0
    try:
        n = cs.LayerCount
    except Exception:
        return 0, 0
    with_mat = 0
    for i in range(n):
        try:
            mid = cs.GetMaterialId(i)
            if mid is not None and mid.IntegerValue > 0:
                with_mat += 1
        except Exception:
            pass
    return n, with_mat


def get_fire_rating(elem, etype, pname):
    """Fire Rating value: named param on type then instance, else BIP.FIRE_RATING."""
    for holder in (etype, elem):
        if holder is None:
            continue
        try:
            p = holder.LookupParameter(pname)
            if p is not None and p.HasValue:
                v = param_display_value(p)
                if v.strip():
                    return v
        except Exception:
            pass
    bip = getattr(BIP, "FIRE_RATING", None)
    if bip is not None:
        for holder in (etype, elem):
            if holder is None:
                continue
            try:
                p = holder.get_Parameter(bip)
                if p is not None and p.HasValue:
                    v = param_display_value(p)
                    if v.strip():
                        return v
            except Exception:
                pass
    return u""


def grade_element(elem, rule):
    """Return (achieved_lod, has_assembly, fire_val, nlayers, nmat)."""
    etype = get_type(elem)
    nlayers, nmat = analyze_compound(etype)
    min_layers = rule["minlayers"]
    require_mat = rule["requirematerial"]
    has_assembly = (nlayers >= min_layers and
                    ((nmat >= nlayers) if require_mat else True))
    fire_val = get_fire_rating(elem, etype, rule["fireparam"])
    lod = 200
    if has_assembly:
        lod = 300
        if fire_val.strip():
            lod = 350
    return lod, has_assembly, fire_val, nlayers, nmat, etype


def missing_for_target(target, rule, has_assembly, fire_val, nlayers, nmat):
    """Human-readable list of what stops this element reaching the target."""
    m = []
    min_layers = rule["minlayers"]
    require_mat = rule["requirematerial"]
    if target >= 300 and not has_assembly:
        if nlayers == 0:
            m.append(u"Generic type - no compound structure (needs {}+ layers)".format(min_layers))
        elif nlayers < min_layers:
            m.append(u"Only {} layer(s) - needs {}+".format(nlayers, min_layers))
        elif require_mat and nmat < nlayers:
            m.append(u"{} of {} layers have no material".format(nlayers - nmat, nlayers))
    if target >= 350 and not fire_val.strip():
        m.append(u"Missing Fire Rating")
    return m


# ------------------------------------------------------------------
# 4. Collect + grade every element of interest in the active view.
# ------------------------------------------------------------------
records = []                 # per-element dashboard records
green_ids, red_ids = [], []  # ElementIds for the view overrides

for cat_name, bic in CATS:
    rule = RULES[cat_name]
    try:
        collected = (DB.FilteredElementCollector(doc, view.Id)
                     .OfCategory(bic)
                     .WhereElementIsNotElementType()
                     .ToElements())
    except Exception:
        collected = []
    for elem in collected:
        try:
            lod, has_assembly, fire_val, nlayers, nmat, etype = grade_element(elem, rule)
        except Exception:
            continue
        meets = lod >= target_lod
        miss = [] if meets else missing_for_target(
            target_lod, rule, has_assembly, fire_val, nlayers, nmat)
        try:
            tname = to_unicode(etype.Name) if etype is not None else to_unicode(elem.Name)
        except Exception:
            tname = u"(type)"
        records.append({
            "id": get_id(elem),
            "category": cat_name,
            "type": tname,
            "level": element_level_name(elem),
            "lod": lod,
            "meets": bool(meets),
            "layers": nlayers,
            "layersWithMaterial": nmat,
            "fire": fire_val,
            "missing": miss,
        })
        (green_ids if meets else red_ids).append(elem.Id)

if not records:
    forms.alert(
        "No walls, floors, roofs, ceilings or curtain panels are visible in "
        "the active view.\n\nOpen a view that shows some of these elements and "
        "run LOD Check again.",
        title="LOD Check", exitscript=True)


# ------------------------------------------------------------------
# 5. Colour the view - GREEN meets target, RED below. View-only overrides.
# ------------------------------------------------------------------
GREEN = DB.Color(80, 214, 130)
RED = DB.Color(232, 74, 74)

# a solid drafting fill pattern so the surface/cut reads as a solid colour
solid_fill_id = DB.ElementId.InvalidElementId
try:
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                solid_fill_id = fp.Id
                break
        except Exception:
            continue
except Exception:
    pass


def make_ogs(color):
    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(color)
    ogs.SetCutLineColor(color)
    if solid_fill_id != DB.ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternId(solid_fill_id)
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetCutForegroundPatternId(solid_fill_id)
        ogs.SetCutForegroundPatternColor(color)
    return ogs


ogs_green = make_ogs(GREEN)
ogs_red = make_ogs(RED)

colored = 0
t = DB.Transaction(doc, "LOD Check - colour view (view-only)")
t.Start()
try:
    for eid in green_ids:
        try:
            view.SetElementOverrides(eid, ogs_green)
            colored += 1
        except Exception:
            pass
    for eid in red_ids:
        try:
            view.SetElementOverrides(eid, ogs_red)
            colored += 1
        except Exception:
            pass
    t.Commit()
except Exception as ex:
    t.RollBack()
    forms.alert("Could not apply view colours (rolled back):\n{}".format(ex))


# ------------------------------------------------------------------
# 6. Build the dashboard payload.
# ------------------------------------------------------------------
total = len(records)
at_target = sum(1 for r in records if r["meets"])
below = total - at_target

cat_summary = []
for cat_name, _ in CATS:
    rows = [r for r in records if r["category"] == cat_name]
    if not rows:
        continue
    ct_at = sum(1 for r in rows if r["meets"])
    cat_summary.append({
        "category": cat_name,
        "total": len(rows),
        "atTarget": ct_at,
        "below": len(rows) - ct_at,
    })

rules_out = []
for cat_name, _ in CATS:
    r = RULES[cat_name]
    rules_out.append({
        "category": cat_name,
        "lod200": r["lod200"],
        "lod300": r["lod300"],
        "lod350": r["lod350"],
        "minLayers": r["minlayers"],
        "requireMaterial": bool(r["requirematerial"]),
        "fireParam": r["fireparam"],
        "source": r["source"],
    })

below_rows = sorted(
    [r for r in records if not r["meets"]],
    key=lambda r: (r["category"], r["lod"], r["type"]))

data = {
    "project": to_unicode(doc.Title),
    "view": to_unicode(view.Name),
    "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "target": target_lod,
    "summary": {"total": total, "atTarget": at_target, "below": below},
    "categories": cat_summary,
    "rules": rules_out,
    "below": below_rows,
    "rulesSource": {
        "kind": rules_kind,
        "path": to_unicode(rules_path) if rules_path else u"",
        "name": to_unicode(os.path.basename(rules_path)) if rules_path else u"",
        "error": rules_error,
    },
}


# ------------------------------------------------------------------
# 7. Inject into template + open.
# ------------------------------------------------------------------
template_path = os.path.join(HERE, "lod_check.html")
with codecs.open(template_path, "r", encoding="utf-8") as f:
    html = f.read()

html = html.replace("__DATA__", json.dumps(data, ensure_ascii=False,
                                           default=lambda o: int(o)))

out_path = script.get_document_data_file("lod_check", "html")
with codecs.open(out_path, "w", encoding="utf-8") as f:
    f.write(html)


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


src_label = {"csv": "reference CSV", "xlsx": "reference Excel file",
             "default": "built-in defaults"}.get(rules_kind, "defaults")
output.print_md("## LOD Check - LOD {} target".format(target_lod))
output.print_md("- **{}** elements graded · **{}** meet target · **{}** below "
                "· **{}** coloured in *{}*".format(
                    total, at_target, below, colored, to_unicode(view.Name)))
output.print_md("- Rules from **{}**{}".format(
    src_label,
    " (`{}`)".format(to_unicode(os.path.basename(rules_path))) if rules_path else ""))
if rules_error:
    output.print_md(":warning: Could not read the reference file, used defaults: `{}`".format(rules_error))
output.print_md("- Colours are **view-only overrides** - press **Ctrl+Z** (undo) to clear them; the model was not changed.")

if not open_in_browser(out_path):
    forms.alert("LOD dashboard generated but could not open automatically.\n\n{}".format(out_path))
