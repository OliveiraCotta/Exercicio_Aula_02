# -*- coding: utf-8 -*-
"""Element QA/QC - flag elements with empty values for chosen parameters.

Works on ANY model category (Rooms, Doors, Walls, ...): on launch you pick a
category, then the report checks that category's elements for missing values.
Rooms keep their familiar Number / Name / Level identity columns; every other
category uses Mark / Name / Level.
"""

__title__ = "Element\nQA/QC"
__doc__ = "Pick a category, then find elements with missing parameter values."

import os
import json
import codecs
from datetime import datetime

from pyrevit import revit, DB, script, forms

doc = revit.doc

BIP = DB.BuiltInParameter


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
        # ElementId / Integer / Double: use value-string which respects units
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


# ------------------------------------------------------------------
# 1. Let the user pick which category to QA/QC.
#    We tally placed model elements once and offer only categories that
#    actually have something to check (with a live count).
# ------------------------------------------------------------------
def _bic_int(bic):
    try:
        return int(bic)
    except Exception:
        return None


ROOMS_INT = _bic_int(DB.BuiltInCategory.OST_Rooms)

cat_counts = {}   # category name -> [Category, count]
for el in DB.FilteredElementCollector(doc).WhereElementIsNotElementType():
    try:
        cat = el.Category
        if cat is None or cat.CategoryType != DB.CategoryType.Model:
            continue
        cname = to_unicode(cat.Name)
        if not cname:
            continue
        if cname not in cat_counts:
            cat_counts[cname] = [cat, 0]
        cat_counts[cname][1] += 1
    except Exception:
        continue

if not cat_counts:
    forms.alert("No placed model elements found in this project.", exitscript=True)

# Build a searchable list of "Name  (count)" labels, Rooms first if present.
def _sort_key(name):
    is_room = 0 if name.lower() == "rooms" else 1
    return (is_room, name.lower())

names_sorted = sorted(cat_counts.keys(), key=_sort_key)
label_to_name = {}
labels = []
for n in names_sorted:
    lbl = u"{}   ({})".format(n, cat_counts[n][1])
    label_to_name[lbl] = n
    labels.append(lbl)

# Styled category picker - matches the CV house dialog theme
# (navy #0E1526 background, cyan #65E3FF heading, dimmed #7A8FA9 helper text,
#  light #CFE3FF labels) so every pop-up across the tools looks the same.
CATEGORY_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Element QA/QC - Category" Height="Auto" Width="440"
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
    <TextBlock Text="CHOOSE A CATEGORY" FontSize="15" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock Text="Pick which category to QA/QC. The report checks that category's elements for missing parameter values. The number in brackets is how many placed elements were found."
               TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,0,0,6"/>

    <TextBlock Text="Category"/>
    <ComboBox x:Name="cat"/>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
      <Button x:Name="cancel" Content="Cancel" Width="80" Height="28" Margin="0,0,10,0"/>
      <Button x:Name="ok" Content="Check" Width="110" Height="28"/>
    </StackPanel>
  </StackPanel>
</Window>
"""


class CategoryWindow(forms.WPFWindow):
    def __init__(self, xaml, cat_labels, default_label):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.cat.ItemsSource = list(cat_labels)
        self.cat.SelectedItem = default_label
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
        return self.cat.SelectedItem


# names_sorted puts Rooms first when present, so labels[0] is the default
win = CategoryWindow(CATEGORY_XAML, labels, labels[0])
win.ShowDialog()
if not win.confirmed:
    script.exit()

chosen_label = win.selected or labels[0]
category_name = label_to_name.get(chosen_label, names_sorted[0])
chosen_cat = cat_counts[category_name][0]

try:
    is_rooms = (ROOMS_INT is not None
                and chosen_cat.Id.IntegerValue == ROOMS_INT)
except Exception:
    is_rooms = (category_name.lower() == "rooms")


# ------------------------------------------------------------------
# 2. Identity + level helpers (generic across categories)
# ------------------------------------------------------------------
def element_identity(elem):
    """Return (number/mark, name) for the element."""
    if is_rooms:
        number = param_display_value(elem.get_Parameter(BIP.ROOM_NUMBER))
        name = param_display_value(elem.get_Parameter(BIP.ROOM_NAME))
        return number, name
    # generic: Mark as the "number", Element.Name as the name
    mark_p = elem.get_Parameter(BIP.ALL_MODEL_MARK)
    number = param_display_value(mark_p) if mark_p is not None else u""
    try:
        name = to_unicode(elem.Name)
    except Exception:
        name = u""
    return number, name


_LEVEL_BIP_NAMES = ("LEVEL_PARAM", "SCHEDULE_LEVEL_PARAM",
                    "FAMILY_LEVEL_PARAM", "ROOM_LEVEL_ID")


def element_level_name(elem):
    """Best-effort level name for any element, else 'N/A'."""
    # 1. LevelId property (most instances)
    try:
        lid = elem.LevelId
        if lid is not None and lid.IntegerValue != -1:
            lvl = doc.GetElement(lid)
            if lvl is not None:
                return to_unicode(lvl.Name)
    except Exception:
        pass
    # 2. .Level property (rooms, some hosted elements)
    try:
        lvl = getattr(elem, "Level", None)
        if lvl is not None:
            return to_unicode(lvl.Name)
    except Exception:
        pass
    # 3. common level built-in parameters
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
# 3. Collect the chosen category's elements + every parameter value
#    (so the browser can check any parameter).
# ------------------------------------------------------------------
elements = (
    DB.FilteredElementCollector(doc)
    .OfCategoryId(chosen_cat.Id)
    .WhereElementIsNotElementType()
    .ToElements()
)

element_records = []
all_param_names = set()

for elem in elements:
    # rooms: skip unplaced/redundant rooms (no area) - they'd flag everything
    if is_rooms:
        try:
            if elem.Area <= 0:
                continue
        except Exception:
            pass

    number, name = element_identity(elem)
    level = element_level_name(elem)

    params = {}
    for p in elem.Parameters:
        try:
            if p.Definition is None:
                continue
            pname = to_unicode(p.Definition.Name)
            if not pname:
                continue
            params[pname] = param_display_value(p)
            all_param_names.add(pname)
        except Exception:
            continue

    element_records.append({
        "id": get_id(elem),
        "number": number,
        "name": name,
        "level": level,
        "params": params,
    })

if not element_records:
    forms.alert("No checkable '{}' elements found in this project.".format(
        category_name), exitscript=True)


# ------------------------------------------------------------------
# 4. Load saved checks (per category) if the user has saved a set.
#    We look in a stable app-data folder first, then the user's Downloads
#    (so a freshly-downloaded file "just works" without moving it).
#    Each category keeps its own file so, e.g., Door checks never leak
#    into the Rooms view.
# ------------------------------------------------------------------
DOWNLOAD_FILENAME = "qaqc_checks.json"   # fixed name the browser downloads


def _safe_name(s):
    keep = u"".join(ch if ch.isalnum() else u"_" for ch in to_unicode(s))
    return keep or u"category"


data_file = script.get_document_data_file("element_qaqc", "html")
appdata_dir = os.path.dirname(data_file)


def _checks_path_for(cat_name):
    return os.path.join(appdata_dir,
                        u"qaqc_checks__{}.json".format(_safe_name(cat_name)))


def _downloads_path():
    try:
        from System import Environment
        profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        return os.path.join(profile, "Downloads", DOWNLOAD_FILENAME)
    except Exception:
        return None


appdata_file = _checks_path_for(category_name)
downloads_file = _downloads_path()

import json as _json


def _safe_remove(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


# If a fresh file sits in Downloads, process it. A "reset" file (reset flag or
# empty checks) clears that category's saved set; otherwise it's adopted as the
# persistent copy for whatever category it names (falling back to the current).
try:
    if downloads_file and os.path.exists(downloads_file):
        with codecs.open(downloads_file, "r", encoding="utf-8") as f:
            payload = _json.loads(f.read())
        target_cat = to_unicode(payload.get("category") or category_name)
        target_file = _checks_path_for(target_cat)
        is_reset = bool(payload.get("reset")) or (
            isinstance(payload.get("checks"), list) and len(payload.get("checks")) == 0)
        if is_reset:
            _safe_remove(target_file)
            _safe_remove(downloads_file)
        elif isinstance(payload.get("checks"), list):
            with codecs.open(target_file, "w", encoding="utf-8") as f:
                f.write(_json.dumps(payload, ensure_ascii=False))
            _safe_remove(downloads_file)   # consume so it isn't re-adopted later
except Exception:
    pass

# Now read the persistent copy for the CURRENT category
saved_checks = []
try:
    if os.path.exists(appdata_file):
        with codecs.open(appdata_file, "r", encoding="utf-8") as f:
            payload = _json.loads(f.read())
        if payload.get("reset"):
            _safe_remove(appdata_file)
            saved_checks = []
        else:
            c = payload.get("checks")
            if isinstance(c, list):
                saved_checks = [to_unicode(x) for x in c if to_unicode(x).strip()]
except Exception:
    saved_checks = []

# Sensible defaults per category. Rooms carry room data (Department, Occupancy);
# walls carry Fire Rating (a wall/door property, not a room one). Anything else
# falls back to the near-universal "Comments".
if is_rooms:
    default_checks = ["Department", "Occupancy"]
elif category_name == "Walls":
    default_checks = ["Fire Rating"]
else:
    default_checks = ["Comments"]

data = {
    "project": to_unicode(doc.Title),
    "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "category": category_name,
    "isRooms": bool(is_rooms),
    # identity column header + export id column name
    "idLabel": "Number" if is_rooms else "Mark",
    "idColumn": "RoomID" if is_rooms else "ElementID",
    "elements": element_records,
    "paramNames": sorted(all_param_names, key=lambda s: s.lower()),
    # default checks; browser matches these case-insensitively and will
    # simply report "not present" if this category doesn't have one of them
    "defaultChecks": default_checks,
    # saved set for this category (empty if none) - browser prefers this
    "savedChecks": saved_checks,
    # where to drop a downloaded qaqc_checks.json for it to persist
    "checksFolder": appdata_dir,
}

# ------------------------------------------------------------------
# 5. Inject into template + open
# ------------------------------------------------------------------
template_path = os.path.join(os.path.dirname(__file__), "qaqc.html")
with codecs.open(template_path, "r", encoding="utf-8") as f:
    html = f.read()

html = html.replace("__DATA__", json.dumps(data, ensure_ascii=False,
                                            default=lambda o: int(o)))

out_path = script.get_document_data_file("element_qaqc", "html")
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


if not open_in_browser(out_path):
    forms.alert("QA/QC report generated but could not open automatically.\n\n{}".format(out_path))
