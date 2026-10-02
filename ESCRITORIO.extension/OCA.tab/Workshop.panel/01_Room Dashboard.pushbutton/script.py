# -*- coding: utf-8 -*-
"""Room Dashboard - opens a rich HTML room analytics dashboard in the browser."""

__title__ = "Room\nDashboard"
__doc__ = "Collects room data and opens a futuristic HTML dashboard."

import os
import json
import codecs
from datetime import datetime

from pyrevit import revit, DB, script, forms

doc = revit.doc


def to_unicode(value):
    """Coerce any Revit / .NET / byte value into clean unicode text.

    IronPython can hand back byte strings encoded in the Windows code page
    (cp1252) for parameters that contain characters like Ue, oe, e-acute, etc.
    json.dumps then chokes when it tries to escape them, so we normalise
    everything to real unicode up front.
    """
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


output = script.get_output()
output.set_title("Room Dashboard - debug log")

output.print_md("**[1/5] Script started.** Engine OK, document: `{}`".format(doc.Title))

# ------------------------------------------------------------------
# 1. Collect levels
# ------------------------------------------------------------------
levels = (
    DB.FilteredElementCollector(doc)
    .OfClass(DB.Level)
    .ToElements()
)
level_elev = {}
for lv in levels:
    level_elev[to_unicode(lv.Name)] = float(DB.UnitUtils.ConvertFromInternalUnits(
        lv.Elevation, DB.UnitTypeId.Meters
    ))

# ------------------------------------------------------------------
# 2. Collect rooms
# ------------------------------------------------------------------
rooms = (
    DB.FilteredElementCollector(doc)
    .OfCategory(DB.BuiltInCategory.OST_Rooms)
    .WhereElementIsNotElementType()
    .ToElements()
)


def get_param(room, bip):
    p = room.get_Parameter(bip)
    if p and p.HasValue:
        return to_unicode(p.AsString() or p.AsValueString() or "")
    return u""


def to_sqm(internal_area):
    return float(DB.UnitUtils.ConvertFromInternalUnits(
        internal_area, DB.UnitTypeId.SquareMeters
    ))


def get_id(elem):
    eid = elem.Id
    raw = eid.Value if hasattr(eid, "Value") else eid.IntegerValue
    return int(raw)


def get_value(room, pname):
    """Read any room parameter by its display name (handles text + key params)."""
    if not pname:
        return u""
    p = room.LookupParameter(pname)
    if not p or not p.HasValue:
        return u""
    if p.StorageType == DB.StorageType.String:
        return to_unicode(p.AsString() or u"")
    return to_unicode(p.AsValueString() or u"")


def collect_text_param_names(room_elems):
    """Union of text-like parameter names across all rooms, for the picker."""
    names = set()
    text_types = (DB.StorageType.String, DB.StorageType.ElementId)
    for r in room_elems:
        for p in r.Parameters:
            try:
                if p.StorageType in text_types and p.Definition is not None:
                    nm = p.Definition.Name
                    if nm:
                        names.add(to_unicode(nm))
            except Exception:
                continue
    return sorted(names, key=lambda s: s.lower())


def collect_numeric_param_names(room_elems):
    """Union of numeric (Double) parameter names - candidates for a brief area."""
    names = set()
    for r in room_elems:
        for p in r.Parameters:
            try:
                if p.StorageType == DB.StorageType.Double and p.Definition is not None:
                    nm = p.Definition.Name
                    if nm:
                        names.add(to_unicode(nm))
            except Exception:
                continue
    return sorted(names, key=lambda s: s.lower())


def get_area_value(room, pname):
    """Read a brief/target AREA parameter and return m2.

    Assumes an Area-type parameter (internal units -> m2). If it's a plain
    number parameter already in m2 the value may need no conversion, but
    Area-type is the common case for a brief/programme target.
    """
    if not pname:
        return 0.0
    p = room.LookupParameter(pname)
    if not p or not p.HasValue:
        return 0.0
    try:
        return round(to_sqm(p.AsDouble()), 2)
    except Exception:
        return 0.0


# ------------------------------------------------------------------
# 2b. Ask which parameters map to Category / Sub / Sub-sub
# ------------------------------------------------------------------
NONE_ITEM = u"\u2014  (none / stop here)  \u2014"
CAT_KEYS = ["cat1", "cat2", "cat3"]
CAT_LABELS = ["Category", "Sub-category", "Sub-sub-category"]

param_names = collect_text_param_names(rooms)
area_names = collect_numeric_param_names(rooms)
mapping = {"cat1": None, "cat2": None, "cat3": None}
brief_param = None

config = script.get_config()
saved = config.get_option("cat_map", [])            # previously chosen names
saved_brief = config.get_option("brief_param", u"")  # previously chosen brief param

# ------------------------------------------------------------------
# 2b + 2c. One combined mapping panel (three category dropdowns +
#          optional brief-area dropdown), instead of 4 pop-ups.
# ------------------------------------------------------------------
MAP_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Room Dashboard - Mapping" Height="Auto" Width="440"
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
    <TextBlock Text="MAP ROOM PARAMETERS" FontSize="15" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock Text="Choose which room parameters drive the category hierarchy. Leave a row on (none) to stop the hierarchy there."
               TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,0,0,6"/>

    <TextBlock Text="Category"/>
    <ComboBox x:Name="cat1"/>
    <TextBlock Text="Sub-category"/>
    <ComboBox x:Name="cat2"/>
    <TextBlock Text="Sub-sub-category"/>
    <ComboBox x:Name="cat3"/>

    <TextBlock Text="Brief / target AREA parameter (optional)" Margin="0,16,0,3"/>
    <ComboBox x:Name="briefp"/>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
      <Button x:Name="cancel" Content="Cancel" Width="80" Height="28" Margin="0,0,10,0"/>
      <Button x:Name="ok" Content="Generate" Width="110" Height="28"/>
    </StackPanel>
  </StackPanel>
</Window>
"""


class MappingWindow(forms.WPFWindow):
    def __init__(self, xaml, text_names, area_names, saved_cats, saved_brief):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False

        cat_items = [NONE_ITEM] + list(text_names)
        for idx, box in enumerate((self.cat1, self.cat2, self.cat3)):
            box.ItemsSource = list(cat_items)
            pref = saved_cats[idx] if idx < len(saved_cats) else u""
            box.SelectedItem = pref if pref in text_names else NONE_ITEM

        brief_items = [NONE_ITEM] + list(area_names)
        self.briefp.ItemsSource = brief_items
        self.briefp.SelectedItem = (saved_brief
                                    if saved_brief in area_names else NONE_ITEM)
        if not area_names:
            self.briefp.IsEnabled = False

        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    def _ok(self, sender, args):
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.confirmed = False
        self.Close()

    @staticmethod
    def _val(item):
        return None if (item is None or item == NONE_ITEM) else item

    @property
    def result(self):
        return (
            [self._val(self.cat1.SelectedItem),
             self._val(self.cat2.SelectedItem),
             self._val(self.cat3.SelectedItem)],
            self._val(self.briefp.SelectedItem),
        )


if param_names or area_names:
    win = MappingWindow(MAP_XAML, param_names, area_names, saved, saved_brief)
    win.ShowDialog()
    if not win.confirmed:
        script.exit()   # user cancelled
    cats, brief_param = win.result

    # collapse gaps so cat2 can't be set while cat1 is empty, and drop dupes
    picked, seen = [], set()
    for c in cats:
        if c and c not in seen:
            picked.append(c)
            seen.add(c)
    for i, key in enumerate(CAT_KEYS):
        mapping[key] = picked[i] if i < len(picked) else None

# remember this run's choices
config.cat_map = [mapping[k] or u"" for k in CAT_KEYS]
config.brief_param = brief_param or u""
script.save_config()

output.print_md("**[2b/5] Category mapping:** {}".format(
    " / ".join([mapping[k] for k in CAT_KEYS if mapping[k]]) or "(none - grouping by level)"))
output.print_md("**[2c/5] Brief area parameter:** {}".format(
    brief_param or "(none - use the dashboard's CSV import instead)"))


room_list = []
unplaced_list = []

for room in rooms:
    name = get_param(room, DB.BuiltInParameter.ROOM_NAME)
    number = get_param(room, DB.BuiltInParameter.ROOM_NUMBER)
    level = to_unicode(room.Level.Name) if room.Level else u"N/A"

    if room.Area > 0:
        room_list.append({
            "id": get_id(room),
            "number": number,
            "name": name,
            "level": level,
            "cat1": get_value(room, mapping["cat1"]),
            "cat2": get_value(room, mapping["cat2"]),
            "cat3": get_value(room, mapping["cat3"]),
            "brief": get_area_value(room, brief_param),
            "area": round(to_sqm(room.Area), 2),
        })
    else:
        unplaced_list.append({
            "number": number, "name": name, "level": level,
        })

output.print_md("**[2/5] Collected:** {} placed rooms, {} unplaced, {} levels".format(
    len(room_list), len(unplaced_list), len(level_elev)))

if not room_list and not unplaced_list:
    forms.alert("No rooms found in this project.", exitscript=True)

data = {
    "project": to_unicode(doc.Title),
    "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "mapping": {
        "cat1": mapping["cat1"] or u"",
        "cat2": mapping["cat2"] or u"",
        "cat3": mapping["cat3"] or u"",
        "briefParam": brief_param or u"",
    },
    "rooms": room_list,
    "unplaced": unplaced_list,
    "levels": sorted(
        [{"name": n, "elevation": e} for n, e in level_elev.items()],
        key=lambda x: x["elevation"],
    ),
}

# ------------------------------------------------------------------
# 3. Inject data into the HTML template
# ------------------------------------------------------------------
template_path = os.path.join(os.path.dirname(__file__), "dashboard.html")
output.print_md("**[3/5] Template path:** `{}` - exists: **{}**".format(
    template_path, os.path.exists(template_path)))

with codecs.open(template_path, "r", encoding="utf-8") as f:
    html = f.read()

html = html.replace("__DATA__", json.dumps(
    data, ensure_ascii=False, default=lambda o: int(o)))

out_path = script.get_document_data_file("room_dashboard", "html")
with codecs.open(out_path, "w", encoding="utf-8") as f:
    f.write(html)

exists = os.path.exists(out_path)
size = os.path.getsize(out_path) if exists else 0
output.print_md("**[4/5] Dashboard written:** `{}` - exists: **{}**, size: **{} KB**".format(
    out_path, exists, size // 1024))
output.print_md("Open manually if needed: {}".format(
    output.linkify_path(out_path) if hasattr(output, "linkify_path")
    else "file:///" + out_path.replace("\\", "/")))

# ------------------------------------------------------------------
# 4. Open in the default browser - log every attempt
# ------------------------------------------------------------------
attempts = []

def try_open(label, fn):
    try:
        fn()
        attempts.append((label, "OK", ""))
        return True
    except Exception as e:
        attempts.append((label, "FAILED", str(e)))
        return False


def m1():
    from System.Diagnostics import Process, ProcessStartInfo
    psi = ProcessStartInfo(out_path)
    psi.UseShellExecute = True
    Process.Start(psi)

def m2():
    from System.Diagnostics import Process
    Process.Start("explorer.exe", '"{}"'.format(out_path))

def m3():
    import webbrowser
    if not webbrowser.open("file:///" + out_path.replace("\\", "/")):
        raise Exception("webbrowser.open returned False")


opened = try_open("ProcessStartInfo + ShellExecute", m1) \
      or try_open("explorer.exe", m2) \
      or try_open("webbrowser module", m3)

output.print_md("**[5/5] Open attempts:**")
for label, status, err in attempts:
    output.print_md("- `{}` → **{}** {}".format(label, status, err))

if not opened:
    output.print_md(":warning: **All open methods failed.** "
                    "Use the manual link above.")
