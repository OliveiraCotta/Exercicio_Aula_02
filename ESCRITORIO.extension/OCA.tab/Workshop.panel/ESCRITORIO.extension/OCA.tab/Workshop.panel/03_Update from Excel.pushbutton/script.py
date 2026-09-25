# -*- coding: utf-8 -*-
"""Update from Excel - write filled QA/QC values back into element parameters."""

__title__ = "Update\nfrom Excel"
__doc__ = ("Reads a filled QA/QC Excel export (any category) and writes the "
           "values back into the matching element parameters. Handles text and "
           "numeric/area parameters (numbers are read in the parameter's display "
           "units, e.g. m2, and converted to Revit internal units). Blanks are "
           "filled and differing values are overwritten (with a preview before any "
           "overwrite). Rows are matched on the RoomID or ElementID column.")

import os

from pyrevit import revit, DB, script, forms

BIP = DB.BuiltInParameter

import clr
clr.AddReference("System.Xml")
clr.AddReference("System.IO.Compression")
clr.AddReference("System.IO.Compression.FileSystem")
from System.IO.Compression import ZipFile
from System.Xml import XmlDocument

doc = revit.doc
output = script.get_output()
output.set_title("Update from Excel")

# Columns that are identity/among the export but are never writable parameters.
# "number" (rooms) and "mark" (other categories) are identity context columns.
NON_PARAM_COLS = set(["roomid", "elementid", "number", "mark",
                      "name", "level", "missing count"])


def to_unicode(value):
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


# ------------------------------------------------------------------
# 1. Pick the file
# ------------------------------------------------------------------
xlsx_path = forms.pick_file(file_ext="xlsx",
                            title="Select the filled QA/QC Excel file")
if not xlsx_path:
    script.exit()


# ------------------------------------------------------------------
# 2. Read the xlsx (ZIP of XML) with .NET - no pandas/openpyxl needed
# ------------------------------------------------------------------
def col_to_index(cell_ref):
    """'B7' -> 1 (zero-based column)."""
    letters = ""
    for ch in cell_ref:
        if ch.isalpha():
            letters += ch
        else:
            break
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch.upper()) - 64)
    return idx - 1


def read_xlsx(path):
    """Return a list of rows; each row is a list of unicode cell strings."""
    archive = ZipFile.OpenRead(path)
    try:
        entries = {}
        for e in archive.Entries:
            entries[e.FullName.replace("\\", "/")] = e

        def read_entry(name):
            e = entries.get(name)
            if e is None:
                return None
            stream = e.Open()
            try:
                from System.IO import StreamReader
                from System.Text import Encoding
                reader = StreamReader(stream, Encoding.UTF8)
                try:
                    return reader.ReadToEnd()
                finally:
                    reader.Dispose()
            finally:
                stream.Dispose()

        # shared strings
        shared = []
        ss_xml = read_entry("xl/sharedStrings.xml")
        if ss_xml:
            d = XmlDocument()
            d.LoadXml(ss_xml)
            for si in d.GetElementsByTagName("si"):
                # concatenate all <t> descendants
                parts = []
                for t in si.GetElementsByTagName("t"):
                    parts.append(t.InnerText)
                shared.append(to_unicode("".join(parts)))

        def parse_sheet(sheet_xml):
            d = XmlDocument()
            d.LoadXml(sheet_xml)
            rows = []
            for row_el in d.GetElementsByTagName("row"):
                cells = {}
                max_c = -1
                for c in row_el.GetElementsByTagName("c"):
                    ref = c.GetAttribute("r")           # e.g. "B7"
                    cidx = col_to_index(ref) if ref else (max(cells.keys()) + 1 if cells else 0)
                    ctype = c.GetAttribute("t")
                    val = u""
                    if ctype == "inlineStr":
                        is_nodes = c.GetElementsByTagName("t")
                        if is_nodes.Count > 0:
                            val = to_unicode(is_nodes[0].InnerText)
                    else:
                        v_nodes = c.GetElementsByTagName("v")
                        if v_nodes.Count > 0:
                            raw = v_nodes[0].InnerText
                            if ctype == "s":
                                try:
                                    val = shared[int(raw)]
                                except Exception:
                                    val = u""
                            else:
                                val = to_unicode(raw)
                    cells[cidx] = val
                    if cidx > max_c:
                        max_c = cidx
                row = [cells.get(i, u"") for i in range(max_c + 1)]
                rows.append(row)
            return rows

        # A workbook can hold several tabs (e.g. the D2 data file). Read every
        # worksheet and pick the one whose header row carries a RoomID/ElementID
        # column - that's the QA/QC-style sheet we can match on. This lets the
        # user point at a multi-tab file and still land on the right sheet.
        sheet_parts = sorted(k for k in entries
                             if k.startswith("xl/worksheets/sheet") and k.endswith(".xml"))
        if not sheet_parts:
            return []

        first_rows = None
        for part in sheet_parts:
            xml = read_entry(part)
            if not xml:
                continue
            rows = parse_sheet(xml)
            if not rows:
                continue
            if first_rows is None:
                first_rows = rows
            header_lower = [to_unicode(h).strip().lower() for h in rows[0]]
            if "roomid" in header_lower or "elementid" in header_lower:
                return rows                      # matched the id-bearing sheet
        # no id column found on any tab - fall back to the first non-empty sheet
        return first_rows or []
    finally:
        archive.Dispose()


try:
    table = read_xlsx(xlsx_path)
except Exception as ex:
    forms.alert("Could not read the Excel file:\n{}".format(ex), exitscript=True)

if not table or len(table) < 2:
    forms.alert("The file has no data rows.", exitscript=True)

header = [to_unicode(h).strip() for h in table[0]]
header_lower = [h.lower() for h in header]

if "roomid" in header_lower:
    id_col = header_lower.index("roomid")
elif "elementid" in header_lower:
    id_col = header_lower.index("elementid")
else:
    forms.alert("No 'RoomID' or 'ElementID' column found. "
                "Please use a QA/QC export file.", exitscript=True)
# parameter columns = every header that isn't an identity/summary column
param_cols = [(i, header[i]) for i, h in enumerate(header_lower)
              if h not in NON_PARAM_COLS and h != ""]

if not param_cols:
    forms.alert("No parameter columns found to update.", exitscript=True)


# ------------------------------------------------------------------
# 3. Build an element lookup by ElementId value (any category)
# ------------------------------------------------------------------
def get_id_value(elem):
    eid = elem.Id
    return int(eid.Value if hasattr(eid, "Value") else eid.IntegerValue)


all_elems = (
    DB.FilteredElementCollector(doc)
    .WhereElementIsNotElementType()
    .ToElements()
)
elem_by_id = {}
for e in all_elems:
    try:
        elem_by_id[get_id_value(e)] = e
    except Exception:
        continue


def current_value(p):
    if p is None or not p.HasValue:
        return u""
    try:
        if p.StorageType == DB.StorageType.String:
            return to_unicode(p.AsString() or u"")
        vs = p.AsValueString()
        return to_unicode(vs or u"")
    except Exception:
        return u""


def can_write_string(p):
    return (p is not None and not p.IsReadOnly
            and p.StorageType == DB.StorageType.String)


def can_write(p):
    """Writable text, number, or integer parameter."""
    return (p is not None and not p.IsReadOnly and p.StorageType in (
        DB.StorageType.String, DB.StorageType.Double, DB.StorageType.Integer))


def _unit_id(p):
    """Return the parameter's display unit (ForgeTypeId on 2021+, else the old
    DisplayUnitType), or None if the value has no unit / can't be determined."""
    try:
        return p.GetUnitTypeId()          # Revit 2021+
    except Exception:
        pass
    try:
        return p.DisplayUnitType          # Revit <= 2020
    except Exception:
        return None


def to_internal(p, display_value):
    """Convert a value typed in the parameter's DISPLAY units (e.g. m2, mm)
    into Revit's internal units. Falls back to the raw number if unitless."""
    uid = _unit_id(p)
    if uid is None:
        return display_value
    try:
        return DB.UnitUtils.ConvertToInternalUnits(display_value, uid)
    except Exception:
        return display_value


def display_number(p):
    """Current numeric value of a Double/Integer param in its display units,
    or None if it has no value."""
    if p is None or not p.HasValue:
        return None
    try:
        if p.StorageType == DB.StorageType.Integer:
            return float(p.AsInteger())
        raw = p.AsDouble()
        uid = _unit_id(p)
        if uid is None:
            return raw
        return DB.UnitUtils.ConvertFromInternalUnits(raw, uid)
    except Exception:
        return None


def parse_number(text):
    """'1,234.5 m2' -> 1234.5, or None if not a number."""
    if text is None:
        return None
    t = to_unicode(text).strip().replace(",", "")
    if t == u"":
        return None
    # keep leading sign, digits, and a single decimal point; drop unit suffixes
    out = []
    for ch in t:
        if ch.isdigit() or ch in u".-+":
            out.append(ch)
        elif out:
            break
    try:
        return float(u"".join(out))
    except Exception:
        return None


def numbers_match(p, want_display):
    """True if the param's current value already equals want_display (in display
    units), within a small tolerance."""
    cur = display_number(p)
    if cur is None:
        cur = 0.0
    tol = max(0.01, abs(want_display) * 1e-4)
    return abs(cur - want_display) <= tol


# Map common room parameter display names to their BuiltInParameter, so we
# resolve the correct WRITABLE built-in instead of relying on LookupParameter
# (which is unreliable for built-in room finish/identity fields and can return
# a read-only duplicate or miss them entirely).
BIP_BY_NAME = {
    u"department": BIP.ROOM_DEPARTMENT,
    u"occupancy": BIP.ROOM_OCCUPANCY,
    u"floor finish": BIP.ROOM_FINISH_FLOOR,
    u"ceiling finish": BIP.ROOM_FINISH_CEILING,
    u"wall finish": BIP.ROOM_FINISH_WALL,
    u"base finish": BIP.ROOM_FINISH_BASE,
    u"comments": BIP.ALL_MODEL_INSTANCE_COMMENTS,
    u"name": BIP.ROOM_NAME,
    u"number": BIP.ROOM_NUMBER,
}


def resolve_param(room, pname):
    """Resolve a writable String parameter for a column header.

    Try the known BuiltInParameter first (correct for finish/identity fields),
    then fall back to a LookupParameter by name for custom/shared parameters.
    Prefer a writable string parameter; if the built-in isn't writable, try
    the lookup as a second chance (covers custom params sharing a name).
    """
    candidates = []
    bip = BIP_BY_NAME.get(pname.strip().lower())
    if bip is not None:
        try:
            p = room.get_Parameter(bip)
            if p is not None:
                candidates.append(p)
        except Exception:
            pass
    # LookupParameter may return more than one match on some versions; get all
    try:
        found = room.LookupParameter(pname)
        if found is not None:
            candidates.append(found)
    except Exception:
        pass
    # also scan the full parameter set for a writable string match by name
    try:
        for p in room.Parameters:
            if p.Definition is not None and to_unicode(p.Definition.Name).strip().lower() == pname.strip().lower():
                candidates.append(p)
    except Exception:
        pass
    # pick the first writable parameter - prefer a writable String (finish /
    # identity fields), then any writable number/area/integer parameter.
    for p in candidates:
        if can_write_string(p):
            return p, True
    for p in candidates:
        if can_write(p):
            return p, True
    # none writable - return the first candidate (for reporting) flagged not writable
    return (candidates[0] if candidates else None), False


# ------------------------------------------------------------------
# 4. Compute planned changes
#    - fill blanks and overwrite differing values
#    - collect overwrites separately for the preview
# ------------------------------------------------------------------
fills = []        # (room, param_name, param, new_value)
overwrites = []   # (room, param_name, param, old_value, new_value)
skipped_ro = []   # (room, param_name)  - read-only / non-string / missing param
unmatched_ids = []
blank_cells = 0

for row in table[1:]:
    if id_col >= len(row):
        continue
    raw_id = to_unicode(row[id_col]).strip()
    if not raw_id:
        continue
    try:
        rid = int(float(raw_id))
    except Exception:
        continue
    room = elem_by_id.get(rid)
    if room is None:
        unmatched_ids.append(raw_id)
        continue

    for ci, pname in param_cols:
        new_val = to_unicode(row[ci]).strip() if ci < len(row) else u""

        p, writable = resolve_param(room, pname)
        if not writable:
            # only report a skip if the cell actually asked for a change
            if new_val != u"":
                skipped_ro.append((room, pname))
            continue

        old_val = current_value(p).strip()
        is_number = p.StorageType in (DB.StorageType.Double, DB.StorageType.Integer)

        if is_number:
            # numeric / area parameter: compare in display units, not as text
            cur_num = display_number(p)
            has_cur = cur_num is not None and abs(cur_num) > 1e-9

            if new_val == u"":
                # blank cell = clear (set to 0)
                if not has_cur:
                    blank_cells += 1
                    continue             # already empty / zero
                overwrites.append((room, pname, p, old_val, u""))
                continue

            want = parse_number(new_val)
            if want is None:
                skipped_ro.append((room, pname))   # not a number - can't write
                continue
            if numbers_match(p, want):
                continue                 # already correct
            if not has_cur:
                fills.append((room, pname, p, new_val))
            else:
                overwrites.append((room, pname, p, old_val, new_val))
            continue

        # text parameter (original behaviour)
        if new_val == u"":
            # blank cell = clear the value in Revit
            if old_val == u"":
                blank_cells += 1
                continue                 # already empty - nothing to do
            # clearing a real value is destructive -> treat as an overwrite (preview + confirm)
            overwrites.append((room, pname, p, old_val, u""))
            continue

        if old_val == new_val:
            continue                     # already correct
        if old_val == u"":
            fills.append((room, pname, p, new_val))
        else:
            overwrites.append((room, pname, p, old_val, new_val))

total_changes = len(fills) + len(overwrites)
if total_changes == 0:
    forms.alert("Nothing to update - every filled cell already matches the model."
                + ("\n\n{} row(s) had an ID not found in this model.".format(
                    len(unmatched_ids)) if unmatched_ids else ""),
                exitscript=True)


# ------------------------------------------------------------------
# 5. Preview overwrites and confirm (fills are applied without prompting)
# ------------------------------------------------------------------
def rnum(room):
    """A short human label for an element: Number (rooms) / Mark / Name / Id."""
    for bip in (DB.BuiltInParameter.ROOM_NUMBER, DB.BuiltInParameter.ALL_MODEL_MARK):
        try:
            p = room.get_Parameter(bip)
            v = current_value(p)
            if v:
                return v
        except Exception:
            pass
    try:
        nm = to_unicode(room.Name)
        if nm:
            return nm
    except Exception:
        pass
    return to_unicode(get_id_value(room))


if overwrites:
    n_clear = sum(1 for _, _, _, ov, nv in overwrites if nv == u"")
    n_change = len(overwrites) - n_clear
    parts = []
    if n_change:
        parts.append("{} value(s) will be OVERWRITTEN".format(n_change))
    if n_clear:
        parts.append("{} value(s) will be CLEARED".format(n_clear))
    output.print_md("## :warning: {}".format(" and ".join(parts)))
    output.print_md("These elements already have a value that will change:")
    table_rows = []
    for room, pname, p, old_val, new_val in overwrites:
        table_rows.append([
            output.linkify(room.Id),
            rnum(room),
            pname,
            old_val,
            new_val if new_val != u"" else u"(cleared)",
        ])
    output.print_table(
        table_data=table_rows,
        columns=["Element", "ID / Mark", "Parameter", "Current value", "New value"],
    )
    if fills:
        output.print_md("_Plus **{}** blank value(s) that will be filled in._".format(len(fills)))

    proceed = forms.alert(
        "{} overwritten, {} cleared, {} blank(s) filled.\n\n"
        "Review the list in the output window.\n\nApply all changes?".format(
            n_change, n_clear, len(fills)),
        yes=True, no=True)
    if not proceed:
        output.print_md("**Cancelled - no changes written.**")
        script.exit()
else:
    # only blanks being filled - no destructive change, apply directly
    output.print_md("**{}** blank value(s) will be filled in (no existing values affected).".format(
        len(fills)))


# ------------------------------------------------------------------
# 6. Write inside a transaction
# ------------------------------------------------------------------
def set_param(p, display_val):
    """Write a value, converting for the parameter's storage type.
    Numbers are read in the parameter's display units (e.g. m2) and stored in
    Revit internal units; a blank clears a number to 0."""
    st = p.StorageType
    if st == DB.StorageType.Double:
        if display_val == u"":
            p.Set(0.0)
        else:
            p.Set(to_internal(p, parse_number(display_val) or 0.0))
    elif st == DB.StorageType.Integer:
        if display_val == u"":
            p.Set(0)
        else:
            p.Set(int(round(parse_number(display_val) or 0.0)))
    else:
        p.Set(display_val)


written = 0
errors = []
t = DB.Transaction(doc, "Update parameters from Excel")
t.Start()
try:
    for room, pname, p, new_val in fills:
        try:
            set_param(p, new_val)
            written += 1
        except Exception as ex:
            errors.append((rnum(room), pname, str(ex)))
    for room, pname, p, old_val, new_val in overwrites:
        try:
            set_param(p, new_val)
            written += 1
        except Exception as ex:
            errors.append((rnum(room), pname, str(ex)))
    t.Commit()
except Exception as ex:
    t.RollBack()
    forms.alert("Update failed and was rolled back:\n{}".format(ex), exitscript=True)


# ------------------------------------------------------------------
# 7. Report
# ------------------------------------------------------------------
output.print_md("## Update complete")
n_clear_done = sum(1 for _, _, _, ov, nv in overwrites if nv == u"")
n_over_done = len(overwrites) - n_clear_done
output.print_md("- **{}** parameter value(s) written "
                "({} filled, {} overwritten, {} cleared)".format(
                    written, len(fills), n_over_done, n_clear_done))

# per-parameter breakdown so you can confirm each column actually updated
from collections import defaultdict
per_param = defaultdict(lambda: [0, 0, 0])   # name -> [filled, overwritten, cleared]
for room, pname, p, new_val in fills:
    per_param[pname][0] += 1
for room, pname, p, old_val, new_val in overwrites:
    if new_val == u"":
        per_param[pname][2] += 1
    else:
        per_param[pname][1] += 1
if per_param:
    output.print_md("**By parameter:**")
    output.print_table(
        table_data=[[pn, c[0], c[1], c[2], c[0] + c[1] + c[2]]
                    for pn, c in sorted(per_param.items())],
        columns=["Parameter", "Filled", "Overwritten", "Cleared", "Total"],
    )

if skipped_ro:
    # count skips per parameter so a read-only/mismatched column is obvious
    ro_count = defaultdict(int)
    for _, pn in skipped_ro:
        ro_count[pn] += 1
    detail = ", ".join("{} ({})".format(pn, n) for pn, n in sorted(ro_count.items()))
    output.print_md("- :warning: **{}** cell(s) skipped - parameter is read-only, "
                    "not found on the element, or the cell was not a number for a "
                    "numeric parameter: {}".format(len(skipped_ro), detail))
    output.print_md("  _If a parameter you expected to update is listed here, it may be a "
                    "calculated/read-only field or a type parameter rather than an instance one._")
if unmatched_ids:
    output.print_md("- **{}** row(s) skipped - ID not found in this model".format(
        len(unmatched_ids)))
if errors:
    output.print_md("### Some writes errored:")
    output.print_table(table_data=[[n, pn, e] for n, pn, e in errors],
                       columns=["Element", "Parameter", "Error"])
