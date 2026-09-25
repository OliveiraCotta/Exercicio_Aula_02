# -*- coding: utf-8 -*-
"""Room Finish Keynote Automation.

Reads the Keynotes of the elements that bound each Room and writes them to the
Room finish parameters (Wall / Floor / Ceiling / Base Finish).

    READ -> ANALYZE -> PREVIEW -> USER CONFIRMATION -> TRANSACTION -> WRITE -> REPORT

Where each finish comes from (all read-only, nothing but the 4 Room
parameters is ever written):

  Wall Finish    Room.GetBoundarySegments()  -> BoundarySegment.ElementId
                 (+ LinkElementId for linked walls), filtered to Walls.
  Floor Finish   SpatialElementGeometryCalculator -> Bottom subfaces whose
                 bounding element is a Floor.
  Ceiling Finish SpatialElementGeometryCalculator -> Top subfaces whose
                 bounding element is a Ceiling (coverage is measured, so a
                 room that is partly open to the slab is reported).
  Base Finish    Wall Sweeps (standalone WallSweep elements hosted on a
                 bounding wall, or sweeps built into the wall type), on the
                 side of the wall that faces the room, near the wall base.
                 Anything else is reported as "Not detected".

Keynote lookup, per element (see KEYNOTE_SOURCES):
  1. Keynote on the element itself (rare - most categories only have it on
     the type), then the Keynote of its type (BuiltInParameter.KEYNOTE_PARAM).
  2. Keynote of the material on the face that actually touches the room
     (painted material first, then the face material), read through the
     room geometry calculator - so compound-wall layers and the Paint tool
     are both honoured.

Engine: IronPython 2.7 (pyRevit default). The syntax is kept py2/py3 neutral,
but the WPF window relies on pyRevit's WPFWindow + DataTable binding, which is
only exercised on IronPython.
"""

__title__ = "Room Finish\nKeynotes"
__doc__ = ("Automatically reads finish Keynotes from Room boundaries and "
           "writes them to Room finish parameters.")

import os
import re
import json
import codecs
from datetime import datetime
from collections import OrderedDict

import clr
clr.AddReference("System.Data")
from System import Boolean, String
from System.Data import DataTable

from pyrevit import revit, DB, script, forms

try:
    unicode
except NameError:          # IronPython 3 / CPython
    unicode = str

doc = revit.doc
BIP = DB.BuiltInParameter
output = script.get_output()
output.set_title("Room Finish Keynotes - log")


# ==================================================================
# Settings
# ==================================================================
FINISH_PARAMS = OrderedDict([
    ("wall", u"Wall Finish"),
    ("floor", u"Floor Finish"),
    ("ceiling", u"Ceiling Finish"),
    ("base", u"Base Finish"),
])
# Revit Rooms already carry built-in parameters with these exact English
# names. If a project parameter with the same name was also created, the room
# has TWO parameters called e.g. "Wall Finish" - the UI asks which one to use.
FINISH_BIPS = {
    "wall": "ROOM_FINISH_WALL",
    "floor": "ROOM_FINISH_FLOOR",
    "ceiling": "ROOM_FINISH_CEILING",
    "base": "ROOM_FINISH_BASE",
}
SEPARATOR = u" / "

# Boundary location for the 2D wall boundaries. Finish = the room-facing
# face of the wall, which is what a finish schedule describes.
BOUNDARY_LOCATION = DB.SpatialElementBoundaryLocation.Finish

# A wall sweep counts as a skirting / base finish when it is measured from the
# wall base and starts no higher than this above it.
BASE_MAX_OFFSET_M = 0.30
BASE_MAX_OFFSET_FT = BASE_MAX_OFFSET_M / 0.3048

# Ceiling must cover this share of the room's top surface to count as "OK".
CEILING_FULL_COVERAGE = 0.98

KEYNOTE_SOURCES = [
    ("auto", u"Element/Type Keynote, then room-facing material Keynote (recommended)"),
    ("type", u"Element/Type Keynote only"),
    ("material", u"Room-facing material Keynote only"),
]


def _bic(name):
    try:
        return int(getattr(DB.BuiltInCategory, name))
    except Exception:
        return None


CAT_WALLS = _bic("OST_Walls")
CAT_FLOORS = _bic("OST_Floors")
CAT_CEILINGS = _bic("OST_Ceilings")
CAT_ROOM_SEP = _bic("OST_RoomSeparationLines")


# ==================================================================
# Small helpers
# ==================================================================
def to_unicode(value):
    """Coerce any Revit / .NET / byte value into clean unicode text."""
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    if isinstance(value, bytes):
        for enc in ("utf-8", "cp1252", "latin-1"):
            try:
                return value.decode(enc)
            except (UnicodeDecodeError, LookupError):
                continue
        return value.decode("ascii", "replace")
    try:
        return unicode(value)
    except Exception:
        return u""


def eid_int(eid):
    """ElementId -> int on every Revit version (Value is 2024+)."""
    try:
        return int(eid.Value)
    except AttributeError:
        return int(eid.IntegerValue)


def is_valid_id(eid):
    return eid is not None and eid != DB.ElementId.InvalidElementId


def cat_int(elem):
    try:
        return eid_int(elem.Category.Id)
    except Exception:
        return None


def cat_name(elem):
    try:
        return to_unicode(elem.Category.Name)
    except Exception:
        return u"(no category)"


def param_str(p):
    if p is None or not p.HasValue:
        return u""
    try:
        return to_unicode(p.AsString() or u"").strip()
    except Exception:
        return u""


def natural_key(text):
    """RE2 < RE10; ties broken on the raw text so the order is total."""
    parts = re.split(r"(\d+)", text)
    return ([int(t) if t.isdigit() else t.lower() for t in parts], text)


def unique_sorted(values):
    """Remove duplicates and return a deterministic, natural order."""
    return sorted(set(v for v in values if v), key=natural_key)


def join_keys(keys):
    return SEPARATOR.join(keys)


def ft_to_m(v):
    return v * 0.3048


# ==================================================================
# Model index: every GetElement / keynote lookup goes through a cache.
# doc_key 0 = host document; otherwise the RevitLinkInstance id.
# ==================================================================
class ModelIndex(object):
    def __init__(self, host):
        self.host = host
        self._docs = {0: host}
        self._link_names = {}
        self._elems = {}
        self._elem_kn = {}
        self._mat_kn = {}
        self._mat_names = {}

    def doc_for(self, doc_key):
        if doc_key in self._docs:
            return self._docs[doc_key]
        ldoc = None
        try:
            inst = self.host.GetElement(DB.ElementId(doc_key))
            self._link_names[doc_key] = to_unicode(inst.Name) if inst else u"link"
            ldoc = inst.GetLinkDocument() if inst else None   # None when unloaded
        except Exception:
            ldoc = None
        self._docs[doc_key] = ldoc
        return ldoc

    def link_name(self, doc_key):
        if not doc_key:
            return u""
        self.doc_for(doc_key)
        return self._link_names.get(doc_key, u"link")

    def element(self, doc_key, eid):
        if not is_valid_id(eid):
            return None
        k = (doc_key, eid_int(eid))
        if k not in self._elems:
            d = self.doc_for(doc_key)
            self._elems[k] = d.GetElement(eid) if d is not None else None
        return self._elems[k]

    def element_keynote(self, doc_key, elem):
        """(keynote, source) from the element, then its type. Never writes."""
        k = (doc_key, eid_int(elem.Id))
        if k in self._elem_kn:
            return self._elem_kn[k]
        result = (u"", None)
        kn = param_str(elem.get_Parameter(BIP.KEYNOTE_PARAM))
        if kn:
            result = (kn, "element")
        else:
            try:
                tid = elem.GetTypeId()
            except Exception:
                tid = None
            typ = self.element(doc_key, tid) if is_valid_id(tid) else None
            if typ is not None:
                kn = param_str(typ.get_Parameter(BIP.KEYNOTE_PARAM))
                if kn:
                    result = (kn, "type")
        self._elem_kn[k] = result
        return result

    def type_name(self, doc_key, elem):
        try:
            typ = self.element(doc_key, elem.GetTypeId())
            if typ is not None:
                return to_unicode(DB.Element.Name.GetValue(typ))
        except Exception:
            pass
        return u""

    def material_keynote(self, doc_key, mat_id):
        if not is_valid_id(mat_id):
            return u""
        k = (doc_key, eid_int(mat_id))
        if k not in self._mat_kn:
            mat = self.element(doc_key, mat_id)
            kn = u""
            name = u""
            if mat is not None:
                name = to_unicode(mat.Name)
                kn = param_str(mat.get_Parameter(BIP.KEYNOTE_PARAM))
            self._mat_kn[k] = kn
            self._mat_names[k] = name
        return self._mat_kn[k]

    def material_name(self, doc_key, mat_id):
        self.material_keynote(doc_key, mat_id)
        return self._mat_names.get((doc_key, eid_int(mat_id)), u"")


def load_keynote_table(d):
    """{key: text} from the loaded keynote file - read only, used for tooltips
    and to flag keys that don't exist in the file. Empty if not available."""
    out = {}
    try:
        table = DB.KeynoteTable.GetKeynoteTable(d)
        for entry in table.GetKeyBasedTreeEntries():
            key = to_unicode(entry.Key)
            if key:
                out[key] = to_unicode(getattr(entry, "KeynoteText", u""))
    except Exception:
        pass
    return out


# ==================================================================
# Room parameters: check they exist, detect duplicated names
# ==================================================================
class ParamTarget(object):
    def __init__(self, key, name, param, kind):
        self.key = key
        self.name = name
        self.kind = kind                   # "built-in" | "shared" | "project"
        self.definition = param.Definition
        self.bip = None
        if kind == "built-in":
            self.bip = self.definition.BuiltInParameter

    def get(self, room):
        if self.bip is not None:
            return room.get_Parameter(self.bip)
        return room.get_Parameter(self.definition)

    def read(self, room):
        return param_str(self.get(room))


def param_kind(p):
    try:
        bip = p.Definition.BuiltInParameter
        if bip != BIP.INVALID:
            return "built-in"
    except Exception:
        pass
    return "shared" if p.IsShared else "project"


def inspect_room_params(sample_room):
    """Returns {key: [candidate Parameter, ...]} by exact name."""
    wanted = set(FINISH_PARAMS.values())
    by_name = {}
    for p in sample_room.Parameters:
        try:
            n = to_unicode(p.Definition.Name)
        except Exception:
            continue
        if n in wanted:
            by_name.setdefault(n, []).append(p)
    return dict((k, by_name.get(n, [])) for k, n in FINISH_PARAMS.items())


def resolve_targets(candidates, prefer_builtin):
    """-> (targets {key: ParamTarget}, problems [str])"""
    targets = {}
    problems = []
    for key, name in FINISH_PARAMS.items():
        cands = [p for p in candidates.get(key, [])]
        if not cands:
            problems.append(u"Parameter '{}' was not found.".format(name))
            continue
        usable = [p for p in cands
                  if p.StorageType == DB.StorageType.String and not p.IsReadOnly]
        if not usable:
            problems.append(u"Parameter '{}' exists but is not a writable text parameter.".format(name))
            continue
        builtin = [p for p in usable if param_kind(p) == "built-in"]
        custom = [p for p in usable if param_kind(p) != "built-in"]
        if builtin and custom:
            p = builtin[0] if prefer_builtin else custom[0]
        else:
            p = usable[0]
        targets[key] = ParamTarget(key, name, p, param_kind(p))
    return targets, problems


def has_duplicates(candidates):
    return [FINISH_PARAMS[k] for k, c in candidates.items() if len(c) > 1]


# ==================================================================
# Wall sweeps (base finish candidates)
# ==================================================================
def build_sweep_index(d):
    """host wall id -> [sweep dict]. One pass over all WallSweep elements."""
    idx = {}
    try:
        sweeps = DB.FilteredElementCollector(d).OfClass(DB.WallSweep).ToElements()
    except Exception:
        return idx
    for sw in sweeps:
        try:
            info = sw.GetWallSweepInfo()
            if info.WallSweepType != DB.WallSweepType.Sweep:
                continue                        # reveals are not finishes
            rec = {
                "elem": sw,
                "side": info.WallSide,
                "from": info.DistanceMeasuredFrom,
                "dist": info.Distance,
                "mat": info.MaterialId,
                "bb": sw.get_BoundingBox(None),
            }
            for hid in sw.GetHostIds():
                idx.setdefault(eid_int(hid), []).append(rec)
        except Exception:
            continue
    return idx


def is_base_height(sweep_from, dist):
    return sweep_from == DB.DistanceMeasuredFrom.Base and dist <= BASE_MAX_OFFSET_FT + 1e-6


def line_interval(line, pts):
    o = line.GetEndPoint(0)
    dv = line.Direction
    ts = [(p.X - o.X) * dv.X + (p.Y - o.Y) * dv.Y for p in pts]
    return min(ts), max(ts)


def sweep_overlaps_room(wall, sweep_bb, seg_points):
    """Does a standalone sweep run along the part of the wall that bounds this
    room? Straight walls only; curved walls are assumed to overlap."""
    if sweep_bb is None or not seg_points:
        return True
    try:
        curve = wall.Location.Curve
    except Exception:
        return True
    if not isinstance(curve, DB.Line):
        return True
    mn, mx = sweep_bb.Min, sweep_bb.Max
    corners = [DB.XYZ(mn.X, mn.Y, 0), DB.XYZ(mx.X, mn.Y, 0),
               DB.XYZ(mn.X, mx.Y, 0), DB.XYZ(mx.X, mx.Y, 0)]
    s0, s1 = line_interval(curve, corners)
    r0, r1 = line_interval(curve, seg_points)
    return min(s1, r1) - max(s0, r0) > 0.05     # > ~15 mm of shared length


# ==================================================================
# Geometry helpers
# ==================================================================
def face_normal(face):
    try:
        bb = face.GetBoundingBox()
        uv = DB.UV((bb.Min.U + bb.Max.U) / 2.0, (bb.Min.V + bb.Max.V) / 2.0)
        return face.ComputeNormal(uv)
    except Exception:
        return None


def face_material_id(fdoc, elem, face):
    """Painted material wins over the face (layer) material."""
    if fdoc is None or face is None:
        return None
    try:
        if fdoc.IsPainted(elem.Id, face):
            mid = fdoc.GetPaintedMaterial(elem.Id, face)
            if is_valid_id(mid):
                return mid
    except Exception:
        pass
    try:
        mid = face.MaterialElementId
        if is_valid_id(mid):
            return mid
    except Exception:
        pass
    return None


def wall_side_facing(wall, normal):
    """WallSide the room is on, from the normal of the wall face touching it."""
    if normal is None:
        return None
    try:
        o = wall.Orientation            # points to the wall's exterior side
    except Exception:
        return None
    dot = normal.X * o.X + normal.Y * o.Y
    if dot > 0.3:
        return DB.WallSide.Exterior
    if dot < -0.3:
        return DB.WallSide.Interior
    return None                         # end face - no side


# ==================================================================
# Scanner: READ + ANALYZE (no Transaction anywhere in here)
# ==================================================================
class Scanner(object):
    def __init__(self, d, mode):
        self.doc = d
        self.mode = mode
        self.idx = ModelIndex(d)
        self.seg_opts = DB.SpatialElementBoundaryOptions()
        self.seg_opts.SpatialElementBoundaryLocation = BOUNDARY_LOCATION
        calc_opts = DB.SpatialElementBoundaryOptions()
        calc_opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
        calc_opts.StoreFreeBoundaryFaces = True
        self.calc = DB.SpatialElementGeometryCalculator(d, calc_opts)
        self.sweeps = build_sweep_index(d)
        self.keynote_texts = load_keynote_table(d)

    # ---------------- raw relationships ----------------
    def _ref(self, host_or_link_id, linked_id):
        """(doc_key, element) from a boundary reference."""
        if is_valid_id(linked_id):
            dk = eid_int(host_or_link_id)
            return dk, self.idx.element(dk, linked_id)
        return 0, self.idx.element(0, host_or_link_id)

    def collect(self, room):
        raw = {
            "walls": OrderedDict(), "separation_lines": 0, "other_bounding": {},
            "floors": OrderedDict(), "bottom_other": {}, "bottom_free": 0.0,
            "ceilings": OrderedDict(), "top_other": {}, "top_free": 0.0, "top_total": 0.0,
            "calc_error": None,
        }
        # 1) 2D boundary: which elements really bound the room
        loops = room.GetBoundarySegments(self.seg_opts) or []
        for loop in loops:
            for seg in loop:
                eid = seg.ElementId
                if not is_valid_id(eid):
                    continue
                dk, el = self._ref(eid, seg.LinkElementId)
                if el is None:
                    continue
                c = cat_int(el)
                if isinstance(el, DB.Wall):
                    k = (dk, eid_int(el.Id))
                    w = raw["walls"].get(k)
                    if w is None:
                        w = {"doc_key": dk, "elem": el, "pts": [], "mats": set(), "sides": set()}
                        raw["walls"][k] = w
                    try:
                        crv = seg.GetCurve()
                        w["pts"].extend([crv.GetEndPoint(0), crv.GetEndPoint(1)])
                    except Exception:
                        pass
                elif c == CAT_ROOM_SEP:
                    raw["separation_lines"] += 1
                else:
                    n = cat_name(el)
                    raw["other_bounding"][n] = raw["other_bounding"].get(n, 0) + 1

        # 2) 3D boundary: floor below, ceiling above, room-facing wall faces
        try:
            res = self.calc.CalculateSpatialElementGeometry(room)
            solid = res.GetGeometry()
        except Exception as ex:
            raw["calc_error"] = to_unicode(ex)
            return raw

        for face in solid.Faces:
            subs = list(res.GetBoundaryFaceInfo(face) or [])
            if not subs:
                n = face_normal(face)
                if n is not None and n.Z > 0.9:
                    raw["top_free"] += face.Area
                    raw["top_total"] += face.Area
                elif n is not None and n.Z < -0.9:
                    raw["bottom_free"] += face.Area
                continue
            for sub in subs:
                stype = sub.SubfaceType
                try:
                    area = sub.GetSubface().Area
                except Exception:
                    area = 0.0
                lid = sub.SpatialBoundaryElement
                if is_valid_id(lid.LinkInstanceId):
                    dk, el = self._ref(lid.LinkInstanceId, lid.LinkedElementId)
                else:
                    dk, el = self._ref(lid.HostElementId, None)
                if stype == DB.SubfaceType.Top:
                    raw["top_total"] += area
                if el is None:
                    if stype == DB.SubfaceType.Top:
                        raw["top_free"] += area
                    elif stype == DB.SubfaceType.Bottom:
                        raw["bottom_free"] += area
                    continue
                try:
                    bface = sub.GetBoundingElementFace()
                except Exception:
                    bface = None
                mat = face_material_id(self.idx.doc_for(dk), el, bface)
                c = cat_int(el)
                k = (dk, eid_int(el.Id))
                if stype == DB.SubfaceType.Bottom:
                    if c == CAT_FLOORS:
                        f = raw["floors"].setdefault(k, {"doc_key": dk, "elem": el, "area": 0.0, "mats": set()})
                        f["area"] += area
                        if mat is not None:
                            f["mats"].add(mat)
                    else:
                        n = cat_name(el)
                        raw["bottom_other"][n] = raw["bottom_other"].get(n, 0.0) + area
                elif stype == DB.SubfaceType.Top:
                    if c == CAT_CEILINGS:
                        f = raw["ceilings"].setdefault(k, {"doc_key": dk, "elem": el, "area": 0.0, "mats": set()})
                        f["area"] += area
                        if mat is not None:
                            f["mats"].add(mat)
                    else:
                        n = cat_name(el)
                        raw["top_other"][n] = raw["top_other"].get(n, 0.0) + area
                else:   # Side
                    w = raw["walls"].get(k)
                    if w is not None:
                        if mat is not None:
                            w["mats"].add(mat)
                        side = wall_side_facing(el, face_normal(bface) if bface else None)
                        if side is not None:
                            w["sides"].add(side)
        return raw

    # ---------------- keynote resolution ----------------
    def element_keys(self, dk, el, mats):
        """-> (list of (keynote, source), why-missing text)"""
        why = []
        if self.mode in ("auto", "type"):
            kn, src = self.idx.element_keynote(dk, el)
            if kn:
                return [(kn, src)], u""
            tname = self.idx.type_name(dk, el)
            why.append(u"type '{}' has no Keynote".format(tname) if tname else u"no type Keynote")
            if self.mode == "type":
                return [], u"; ".join(why)
        found = []
        for mid in sorted(mats or [], key=eid_int):
            kn = self.idx.material_keynote(dk, mid)
            if kn:
                found.append((kn, "material"))
            else:
                why.append(u"material '{}' has no Keynote".format(self.idx.material_name(dk, mid)))
        if not mats:
            why.append(u"no room-facing material found")
        return found, u"; ".join(why)

    def _item(self, dk, el, keys):
        return {
            "id": eid_int(el.Id),
            "cat": cat_name(el),
            "link": self.idx.link_name(dk),
            "kn": [k for k, _ in keys],
            "src": sorted(set(s for _, s in keys)),
        }

    def _finish_from(self, entries):
        """Common Wall/Floor/Ceiling logic. entries: ordered list of dicts with
        doc_key / elem / mats."""
        items, issues, keys = [], [], []
        for e in sorted(entries, key=lambda x: (x["doc_key"], eid_int(x["elem"].Id))):
            dk, el = e["doc_key"], e["elem"]
            found, why = self.element_keys(dk, el, e["mats"])
            items.append(self._item(dk, el, found))
            if found:
                keys.extend(k for k, _ in found)
            else:
                issues.append({"id": eid_int(el.Id), "cat": cat_name(el),
                               "link": self.idx.link_name(dk),
                               "msg": u"Keynote missing ({})".format(why)})
        keys = unique_sorted(keys)
        if not entries:
            state = "not_detected"
        elif not keys:
            state = "missing"
        elif issues:
            state = "partial"
        else:
            state = "ok"
        return {"state": state, "keys": keys, "items": items, "issues": issues,
                "notes": [], "warnings": []}

    def finish_walls(self, raw):
        fin = self._finish_from(list(raw["walls"].values()))
        if raw["calc_error"] and self.mode != "type":
            fin["notes"].append(u"Room geometry could not be calculated, so room-facing wall materials "
                                u"are unknown: " + raw["calc_error"])
        if not raw["walls"]:
            why = u"No Wall bounds this room"
            if raw["separation_lines"]:
                why += u" (bounded by Room Separation Lines only)"
            fin["notes"].append(why + u".")
        if raw["separation_lines"] and raw["walls"]:
            fin["notes"].append(u"{} boundary segment(s) are Room Separation Lines (no finish).".format(raw["separation_lines"]))
        if raw["other_bounding"]:
            fin["notes"].append(u"Also bounded by: " + u", ".join(
                u"{} ({})".format(k, v) for k, v in sorted(raw["other_bounding"].items())) + u" - not counted as walls.")
        return fin

    def finish_floor(self, raw):
        fin = self._finish_from(list(raw["floors"].values()))
        if raw["calc_error"]:
            fin["notes"].append(u"Room geometry could not be calculated: " + raw["calc_error"])
        elif not raw["floors"]:
            if raw["bottom_other"]:
                fin["notes"].append(u"Room bottom is bounded by " + u", ".join(sorted(raw["bottom_other"])) + u", not by a Floor.")
            else:
                fin["notes"].append(u"Room bottom is not bounded by a Floor. Check the floor is Room Bounding and its top is not below the room base.")
        return fin

    def finish_ceiling(self, raw):
        fin = self._finish_from(list(raw["ceilings"].values()))
        if raw["calc_error"]:
            fin["notes"].append(u"Room geometry could not be calculated: " + raw["calc_error"])
            return fin
        ceil_area = sum(c["area"] for c in raw["ceilings"].values())
        total = raw["top_total"] or 0.0
        if not raw["ceilings"]:
            if raw["top_other"]:
                fin["notes"].append(u"Room top is bounded by " + u", ".join(sorted(raw["top_other"])) + u" (no ceiling).")
            else:
                fin["notes"].append(u"Room top is not bounded by a Ceiling. If the room has a ceiling, raise the room Upper Limit / Limit Offset above it.")
        elif total > 0 and ceil_area / total < CEILING_FULL_COVERAGE:
            rest = []
            if raw["top_other"]:
                rest.append(u", ".join(sorted(raw["top_other"])))
            if raw["top_free"] > 0:
                rest.append(u"unbounded (room top below the ceiling)")
            fin["warnings"].append(u"Ceilings cover {:.0f}% of the room top; the rest is {}.".format(
                100.0 * ceil_area / total, u" / ".join(rest) or u"other elements"))
            if fin["state"] == "ok":
                fin["state"] = "partial"
        return fin

    def finish_base(self, raw):
        items, issues, keys, notes = [], [], [], []
        found_any = False
        linked_walls = 0
        for w in sorted(raw["walls"].values(), key=lambda x: (x["doc_key"], eid_int(x["elem"].Id))):
            if w["doc_key"]:
                linked_walls += 1
                continue
            wall = w["elem"]
            sides = w["sides"]
            if not sides:
                continue
            members = [wall]
            try:
                if wall.IsStackedWall:
                    members = [self.idx.element(0, m) for m in wall.GetStackedWallMemberIds()]
                    members = [m for m in members if m is not None]
            except Exception:
                pass
            for m in members:
                # a) standalone wall sweeps hosted on this wall
                for sw in self.sweeps.get(eid_int(m.Id), []):
                    if sw["side"] not in sides or not is_base_height(sw["from"], sw["dist"]):
                        continue
                    if not sweep_overlaps_room(m, sw["bb"], w["pts"]):
                        continue
                    found_any = True
                    el = sw["elem"]
                    mats = set([sw["mat"]]) if is_valid_id(sw["mat"]) else set()
                    got, why = self.element_keys(0, el, mats)
                    items.append(self._item(0, el, got))
                    if got:
                        keys.extend(k for k, _ in got)
                    else:
                        issues.append({"id": eid_int(el.Id), "cat": cat_name(el), "link": u"",
                                       "msg": u"Keynote missing ({})".format(why)})
                # b) sweeps built into the wall type's structure
                try:
                    cs = m.WallType.GetCompoundStructure()
                    infos = list(cs.GetWallSweepsInfo(DB.WallSweepType.Sweep)) if cs else []
                except Exception:
                    infos = []
                for info in infos:
                    if info.WallSide not in sides or not is_base_height(info.DistanceMeasuredFrom, info.Distance):
                        continue
                    found_any = True
                    kn = self.idx.material_keynote(0, info.MaterialId)
                    items.append({"id": eid_int(m.Id), "cat": u"Wall type sweep", "link": u"",
                                  "kn": [kn] if kn else [], "src": ["integral sweep material"] if kn else []})
                    if kn:
                        keys.append(kn)
                    else:
                        issues.append({"id": eid_int(m.Id), "cat": u"Walls", "link": u"",
                                       "msg": u"Integral sweep in wall type - material '{}' has no Keynote".format(
                                           self.idx.material_name(0, info.MaterialId))})
        keys = unique_sorted(keys)
        if not found_any:
            state = "not_detected"
            if raw["calc_error"]:
                notes.append(u"Room geometry could not be calculated, so the room-facing side of the walls is unknown.")
            else:
                notes.append(u"No wall sweep found at the base (<= {:.2f} m) of the room-facing side of the bounding walls.".format(BASE_MAX_OFFSET_M))
        elif not keys:
            state = "missing"
        elif issues:
            state = "partial"
        else:
            state = "ok"
        if linked_walls:
            notes.append(u"{} linked wall(s) not evaluated for base finish.".format(linked_walls))
        return {"state": state, "keys": keys, "items": items, "issues": issues,
                "notes": notes, "warnings": []}

    # ---------------- one room ----------------
    def analyze(self, room):
        rec = {"room": room, "boundary": "ok", "finishes": {}, "error": None}
        try:
            if room.Location is None:
                rec["boundary"] = "unplaced"
                return rec
            if room.Area <= 0:
                rec["boundary"] = "unenclosed"
                return rec
            raw = self.collect(room)
            if not raw["walls"] and not raw["separation_lines"] and not raw["other_bounding"]:
                rec["boundary"] = "no_segments"
            rec["finishes"] = OrderedDict([
                ("wall", self.finish_walls(raw)),
                ("floor", self.finish_floor(raw)),
                ("ceiling", self.finish_ceiling(raw)),
                ("base", self.finish_base(raw)),
            ])
            # keys that don't exist in the loaded keynote file
            if self.keynote_texts:
                for fk, fin in rec["finishes"].items():
                    for k in fin["keys"]:
                        if k not in self.keynote_texts:
                            fin["warnings"].append(u"'{}' is not in the loaded keynote file.".format(k))
        except Exception as ex:
            rec["boundary"] = "error"
            rec["error"] = to_unicode(ex)
        return rec


# ==================================================================
# Room records -> status, preview rows, JSON
# ==================================================================
def room_info(room):
    try:
        level = to_unicode(room.Level.Name) if room.Level else u"(unplaced)"
        elev = room.Level.Elevation if room.Level else 0.0
    except Exception:
        level, elev = u"(unplaced)", 0.0
    phase = u""
    try:
        phase = to_unicode(room.get_Parameter(BIP.ROOM_PHASE).AsValueString())
    except Exception:
        pass
    return {
        "id": eid_int(room.Id),
        "number": to_unicode(room.Number),
        "name": param_str(room.get_Parameter(BIP.ROOM_NAME)),
        "level": level,
        "elev": elev,
        "phase": phase,
    }


def room_status(rec):
    if rec["boundary"] != "ok" or rec.get("write_error"):
        return "ERROR"
    # notes are informational (e.g. a separation line); warnings and any
    # state other than "ok" make the room incomplete
    if all(f["state"] == "ok" and not f["warnings"] for f in rec["finishes"].values()):
        return "OK"
    return "WARNING"


def sort_records(records):
    return sorted(records, key=lambda r: (r["info"]["elev"], r["info"]["level"],
                                          natural_key(r["info"]["number"]), r["info"]["id"]))


def compute_changes(records, targets):
    """Current vs new value for every room/finish. Nothing is written here."""
    for rec in records:
        rec["changes"] = OrderedDict()
        room = rec["room"]
        for key in FINISH_PARAMS:
            t = targets.get(key)
            current = t.read(room) if t else u""
            fin = rec["finishes"].get(key)
            new = join_keys(fin["keys"]) if fin else u""
            if t is None:
                kind = "No parameter"
            elif rec["boundary"] != "ok":
                kind = "Skipped"
            elif not new:
                kind = "Keep" if current else "Nothing found"
            elif new == current:
                kind = "No change"
            elif not current:
                kind = "Fill"
            else:
                kind = "Overwrite"
            rec["changes"][key] = {"current": current, "new": new, "kind": kind}


def worksharing_block(room):
    """Reason this room can't be edited right now, or u''."""
    if not doc.IsWorkshared:
        return u""
    try:
        st = DB.WorksharingUtils.GetCheckoutStatus(doc, room.Id)
        if st == DB.CheckoutStatus.OwnedByOtherUser:
            owner = u""
            try:
                owner = to_unicode(DB.WorksharingUtils.GetWorksharingTooltipInfo(doc, room.Id).Owner)
            except Exception:
                pass
            return u"Owned by {}".format(owner or u"another user")
        upd = DB.WorksharingUtils.GetModelUpdatesStatus(doc, room.Id)
        if upd == DB.ModelUpdatesStatus.UpdatedInCentral:
            return u"Updated in central - reload latest first"
        if upd == DB.ModelUpdatesStatus.DeletedInCentral:
            return u"Deleted in central"
    except Exception:
        pass
    return u""


def build_json(records, targets, mode, keynote_texts, problems):
    rooms = []
    for rec in records:
        info = rec["info"]
        fins = OrderedDict()
        issues = []
        for key in FINISH_PARAMS:
            fin = rec["finishes"].get(key)
            ch = rec.get("changes", {}).get(key, {})
            if fin is None:
                fins[key] = {"state": "n/a", "keys": [], "value": u"", "current": ch.get("current", u""),
                             "items": [], "notes": [], "warnings": [], "change": ch.get("kind", u"")}
                continue
            fins[key] = {
                "state": fin["state"], "keys": fin["keys"], "value": join_keys(fin["keys"]),
                "current": ch.get("current", u""), "change": ch.get("kind", u""),
                "items": fin["items"], "notes": fin["notes"], "warnings": fin["warnings"],
            }
            for iss in fin["issues"]:
                d = dict(iss)
                d["finish"] = key
                issues.append(d)
        boundary_msg = {
            "unplaced": u"Room is not placed.",
            "unenclosed": u"Room is not enclosed or is redundant (area = 0).",
            "no_segments": u"Room has no bounding elements.",
            "error": u"Scan failed: {}".format(rec.get("error") or u""),
        }.get(rec["boundary"], u"")
        rooms.append({
            "id": info["id"], "number": info["number"], "name": info["name"],
            "level": info["level"], "phase": info["phase"],
            "boundary": rec["boundary"], "boundaryMsg": boundary_msg,
            "status": rec["status"], "updated": bool(rec.get("updated")),
            "writeError": rec.get("write_error") or u"",
            "finishes": fins, "issues": issues,
        })
    return {
        "project": to_unicode(doc.Title),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "keynoteSource": dict(KEYNOTE_SOURCES).get(mode, mode),
        "boundaryLocation": to_unicode(BOUNDARY_LOCATION),
        "separator": SEPARATOR,
        "params": OrderedDict((k, {"name": n, "kind": targets[k].kind if k in targets else None})
                              for k, n in FINISH_PARAMS.items()),
        "paramProblems": problems,
        "keynotes": keynote_texts,
        "rooms": rooms,
    }


TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "room_finish_dashboard.html")
REPORT_PATH = script.get_document_data_file("room_finish_dashboard", "html")


def write_report(data):
    with codecs.open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html = f.read()
    payload = to_unicode(json.dumps(data, ensure_ascii=False))
    # keep the JSON from closing the <script> tag or breaking JS string rules
    payload = payload.replace(u"</", u"<\\/").replace(u"\u2028", u"\\u2028").replace(u"\u2029", u"\\u2029")
    html = html.replace(u"__DATA__", payload)
    with codecs.open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    return REPORT_PATH


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
        import webbrowser
        return webbrowser.open("file:///" + path.replace("\\", "/"))
    except Exception:
        return False


def log_link(item):
    """Clickable element id in the output window (host elements only)."""
    if item.get("link"):
        return u"{} (linked: {})".format(item["id"], item["link"])
    try:
        return output.linkify(DB.ElementId(item["id"]))
    except Exception:
        return to_unicode(item["id"])


# ==================================================================
# WPF window
# ==================================================================
XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Room Finish Keynote Automation" Width="1120" Height="760"
        MinWidth="860" MinHeight="560" WindowStartupLocation="CenterScreen"
        Background="#0E1526">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="26"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Height" Value="30"/>
      <Setter Property="MinWidth" Value="120"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
      <Setter Property="Padding" Value="12,0,12,0"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#141E33"/>
      <Setter Property="Foreground" Value="#65E3FF"/>
      <Setter Property="Padding" Value="8,5,8,5"/>
      <Setter Property="BorderBrush" Value="#23324F"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Background" Value="#0E1526"/>
      <Setter Property="Foreground" Value="#D9E8F5"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding Change}" Value="Overwrite">
          <Setter Property="Foreground" Value="#FFB454"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Editable}" Value="False">
          <Setter Property="Foreground" Value="#7A8FA9"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>
  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Text="ROOM FINISH KEYNOTE AUTOMATION" FontSize="16"
               FontWeight="SemiBold" Foreground="#65E3FF"/>
    <TextBlock Grid.Row="1" TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,4,0,12"
               Text="Automatically reads finish Keynotes from Room boundaries and writes them to Room finish parameters."/>

    <Grid Grid.Row="2" Margin="0,0,0,10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="18"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <StackPanel Grid.Column="0">
        <TextBlock Text="Keynote source" Margin="0,0,0,3"/>
        <ComboBox x:Name="cb_source"/>
        <TextBlock Text="If a Room has two parameters with the same name" Margin="0,10,0,3"/>
        <ComboBox x:Name="cb_dup"/>
      </StackPanel>
      <StackPanel Grid.Column="2">
        <TextBlock Text="Room parameters" Margin="0,0,0,3"/>
        <TextBlock x:Name="tb_params" TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9"/>
        <CheckBox x:Name="chk_overwrite" Margin="0,10,0,0"
                  Content="Allow overwriting existing values that differ (otherwise only empty parameters are filled)"/>
      </StackPanel>
    </Grid>

    <TextBlock Grid.Row="3" x:Name="tb_status" TextWrapping="Wrap" FontSize="12"
               Foreground="#3DDCB4" Margin="0,4,0,8" Text="Click Scan Model to start. Nothing is written until you click Update Rooms."/>

    <DataGrid Grid.Row="4" x:Name="grid" AutoGenerateColumns="False" CanUserAddRows="False"
              CanUserDeleteRows="False" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
              HorizontalGridLinesBrush="#1B2740" Background="#0B1120" BorderBrush="#23324F"
              RowHeaderWidth="0" SelectionMode="Extended">
      <DataGrid.Columns>
        <DataGridTemplateColumn Header="Apply" Width="56">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                        IsChecked="{Binding Apply, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                        IsEnabled="{Binding Editable}"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
        <DataGridTextColumn Header="Room" Binding="{Binding Room}" IsReadOnly="True" Width="70"/>
        <DataGridTextColumn Header="Name" Binding="{Binding Name}" IsReadOnly="True" Width="150"/>
        <DataGridTextColumn Header="Level" Binding="{Binding Level}" IsReadOnly="True" Width="100"/>
        <DataGridTextColumn Header="Parameter" Binding="{Binding Parameter}" IsReadOnly="True" Width="105"/>
        <DataGridTextColumn Header="Current" Binding="{Binding Current}" IsReadOnly="True" Width="*"/>
        <DataGridTextColumn Header="New" Binding="{Binding New}" IsReadOnly="True" Width="*"/>
        <DataGridTextColumn Header="Change" Binding="{Binding Change}" IsReadOnly="True" Width="85"/>
        <DataGridTextColumn Header="Note" Binding="{Binding Note}" IsReadOnly="True" Width="160"/>
      </DataGrid.Columns>
    </DataGrid>

    <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="btn_scan" Content="Scan Model"/>
      <Button x:Name="btn_preview" Content="Preview Changes" IsEnabled="False"/>
      <Button x:Name="btn_update" Content="Update Rooms" IsEnabled="False"/>
      <Button x:Name="btn_report" Content="Open HTML Report" IsEnabled="False"/>
      <Button x:Name="btn_cancel" Content="Cancel" Margin="0"/>
    </StackPanel>
  </Grid>
</Window>
"""


class RoomFinishWindow(forms.WPFWindow):
    def __init__(self, rooms):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)
        self.rooms = rooms
        self.records = []
        self.row_keys = []          # table row index -> (record, finish key)
        self.table = None
        self.mode = "auto"
        self.keynote_texts = {}

        self.cb_source.ItemsSource = [label for _, label in KEYNOTE_SOURCES]
        self.cb_source.SelectedIndex = 0
        self.cb_dup.ItemsSource = [u"Write to the built-in Room parameter",
                                   u"Write to the project / shared parameter"]
        self.cb_dup.SelectedIndex = 0

        self.candidates = inspect_room_params(rooms[0])
        self.dups = has_duplicates(self.candidates)
        self.cb_dup.IsEnabled = bool(self.dups)
        self._resolve_params()

        self.btn_scan.Click += self.on_scan
        self.btn_preview.Click += self.on_preview
        self.btn_update.Click += self.on_update
        self.btn_report.Click += self.on_report
        self.btn_cancel.Click += self.on_cancel
        self.chk_overwrite.Checked += self.on_overwrite_toggle
        self.chk_overwrite.Unchecked += self.on_overwrite_toggle
        self.cb_dup.SelectionChanged += self.on_dup_changed

    # ---------------- parameters ----------------
    def _resolve_params(self):
        self.targets, self.problems = resolve_targets(self.candidates, self.cb_dup.SelectedIndex == 0)
        lines = []
        for key, name in FINISH_PARAMS.items():
            t = self.targets.get(key)
            if t is None:
                continue
            dup = u" - duplicated name!" if name in self.dups else u""
            lines.append(u"{}: {} parameter{}".format(name, t.kind, dup))
        lines.extend(self.problems)
        self.tb_params.Text = u"\n".join(lines)

    def on_dup_changed(self, sender, args):
        self._resolve_params()
        self._invalidate_preview()

    def _invalidate_preview(self):
        self.grid.ItemsSource = None
        self.table = None
        self.row_keys = []
        self.btn_update.IsEnabled = False

    # ---------------- READ + ANALYZE ----------------
    def on_scan(self, sender, args):
        self.mode = KEYNOTE_SOURCES[max(0, self.cb_source.SelectedIndex)][0]
        self._invalidate_preview()
        output.print_md(u"**[2/6] Rooms collected:** {} room(s).".format(len(self.rooms)))
        try:
            scanner = Scanner(doc, self.mode)
        except Exception as ex:
            output.print_md(u"**ERROR:** could not start the scan: `{}`".format(to_unicode(ex)))
            forms.alert(u"Could not start the scan:\n\n{}".format(to_unicode(ex)))
            return
        self.keynote_texts = scanner.keynote_texts
        records = []
        cancelled = False
        n = len(self.rooms)
        with forms.ProgressBar(title="Scanning rooms ({value} of {max_value})", cancellable=True) as pb:
            for i, room in enumerate(self.rooms):
                if pb.cancelled:
                    cancelled = True
                    break
                rec = scanner.analyze(room)
                rec["info"] = room_info(room)
                records.append(rec)
                if i % 5 == 0 or i == n - 1:
                    pb.update_progress(i + 1, n)
        if cancelled:
            output.print_md(u"Scan cancelled after {} room(s) - nothing was reported.".format(len(records)))
            self.tb_status.Text = u"Scan cancelled."
            return

        self.records = sort_records(records)
        no_bound = [r for r in self.records if r["boundary"] != "ok"]
        output.print_md(u"**[3/6] Room boundaries analyzed:** {} room(s) with boundaries, {} without "
                        u"(boundary location: {}).".format(len(self.records) - len(no_bound), len(no_bound),
                                                           to_unicode(BOUNDARY_LOCATION)))
        for r in no_bound[:30]:
            output.print_md(u"- Room {} `{}` - {}".format(
                log_link({"id": r["info"]["id"], "link": u""}), r["info"]["number"], r["boundary"]))

        compute_changes(self.records, self.targets)
        for r in self.records:
            r["status"] = room_status(r)
        self._log_keynotes()
        self._write_report()
        self.btn_preview.IsEnabled = True
        self.btn_report.IsEnabled = True
        self.tb_status.Text = self._summary()

    def _log_keynotes(self):
        issues = []
        for r in self.records:
            for key, fin in r["finishes"].items():
                for iss in fin["issues"]:
                    issues.append((r, iss))
        output.print_md(u"**[4/6] Keynotes collected** (source: {}). {} element(s) without Keynote.".format(
            dict(KEYNOTE_SOURCES)[self.mode], len(issues)))
        if not self.keynote_texts:
            output.print_md(u"_Keynote file not loaded or empty - keys are not validated against it._")
        for r, iss in issues[:60]:
            output.print_md(u"- Room `{}`: {} {} - {}".format(
                r["info"]["number"], iss["cat"], log_link(iss), iss["msg"]))
        if len(issues) > 60:
            output.print_md(u"- ... and {} more (see the HTML report).".format(len(issues) - 60))

    def _summary(self):
        recs = self.records
        st = [r["status"] for r in recs]
        return (u"Scanned {} rooms - {} OK, {} warning, {} error. "
                u"Review, then click Preview Changes.".format(
                    len(recs), st.count("OK"), st.count("WARNING"), st.count("ERROR")))

    def _write_report(self):
        try:
            data = build_json(self.records, self.targets, self.mode, self.keynote_texts, self.problems)
            write_report(data)
        except Exception as ex:
            output.print_md(u"**ERROR:** could not write the HTML report: `{}`".format(to_unicode(ex)))

    # ---------------- PREVIEW ----------------
    def on_preview(self, sender, args):
        if not self.records:
            return
        compute_changes(self.records, self.targets)       # re-read current values
        allow_ow = bool(self.chk_overwrite.IsChecked)
        t = DataTable("preview")
        for col, typ in (("Apply", Boolean), ("Editable", Boolean), ("Room", String), ("Name", String),
                         ("Level", String), ("Parameter", String), ("Current", String), ("New", String),
                         ("Change", String), ("Note", String)):
            t.Columns.Add(col, clr.GetClrType(typ))
        self.row_keys = []
        counts = {"Fill": 0, "Overwrite": 0, "Keep": 0}
        for rec in self.records:
            block = None
            for key, ch in rec["changes"].items():
                if ch["kind"] == "Keep":
                    counts["Keep"] += 1
                if ch["kind"] not in ("Fill", "Overwrite"):
                    continue
                if block is None:
                    block = worksharing_block(rec["room"])
                counts[ch["kind"]] += 1
                row = t.NewRow()
                editable = not block
                row["Editable"] = editable
                row["Apply"] = editable and (ch["kind"] == "Fill" or allow_ow)
                row["Room"] = rec["info"]["number"]
                row["Name"] = rec["info"]["name"]
                row["Level"] = rec["info"]["level"]
                row["Parameter"] = FINISH_PARAMS[key]
                row["Current"] = ch["current"]
                row["New"] = ch["new"]
                row["Change"] = ch["kind"]
                note = block or u""
                if rec["finishes"][key]["state"] == "partial":
                    note = (note + u"; " if note else u"") + u"some elements have no Keynote"
                row["Note"] = note
                t.Rows.Add(row)
                self.row_keys.append((rec, key))
        self.table = t
        self.grid.ItemsSource = t.DefaultView
        self.btn_update.IsEnabled = t.Rows.Count > 0
        output.print_md(u"**[5/6] Preview generated:** {} to fill, {} would overwrite a different value, "
                        u"{} existing value(s) kept where nothing was detected.".format(
                            counts["Fill"], counts["Overwrite"], counts["Keep"]))
        if t.Rows.Count == 0:
            self.tb_status.Text = u"Nothing to update - every detected value is already in the rooms."
        else:
            self.tb_status.Text = (u"{} value(s) to fill, {} overwrite(s) (amber - ticked only if overwriting "
                                   u"is allowed). Existing values are never cleared.".format(
                                       counts["Fill"], counts["Overwrite"]))

    def on_overwrite_toggle(self, sender, args):
        if self.table is None:
            return
        allow = bool(self.chk_overwrite.IsChecked)
        for row in self.table.Rows:
            if row["Change"] == "Overwrite" and row["Editable"]:
                row["Apply"] = allow

    # ---------------- CONFIRM + TRANSACTION + WRITE ----------------
    def on_update(self, sender, args):
        if self.table is None:
            return
        try:
            self.grid.CommitEdit()
        except Exception:
            pass
        todo = []
        for i in range(self.table.Rows.Count):
            row = self.table.Rows[i]
            if bool(row["Apply"]) and bool(row["Editable"]):
                rec, key = self.row_keys[i]
                todo.append((rec, key, rec["changes"][key]["new"], rec["changes"][key]["kind"]))
        if not todo:
            forms.alert(u"No rows are ticked in the preview.")
            return
        n_rooms = len(set(eid_int(r["room"].Id) for r, _, _, _ in todo))
        n_ow = len([1 for _, _, _, k in todo if k == "Overwrite"])
        msg = u"Write {} value(s) to {} room(s)?".format(len(todo), n_rooms)
        if n_ow:
            msg += u"\n\n{} of them REPLACE an existing different value.".format(n_ow)
        if not forms.alert(msg, yes=True, no=True):
            return

        written, errors = [], []
        tx = DB.Transaction(doc, "Room Finish Keynotes - update rooms")
        try:
            tx.Start()
            for rec, key, value, kind in todo:
                target = self.targets[key]
                try:
                    p = target.get(rec["room"])
                    if p is None or p.IsReadOnly:
                        errors.append((rec, key, u"parameter missing or read-only on this room"))
                    elif p.Set(value):
                        written.append((rec, key))
                    else:
                        errors.append((rec, key, u"Revit refused the value"))
                except Exception as ex:
                    errors.append((rec, key, to_unicode(ex)))
            status = tx.Commit()
            if status != DB.TransactionStatus.Committed:
                raise Exception(u"Transaction ended with status {}".format(status))
        except Exception as ex:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
            output.print_md(u"**ERROR:** update rolled back - nothing was written. `{}`".format(to_unicode(ex)))
            forms.alert(u"The update failed and was rolled back. Nothing was written.\n\n{}".format(to_unicode(ex)))
            return
        finally:
            tx.Dispose()

        for rec, key in written:
            rec["updated"] = True
        for rec, key, err in errors:
            rec["write_error"] = (rec.get("write_error") or u"") + u"{}: {}. ".format(FINISH_PARAMS[key], err)
        output.print_md(u"**[6/6] Rooms updated:** {} value(s) written to {} room(s), {} error(s).".format(
            len(written), len(set(eid_int(r["room"].Id) for r, _ in written)), len(errors)))
        for rec, key, err in errors[:40]:
            output.print_md(u"- **ERROR** Room {} `{}` {}: {}".format(
                log_link({"id": rec["info"]["id"], "link": u""}), rec["info"]["number"], FINISH_PARAMS[key], err))

        compute_changes(self.records, self.targets)      # read back what is in Revit now
        for r in self.records:
            r["status"] = room_status(r)
        self._write_report()
        self._invalidate_preview()
        self.tb_status.Text = u"Updated {} value(s){}. Open the HTML report to review.".format(
            len(written), u" - {} error(s), see the log".format(len(errors)) if errors else u"")

    # ---------------- report / close ----------------
    def on_report(self, sender, args):
        if not os.path.exists(REPORT_PATH):
            self._write_report()
        if not open_in_browser(REPORT_PATH):
            forms.alert(u"Report saved but could not be opened automatically:\n\n{}".format(REPORT_PATH))

    def on_cancel(self, sender, args):
        self.Close()


# ==================================================================
# Entry point
# ==================================================================
def collect_rooms():
    rooms = []
    for el in (DB.FilteredElementCollector(doc)
               .OfCategory(DB.BuiltInCategory.OST_Rooms)
               .WhereElementIsNotElementType()):
        if isinstance(el, DB.Architecture.Room):
            rooms.append(el)
    return rooms


def main():
    output.print_md(u"**[1/6] Script started.** Document: `{}`".format(to_unicode(doc.Title)))
    if doc.IsFamilyDocument:
        forms.alert(u"Open a project, not a family.", exitscript=True)
    rooms = collect_rooms()
    if not rooms:
        forms.alert(u"No Rooms found in this project.", exitscript=True)
    win = RoomFinishWindow(rooms)
    for msg in win.problems:
        output.print_md(u"**ERROR:** {}".format(msg))
    if win.problems:
        forms.alert(u"\n".join(win.problems) +
                    u"\n\nThese parameters will be reported but not written. "
                    u"Nothing is created automatically.")
    win.ShowDialog()


main()
