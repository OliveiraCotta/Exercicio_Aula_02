# -*- coding: utf-8 -*-
"""Room Finish Keynote Automation.

Reads the Keynotes of the elements that bound each Room and writes them to the
Room finish parameters (Wall / Floor / Ceiling / Base Finish).

    READ -> ANALYZE -> PREVIEW -> USER CONFIRMATION -> TRANSACTION -> WRITE -> REPORT

Which parameter a value goes to is decided ONLY by the Keynote prefix
(case-insensitive, value normalised to upper case), never by the element's
category:

  RD... -> Base Finish     FR... -> Ceiling Finish
  RE... -> Wall Finish     PI... -> Floor Finish      anything else: ignored

The category is only used afterwards to REPORT mismatches (e.g. RE01 on a
Floor). Nothing but the 4 Room parameters is ever written.

Elements considered for a room (all through Revit API relationships):
  - Room.GetBoundarySegments(): BoundarySegment.ElementId (+ LinkElementId
    for linked models) - the elements forming the room outline.
  - SpatialElementGeometryCalculator: every element bounding the room volume
    (side / top / bottom subfaces) and the material of the face touching it.
  - Wall sweeps on the room-facing side of the bounding walls (standalone
    WallSweep elements and sweeps built into the wall type).
  - Family instances located in the room (FamilyInstance.Room for the room's
    phase) - only those whose Keynote has one of the four prefixes.
  - Geometric search of the room's surfaces: the room outline (finish
    boundary) extruded from just below the room base to SEARCH_ABOVE_M above
    its top, tested with ElementIntersectsSolidFilter against every element
    (host + loaded links) whose Keynote has one of the four prefixes. This is
    what finds skirtings (low walls, sweeps, families) and ceilings that the
    Revit room relations don't report: elements below the room computation
    height, not Room Bounding, or above a low room Limit Offset.
  Height guard: a value is only used if the element sits where that finish
  can be for THIS room - PI starts below the room's mid-height, FR ends above
  it, RE / RD overlap the room's height. This keeps the floor finish,
  skirting and walls of the storey above out of the room below. Rejected
  values are reported, not written.

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
from System.Collections.Generic import List

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

# The ONLY rule that decides which Room parameter a Keynote fills.
PREFIX_RULES = OrderedDict([
    ("RD", "base"),
    ("FR", "ceiling"),
    ("RE", "wall"),
    ("PI", "floor"),
])

# FR elements must cover this share of the room's top surface, else warning.
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


CAT_ROOM_SEP = _bic("OST_RoomSeparationLines")
CAT_CORNICES = _bic("OST_Cornices")      # Wall Sweeps

# Categories each finish is EXPECTED on. Used only to report mismatches -
# never to classify. Edit freely to match your modelling standard.
EXPECTED_CATEGORIES = {
    "wall": set([_bic("OST_Walls")]),
    "floor": set([_bic("OST_Floors")]),
    "ceiling": set([_bic("OST_Ceilings")]),
    "base": set([CAT_CORNICES, _bic("OST_Walls"), _bic("OST_GenericModel")]),
}

# Geometric search zone around the room surfaces: the finish outline extruded
# from SEARCH_BELOW_M under the room base to SEARCH_ABOVE_M over the room top
# (so a ceiling above a room whose Limit Offset is too low is still found).
SEARCH_BELOW_M = 0.02
SEARCH_ABOVE_M = 1.00

# Fallback when an element has no bounding box (height guard not possible):
# a Keynote of an element that ONLY bounds the room from this side belongs to
# the neighbouring room (floor slab above, ceiling below) - not written.
POSITION_GUARD = {"floor": "top", "ceiling": "bottom"}


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
        self._type_kn = {}
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
            if is_valid_id(tid):
                tk = (doc_key, eid_int(tid))
                if tk not in self._type_kn:
                    typ = self.element(doc_key, tid)
                    self._type_kn[tk] = param_str(typ.get_Parameter(BIP.KEYNOTE_PARAM)) if typ else u""
                if self._type_kn[tk]:
                    result = (self._type_kn[tk], "type")
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
# Finish classification - by Keynote PREFIX only
# ==================================================================
def classify(keynote):
    """-> (normalised keynote, finish key or None). Case-insensitive prefix
    match; the value is returned in upper case. Unmapped prefixes -> None."""
    norm = to_unicode(keynote).strip().upper()
    for prefix, finish in PREFIX_RULES.items():
        if norm.startswith(prefix):
            return norm, finish
    return norm, None


# ==================================================================
# Wall sweeps on bounding walls
# ==================================================================
def build_sweep_index(d):
    """host wall id -> [sweep dict]. One pass over all WallSweep elements
    (sweeps and reveals - the Keynote prefix decides what they are)."""
    idx = {}
    try:
        sweeps = DB.FilteredElementCollector(d).OfClass(DB.WallSweep).ToElements()
    except Exception:
        return idx
    for sw in sweeps:
        try:
            info = sw.GetWallSweepInfo()
            rec = {
                "elem": sw,
                "side": info.WallSide,
                "mat": info.MaterialId,
                "bb": sw.get_BoundingBox(None),
            }
            for hid in sw.GetHostIds():
                idx.setdefault(eid_int(hid), []).append(rec)
        except Exception:
            continue
    return idx


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


REL_LABEL = {
    "boundary": u"room boundary",
    "side": u"side of room volume",
    "top": u"above the room",
    "bottom": u"below the room",
    "sweep": u"sweep on bounding wall",
    "inside": u"family instance in room",
    "surface": u"on the room surfaces (geometry)",
}


def _subface_rel(stype):
    if stype == DB.SubfaceType.Top:
        return "top"
    if stype == DB.SubfaceType.Bottom:
        return "bottom"
    return "side"


# ==================================================================
# Scanner: READ + ANALYZE (no Transaction anywhere in here)
#
# 1. collect(): every element related to the room through the Revit API
#    (2D boundary, 3D room volume faces, sweeps on its walls, family
#    instances located in it) - WITHOUT looking at categories.
# 2. classify_room(): read each element's Keynote and route it to a finish
#    parameter purely by prefix (PREFIX_RULES). Categories are only used
#    afterwards, to report prefix/category mismatches.
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
        # keys are matched upper-case, like the normalised values
        self.keynote_texts = dict((k.upper(), v) for k, v in load_keynote_table(d).items())
        self._fi_candidates = None
        self._inside = {}           # phase id -> {room id: [FamilyInstance]}
        self._cands = {}            # doc_key -> List[ElementId] with a finish-prefix Keynote
        self._links = None          # [(doc_key, link doc, total transform)]

    # ---------------- geometric search of the room surfaces ----------------
    def _finish_candidates(self, dk, d):
        """Placed model elements whose type (or own) Keynote has a finish
        prefix. Built once per document; only these are tested geometrically,
        so the per-room search stays cheap on big models."""
        if dk in self._cands:
            return self._cands[dk]
        type_ok = set()
        for t in DB.FilteredElementCollector(d).WhereElementIsElementType():
            try:
                if classify(param_str(t.get_Parameter(BIP.KEYNOTE_PARAM)))[1]:
                    type_ok.add(eid_int(t.Id))
            except Exception:
                continue
        ids = List[DB.ElementId]()
        for el in DB.FilteredElementCollector(d).WhereElementIsNotElementType():
            try:
                cat = el.Category
                if cat is None or cat.CategoryType != DB.CategoryType.Model:
                    continue
                tid = el.GetTypeId()
                if (is_valid_id(tid) and eid_int(tid) in type_ok) or \
                        classify(param_str(el.get_Parameter(BIP.KEYNOTE_PARAM)))[1]:
                    ids.Add(el.Id)
            except Exception:
                continue
        self._cands[dk] = ids
        return ids

    def _search_docs(self):
        """[(doc_key, document, transform or None)] - host plus loaded links."""
        if self._links is None:
            self._links = []
            try:
                for inst in DB.FilteredElementCollector(self.doc).OfClass(DB.RevitLinkInstance):
                    ld = inst.GetLinkDocument()
                    if ld is not None:
                        dk = eid_int(inst.Id)
                        self.idx.doc_for(dk)
                        self._links.append((dk, ld, inst.GetTotalTransform()))
            except Exception:
                pass
        return [(0, self.doc, None)] + self._links

    def _link_xf(self, dk):
        for k, _, xf in self._links or []:
            if k == dk:
                return xf
        return None

    def room_zrange(self, room):
        """(base, top) elevation of the room volume, or None."""
        try:
            bb = room.get_BoundingBox(None)
            if bb is not None and bb.Max.Z > bb.Min.Z:
                return bb.Min.Z, bb.Max.Z
        except Exception:
            pass
        return None

    def elem_zrange(self, dk, el):
        """(low, high) elevation of an element in host coordinates, or None."""
        try:
            bb = el.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is None:
            return None
        lo, hi = bb.Min, bb.Max
        xf = self._link_xf(dk) if dk else None
        if xf is not None:
            lo, hi = xf.OfPoint(lo), xf.OfPoint(hi)
        return min(lo.Z, hi.Z), max(lo.Z, hi.Z)

    def search_zone(self, loops, zrange):
        """Room outline extruded around the room height, as a Solid."""
        base, top = zrange
        z0 = base - SEARCH_BELOW_M / 0.3048
        z1 = top + SEARCH_ABOVE_M / 0.3048
        curve_loops = []
        for loop in loops:
            curves = [seg.GetCurve() for seg in loop]
            if not curves:
                continue
            cl = DB.CurveLoop()
            for c in curves:
                cl.Append(c)
            move = DB.Transform.CreateTranslation(DB.XYZ(0, 0, z0 - curves[0].GetEndPoint(0).Z))
            curve_loops.append(DB.CurveLoop.CreateViaTransform(cl, move))
        if not curve_loops:
            return None
        for attempt in (curve_loops, curve_loops[:1]):     # all loops, else outer only
            try:
                return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
                    List[DB.CurveLoop](attempt), DB.XYZ.BasisZ, z1 - z0)
            except Exception:
                continue
        return None

    def elements_on_surfaces(self, solid):
        """(doc_key, element) for every finish-keynoted element intersecting
        the search zone, host and linked models."""
        found = []
        for dk, d, xf in self._search_docs():
            ids = self._finish_candidates(dk, d)
            if ids.Count == 0:
                continue
            s = solid if xf is None else DB.SolidUtils.CreateTransformed(solid, xf.Inverse)
            bb = s.GetBoundingBox()
            p0, p1 = bb.Transform.OfPoint(bb.Min), bb.Transform.OfPoint(bb.Max)
            outline = DB.Outline(DB.XYZ(min(p0.X, p1.X), min(p0.Y, p1.Y), min(p0.Z, p1.Z)),
                                 DB.XYZ(max(p0.X, p1.X), max(p0.Y, p1.Y), max(p0.Z, p1.Z)))
            col = (DB.FilteredElementCollector(d, ids)
                   .WherePasses(DB.BoundingBoxIntersectsFilter(outline))
                   .WherePasses(DB.ElementIntersectsSolidFilter(s)))
            for el in col:
                found.append((dk, el))
        return found

    # ---------------- family instances located in a room ----------------
    def _inside_candidates(self):
        """Family instances whose own/type Keynote has a finish prefix. The
        Keynote check is cheap (cached per type), so FamilyInstance.Room - a
        point-in-room test - only runs for the few instances that matter."""
        if self._fi_candidates is None:
            out = []
            for fi in (DB.FilteredElementCollector(self.doc)
                       .OfClass(DB.FamilyInstance)
                       .WhereElementIsNotElementType()):
                kn, _ = self.idx.element_keynote(0, fi)
                if kn and classify(kn)[1]:
                    out.append(fi)
            self._fi_candidates = out
        return self._fi_candidates

    def inside_room(self, room):
        try:
            phase_id = room.get_Parameter(BIP.ROOM_PHASE).AsElementId()
        except Exception:
            phase_id = None
        pk = eid_int(phase_id) if is_valid_id(phase_id) else -1
        if pk not in self._inside:
            phase = self.doc.GetElement(phase_id) if pk != -1 else None
            by_room = {}
            for fi in self._inside_candidates():
                r = None
                try:
                    r = fi.get_Room(phase) if phase is not None else fi.Room
                except Exception:
                    try:
                        r = fi.Room
                    except Exception:
                        r = None
                if r is not None:
                    by_room.setdefault(eid_int(r.Id), []).append(fi)
            self._inside[pk] = by_room
        return self._inside[pk].get(eid_int(room.Id), [])

    # ---------------- raw relationships ----------------
    def _ref(self, host_or_link_id, linked_id):
        """(doc_key, element) from a boundary reference."""
        if is_valid_id(linked_id):
            dk = eid_int(host_or_link_id)
            return dk, self.idx.element(dk, linked_id)
        return 0, self.idx.element(0, host_or_link_id)

    def collect(self, room):
        raw = {
            "elements": OrderedDict(),   # (doc_key, id) -> element record
            "walls": {},                 # (doc_key, id) -> {"pts", "sides"} for sweeps
            "integral": [],              # sweeps built into wall types
            "separation_lines": 0,
            "top_total": 0.0, "top_free": 0.0, "bottom_free": 0.0,
            "calc_error": None, "zrange": self.room_zrange(room), "search_error": None,
        }

        def touch(dk, el, rel):
            k = (dk, eid_int(el.Id))
            e = raw["elements"].get(k)
            if e is None:
                e = {"doc_key": dk, "elem": el, "rels": set(), "mats": set(), "top_area": 0.0}
                raw["elements"][k] = e
            e["rels"].add(rel)
            return e

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
                if cat_int(el) == CAT_ROOM_SEP:
                    raw["separation_lines"] += 1
                    continue
                touch(dk, el, "boundary")
                if isinstance(el, DB.Wall):
                    w = raw["walls"].setdefault((dk, eid_int(el.Id)), {"pts": [], "sides": set()})
                    try:
                        crv = seg.GetCurve()
                        w["pts"].extend([crv.GetEndPoint(0), crv.GetEndPoint(1)])
                    except Exception:
                        pass

        # 2) 3D boundary: every element bounding the room volume, with the
        #    material of the face that touches the room
        try:
            res = self.calc.CalculateSpatialElementGeometry(room)
            solid = res.GetGeometry()
        except Exception as ex:
            raw["calc_error"] = to_unicode(ex)
            solid = None
        if solid is not None:
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
                    rel = _subface_rel(sub.SubfaceType)
                    try:
                        area = sub.GetSubface().Area
                    except Exception:
                        area = 0.0
                    lid = sub.SpatialBoundaryElement
                    if is_valid_id(lid.LinkInstanceId):
                        dk, el = self._ref(lid.LinkInstanceId, lid.LinkedElementId)
                    else:
                        dk, el = self._ref(lid.HostElementId, None)
                    if rel == "top":
                        raw["top_total"] += area
                    if el is None:
                        if rel == "top":
                            raw["top_free"] += area
                        elif rel == "bottom":
                            raw["bottom_free"] += area
                        continue
                    try:
                        bface = sub.GetBoundingElementFace()
                    except Exception:
                        bface = None
                    e = touch(dk, el, rel)
                    mat = face_material_id(self.idx.doc_for(dk), el, bface)
                    if mat is not None:
                        e["mats"].add(mat)
                    if rel == "top":
                        e["top_area"] += area
                    if rel == "side" and isinstance(el, DB.Wall):
                        side = wall_side_facing(el, face_normal(bface) if bface else None)
                        w = raw["walls"].setdefault((dk, eid_int(el.Id)), {"pts": [], "sides": set()})
                        if side is not None:
                            w["sides"].add(side)

        # 3) sweeps on the room-facing side of the bounding walls (host model)
        for (dk, wid), w in sorted(raw["walls"].items()):
            if dk or not w["sides"]:
                continue
            wall = raw["elements"][(dk, wid)]["elem"]
            members = [wall]
            try:
                if wall.IsStackedWall:
                    members = [self.idx.element(0, m) for m in wall.GetStackedWallMemberIds()]
                    members = [m for m in members if m is not None]
            except Exception:
                pass
            for m in members:
                for sw in self.sweeps.get(eid_int(m.Id), []):
                    if sw["side"] in w["sides"] and sweep_overlaps_room(m, sw["bb"], w["pts"]):
                        e = touch(0, sw["elem"], "sweep")
                        if is_valid_id(sw["mat"]):
                            e["mats"].add(sw["mat"])
                try:
                    cs = m.WallType.GetCompoundStructure()
                    infos = list(cs.GetWallSweepsInfo(DB.WallSweepType.Sweep)) if cs else []
                except Exception:
                    infos = []
                for info in infos:
                    if info.WallSide in w["sides"]:
                        raw["integral"].append({"wall": m, "mat": info.MaterialId})

        # 4) family instances located in the room (e.g. skirting families)
        for fi in self.inside_room(room):
            touch(0, fi, "inside")

        # 5) everything else on the room surfaces: skirtings below the room
        #    computation height, non-bounding finish walls / ceilings,
        #    ceilings above a low room Limit Offset
        if raw["zrange"] is None:
            raw["search_error"] = u"room has no volume bounding box"
        else:
            try:
                zone = self.search_zone(loops, raw["zrange"])
                if zone is None:
                    raw["search_error"] = u"could not build the room search zone from its boundary"
                else:
                    for dk, el in self.elements_on_surfaces(zone):
                        touch(dk, el, "surface")
            except Exception as ex:
                raw["search_error"] = to_unicode(ex)
        return raw

    # ---------------- keynote resolution ----------------
    def element_keys(self, dk, el, mats):
        """-> (list of (keynote, source), why-missing text). Read only."""
        why = []
        found = []
        if self.mode in ("auto", "type"):
            kn, src = self.idx.element_keynote(dk, el)
            if kn:
                found.append((kn, src))
                # a finish-prefixed element/type Keynote is the answer; an
                # unmapped one (e.g. a structural code) lets auto mode go on
                # to the room-facing material
                if self.mode == "type" or classify(kn)[1]:
                    return found, u""
            else:
                tname = self.idx.type_name(dk, el)
                why.append(u"type '{}' has no Keynote".format(tname) if tname else u"no type Keynote")
                if self.mode == "type":
                    return [], u"; ".join(why)
        for mid in sorted(mats or [], key=eid_int):
            kn = self.idx.material_keynote(dk, mid)
            if kn:
                found.append((kn, "material"))
            else:
                why.append(u"material '{}' has no Keynote".format(self.idx.material_name(dk, mid)))
        if not mats and self.mode == "material":
            why.append(u"no room-facing material found")
        return found, u"; ".join(why)

    def _base_item(self, dk, el, rels, cat=None):
        return {
            "id": eid_int(el.Id),
            "cat": cat or cat_name(el),
            "link": self.idx.link_name(dk),
            "rels": [REL_LABEL.get(r, r) for r in sorted(rels)],
        }

    def classify_room(self, raw):
        fins = OrderedDict((k, {"keys": [], "items": [], "warnings": [], "mismatches": [], "notes": []})
                           for k in FINISH_PARAMS)
        issues, others = [], []
        fr_top_area = 0.0
        fr_keys = set()             # elements that contributed a ceiling (FR) value

        room_z = raw["zrange"]

        def height_reject(fk, zr, rels):
            """Why this element can't carry finish fk for THIS room, or None."""
            if zr is not None and room_z is not None:
                base, top = room_z
                mid, tol = (base + top) / 2.0, 0.03     # tol ~ 1 cm
                lo, hi = zr
                if fk == "floor" and lo >= mid:
                    return u"it is above the room (floor of the storey above)"
                if fk == "ceiling" and hi <= mid:
                    return u"it is below the room (ceiling of the storey below)"
                if fk in ("wall", "base") and (lo >= top - tol or hi <= base + tol):
                    return u"it is outside the room height (belongs to another storey)"
                return None
            pos = POSITION_GUARD.get(fk)
            if pos and rels == set([pos]):
                return u"it only bounds the room from {}".format(u"above" if pos == "top" else u"below")
            return None

        def route(norm, fk, item, cat_id, rels, zr=None):
            """Put one classified keynote into its finish; returns True if used."""
            why = height_reject(fk, zr, rels)
            if why:
                fins[fk]["notes"].append(u"{} on {} {} is ignored: {}.".format(
                    norm, item["cat"], item["id"], why))
                return False
            fins[fk]["keys"].append(norm)
            if cat_id not in EXPECTED_CATEGORIES[fk]:
                fins[fk]["mismatches"].append(
                    u"{} found on a {} element ({}) - classified as {} by the prefix rule.".format(
                        norm, item["cat"], item["id"], FINISH_PARAMS[fk]))
            return True

        for key in sorted(raw["elements"]):
            e = raw["elements"][key]
            dk, el, rels = e["doc_key"], e["elem"], e["rels"]
            item = self._base_item(dk, el, rels)
            found, why = self.element_keys(dk, el, e["mats"])
            # elements only found by position (in the room / geometric search)
            # are candidates, not boundaries: no "missing Keynote" noise
            bounding = rels - set(["inside", "surface"])
            if not found:
                if bounding:
                    issues.append(dict(item, msg=u"Keynote missing ({})".format(why)))
                continue
            zr = self.elem_zrange(dk, el)
            used = OrderedDict()
            ignored = []
            for kn, src in found:
                norm, fk = classify(kn)
                if fk is None:
                    ignored.append(norm)
                elif route(norm, fk, item, cat_int(el), rels, zr):
                    used.setdefault(fk, []).append((norm, src))
            for fk, lst in used.items():
                fins[fk]["items"].append(dict(item, kn=unique_sorted(k for k, _ in lst),
                                              src=sorted(set(s for _, s in lst))))
                if fk == "ceiling":
                    fr_top_area += e["top_area"]
                    fr_keys.add(key)
            if ignored and not used and bounding:
                others.append(dict(item, kn=unique_sorted(ignored),
                                   msg=u"prefix is not RD / FR / RE / PI - ignored"))

        # sweeps built into wall types: the only Keynote they can carry is
        # their material's
        for it in raw["integral"]:
            wall = it["wall"]
            item = {"id": eid_int(wall.Id), "cat": u"Wall type sweep", "link": u"",
                    "rels": [u"sweep in wall type"]}
            kn = self.idx.material_keynote(0, it["mat"])
            if not kn:
                issues.append(dict(item, msg=u"Keynote missing (sweep material '{}' has no Keynote)".format(
                    self.idx.material_name(0, it["mat"]))))
                continue
            norm, fk = classify(kn)
            if fk is None:
                others.append(dict(item, kn=[norm], msg=u"prefix is not RD / FR / RE / PI - ignored"))
            elif route(norm, fk, item, CAT_CORNICES, set(["sweep"])):
                fins[fk]["items"].append(dict(item, kn=[norm], src=["integral sweep material"]))

        for fk, fin in fins.items():
            fin["keys"] = unique_sorted(fin["keys"])
            fin["state"] = "ok" if fin["keys"] else "missing"
            fin["notes"] = unique_sorted(fin["notes"])
            fin["mismatches"] = sorted(set(fin["mismatches"]), key=natural_key)
            if self.keynote_texts:
                for k in fin["keys"]:
                    if k not in self.keynote_texts:
                        fin["warnings"].append(u"'{}' is not in the loaded keynote file.".format(k))
        self._diagnose(raw, fins, fr_top_area, fr_keys)
        return fins, issues, others

    def _diagnose(self, raw, fins, fr_top_area, fr_keys):
        """Explain an empty finish - diagnostics only, never a value."""
        def cats_with(rel, skip=()):
            return sorted(set(cat_name(e["elem"]) for k, e in raw["elements"].items()
                              if rel in e["rels"] and k not in skip))
        err = raw["calc_error"]
        if err:
            for fk in ("floor", "ceiling", "base"):
                fins[fk]["notes"].append(u"Room geometry could not be calculated: " + err)
        if raw["search_error"]:
            for fk in FINISH_PARAMS:
                if fins[fk]["state"] == "missing":
                    fins[fk]["notes"].append(u"Geometric search of the room surfaces failed: " + raw["search_error"])
        if fins["wall"]["state"] == "missing":
            if not raw["walls"]:
                fins["wall"]["notes"].append(u"No Wall bounds this room{}.".format(
                    u" (Room Separation Lines only)" if raw["separation_lines"] else u""))
            else:
                fins["wall"]["notes"].append(u"No Keynote starting with RE on the elements around this room.")
        if fins["floor"]["state"] == "missing" and not err:
            below = cats_with("bottom")
            fins["floor"]["notes"].append(
                u"No Keynote starting with PI. Below the room: {}.".format(u", ".join(below))
                if below else
                u"Room bottom is not bounded by any element. Check the floor is Room Bounding "
                u"and its top is not below the room base.")
        if not err:
            total = raw["top_total"]
            if fins["ceiling"]["state"] == "missing":
                above = cats_with("top")
                fins["ceiling"]["notes"].append(
                    u"No Keynote starting with FR on the room surfaces or up to {:.2f} m above the room top "
                    u"(room top bounded by: {}). If the ceiling is higher, raise the room Upper Limit / "
                    u"Limit Offset or SEARCH_ABOVE_M.".format(SEARCH_ABOVE_M, u", ".join(above) or u"nothing"))
            elif fr_top_area > 0 and total > 0 and fr_top_area / total < CEILING_FULL_COVERAGE:
                rest = []
                other_top = cats_with("top", fr_keys)
                if other_top:
                    rest.append(u", ".join(other_top))
                if raw["top_free"] > 0:
                    rest.append(u"unbounded (room top below the ceiling)")
                fins["ceiling"]["warnings"].append(
                    u"FR elements cover {:.0f}% of the room top; the rest is {}.".format(
                        100.0 * fr_top_area / total, u" / ".join(rest) or u"other elements"))
        if fins["base"]["state"] == "missing":
            fins["base"]["notes"].append(
                u"No Keynote starting with RD (checked: bounding elements, sweeps on the room-facing "
                u"side of the bounding walls, sweeps in wall types, family instances in the room, "
                u"elements on the room surfaces).")
        linked = len([1 for (dk, _) in raw["walls"] if dk])
        if linked:
            fins["base"]["notes"].append(u"{} linked wall(s): their sweeps are not evaluated.".format(linked))
        if raw["separation_lines"] and raw["walls"]:
            fins["wall"]["notes"].append(u"{} boundary segment(s) are Room Separation Lines (no finish).".format(
                raw["separation_lines"]))

    # ---------------- one room ----------------
    def analyze(self, room):
        rec = {"room": room, "boundary": "ok", "finishes": {}, "issues": [], "others": [], "error": None}
        try:
            if room.Location is None:
                rec["boundary"] = "unplaced"
                return rec
            if room.Area <= 0:
                rec["boundary"] = "unenclosed"
                return rec
            raw = self.collect(room)
            if not raw["elements"] and not raw["separation_lines"]:
                rec["boundary"] = "no_segments"
            rec["finishes"], rec["issues"], rec["others"] = self.classify_room(raw)
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
    # notes are informational (e.g. a separation line); an empty finish, a
    # warning or a prefix/category mismatch makes the room incomplete
    if all(f["state"] == "ok" and not f["warnings"] and not f["mismatches"]
           for f in rec["finishes"].values()):
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
        for key in FINISH_PARAMS:
            fin = rec["finishes"].get(key)
            ch = rec.get("changes", {}).get(key, {})
            if fin is None:
                fins[key] = {"state": "n/a", "keys": [], "value": u"", "current": ch.get("current", u""),
                             "items": [], "notes": [], "warnings": [], "mismatches": [],
                             "change": ch.get("kind", u"")}
                continue
            fins[key] = {
                "state": fin["state"], "keys": fin["keys"], "value": join_keys(fin["keys"]),
                "current": ch.get("current", u""), "change": ch.get("kind", u""),
                "items": fin["items"], "notes": fin["notes"], "warnings": fin["warnings"],
                "mismatches": fin["mismatches"],
            }
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
            "finishes": fins, "issues": rec.get("issues", []), "others": rec.get("others", []),
        })
    return {
        "project": to_unicode(doc.Title),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "keynoteSource": dict(KEYNOTE_SOURCES).get(mode, mode),
        "boundaryLocation": to_unicode(BOUNDARY_LOCATION),
        "separator": SEPARATOR,
        "prefixRules": OrderedDict((p, FINISH_PARAMS[f]) for p, f in PREFIX_RULES.items()),
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
        issues = [(r, iss) for r in self.records for iss in r.get("issues", [])]
        mismatches = [(r, m) for r in self.records for f in r["finishes"].values() for m in f["mismatches"]]
        output.print_md(u"**[4/6] Keynotes collected** (source: {}; classified by prefix: {}). "
                        u"{} element(s) without Keynote, {} prefix/category mismatch(es).".format(
                            dict(KEYNOTE_SOURCES)[self.mode],
                            u", ".join(u"{} = {}".format(p, FINISH_PARAMS[f]) for p, f in PREFIX_RULES.items()),
                            len(issues), len(mismatches)))
        if not self.keynote_texts:
            output.print_md(u"_Keynote file not loaded or empty - keys are not validated against it._")
        for r, iss in issues[:60]:
            output.print_md(u"- Room `{}`: {} {} - {}".format(
                r["info"]["number"], iss["cat"], log_link(iss), iss["msg"]))
        if len(issues) > 60:
            output.print_md(u"- ... and {} more (see the HTML report).".format(len(issues) - 60))
        for r, m in mismatches[:40]:
            output.print_md(u"- **Mismatch** Room `{}`: {}".format(r["info"]["number"], m))
        if len(mismatches) > 40:
            output.print_md(u"- ... and {} more mismatches (see the HTML report).".format(len(mismatches) - 40))

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
                if rec["finishes"][key]["mismatches"]:
                    note = (note + u"; " if note else u"") + u"prefix/category mismatch - see report"
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
