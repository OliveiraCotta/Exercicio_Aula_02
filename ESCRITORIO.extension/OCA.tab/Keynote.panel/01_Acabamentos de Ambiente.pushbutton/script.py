# -*- coding: utf-8 -*-
"""Room Finish Keynote Automation.

Reads the Keynotes of the elements that bound each Room and writes them to the
Room finish parameters (shared parameters "Acabamento de Soleira/Rodateto/
Piso/Teto/Rodapé/Parede 01..05"). Each finish has numbered fields that its
unique Keynotes fill in order (see FINISH_SLOTS / compute_changes).

    READ -> ANALYZE -> PREVIEW -> USER CONFIRMATION -> TRANSACTION -> WRITE -> REPORT

Which parameter a value goes to is decided ONLY by the Keynote prefix
(case-insensitive, value normalised to upper case), never by the element's
category:

  SL -> Soleira   RT -> Rodateto   PI -> Piso   FR / CB -> Teto
  RD -> Rodapé    RE -> Parede     anything else: ignored

The category is only used afterwards to REPORT mismatches (e.g. RE01 on a
Floor). Nothing but the Room finish parameters is ever written.

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
  The zone is grown 2 cm outwards so elements flush with the room faces
  count too, and every Wall Sweep and Ceiling is tested even without a
  finish-prefix type Keynote, so one without a Keynote is reported.
  - Room containment (Document.GetRoomAtPoint) for every element with a
    finish Keynote that no room claimed above, and for EVERY wall sweep:
    points are sampled inside the element's own geometry along its whole
    length, so a sweep running through several rooms goes to all of them.
    Elements with a finish Keynote that end up in no room are listed in the
    log and the report ("not assigned to any room").
  Exception to the prefix rule: a Wall Sweep ("Moldura de parede") fills
  Rodapé (BASE_CATEGORIES) unless its Keynote starts with RT or RD.
  Height guard: a value is only used if the element sits where that finish
  can be for THIS room - PI starts below the room's mid-height, FR is not
  entirely below the room floor, RE / RD overlap the room's height. This keeps the floor finish,
  skirting and walls of the storey above out of the room below. Rejected
  values are reported, not written.

Keynote lookup, per element (KEYNOTE_MODE = "type": only step 1 is used;
the other modes in KEYNOTE_SOURCES are kept in code but not offered in the UI):
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

__title__ = "Acabamentos\nde Ambiente"
__doc__ = u"""Preenche automaticamente os acabamentos dos ambientes com base nas Keynotes dos elementos ao redor e ajusta o tipo do identificador conforme o número de linhas de acabamento.

Analise, confira a pré-visualização e confirme."""

import os
import re
import math
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
output.set_title(u"Acabamentos de Ambiente - log")


# ==================================================================
# Settings
# ==================================================================
# Room finish parameters - shared parameters "Parametros de Ambiente"
# (OCA_Parametros_Template.txt). Each finish has fields 01 to 05; its unique
# Keynotes (natural order) fill them in sequence: 01, 02, 03... Looked up by
# GUID, then by exact name.
FINISH_SLOTS = OrderedDict([
    ("sill", [(u"Acabamento de Soleira 01", "53c7ce2b-cb20-4fcf-88ca-92b096d7ac16"),
              (u"Acabamento de Soleira 02", "cd25c084-0b6d-4c3e-b62d-db150376a95f"),
              (u"Acabamento de Soleira 03", "73f3ae00-cda5-4427-9de8-10161b23bed7"),
              (u"Acabamento de Soleira 04", "53701d9a-fc54-437a-b7bd-cdfb346db9d0"),
              (u"Acabamento de Soleira 05", "25070ec2-470d-4d36-b812-089c253fffa7")]),
    ("crown", [(u"Acabamento de Rodateto 01", "af22da2c-43fb-4397-bb0b-480195daa6e1"),
               (u"Acabamento de Rodateto 02", "c0cc8f63-9088-4b6a-b88b-21bfb0747fe2"),
               (u"Acabamento de Rodateto 03", "06179583-d511-4b0e-ab9d-26fe47ba31e2"),
               (u"Acabamento de Rodateto 04", "cf9af1ce-e72b-4fd7-bc0f-d9276ac0fe2f"),
               (u"Acabamento de Rodateto 05", "d789f873-bda6-4362-b446-49d40ca1ab85")]),
    ("floor", [(u"Acabamento de Piso 01", "cc2fa173-836d-4b2f-8ab2-907b932d05c0"),
               (u"Acabamento de Piso 02", "ce590775-fef4-460b-86c2-6798e9b44301"),
               (u"Acabamento de Piso 03", "2100cb48-c4f4-4e4c-94a5-59bfc0ee54f4"),
               (u"Acabamento de Piso 04", "044158c5-0085-49d5-9b5e-c54b363b7f67"),
               (u"Acabamento de Piso 05", "51b8c21a-dd7d-4f2b-8dfd-ad065482421c")]),
    ("ceiling", [(u"Acabamento de Teto 01", "f5a77c4e-a48c-44f0-89ff-fe8ea8242d3e"),
                 (u"Acabamento de Teto 02", "49faa0d2-3f47-4167-b438-3a1f7a817d2e"),
                 (u"Acabamento de Teto 03", "32dfa3eb-1d00-4564-8bb9-2e5878fb3ef4"),
                 (u"Acabamento de Teto 04", "4690b55e-dbca-493a-b898-79bcaeea4544"),
                 (u"Acabamento de Teto 05", "9d6cdfae-2b13-4515-867f-66ea1ed7dd35")]),
    ("base", [(u"Acabamento de Rodapé 01", "7a09d8e3-4501-4ec3-b81f-e1f225b387ee"),
              (u"Acabamento de Rodapé 02", "3ce3c29f-4607-41c7-b2b9-80c856043b55"),
              (u"Acabamento de Rodapé 03", "659e50fb-df51-44f1-a544-726c6cc29485"),
              (u"Acabamento de Rodapé 04", "6949f296-1c82-4601-b86e-842fe38b8aa4"),
              (u"Acabamento de Rodapé 05", "e02b1c7a-3f10-4af3-9f30-82c3b152e08d")]),
    ("wall", [(u"Acabamento de Parede 01", "896c4028-fc02-4e4e-b01e-67db6b4914bb"),
              (u"Acabamento de Parede 02", "5c6598d4-596b-4578-96ac-092e71a7d338"),
              (u"Acabamento de Parede 03", "481190fa-1b6d-4834-bbff-e895e01a4dd2"),
              (u"Acabamento de Parede 04", "4a1e33ed-f6cd-4823-b3eb-e84829ecce16"),
              (u"Acabamento de Parede 05", "bf744c4b-1ff6-44a9-ab50-e445f0521559")]),
])
SEPARATOR = u" / "

# Rule A - row placeholder. Field NN of every finish is row NN of the room
# identifier. When at least one row has a real value, the empty fields of
# rows 1..N get this text, so the tag prints a dash instead of a blank.
PLACEHOLDER = u"-"

# Rule B - room identifier (Room Tag) type per number of rows. The tag
# family is not named in code: any Room Tag family whose types carry
# "REVESTIMENTOS NN" / "REVESTIMENTOS NX" anywhere in the name is used
# ("OCA_REVESTIMENTOS_3X - 1:50" too), and only tags already placed in the
# project are switched (nothing is created).
TAG_ROWS_RE = re.compile(u"(REVESTIMENTOS?[\\s_.-]*)(\\d{1,2})(\\s*X)?", re.IGNORECASE)
TAG_MAX_ROWS = 5

# Boundary location for the 2D wall boundaries. Finish = the room-facing
# face of the wall, which is what a finish schedule describes.
BOUNDARY_LOCATION = DB.SpatialElementBoundaryLocation.Finish

# The ONLY rule that decides which Room finish a Keynote fills.
PREFIX_RULES = OrderedDict([
    ("SL", "sill"),
    ("RT", "crown"),
    ("PI", "floor"),
    ("FR", "ceiling"),
    ("CB", "ceiling"),
    ("RD", "base"),
    ("RE", "wall"),
])

# FR elements must cover this share of the room's top surface, else warning.
CEILING_FULL_COVERAGE = 0.98

# The keynote is always read from the element itself (its own Keynote, else
# its type's - where Revit stores it). Material Keynotes are not used.
KEYNOTE_MODE = "type"

KEYNOTE_SOURCES = [
    ("auto", u"Keynote do elemento/tipo e, se não houver, do material voltado ao ambiente (recomendado)"),
    ("type", u"Keynote do elemento"),
    ("material", u"Somente Keynote do material voltado ao ambiente"),
]

FINISH_LABELS = OrderedDict([
    ("sill", u"Soleira"),
    ("crown", u"Rodateto"),
    ("floor", u"Piso"),
    ("ceiling", u"Teto"),
    ("base", u"Rodapé"),
    ("wall", u"Parede"),
])


def finish_prefixes(key):
    return [p for p, f in PREFIX_RULES.items() if f == key]


def finish_label(key):
    """'Parede (RE)' / 'Teto (FR/CB)'"""
    return u"{} ({})".format(FINISH_LABELS[key], u"/".join(finish_prefixes(key)))


def _bic(name):
    try:
        return int(getattr(DB.BuiltInCategory, name))
    except Exception:
        return None


CAT_ROOM_SEP = _bic("OST_RoomSeparationLines")
CAT_CORNICES = _bic("OST_Cornices")      # Wall Sweeps ("Molduras de parede")
CAT_CEILINGS = _bic("OST_Ceilings")

# Categories whose Keynote fills Rodapé (treated as RD) unless its prefix is
# RT (rodateto): wall sweeps are skirtings / mouldings. Sweeps built into
# wall types follow the same rule.
BASE_CATEGORIES = set([CAT_CORNICES])

# Categories tested by the geometric search even without a finish-prefix
# type Keynote, so a sweep or ceiling whose Keynote is only on its material
# (or missing) is still found and reported.
ALWAYS_SEARCH = set([CAT_CORNICES, CAT_CEILINGS])

# Categories each finish is EXPECTED on. Used only to report mismatches -
# never to classify. Edit freely to match your modelling standard.
EXPECTED_CATEGORIES = {
    "sill": set([_bic("OST_Floors"), _bic("OST_GenericModel")]),
    "crown": set([CAT_CORNICES, _bic("OST_Walls"), _bic("OST_GenericModel")]),
    "floor": set([_bic("OST_Floors")]),
    "ceiling": set([CAT_CEILINGS, _bic("OST_Roofs")]),
    "base": set([CAT_CORNICES, _bic("OST_Walls"), _bic("OST_GenericModel")]),
    "wall": set([_bic("OST_Walls")]),
}

# Finishes that may come from an element that only touches the room from
# outside (flush): the ceiling, and the door sill (soleira), which sits in the
# door opening, outside the volume of both rooms.
FLUSH_FINISHES = set(["ceiling", "sill"])

# Geometric search zone around the room surfaces: the finish outline extruded
# from SEARCH_BELOW_M under the room base to SEARCH_ABOVE_M over the room top
# (so a ceiling above a room whose Limit Offset is too low is still found).
SEARCH_BELOW_M = 0.02
SEARCH_ABOVE_M = 1.00
# The outline is also grown outwards by this much, so elements that are only
# flush with the room faces (touching, zero overlap) are found too.
SEARCH_OUTWARD_M = 0.02

# Fallback when an element has no bounding box (height guard not possible):
# a Keynote of an element that ONLY bounds the room from this side belongs to
# the neighbouring room (floor slab above, ceiling below) - not written.
POSITION_GUARD = {"floor": "top", "sill": "top", "ceiling": "bottom", "crown": "bottom"}


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
        return u"(sem categoria)"


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


def is_real(value):
    """A finish value found by the command - not empty, not the '-' placeholder."""
    v = (value or u"").strip()
    return bool(v) and v != PLACEHOLDER


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
            self._link_names[doc_key] = to_unicode(inst.Name) if inst else u"vínculo"
            ldoc = inst.GetLinkDocument() if inst else None   # None when unloaded
        except Exception:
            ldoc = None
        self._docs[doc_key] = ldoc
        return ldoc

    def link_name(self, doc_key):
        if not doc_key:
            return u""
        self.doc_for(doc_key)
        return self._link_names.get(doc_key, u"vínculo")

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
            result = (kn, u"elemento")
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
                    result = (self._type_kn[tk], u"tipo")
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

    def element_materials(self, doc_key, elem):
        """All materials of an element, painted ones included."""
        out = set()
        for painted in (False, True):
            try:
                for mid in elem.GetMaterialIds(painted):
                    if is_valid_id(mid):
                        out.add(mid)
            except Exception:
                pass
        return out

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
    """One finish field (e.g. 'Acabamento de Parede 02') on the rooms."""
    def __init__(self, key, slot, name, param, guid=None):
        self.key = key
        self.slot = slot
        self.name = name
        self.kind = param_kind(param)      # "built-in" | "shared" | "project"
        self.definition = param.Definition
        self.guid = guid                   # System.Guid when found by GUID

    def get(self, room):
        if self.guid is not None:
            return room.get_Parameter(self.guid)
        return room.get_Parameter(self.definition)

    def read(self, room):
        return param_str(self.get(room))


KIND_LABELS = {"built-in": u"nativo", "shared": u"compartilhado", "project": u"de projeto"}


def param_kind(p):
    try:
        bip = p.Definition.BuiltInParameter
        if bip != BIP.INVALID:
            return "built-in"
    except Exception:
        pass
    return "shared" if p.IsShared else "project"


def resolve_targets(sample_room):
    """Find every finish field on the rooms: by GUID (shared parameter file),
    else by exact name. -> (targets {key: [ParamTarget or None per field]},
    problems [str]). Nothing is created."""
    from System import Guid
    by_name = {}
    for p in sample_room.Parameters:
        try:
            by_name.setdefault(to_unicode(p.Definition.Name), []).append(p)
        except Exception:
            continue
    targets, problems = OrderedDict(), []
    for key, slots in FINISH_SLOTS.items():
        found = []
        for i, (name, guid) in enumerate(slots):
            p, g = None, None
            try:
                g = Guid(guid)
                p = sample_room.get_Parameter(g)
            except Exception:
                p = None
            if p is None:
                g = None
                cands = [c for c in by_name.get(name, []) if c.StorageType == DB.StorageType.String]
                p = cands[0] if cands else None
            if p is None:
                problems.append(u"Parâmetro '{}' não foi encontrado.".format(name))
                found.append(None)
            elif p.StorageType != DB.StorageType.String or p.IsReadOnly:
                problems.append(u"Parâmetro '{}' existe, mas não é um parâmetro de texto editável.".format(name))
                found.append(None)
            else:
                found.append(ParamTarget(key, i, name, p, g))
        targets[key] = found
    return targets, problems


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


def classify_element(keynote, cat_id):
    """classify() plus the category exception: a Keynote on a BASE_CATEGORIES
    element (wall sweep) goes to Rodapé unless its prefix is RT or RD.
    -> (normalised keynote, finish key or None, forced)"""
    norm, finish = classify(keynote)
    if norm and cat_id in BASE_CATEGORIES and finish not in ("base", "crown"):
        return norm, "base", True
    return norm, finish, False


# ==================================================================
# Wall sweeps on bounding walls
# ==================================================================
def build_sweep_index(d):
    """host wall id -> [sweep dict]. One pass over all WallSweep elements
    (sweeps and reveals - the Keynote prefix decides what they are)."""
    idx = {}
    sweeps = []
    for bic in ("OST_Cornices", "OST_Reveals"):          # by category, not OfClass
        try:
            sweeps.extend(DB.FilteredElementCollector(d)
                          .OfCategory(getattr(DB.BuiltInCategory, bic))
                          .WhereElementIsNotElementType().ToElements())
        except Exception:
            continue
    for sw in sweeps:
        if not isinstance(sw, DB.WallSweep):
            continue
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


def touch_element(raw, dk, el, rel):
    """Add (or extend) an element record in a room's raw element set."""
    k = (dk, eid_int(el.Id))
    e = raw["elements"].get(k)
    if e is None:
        e = {"doc_key": dk, "elem": el, "rels": set(), "mats": set(), "top_area": 0.0}
        raw["elements"][k] = e
    e["rels"].add(rel)
    return e


# ==================================================================
# Sampling points inside an element (for Document.GetRoomAtPoint)
# ==================================================================
SAMPLE_STEP_FT = 0.50 / 0.3048      # one sample every ~50 cm along each face
SAMPLE_INSET_FT = 0.005 / 0.3048    # pushed 5 mm into the element
SAMPLE_CAP = 300                    # per element


def element_solids(el):
    opts = DB.Options()
    opts.ComputeReferences = False
    opts.IncludeNonVisibleObjects = False
    try:
        opts.DetailLevel = DB.ViewDetailLevel.Fine
    except Exception:
        pass
    out = []

    def walk(geo, depth):
        if geo is None or depth > 3:
            return
        for g in geo:
            if isinstance(g, DB.Solid):
                if g.Volume > 1e-9:
                    out.append(g)
            elif isinstance(g, DB.GeometryInstance):
                walk(g.GetInstanceGeometry(), depth + 1)
    try:
        walk(el.get_Geometry(opts), 0)
    except Exception:
        pass
    return out


def sample_points(el):
    """Points just inside the element's solids, spread over every face so a
    long element is sampled along its whole length."""
    pts = []
    for solid in element_solids(el):
        for f in solid.Faces:
            try:
                bb = f.GetBoundingBox()
                du, dv = bb.Max.U - bb.Min.U, bb.Max.V - bb.Min.V
                nu = int(min(60, max(1, math.ceil(abs(du) / SAMPLE_STEP_FT))))
                nv = int(min(60, max(1, math.ceil(abs(dv) / SAMPLE_STEP_FT))))
                while nu * nv > 120:
                    if nu >= nv:
                        nu -= 1
                    else:
                        nv -= 1
                for i in range(nu):
                    for j in range(nv):
                        uv = DB.UV(bb.Min.U + du * (i + 0.5) / nu, bb.Min.V + dv * (j + 0.5) / nv)
                        if not f.IsInside(uv):
                            continue
                        n = f.ComputeNormal(uv)
                        pts.append(f.Evaluate(uv).Subtract(n.Multiply(SAMPLE_INSET_FT)))
            except Exception:
                continue
            if len(pts) >= SAMPLE_CAP:
                return pts
    if not pts:
        try:
            bb = el.get_BoundingBox(None)
            if bb is not None:
                pts.append(DB.XYZ((bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0,
                                  (bb.Min.Z + bb.Max.Z) / 2.0))
        except Exception:
            pass
    return pts


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
    "boundary": u"contorno do ambiente",
    "side": u"lateral do volume do ambiente",
    "top": u"acima do ambiente",
    "bottom": u"abaixo do ambiente",
    "sweep": u"moldura na parede delimitadora",
    "inside": u"família dentro do ambiente",
    "surface": u"nas superfícies do ambiente (geometria)",
    "flush": u"faceando o ambiente (geometria)",
    "contained": u"dentro do ambiente (ponto no ambiente)",
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
    def __init__(self, d, mode, search_above_m=None):
        self.doc = d
        self.mode = mode
        self.search_above_m = SEARCH_ABOVE_M if search_above_m is None else search_above_m
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
                if eid_int(cat.Id) in ALWAYS_SEARCH or \
                        (is_valid_id(tid) and eid_int(tid) in type_ok) or \
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
        """(grown, inner): the room outline extruded around the room height,
        grown outwards by SEARCH_OUTWARD_M (grown) and as-is (inner). An
        element hitting only the grown zone is flush with the room from
        outside. None if the outline cannot be extruded."""
        base, top = zrange
        z0 = base - SEARCH_BELOW_M / 0.3048
        z1 = top + self.search_above_m / 0.3048
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
        def extrude(sets):
            for attempt in sets:                        # all loops, else outer only
                try:
                    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
                        List[DB.CurveLoop](attempt), DB.XYZ.BasisZ, z1 - z0)
                except Exception:
                    continue
            return None
        inner = extrude([curve_loops, curve_loops[:1]])
        if inner is None:
            return None
        grown_loops = self._grow(curve_loops)
        grown = extrude([grown_loops, grown_loops[:1]]) if grown_loops else None
        return (grown or inner), inner

    def _grow(self, loops):
        """Loops offset outwards by SEARCH_OUTWARD_M (islands shrink by the
        same amount). The offset sign that lengthens the outer loop is the
        outward one. None if the offset is not possible."""
        off = SEARCH_OUTWARD_M / 0.3048
        if off <= 0:
            return None
        try:
            outer_len = loops[0].GetExactLength()
        except Exception:
            return None
        for sign in (1.0, -1.0):
            try:
                test = DB.CurveLoop.CreateViaOffset(loops[0], sign * off, DB.XYZ.BasisZ)
                if test.GetExactLength() > outer_len:
                    return [DB.CurveLoop.CreateViaOffset(cl, sign * off, DB.XYZ.BasisZ) for cl in loops]
            except Exception:
                continue
        return None

    def elements_on_surfaces(self, zone):
        """(doc_key, element, flush_only) for every candidate element hitting
        the grown zone, host and linked models. flush_only = it only touches
        the room from outside (misses the un-grown zone)."""
        grown, inner = zone
        found = []
        for dk, d, xf in self._search_docs():
            ids = self._finish_candidates(dk, d)
            if ids.Count == 0:
                continue
            hits = self._hits(d, ids, grown, xf)
            if not hits:
                continue
            if inner is grown:
                inside = set(eid_int(el.Id) for el in hits)
            else:
                hit_ids = List[DB.ElementId]()
                for el in hits:
                    hit_ids.Add(el.Id)
                inside = set(eid_int(el.Id) for el in self._hits(d, hit_ids, inner, xf))
            for el in hits:
                found.append((dk, el, eid_int(el.Id) not in inside))
        return found

    def _hits(self, d, ids, solid, xf):
        """Elements of ids (in document d) intersecting solid (host coords)."""
        s = solid if xf is None else DB.SolidUtils.CreateTransformed(solid, xf.Inverse)
        bb = s.GetBoundingBox()
        p0, p1 = bb.Transform.OfPoint(bb.Min), bb.Transform.OfPoint(bb.Max)
        outline = DB.Outline(DB.XYZ(min(p0.X, p1.X), min(p0.Y, p1.Y), min(p0.Z, p1.Z)),
                             DB.XYZ(max(p0.X, p1.X), max(p0.Y, p1.Y), max(p0.Z, p1.Z)))
        return list(DB.FilteredElementCollector(d, ids)
                    .WherePasses(DB.BoundingBoxIntersectsFilter(outline))
                    .WherePasses(DB.ElementIntersectsSolidFilter(s)))

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
            return touch_element(raw, dk, el, rel)

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
            raw["search_error"] = u"o ambiente não tem caixa de volume"
        else:
            try:
                zone = self.search_zone(loops, raw["zrange"])
                if zone is None:
                    raw["search_error"] = u"não foi possível montar a zona de busca a partir do contorno"
                else:
                    for dk, el, flush_only in self.elements_on_surfaces(zone):
                        touch(dk, el, "flush" if flush_only else "surface")
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
                why.append(u"o tipo '{}' não tem Keynote".format(tname) if tname else u"tipo sem Keynote")
                if self.mode == "type":
                    return [], u"; ".join(why)
        for mid in sorted(mats or [], key=eid_int):
            kn = self.idx.material_keynote(dk, mid)
            if kn:
                found.append((kn, u"material"))
            else:
                why.append(u"o material '{}' não tem Keynote".format(self.idx.material_name(dk, mid)))
        if not mats and self.mode == "material":
            why.append(u"nenhum material voltado ao ambiente encontrado")
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
                           for k in FINISH_SLOTS)
        issues, others = [], []
        fr_top_area = 0.0
        fr_keys = set()             # elements that contributed a ceiling (FR) value

        room_z = raw["zrange"]

        def height_reject(fk, zr, rels):
            """Why this element can't carry finish fk for THIS room, or None."""
            if rels == set(["flush"]) and fk not in FLUSH_FINISHES:
                return u"apenas faceia o ambiente pelo lado de fora - aceito somente para teto e soleira"
            if zr is not None and room_z is not None:
                base, top = room_z
                mid, tol = (base + top) / 2.0, 0.03     # tol ~ 1 cm
                lo, hi = zr
                if fk in ("floor", "sill") and lo >= mid:
                    return u"está acima do ambiente (pertence ao pavimento de cima)"
                if fk == "ceiling" and hi <= base + tol:
                    return u"está abaixo do piso do ambiente (forro do pavimento de baixo)"
                if fk == "crown" and hi <= mid:
                    return u"está na metade de baixo do ambiente - rodateto fica junto ao teto"
                if fk in ("wall", "base") and (lo >= top - tol or hi <= base + tol):
                    return u"está fora da altura do ambiente (pertence a outro pavimento)"
                return None
            pos = POSITION_GUARD.get(fk)
            if pos and rels == set([pos]):
                return u"só delimita o ambiente por {}".format(u"cima" if pos == "top" else u"baixo")
            return None

        def route(norm, fk, item, cat_id, rels, zr=None, forced=False):
            """Put one classified keynote into its finish; returns True if used."""
            why = height_reject(fk, zr, rels)
            if why:
                fins[fk]["notes"].append(u"{} em {} {} foi ignorada: {}.".format(
                    norm, item["cat"], item["id"], why))
                return False
            fins[fk]["keys"].append(norm)
            if forced:
                prefix_fk = classify(norm)[1]
                msg = u"{} em {} {} - moldura de parede sem prefixo RT vai para o Rodapé".format(norm, item["cat"], item["id"])
                if prefix_fk and prefix_fk != "base":
                    fins[fk]["mismatches"].append(msg + u" (o prefixo indica {}).".format(FINISH_LABELS[prefix_fk]))
                elif not prefix_fk:
                    fins[fk]["notes"].append(msg + u".")
            elif cat_id not in EXPECTED_CATEGORIES[fk]:
                fins[fk]["mismatches"].append(
                    u"{} encontrada em elemento da categoria {} ({}) - classificada como {} conforme a regra do prefixo.".format(
                        norm, item["cat"], item["id"], FINISH_LABELS[fk]))
            return True

        for key in sorted(raw["elements"]):
            e = raw["elements"][key]
            dk, el, rels = e["doc_key"], e["elem"], e["rels"]
            item = self._base_item(dk, el, rels)
            cid = cat_int(el)
            # elements only found by position (in the room / geometric search)
            # are candidates, not boundaries: no "missing Keynote" noise,
            # except sweeps and ceilings, which are always relevant
            bounding = rels - set(["inside", "surface", "flush", "contained"])
            mats = e["mats"]
            if not mats and not bounding:
                # no touching face known: use the element's own materials
                mats = self.idx.element_materials(dk, el)
            found, why = self.element_keys(dk, el, mats)
            if not found:
                if bounding or cid in ALWAYS_SEARCH:
                    issues.append(dict(item, msg=u"Keynote ausente ({})".format(why)))
                continue
            zr = self.elem_zrange(dk, el)
            used = OrderedDict()
            ignored = []
            for kn, src in found:
                norm, fk, forced = classify_element(kn, cid)
                if fk is None:
                    ignored.append(norm)
                elif route(norm, fk, item, cid, rels, zr, forced):
                    used.setdefault(fk, []).append((norm, src))
            for fk, lst in used.items():
                fins[fk]["items"].append(dict(item, kn=unique_sorted(k for k, _ in lst),
                                              src=sorted(set(s for _, s in lst))))
                if fk == "ceiling":
                    fr_top_area += e["top_area"]
                    fr_keys.add(key)
            if ignored and not used and (bounding or cid in ALWAYS_SEARCH):
                others.append(dict(item, kn=unique_sorted(ignored),
                                   msg=u"prefixo não é RD / FR / RE / PI - ignorada"))

        # sweeps built into wall types: the only Keynote they can carry is
        # their material's
        for it in (raw["integral"] if self.mode != "type" else []):
            wall = it["wall"]
            item = {"id": eid_int(wall.Id), "cat": u"Moldura do tipo de parede", "link": u"",
                    "rels": [u"moldura no tipo de parede"]}
            kn = self.idx.material_keynote(0, it["mat"])
            if not kn:
                issues.append(dict(item, msg=u"Keynote ausente (o material da moldura '{}' não tem Keynote)".format(
                    self.idx.material_name(0, it["mat"]))))
                continue
            norm, fk, forced = classify_element(kn, CAT_CORNICES)
            if route(norm, fk, item, CAT_CORNICES, set(["sweep"]), None, forced):
                fins[fk]["items"].append(dict(item, kn=[norm], src=[u"material da moldura do tipo de parede"]))

        for fk, fin in fins.items():
            fin["keys"] = unique_sorted(fin["keys"])
            fin["state"] = "ok" if fin["keys"] else "missing"
            fin["notes"] = unique_sorted(fin["notes"])
            fin["mismatches"] = sorted(set(fin["mismatches"]), key=natural_key)
            if self.keynote_texts:
                for k in fin["keys"]:
                    if k not in self.keynote_texts:
                        fin["warnings"].append(u"'{}' não existe no arquivo de Keynotes carregado.".format(k))
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
                fins[fk]["notes"].append(u"Não foi possível calcular a geometria do ambiente: " + err)
        if raw["search_error"]:
            for fk in FINISH_SLOTS:
                if fins[fk]["state"] == "missing":
                    fins[fk]["notes"].append(u"A busca geométrica nas superfícies do ambiente falhou: " + raw["search_error"])
        if fins["wall"]["state"] == "missing":
            if not raw["walls"]:
                fins["wall"]["notes"].append(u"Nenhuma parede delimita este ambiente{}.".format(
                    u" (somente Room Separation Lines)" if raw["separation_lines"] else u""))
            else:
                fins["wall"]["notes"].append(u"Nenhuma Keynote começando com RE nos elementos ao redor do ambiente.")
        if fins["floor"]["state"] == "missing" and not err:
            below = cats_with("bottom")
            fins["floor"]["notes"].append(
                u"Nenhuma Keynote começando com PI. Abaixo do ambiente: {}.".format(u", ".join(below))
                if below else
                u"O fundo do ambiente não é delimitado por nenhum elemento. Verifique se o piso é "
                u"Room Bounding e se o topo dele não está abaixo da base do ambiente.")
        if not err:
            total = raw["top_total"]
            if fins["ceiling"]["state"] == "missing":
                above = cats_with("top")
                fins["ceiling"]["notes"].append(
                    u"Nenhuma Keynote começando com FR ou CB nas superfícies do ambiente nem até {} m acima do "
                    u"topo (topo delimitado por: {}). Se o forro estiver mais alto, aumente o Limit Offset do "
                    u"ambiente ou a margem de busca do forro.".format(fmt_m(self.search_above_m), u", ".join(above) or u"nada"))
            elif fr_top_area > 0 and total > 0 and fr_top_area / total < CEILING_FULL_COVERAGE:
                rest = []
                other_top = cats_with("top", fr_keys)
                if other_top:
                    rest.append(u", ".join(other_top))
                if raw["top_free"] > 0:
                    rest.append(u"sem delimitação (topo do ambiente abaixo do forro)")
                fins["ceiling"]["warnings"].append(
                    u"Elementos de teto (FR/CB) cobrem {:.0f}% do topo do ambiente; o restante é {}.".format(
                        100.0 * fr_top_area / total, u" / ".join(rest) or u"outros elementos"))
        if fins["base"]["state"] == "missing":
            fins["base"]["notes"].append(
                u"Nenhuma Keynote começando com RD (verificados: elementos delimitadores, molduras no lado "
                u"das paredes voltado ao ambiente, molduras dos tipos de parede, famílias dentro do ambiente "
                u"e elementos nas superfícies do ambiente).")
        if fins["sill"]["state"] == "missing":
            fins["sill"]["notes"].append(
                u"Nenhuma Keynote começando com SL. A soleira fica no vão da porta e é aceita quando faceia "
                u"o ambiente ou está dentro dele.")
        if fins["crown"]["state"] == "missing":
            fins["crown"]["notes"].append(
                u"Nenhuma Keynote começando com RT (verificados: molduras, elementos delimitadores e "
                u"elementos nas superfícies do ambiente, até a margem de busca acima do topo).")
        linked = len([1 for (dk, _) in raw["walls"] if dk])
        if linked:
            fins["base"]["notes"].append(u"{} parede(s) de vínculo: as molduras delas não são avaliadas.".format(linked))
        if raw["separation_lines"] and raw["walls"]:
            fins["wall"]["notes"].append(u"{} trecho(s) do contorno são Room Separation Lines (sem acabamento).".format(
                raw["separation_lines"]))

    # ---------------- one room ----------------
    def collect_room(self, room):
        """READ step for one room: rec with its raw element set."""
        rec = {"room": room, "boundary": "ok", "finishes": {}, "issues": [], "others": [],
               "error": None, "raw": None}
        try:
            if room.Location is None:
                rec["boundary"] = "unplaced"
                return rec
            if room.Area <= 0:
                rec["boundary"] = "unenclosed"
                return rec
            rec["raw"] = self.collect(room)
        except Exception as ex:
            rec["boundary"] = "error"
            rec["error"] = to_unicode(ex)
        return rec

    def finish_room(self, rec):
        """ANALYZE step for one room: classify its raw element set."""
        raw = rec.get("raw")
        if raw is None:
            return rec
        try:
            if not raw["elements"] and not raw["separation_lines"]:
                rec["boundary"] = "no_segments"
            rec["finishes"], rec["issues"], rec["others"] = self.classify_room(raw)
        except Exception as ex:
            rec["boundary"] = "error"
            rec["error"] = to_unicode(ex)
        rec["raw"] = None                     # free the geometry references
        return rec

    def analyze(self, room):
        return self.finish_room(self.collect_room(room))

    def contain_pass(self, records, progress=None):
        """Assign by Document.GetRoomAtPoint every finish-keynoted element no
        room has claimed yet, plus every wall sweep (a sweep may run through
        several rooms). Returns stats and the elements left in no room."""
        by_room, assigned, phases = {}, set(), {}
        for rec in records:
            raw = rec.get("raw")
            if raw is None:
                continue
            by_room[eid_int(rec["room"].Id)] = rec
            assigned.update(raw["elements"].keys())
            try:
                pid = rec["room"].get_Parameter(BIP.ROOM_PHASE).AsElementId()
                if is_valid_id(pid):
                    phases[eid_int(pid)] = self.doc.GetElement(pid)
            except Exception:
                pass
        phase_list = [phases[k] for k in sorted(phases)] or [None]

        todo = []
        for dk, d, xf in self._search_docs():
            for eid in self._finish_candidates(dk, d):
                el = self.idx.element(dk, eid)
                if el is None:
                    continue
                is_sweep = cat_int(el) in BASE_CATEGORIES
                if is_sweep or (dk, eid_int(eid)) not in assigned:
                    todo.append((dk, el, xf, is_sweep))

        stats = {"tested": len(todo), "sweeps": 0, "sweeps_in_rooms": 0, "unassigned": []}
        for i, (dk, el, xf, is_sweep) in enumerate(todo):
            if progress is not None and not progress(i + 1, len(todo)):
                break
            pts = sample_points(el)
            if xf is not None:
                pts = [xf.OfPoint(p) for p in pts]
            rids = set()
            for p in pts:
                for ph in phase_list:
                    try:
                        r = self.doc.GetRoomAtPoint(p, ph) if ph is not None else self.doc.GetRoomAtPoint(p)
                    except Exception:
                        r = None
                    if r is not None:
                        rids.add(eid_int(r.Id))
            hits = [by_room[r] for r in sorted(rids) if r in by_room]
            for rec in hits:
                touch_element(rec["raw"], dk, el, "contained")
            k = (dk, eid_int(el.Id))
            if is_sweep:
                stats["sweeps"] += 1
                if hits or k in assigned:
                    stats["sweeps_in_rooms"] += 1
            if not hits and k not in assigned:
                kn, _ = self.idx.element_keynote(dk, el)
                stats["unassigned"].append({
                    "id": eid_int(el.Id), "cat": cat_name(el), "link": self.idx.link_name(dk),
                    "kn": classify(kn)[0] if kn else u"",
                    "reason": (u"nenhum ponto amostrado está dentro de um ambiente" if pts
                               else u"o elemento não tem geometria sólida"),
                })
        stats["unassigned"].sort(key=lambda u: (u["cat"], u["id"]))
        return stats


# ==================================================================
# Room records -> status, preview rows, JSON
# ==================================================================
def room_info(room):
    try:
        level = to_unicode(room.Level.Name) if room.Level else u"(não colocado)"
        elev = room.Level.Elevation if room.Level else 0.0
    except Exception:
        level, elev = u"(não colocado)", 0.0
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
    if any(ch.get("overflow") for ch in rec.get("changes", {}).values()):
        return "WARNING"
    if all(f["state"] == "ok" and not f["warnings"] and not f["mismatches"]
           for f in rec["finishes"].values()):
        return "OK"
    return "WARNING"


def sort_records(records):
    return sorted(records, key=lambda r: (r["info"]["elev"], r["info"]["level"],
                                          natural_key(r["info"]["number"]), r["info"]["id"]))


def compute_changes(records, targets, tag_index=None):
    """Current vs new value for every finish field of every room. The unique
    Keynotes of a finish fill its fields in order (01, 02, ...). When at
    least one Keynote was found, the finish is rewritten as a whole: fields
    past the last Keynote are cleared ("Clear"). When nothing was found, the
    real values already in the room are kept. Keynotes beyond the last field
    overflow and are reported. Nothing is written here.

    On top of that (step 2, after the values are known):
      Rule A - N = the last row (field number) holding a real value in any
      finish. The empty fields of rows 1..N get PLACEHOLDER ('-'); rows after
      N stay empty (a '-' left there by an earlier run is cleared). With no
      real value at all (N = 0) nothing is filled - previous behaviour.
      Rule B - with tag_index, the room identifier tags are planned to switch
      to the REVESTIMENTOS type for N rows (plan_tags)."""
    for rec in records:
        rec["changes"] = OrderedDict()
        room = rec["room"]
        ok = rec["boundary"] == "ok"
        # step 1 - values from the Keynotes (unchanged rule)
        plan = []
        for key, slots in FINISH_SLOTS.items():
            fields = targets.get(key) or [None] * len(slots)
            fin = rec["finishes"].get(key)
            keys = list(fin["keys"]) if fin else []
            available = [i for i, t in enumerate(fields) if t is not None]
            new_vals = [u""] * len(slots)
            for kv, i in zip(keys, available):
                new_vals[i] = kv
            overflow = keys[len(available):]
            curs = [t.read(room) if t is not None else u"" for t in fields]
            if keys:
                final = new_vals
            else:
                # nothing found for this finish: real values stay, '-' is re-planned
                final = [c if is_real(c) else u"" for c in curs]
            plan.append((key, slots, fields, keys, curs, final, overflow))

        # step 2 - rule A: how many rows have real information
        n_rows = 0
        if ok:
            for _, _, fields, _, _, final, _ in plan:
                for i, v in enumerate(final):
                    if fields[i] is not None and is_real(v):
                        n_rows = max(n_rows, i + 1)
        rec["rows"] = n_rows

        for key, slots, fields, keys, curs, final, overflow in plan:
            out = []
            for i, (name, _) in enumerate(slots):
                t, cur, new = fields[i], curs[i], final[i]
                if t is not None and not new and i < n_rows:
                    new = PLACEHOLDER
                if t is None:
                    kind = "No parameter"
                elif not ok:
                    kind, new = "Skipped", cur
                elif new == cur:
                    kind = "Keep" if (not keys and is_real(cur)) else "No change"
                elif is_real(new):
                    kind = "Overwrite" if is_real(cur) else "Fill"
                elif is_real(cur) or not new:
                    kind = "Clear"          # old value removed, or a stale '-' removed
                else:
                    kind = "Fill"           # '-' into an empty field
                out.append({"name": name, "current": cur, "new": new, "kind": kind})
            rec["changes"][key] = {
                "slots": out, "overflow": overflow,
                "current": join_keys([s["current"] for s in out if is_real(s["current"])]),
                "new": join_keys(keys),
            }
    if tag_index is not None:
        plan_tags(records, tag_index)


def rows_in_room(room, targets):
    """Rows with a real value in the room as it is now (read back after
    writing, so unticked preview rows are respected)."""
    n = 0
    for key, fields in targets.items():
        for i, t in enumerate(fields):
            if t is not None and is_real(t.read(room)):
                n = max(n, i + 1)
    return n


# ==================================================================
# Rule B - room identifier (Room Tag) type by number of rows
# ==================================================================
def type_name(t):
    try:
        return to_unicode(DB.Element.Name.GetValue(t))
    except Exception:
        try:
            return param_str(t.get_Parameter(BIP.SYMBOL_NAME_PARAM))
        except Exception:
            return u""


def tag_rows_in_name(name):
    m = TAG_ROWS_RE.search(name or u"")
    return int(m.group(2)) if m else None


class TagIndex(object):
    """Room tags already placed in the project whose family has REVESTIMENTOS
    types, and those types by number of rows. Read only."""
    def __init__(self, d):
        self.doc = d
        self.families = OrderedDict()     # family name -> {rows: [RoomTagType]}
        self.by_room = {}                 # room id -> [RoomTag]
        self._views = {}
        try:
            types = (DB.FilteredElementCollector(d).OfCategory(DB.BuiltInCategory.OST_RoomTags)
                     .WhereElementIsElementType().ToElements())
        except Exception:
            types = []
        for t in types:
            n = tag_rows_in_name(type_name(t))
            if n is None:
                continue
            fam = to_unicode(getattr(t, "FamilyName", u""))
            self.families.setdefault(fam, {}).setdefault(n, []).append(t)
        if not self.families:
            return
        for tag in (DB.FilteredElementCollector(d).OfCategory(DB.BuiltInCategory.OST_RoomTags)
                    .WhereElementIsNotElementType()):
            try:
                t = d.GetElement(tag.GetTypeId())
                if t is None or to_unicode(t.FamilyName) not in self.families:
                    continue
                try:
                    rid = tag.TaggedLocalRoomId
                except Exception:
                    rid = tag.Room.Id if tag.Room is not None else None
                if is_valid_id(rid):
                    self.by_room.setdefault(eid_int(rid), []).append(tag)
            except Exception:
                continue

    def tags_for(self, room):
        return self.by_room.get(eid_int(room.Id), [])

    def family_of(self, tag):
        t = self.doc.GetElement(tag.GetTypeId())
        return to_unicode(t.FamilyName) if t is not None else u""

    def current_name(self, tag):
        t = self.doc.GetElement(tag.GetTypeId())
        return type_name(t) if t is not None else u""

    def view_name(self, tag):
        k = eid_int(tag.OwnerViewId)
        if k not in self._views:
            v = self.doc.GetElement(tag.OwnerViewId)
            self._views[k] = to_unicode(v.Name) if v is not None else u""
        return self._views[k]

    def missing(self):
        """[(family, [missing type labels])] for rows 1..TAG_MAX_ROWS."""
        out = []
        for fam, by_n in self.families.items():
            gone = [u"REVESTIMENTOS {:02d}".format(n) for n in range(1, TAG_MAX_ROWS + 1) if n not in by_n]
            if gone:
                out.append((fam, gone))
        return out

    def target(self, tag, n):
        """Type of the tag's own family for n rows -> (type or None, name, message).
        The current type name with its number swapped is preferred, so
        prefixes / suffixes in the names ('1:50', 'X', ...) are kept."""
        cur_type = self.doc.GetElement(tag.GetTypeId())
        fam = to_unicode(cur_type.FamilyName) if cur_type is not None else u""
        cur_name = type_name(cur_type) if cur_type is not None else u""
        by_n = self.families.get(fam, {})
        wanted = None
        m = TAG_ROWS_RE.search(cur_name)
        if m:
            wanted = (cur_name[:m.start()] + m.group(1) + str(n).zfill(len(m.group(2))) +
                      (m.group(3) or u"") + cur_name[m.end():])
            for t in by_n.get(n, []):
                if type_name(t) == wanted:
                    return t, wanted, u""
        cands = sorted(by_n.get(n, []), key=type_name)
        if cands:
            note = u"" if len(cands) == 1 else u"{} tipos para {} linha(s); usado '{}'".format(
                len(cands), n, type_name(cands[0]))
            return cands[0], type_name(cands[0]), note
        label = wanted or u"REVESTIMENTOS {:02d}".format(n)
        return None, label, u"O tipo '{}' não foi encontrado na família de identificador '{}'.".format(label, fam)


def plan_tags(records, tag_index):
    """rec['tags'] = identifier tags of the room that need another type."""
    for rec in records:
        rec["tags"] = []
        n = rec.get("rows", 0)
        if rec["boundary"] != "ok" or not n:
            continue
        for tag in tag_index.tags_for(rec["room"]):
            target, name, msg = tag_index.target(tag, n)
            if target is not None and eid_int(target.Id) == eid_int(tag.GetTypeId()):
                continue                                     # already the right type
            rec["tags"].append({
                "tag": tag, "id": eid_int(tag.Id), "view": tag_index.view_name(tag),
                "current": tag_index.current_name(tag), "new": name,
                "type": target, "msg": msg, "rows": n,
            })


def worksharing_block(room):
    """Reason this element (room or tag) can't be edited right now, or u''."""
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
            return u"Em uso por {}".format(owner or u"outro usuário")
        upd = DB.WorksharingUtils.GetModelUpdatesStatus(doc, room.Id)
        if upd == DB.ModelUpdatesStatus.UpdatedInCentral:
            return u"Alterado no central - faça Reload Latest antes"
        if upd == DB.ModelUpdatesStatus.DeletedInCentral:
            return u"Excluído no central"
    except Exception:
        pass
    return u""


def build_json(records, targets, mode, keynote_texts, problems, contain_stats=None, search_above_m=SEARCH_ABOVE_M,
               tag_index=None):
    rooms = []
    for rec in records:
        info = rec["info"]
        fins = OrderedDict()
        for key in FINISH_SLOTS:
            fin = rec["finishes"].get(key)
            ch = rec.get("changes", {}).get(key, {})
            slots = [dict(s, change=CHANGE_LABELS.get(s["kind"], u"")) for s in ch.get("slots", [])]
            if fin is None:
                fins[key] = {"state": "n/a", "keys": [], "value": u"", "current": ch.get("current", u""),
                             "slots": slots, "overflow": [], "items": [], "notes": [], "warnings": [],
                             "mismatches": []}
                continue
            fins[key] = {
                "state": fin["state"], "keys": fin["keys"], "value": join_keys(fin["keys"]),
                "current": ch.get("current", u""), "slots": slots, "overflow": ch.get("overflow", []),
                "items": fin["items"], "notes": fin["notes"], "warnings": fin["warnings"],
                "mismatches": fin["mismatches"],
            }
        boundary_msg = {
            "unplaced": u"Ambiente não colocado.",
            "unenclosed": u"Ambiente não fechado ou redundante (área = 0).",
            "no_segments": u"O ambiente não tem elementos delimitadores.",
            "error": u"A análise falhou: {}".format(rec.get("error") or u""),
        }.get(rec["boundary"], u"")
        rooms.append({
            "id": info["id"], "number": info["number"], "name": info["name"],
            "level": info["level"], "phase": info["phase"],
            "boundary": rec["boundary"], "boundaryMsg": boundary_msg,
            "status": rec["status"], "updated": bool(rec.get("updated")),
            "writeError": rec.get("write_error") or u"",
            "finishes": fins, "issues": rec.get("issues", []), "others": rec.get("others", []),
            "rows": rec.get("rows", 0),
            "tags": [{"id": t["id"], "view": t["view"], "current": t["current"], "new": t["new"],
                      "missing": t["type"] is None, "msg": t["msg"]} for t in rec.get("tags", [])],
        })
    return {
        "project": to_unicode(doc.Title),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "keynoteSource": dict(KEYNOTE_SOURCES).get(mode, mode),
        "boundaryLocation": u"face de acabamento" if BOUNDARY_LOCATION == DB.SpatialElementBoundaryLocation.Finish else to_unicode(BOUNDARY_LOCATION),
        "searchAboveM": search_above_m,
        "separator": SEPARATOR,
        "prefixRules": OrderedDict((p, finish_label(f)) for p, f in PREFIX_RULES.items()),
        "groups": [{"key": k, "label": FINISH_LABELS[k], "prefixes": finish_prefixes(k),
                    "slots": [{"name": n, "found": bool(targets.get(k) and targets[k][i] is not None)}
                              for i, (n, _) in enumerate(slots)]}
                   for k, slots in FINISH_SLOTS.items()],
        "paramProblems": problems,
        "keynotes": keynote_texts,
        "rooms": rooms,
        "tagFamilies": list(tag_index.families.keys()) if tag_index else [],
        "tagMissing": [{"family": f, "types": m} for f, m in (tag_index.missing() if tag_index else [])],
        "placeholder": PLACEHOLDER,
        "containStats": dict((k, v) for k, v in (contain_stats or {}).items() if k != "unassigned"),
        "unassigned": (contain_stats or {}).get("unassigned", []),
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
        return u"{} (vínculo: {})".format(item["id"], item["link"])
    try:
        return output.linkify(DB.ElementId(item["id"]))
    except Exception:
        return to_unicode(item["id"])


STATUS_LABELS = {"OK": u"OK", "WARNING": u"ATENÇÃO", "ERROR": u"ERRO"}
CHANGE_LABELS = {"Fill": u"Preencher", "Overwrite": u"Sobrescrever", "Clear": u"Limpar",
                 "Tag": u"Trocar tipo", "TagMissing": u"Tipo ausente"}
BOUNDARY_LABELS = {
    "unplaced": u"não colocado",
    "unenclosed": u"não fechado ou redundante",
    "no_segments": u"sem elementos delimitadores",
    "error": u"erro na análise",
}


def fmt_m(value):
    """1.0 -> u'1,00'"""
    return u"{:.2f}".format(value).replace(u".", u",")


def parse_m(text):
    """u'1,5' / u'1.5' -> 1.5 ; None if invalid or outside 0-10 m."""
    try:
        v = float(to_unicode(text).strip().replace(u",", u"."))
    except Exception:
        return None
    return v if 0.0 <= v <= 10.0 else None


def help_rules_text():
    lines = []
    for key, slots in FINISH_SLOTS.items():
        lines.append(u"{}…  →  {} a {}".format(u" / ".join(finish_prefixes(key)), slots[0][0], slots[-1][0][-2:]))
    lines.append(u"")
    lines.append(u"Cada Keynote diferente ocupa um campo, em ordem: por exemplo, paredes RE02, RE02, RE05 e "
                 u"RE08 ficam Parede 01 = RE02, Parede 02 = RE05, Parede 03 = RE08. Repetições contam uma vez; "
                 u"maiúsculas e minúsculas não importam (o valor é gravado em maiúsculas).")
    lines.append(u"")
    lines.append(u"Keynotes com outros prefixos são ignoradas.")
    lines.append(u"")
    lines.append(u"Exceção: moldura de parede (Wall Sweep) vai para o Rodapé, a não ser que a Keynote dela "
                 u"comece com RT (vai para o Rodateto).")
    return u"\n".join(lines)


# ==================================================================
# WPF window
# ==================================================================
XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Acabamentos de Ambiente" Width="1160" Height="800"
        MinWidth="900" MinHeight="600" WindowStartupLocation="CenterScreen"
        Background="#0E1526">
  <Window.Resources>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
    </Style>
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
    </Style>
    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#7A8FA9"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="HelpTitle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#65E3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Margin" Value="0,16,0,6"/>
    </Style>
    <Style x:Key="HelpBody" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="LineHeight" Value="19"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="26"/>
      <Setter Property="Padding" Value="4,2,4,2"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="26"/>
      <Setter Property="Padding" Value="6,3,6,3"/>
      <Setter Property="Background" Value="#0B1120"/>
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="BorderBrush" Value="#2A4A6E"/>
      <Setter Property="CaretBrush" Value="#65E3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Height" Value="32"/>
      <Setter Property="MinWidth" Value="140"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
      <Setter Property="Foreground" Value="#65E3FF"/>
      <Setter Property="Background" Value="#101B30"/>
      <Setter Property="BorderBrush" Value="#2A4A6E"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="3" Padding="14,0">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#65E3FF"/>
                <Setter TargetName="bd" Property="Background" Value="#16304A"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="#7A8FA9"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" BorderBrush="#23324F" BorderThickness="1,1,1,0"
                    Padding="16,7" Margin="0,0,4,0" CornerRadius="3,3,0,0">
              <ContentPresenter ContentSource="Header" TextElement.Foreground="{TemplateBinding Foreground}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#16233B"/>
                <Setter Property="Foreground" Value="#65E3FF"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
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
        <DataTrigger Binding="{Binding Kind}" Value="Overwrite">
          <Setter Property="Foreground" Value="#FFB454"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Kind}" Value="Clear">
          <Setter Property="Foreground" Value="#FF8FB5"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Kind}" Value="Tag">
          <Setter Property="Foreground" Value="#9D8CFF"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Kind}" Value="TagMissing">
          <Setter Property="Foreground" Value="#FF4F9A"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Editable}" Value="False">
          <Setter Property="Foreground" Value="#7A8FA9"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="RoomGroup" TargetType="Expander">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Expander">
            <DockPanel>
              <ToggleButton DockPanel.Dock="Top" Cursor="Hand" Content="{TemplateBinding Header}"
                            IsChecked="{Binding IsExpanded, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="bd" Background="#141E33" BorderBrush="#23324F" BorderThickness="0,0,0,1" Padding="8,6">
                      <DockPanel>
                        <TextBlock x:Name="arrow" Text="&#x25B6;" Foreground="#65E3FF" Width="20" FontSize="10"
                                   VerticalAlignment="Center"/>
                        <ContentPresenter VerticalAlignment="Center"/>
                      </DockPanel>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsChecked" Value="True">
                        <Setter TargetName="arrow" Property="Text" Value="&#x25BC;"/>
                      </Trigger>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="bd" Property="Background" Value="#16304A"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter x:Name="body" Visibility="Collapsed"/>
            </DockPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsExpanded" Value="True">
                <Setter TargetName="body" Property="Visibility" Value="Visible"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Text="AUTOMAÇÃO DE ACABAMENTOS POR KEYNOTE" FontSize="16"
               FontWeight="SemiBold" Foreground="#65E3FF"/>
    <TextBlock Grid.Row="1" Style="{StaticResource Hint}" Margin="0,4,0,12"
               Text="Lê automaticamente as Keynotes dos elementos ao redor de cada ambiente e grava nos parâmetros de acabamento do ambiente."/>

    <TabControl Grid.Row="2" Background="Transparent" BorderBrush="#23324F" BorderThickness="1" Padding="14">

      <!-- ============ TAB 1: update ============ -->
      <TabItem Header="Atualizar ambientes">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="22"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
              <TextBlock Style="{StaticResource Label}" Text="Margem de busca do forro acima do ambiente (m)"/>
              <StackPanel Orientation="Horizontal">
                <TextBox x:Name="tb_margin" Width="80" ToolTipService.ShowDuration="20000"
                         ToolTip="Até quantos metros acima do topo do ambiente a ferramenta procura o forro. Aumente se o forro estiver acima do Limit Offset do ambiente. Depois de alterar, clique em Analisar Modelo de novo."/>
              </StackPanel>
            </StackPanel>

            <StackPanel Grid.Column="2">
              <TextBlock Style="{StaticResource Label}" Text="Parâmetros do ambiente"/>
              <TextBlock x:Name="tb_params" Style="{StaticResource Hint}" LineHeight="18"/>
            </StackPanel>
          </Grid>

          <DockPanel Grid.Row="1" Margin="0,4,0,8">
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Bottom" Margin="12,0,0,0">
              <Button x:Name="btn_expand" Content="Expandir todos" MinWidth="0" Height="26" Margin="0,0,6,0" FontSize="11"
                      IsEnabled="False" ToolTip="Mostra os campos de todos os ambientes."/>
              <Button x:Name="btn_collapse" Content="Recolher todos" MinWidth="0" Height="26" Margin="0" FontSize="11"
                      IsEnabled="False" ToolTip="Mostra só uma linha por ambiente."/>
            </StackPanel>
            <TextBlock x:Name="tb_status" TextWrapping="Wrap" FontSize="12" Foreground="#3DDCB4" VerticalAlignment="Bottom"
                       Text="Clique em Analisar Modelo para começar. Nada é gravado até você clicar em Atualizar Ambientes."/>
          </DockPanel>

          <DataGrid Grid.Row="2" x:Name="grid" AutoGenerateColumns="False" CanUserAddRows="False"
                    CanUserDeleteRows="False" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                    HorizontalGridLinesBrush="#1B2740" Background="#0B1120" BorderBrush="#23324F"
                    RowHeaderWidth="0" SelectionMode="Extended" ToolTipService.ShowDuration="20000"
                    ToolTip="Pré-visualização agrupada por ambiente: clique no ambiente para ver o que muda nele. Todas as linhas vêm marcadas; desmarque uma linha para não gravar aquele valor. Âmbar = substitui um valor existente; rosa = limpa um campo que sobrou; lilás = troca o tipo do identificador do ambiente.">
            <DataGrid.GroupStyle>
              <GroupStyle>
                <GroupStyle.ContainerStyle>
                  <Style TargetType="GroupItem">
                    <Setter Property="Template">
                      <Setter.Value>
                        <ControlTemplate TargetType="GroupItem">
                          <Expander Style="{StaticResource RoomGroup}" IsExpanded="False">
                            <Expander.Header>
                              <TextBlock Text="{Binding Name}" Foreground="#D9E8F5" FontWeight="SemiBold"/>
                            </Expander.Header>
                            <ItemsPresenter/>
                          </Expander>
                        </ControlTemplate>
                      </Setter.Value>
                    </Setter>
                  </Style>
                </GroupStyle.ContainerStyle>
              </GroupStyle>
            </DataGrid.GroupStyle>
            <DataGrid.Columns>
              <DataGridTemplateColumn Header="Aplicar" Width="60">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                              IsChecked="{Binding Apply, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                              IsEnabled="{Binding Editable}"/>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Header="Ambiente" Binding="{Binding Room}" IsReadOnly="True" Width="75"/>
              <DataGridTextColumn Header="Nome" Binding="{Binding Name}" IsReadOnly="True" Width="140"/>
              <DataGridTextColumn Header="Pavimento" Binding="{Binding Level}" IsReadOnly="True" Width="100"/>
              <DataGridTextColumn Header="Parâmetro" Binding="{Binding Parameter}" IsReadOnly="True" Width="190"/>
              <DataGridTextColumn Header="Valor atual" Binding="{Binding Current}" IsReadOnly="True" Width="*"/>
              <DataGridTextColumn Header="Novo valor" Binding="{Binding New}" IsReadOnly="True" Width="*"/>
              <DataGridTextColumn Header="Alteração" Binding="{Binding Change}" IsReadOnly="True" Width="95"/>
              <DataGridTextColumn Header="Observação" Binding="{Binding Note}" IsReadOnly="True" Width="170"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </TabItem>

      <!-- ============ TAB 2: help ============ -->
      <TabItem Header="Como funciona">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel Margin="4,0,12,12" MaxWidth="900" HorizontalAlignment="Left">
            <TextBlock Style="{StaticResource HelpTitle}" Margin="0,0,0,6" Text="O QUE A FERRAMENTA FAZ"/>
            <TextBlock Style="{StaticResource HelpBody}" Text="Para cada ambiente do modelo, encontra os elementos construtivos ao redor (paredes, pisos, forros, rodapés e molduras), lê a Keynote de cada um e grava os códigos nos quatro parâmetros de acabamento do ambiente. A Keynote dos elementos só é lida, nunca alterada."/>

            <TextBlock Style="{StaticResource HelpTitle}" Text="REGRA DE CLASSIFICAÇÃO (PELO PREFIXO DA KEYNOTE)"/>
            <TextBlock x:Name="tb_help_rules" Style="{StaticResource HelpBody}"/>

            <TextBlock Style="{StaticResource HelpTitle}" Text="ONDE OS ELEMENTOS SÃO PROCURADOS"/>
            <TextBlock x:Name="tb_help_search" Style="{StaticResource HelpBody}"/>

            <TextBlock Style="{StaticResource HelpTitle}" Text="LINHAS COM &quot;-&quot; E IDENTIFICADOR DO AMBIENTE"/>
            <TextBlock Style="{StaticResource HelpBody}" Text="• Cada número de campo é uma linha do identificador: linha 01 = Parede 01, Piso 01, Teto 01, Rodapé 01, Rodateto 01 e Soleira 01; linha 02 = os campos 02; e assim por diante.&#10;• Se pelo menos uma linha tiver valor, os campos vazios das linhas preenchidas recebem &quot;-&quot;. Linhas sem nenhum valor continuam vazias. Ambiente sem nenhum acabamento: nada muda.&#10;• O número de linhas com valor escolhe o tipo do identificador: 1 linha = REVESTIMENTOS 01, 2 linhas = REVESTIMENTOS 02 ... até 05. Vale para nomes como &quot;REVESTIMENTOS 3X&quot; ou com outros textos antes e depois.&#10;• Só os identificadores já colocados no projeto são trocados, sempre dentro da mesma família. Nada é criado.&#10;• Se o tipo necessário não existir na família, a ferramenta avisa qual tipo está faltando e não troca aquele identificador."/>

            <TextBlock Style="{StaticResource HelpTitle}" Text="EXCEL"/>
            <TextBlock Style="{StaticResource HelpBody}" Text="• No relatório HTML, Baixar Excel gera as abas Resumo e Ambientes (Número, Nome do ambiente, Pavimento, Parede, Piso, Rodapé, Soleira, Teto e Rodateto). Só os ambientes visíveis na tabela entram.&#10;• Cada célula de acabamento traz os valores separados por &quot; / &quot; (ex.: RE01 / RE02), na ordem dos campos 01, 02, 03...&#10;• A planilha é só para consulta: ela não é lida de volta pelo Revit."/>

            <TextBlock Style="{StaticResource HelpTitle}" Text="PROTEÇÕES"/>
            <TextBlock Style="{StaticResource HelpBody}" Text="• Piso, rodapé e parede do pavimento de cima ou de baixo não entram no ambiente.&#10;• Nada é gravado até você clicar em Atualizar Ambientes.&#10;• Valores diferentes dos atuais aparecem em âmbar na pré-visualização e na confirmação antes de gravar; desmarque a linha para manter o valor atual.&#10;• Cada acabamento é regravado por inteiro: se sobrar um campo com valor antigo (ex.: Parede 04 quando agora só há 3 revestimentos), ele aparece em rosa como Limpar.&#10;• Se houver mais Keynotes do que campos, as que sobrarem aparecem como aviso - nada é descartado sem aviso.&#10;• Se nada for encontrado para um parâmetro, o valor atual é mantido - nunca é apagado.&#10;• Tudo é gravado em uma única transação: Ctrl+Z no Revit desfaz a atualização inteira.&#10;• Ambientes em uso por outro usuário (workset) aparecem bloqueados na pré-visualização."/>

            <TextBlock Style="{StaticResource HelpTitle}" Text="PASSO A PASSO"/>
            <TextBlock Style="{StaticResource HelpBody}" Text="1. Analisar Modelo - lê todos os ambientes e já mostra a pré-visualização: valor atual e novo valor de cada parâmetro. Não altera nada.&#10;2. Confira a lista. Todas as linhas vêm marcadas; desmarque as que não quer gravar. Linhas em âmbar substituem um valor existente.&#10;3. Atualizar Ambientes - grava as linhas marcadas.&#10;4. Abrir Relatório HTML - resumo, filtros e o detalhe de onde veio cada Keynote."/>

            <TextBlock Style="{StaticResource HelpTitle}" Text="QUANDO ALGO NÃO APARECE"/>
            <TextBlock Style="{StaticResource HelpBody}" Text="• Forro não encontrado: aumente a margem de busca do forro ou o Limit Offset do ambiente e analise de novo.&#10;• Elemento sem Keynote: aparece no relatório como &quot;Keynote ausente&quot;.&#10;• Elemento com Keynote de acabamento fora de qualquer ambiente: seção &quot;Não atribuídos a nenhum ambiente&quot; no relatório.&#10;• Prefixo que não combina com a categoria (ex.: RE01 num piso): aparece como inconsistência, mas o valor é gravado pela regra do prefixo.&#10;• No relatório, clique num ambiente para ver cada elemento encontrado, a Keynote e o motivo de cada valor ignorado."/>
          </StackPanel>
        </ScrollViewer>
      </TabItem>
    </TabControl>

    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="btn_scan" Content="Analisar Modelo" ToolTipService.ShowDuration="20000"
              ToolTip="Lê o modelo, analisa todos os ambientes e mostra a pré-visualização com o valor atual e o novo valor de cada parâmetro. Não altera nada."/>
      <Button x:Name="btn_update" Content="Atualizar Ambientes" IsEnabled="False" Foreground="#3DDCB4" BorderBrush="#2F7F6B"
              ToolTipService.ShowDuration="20000"
              ToolTip="Grava as linhas marcadas na pré-visualização, em uma única transação (Ctrl+Z desfaz)."/>
      <Button x:Name="btn_report" Content="Abrir Relatório HTML" IsEnabled="False" ToolTipService.ShowDuration="20000"
              ToolTip="Abre o relatório com todos os ambientes, filtros e o detalhe de onde veio cada Keynote."/>
      <Button x:Name="btn_cancel" Content="Cancelar" Margin="0" MinWidth="110"
              ToolTip="Fecha a janela sem gravar nada."/>
    </StackPanel>
  </Grid>
</Window>
"""


def group_header(rec, rows):
    """Preview group title: 'Ambiente 1 · TESTE 02 · TÉRREO - 3 alterações:
    2 a preencher, 1 identificador a trocar'."""
    info = rec["info"]
    kinds = [r["Kind"] for r in rows]
    parts = []
    for kind, one, many in (("Fill", u"a preencher", u"a preencher"), ("Overwrite", u"a substituir", u"a substituir"),
                            ("Clear", u"a limpar", u"a limpar"),
                            ("Tag", u"identificador a trocar", u"identificadores a trocar"),
                            ("TagMissing", u"tipo de identificador ausente", u"tipos de identificador ausentes")):
        n = kinds.count(kind)
        if n:
            parts.append(u"{} {}".format(n, one if n == 1 else many))
    blocked = len([1 for r in rows if not bool(r["Editable"]) and r["Kind"] != "TagMissing"])
    if blocked:
        parts.append(u"{} bloqueada(s)".format(blocked))
    title = u" · ".join(x for x in (u"Ambiente " + info["number"], info["name"], info["level"]) if x)
    return u"{}   —   {} {}: {}".format(title, len(rows), u"alteração" if len(rows) == 1 else u"alterações",
                                       u", ".join(parts))


class RoomFinishWindow(forms.WPFWindow):
    def __init__(self, rooms):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)
        self.rooms = rooms
        self.records = []
        self.row_keys = []          # table row index -> (record, finish key)
        self.table = None
        self.mode = KEYNOTE_MODE
        self.search_above_m = SEARCH_ABOVE_M
        self.keynote_texts = {}
        self.contain_stats = {"tested": 0, "sweeps": 0, "sweeps_in_rooms": 0, "unassigned": []}

        self.tb_margin.Text = fmt_m(SEARCH_ABOVE_M)
        self.tb_help_rules.Text = help_rules_text()
        self._fill_help_search()

        self.sample_room = rooms[0]
        try:
            self.tag_index = TagIndex(doc)
        except Exception as ex:
            self.tag_index = None
            output.print_md(u"**ERRO:** não foi possível ler os identificadores de ambiente: `{}`".format(to_unicode(ex)))
        self._resolve_params()

        self.btn_scan.Click += self.on_scan
        self.btn_expand.Click += self.on_expand
        self.btn_collapse.Click += self.on_collapse
        self.btn_update.Click += self.on_update
        self.btn_report.Click += self.on_report
        self.btn_cancel.Click += self.on_cancel
        self.tb_margin.TextChanged += self.on_settings_changed

    def _fill_help_search(self):
        m = parse_m(self.tb_margin.Text)
        self.tb_help_search.Text = (
            u"• Contorno do ambiente: os elementos que delimitam o ambiente em planta (face de acabamento).\n"
            u"• Volume do ambiente: o piso abaixo, o forro ou a laje acima e as faces laterais.\n"
            u"• Superfícies do ambiente: tudo o que estiver dentro do ambiente, de 2 cm abaixo do piso "
            u"até {} m acima do topo (margem de busca do forro).\n"
            u"• Elementos que apenas faceiam o ambiente pelo lado de fora contam somente para teto e soleira "
            u"(a soleira fica no vão da porta e entra nos dois ambientes que ela separa).\n"
            u"• Molduras de parede e demais elementos com Keynote de acabamento: pelo ambiente que contém "
            u"o elemento. Uma moldura que passa por vários ambientes entra em todos.\n"
            u"• Paredes de acabamento que não são Room Bounding também são encontradas, pela busca nas "
            u"superfícies.".format(fmt_m(m if m is not None else SEARCH_ABOVE_M)))

    # ---------------- parameters ----------------
    def _resolve_params(self):
        self.targets, self.problems = resolve_targets(self.sample_room)
        lines = []
        for key, slots in FINISH_SLOTS.items():
            found = len([t for t in self.targets[key] if t is not None])
            lines.append(u"{}: {} de {} campos ({} a {})".format(
                finish_label(key), found, len(slots), slots[0][0], slots[-1][0][-2:]))
        ti = self.tag_index
        if ti is None or not ti.families:
            lines.append(u"Identificador: nenhum tipo REVESTIMENTOS encontrado - os identificadores não serão trocados.")
        else:
            for fam in ti.families:
                n_rooms = len([1 for tags in ti.by_room.values()
                               if any(ti.family_of(t) == fam for t in tags)])
                lines.append(u"Identificador '{}': colocado em {} ambiente(s)".format(fam, n_rooms))
            for fam, gone in ti.missing():
                lines.append(u"Identificador '{}': faltando {}".format(fam, u", ".join(gone)))
        lines.extend(self.problems)
        self.tb_params.Text = u"\n".join(lines)

    def on_settings_changed(self, sender, args):
        """Ceiling margin changed: the scan is out of date."""
        self._fill_help_search()
        if self.records:
            self.records = []
            self._invalidate_preview()
            self.tb_status.Text = u"Configuração alterada - clique em Analisar Modelo novamente."

    def _invalidate_preview(self):
        self.grid.ItemsSource = None
        self.table = None
        self.row_keys = []
        self.btn_update.IsEnabled = False
        self.btn_expand.IsEnabled = False
        self.btn_collapse.IsEnabled = False

    # ---------------- READ + ANALYZE ----------------
    def on_scan(self, sender, args):
        margin = parse_m(self.tb_margin.Text)
        if margin is None:
            forms.alert(u"Margem de busca do forro inválida: '{}'.\n\nUse um valor em metros entre 0 e 10 "
                        u"(ex.: 1,00).".format(to_unicode(self.tb_margin.Text)))
            return
        self.search_above_m = margin
        self.mode = KEYNOTE_MODE
        self._invalidate_preview()
        output.print_md(u"**[2/6] Ambientes coletados:** {} ambiente(s).".format(len(self.rooms)))
        try:
            scanner = Scanner(doc, self.mode, margin)
        except Exception as ex:
            output.print_md(u"**ERRO:** não foi possível iniciar a análise: `{}`".format(to_unicode(ex)))
            forms.alert(u"Não foi possível iniciar a análise:\n\n{}".format(to_unicode(ex)))
            return
        self.keynote_texts = scanner.keynote_texts
        records = []
        cancelled = False
        n = len(self.rooms)
        with forms.ProgressBar(title=u"Analisando ambientes ({value} de {max_value})", cancellable=True) as pb:
            for i, room in enumerate(self.rooms):
                if pb.cancelled:
                    cancelled = True
                    break
                rec = scanner.collect_room(room)
                rec["info"] = room_info(room)
                records.append(rec)
                if i % 5 == 0 or i == n - 1:
                    pb.update_progress(i + 1, n)
        if not cancelled:
            # every element with a finish Keynote (and every wall sweep) that no
            # room claimed yet: which room contains it?
            with forms.ProgressBar(title=u"Associando elementos aos ambientes ({value} de {max_value})",
                                   cancellable=True) as pb:
                def progress(i, total):
                    if pb.cancelled:
                        return False
                    if i % 10 == 0 or i == total:
                        pb.update_progress(i, total)
                    return True
                self.contain_stats = scanner.contain_pass(records, progress)
                cancelled = pb.cancelled
        if cancelled:
            output.print_md(u"Análise cancelada - nenhum relatório foi gerado.")
            self.tb_status.Text = u"Análise cancelada."
            return
        for rec in records:
            scanner.finish_room(rec)
        cs = self.contain_stats
        output.print_md(u"**Ambiente por ponto:** {} elemento(s) testados com GetRoomAtPoint; molduras de parede: "
                        u"{} no modelo, {} dentro de um ambiente. **{} elemento(s) com Keynote de acabamento não "
                        u"estão em nenhum ambiente.**".format(
                            cs["tested"], cs["sweeps"], cs["sweeps_in_rooms"], len(cs["unassigned"])))
        for u in cs["unassigned"][:40]:
            output.print_md(u"- {} {} `{}` - {}".format(u["cat"], log_link(u), u["kn"] or u"(sem Keynote)", u["reason"]))
        if len(cs["unassigned"]) > 40:
            output.print_md(u"- ... e mais {} (veja o relatório HTML).".format(len(cs["unassigned"]) - 40))

        self.records = sort_records(records)
        no_bound = [r for r in self.records if r["boundary"] != "ok"]
        output.print_md(u"**[3/6] Contornos analisados:** {} ambiente(s) com contorno, {} sem contorno "
                        u"(contorno na face de acabamento; margem de busca do forro: {} m).".format(
                            len(self.records) - len(no_bound), len(no_bound), fmt_m(margin)))
        for r in no_bound[:30]:
            output.print_md(u"- Ambiente {} `{}` - {}".format(
                log_link({"id": r["info"]["id"], "link": u""}), r["info"]["number"],
                BOUNDARY_LABELS.get(r["boundary"], r["boundary"])))

        compute_changes(self.records, self.targets, self.tag_index)
        for r in self.records:
            r["status"] = room_status(r)
        self._log_keynotes()
        self._write_report()
        self.btn_report.IsEnabled = True
        self.show_preview()
        self.tb_status.Text = self._summary() + u" " + self.tb_status.Text
        self._alert_missing_tags()

    def _alert_missing_tags(self):
        """Rule B error handling: tell exactly which identifier type is missing."""
        need = OrderedDict()
        for rec in self.records:
            for t in rec.get("tags", []):
                if t["type"] is None:
                    need.setdefault(t["msg"], []).append(rec["info"]["number"])
        if not need:
            return
        lines = []
        for msg, rooms in need.items():
            uniq = sorted(set(rooms), key=natural_key)
            lines.append(u"{}\n   Necessário para {} ambiente(s): {}{}".format(
                msg, len(uniq), u", ".join(uniq[:15]), u"..." if len(uniq) > 15 else u""))
            output.print_md(u"**ERRO - identificador:** {} Ambientes: {}".format(msg, u", ".join(uniq)))
        forms.alert(u"Tipo de identificador não encontrado no projeto:\n\n" + u"\n\n".join(lines) +
                    u"\n\nOs acabamentos desses ambientes podem ser gravados normalmente; apenas o "
                    u"identificador deles não será trocado.")

    def _log_keynotes(self):
        issues = [(r, iss) for r in self.records for iss in r.get("issues", [])]
        mismatches = [(r, m) for r in self.records for f in r["finishes"].values() for m in f["mismatches"]]
        output.print_md(u"**[4/6] Keynotes coletadas** (fonte: {}; classificação pelo prefixo: {}). "
                        u"{} elemento(s) sem Keynote, {} inconsistência(s) de prefixo/categoria.".format(
                            dict(KEYNOTE_SOURCES)[self.mode],
                            u", ".join(u"{} = {}".format(p, finish_label(f)) for p, f in PREFIX_RULES.items()),
                            len(issues), len(mismatches)))
        if not self.keynote_texts:
            output.print_md(u"_Arquivo de Keynotes não carregado ou vazio - as chaves não são validadas._")
        for r, iss in issues[:60]:
            output.print_md(u"- Ambiente `{}`: {} {} - {}".format(
                r["info"]["number"], iss["cat"], log_link(iss), iss["msg"]))
        if len(issues) > 60:
            output.print_md(u"- ... e mais {} (veja o relatório HTML).".format(len(issues) - 60))
        for r, m in mismatches[:40]:
            output.print_md(u"- **Inconsistência** Ambiente `{}`: {}".format(r["info"]["number"], m))
        if len(mismatches) > 40:
            output.print_md(u"- ... e mais {} inconsistências (veja o relatório HTML).".format(len(mismatches) - 40))

    def _summary(self):
        st = [r["status"] for r in self.records]
        return (u"{} ambientes analisados - {} OK, {} com atenção, {} com erro.".format(
                    len(st), st.count("OK"), st.count("WARNING"), st.count("ERROR")))

    def _write_report(self):
        try:
            data = build_json(self.records, self.targets, self.mode, self.keynote_texts, self.problems,
                              self.contain_stats, self.search_above_m, self.tag_index)
            write_report(data)
        except Exception as ex:
            output.print_md(u"**ERRO:** não foi possível gerar o relatório HTML: `{}`".format(to_unicode(ex)))

    # ---------------- PREVIEW ----------------
    def show_preview(self, log=True):
        """Fill the preview grid. Every Fill / Overwrite row starts ticked;
        overwrites are shown in amber and counted in the confirmation."""
        if not self.records:
            return
        compute_changes(self.records, self.targets, self.tag_index)   # re-read current values
        t = DataTable("preview")
        for col, typ in (("Apply", Boolean), ("Editable", Boolean), ("Kind", String), ("Room", String),
                         ("Name", String), ("Level", String), ("Parameter", String), ("Current", String),
                         ("New", String), ("Change", String), ("Note", String), ("Group", String)):
            t.Columns.Add(col, clr.GetClrType(typ))
        self.row_keys = []
        counts = {"Fill": 0, "Overwrite": 0, "Clear": 0, "Keep": 0, "Overflow": 0, "Tag": 0, "TagMissing": 0}
        overflow_log = []
        headers = set()
        for rec in self.records:
            first_row = t.Rows.Count
            block = None
            for key, ch in rec["changes"].items():
                if ch["overflow"]:
                    counts["Overflow"] += len(ch["overflow"])
                    overflow_log.append((rec, key, ch["overflow"]))
                for i, sl in enumerate(ch["slots"]):
                    if sl["kind"] == "Keep":
                        counts["Keep"] += 1
                    if sl["kind"] not in ("Fill", "Overwrite", "Clear"):
                        continue
                    if block is None:
                        block = worksharing_block(rec["room"])
                    counts[sl["kind"]] += 1
                    row = t.NewRow()
                    editable = not block
                    row["Editable"] = editable
                    row["Apply"] = editable
                    row["Kind"] = sl["kind"]
                    row["Room"] = rec["info"]["number"]
                    row["Name"] = rec["info"]["name"]
                    row["Level"] = rec["info"]["level"]
                    row["Parameter"] = sl["name"]
                    row["Current"] = sl["current"]
                    row["New"] = sl["new"]
                    row["Change"] = CHANGE_LABELS.get(sl["kind"], sl["kind"])
                    notes = [block] if block else []
                    if ch["overflow"]:
                        notes.append(u"não couberam: " + join_keys(ch["overflow"]))
                    if rec["finishes"][key]["mismatches"]:
                        notes.append(u"inconsistência de prefixo/categoria - veja o relatório")
                    row["Note"] = u"; ".join(notes)
                    t.Rows.Add(row)
                    self.row_keys.append((rec, key, i))
            # rule B: identifier tags of this room
            for j, tg in enumerate(rec.get("tags", [])):
                missing = tg["type"] is None
                tblock = u"" if missing else worksharing_block(tg["tag"])
                kind = "TagMissing" if missing else "Tag"
                counts[kind] += 1
                row = t.NewRow()
                editable = not missing and not tblock
                row["Editable"] = editable
                row["Apply"] = editable
                row["Kind"] = kind
                row["Room"] = rec["info"]["number"]
                row["Name"] = rec["info"]["name"]
                row["Level"] = rec["info"]["level"]
                row["Parameter"] = u"Identificador (vista: {})".format(tg["view"])
                row["Current"] = tg["current"]
                row["New"] = tg["new"]
                row["Change"] = CHANGE_LABELS[kind]
                row["Note"] = u"; ".join([x for x in (u"{} linha(s)".format(tg["rows"]), tblock, tg["msg"]) if x])
                t.Rows.Add(row)
                self.row_keys.append((rec, "__tag__", j))
            rows = [t.Rows[k] for k in range(first_row, t.Rows.Count)]
            if rows:
                header = group_header(rec, rows)
                if header in headers:                     # same number / name / level twice
                    header += u"  (Id {})".format(rec["info"]["id"])
                headers.add(header)
                for row in rows:
                    row["Group"] = header
        self.table = t
        self.grid.ItemsSource = t.DefaultView
        self._group_by_room()
        self.btn_update.IsEnabled = any(bool(r["Editable"]) for r in t.Rows)
        if log:
            output.print_md(u"**[5/6] Pré-visualização gerada:** {} campo(s) a preencher, {} a substituir, {} a limpar, "
                            u"{} campo(s) mantidos onde nada foi encontrado; {} identificador(es) a trocar de tipo, "
                            u"{} com tipo ausente.".format(
                                counts["Fill"], counts["Overwrite"], counts["Clear"], counts["Keep"],
                                counts["Tag"], counts["TagMissing"]))
            for rec, key, extra in overflow_log[:40]:
                output.print_md(u"- **Sem campo livre** Ambiente `{}` {}: {} (só há {} campos)".format(
                    rec["info"]["number"], finish_label(key), join_keys(extra), len(FINISH_SLOTS[key])))
        msg_over = (u" {} Keynote(s) não couberam nos campos - veja a coluna Observação.".format(counts["Overflow"])
                    if counts["Overflow"] else u"")
        if t.Rows.Count == 0:
            self.tb_status.Text = u"Nada a atualizar - todos os valores encontrados já estão nos ambientes." + msg_over
        else:
            self.tb_status.Text = (u"{} campo(s) a preencher, {} a substituir (âmbar), {} a limpar (rosa) e {} "
                                   u"identificador(es) a trocar (lilás). Desmarque o que não quiser gravar e clique em "
                                   u"Atualizar Ambientes.{}{}".format(
                                       counts["Fill"], counts["Overwrite"], counts["Clear"], counts["Tag"], msg_over,
                                       u" {} identificador(es) sem o tipo necessário - veja a coluna Observação.".format(
                                           counts["TagMissing"]) if counts["TagMissing"] else u""))

    def _group_by_room(self):
        """One collapsible group per room (header = room + summary of its
        changes). If WPF refuses the grouping, the flat list stays usable."""
        grouped = False
        try:
            from System.Windows.Data import CollectionViewSource, PropertyGroupDescription
            view = CollectionViewSource.GetDefaultView(self.grid.ItemsSource)
            if view is not None and view.CanGroup:
                view.GroupDescriptions.Clear()
                view.GroupDescriptions.Add(PropertyGroupDescription("Group"))
                grouped = True
        except Exception as ex:
            output.print_md(u"_Agrupamento por ambiente indisponível ({}) - lista simples._".format(to_unicode(ex)))
        try:
            from System.Windows import Visibility
            # number / name / level are in the group header
            for i in (1, 2, 3):
                self.grid.Columns[i].Visibility = Visibility.Collapsed if grouped else Visibility.Visible
        except Exception:
            pass
        has_rows = self.table is not None and self.table.Rows.Count > 0
        self.btn_expand.IsEnabled = grouped and has_rows
        self.btn_collapse.IsEnabled = grouped and has_rows

    def _expand_all(self, value):
        try:
            from System.Windows.Media import VisualTreeHelper
            from System.Windows.Controls import Expander
            stack = [self.grid]
            while stack:
                el = stack.pop()
                if isinstance(el, Expander):
                    el.IsExpanded = value
                try:
                    n = VisualTreeHelper.GetChildrenCount(el)
                except Exception:
                    continue
                for i in range(n):
                    stack.append(VisualTreeHelper.GetChild(el, i))
        except Exception as ex:
            output.print_md(u"_Não foi possível expandir/recolher: {}_".format(to_unicode(ex)))

    def on_expand(self, sender, args):
        self._expand_all(True)

    def on_collapse(self, sender, args):
        self._expand_all(False)

    # ---------------- CONFIRM + TRANSACTION + WRITE ----------------
    def on_update(self, sender, args):
        if self.table is None:
            return
        try:
            self.grid.CommitEdit()
        except Exception:
            pass
        todo, tag_todo = [], []
        for i in range(self.table.Rows.Count):
            row = self.table.Rows[i]
            if bool(row["Apply"]) and bool(row["Editable"]):
                rec, key, slot = self.row_keys[i]
                if key == "__tag__":
                    tag_todo.append((rec, rec["tags"][slot]))
                    continue
                sl = rec["changes"][key]["slots"][slot]
                todo.append((rec, key, slot, sl["new"], sl["kind"]))
        if not todo and not tag_todo:
            forms.alert(u"Nenhuma linha está marcada na pré-visualização.")
            return
        n_rooms = len(set([eid_int(r["room"].Id) for r, _, _, _, _ in todo] +
                          [eid_int(r["room"].Id) for r, _ in tag_todo]))
        n_ow = len([1 for _, _, _, _, k in todo if k == "Overwrite"])
        n_clear = len([1 for _, _, _, _, k in todo if k == "Clear"])
        msg = u"Gravar {} campo(s) em {} ambiente(s)?".format(len(todo), n_rooms)
        if n_ow:
            msg += u"\n\n{} deles SUBSTITUEM um valor existente diferente.".format(n_ow)
        if n_clear:
            msg += u"\n{} campo(s) com valor antigo serão LIMPOS.".format(n_clear)
        if tag_todo:
            msg += u"\n\n{} identificador(es) de ambiente terão o tipo trocado.".format(len(tag_todo))
        if not forms.alert(msg, yes=True, no=True):
            return

        written, errors, tags_done = [], [], []
        tx = DB.Transaction(doc, u"Acabamentos de Ambiente - atualizar ambientes")
        try:
            tx.Start()
            for rec, key, slot, value, kind in todo:
                target = self.targets[key][slot]
                label = target.name
                try:
                    p = target.get(rec["room"])
                    if p is None or p.IsReadOnly:
                        errors.append((rec, label, u"parâmetro ausente ou somente leitura neste ambiente"))
                    elif p.Set(value):
                        written.append((rec, label))
                    else:
                        errors.append((rec, label, u"o Revit recusou o valor"))
                except Exception as ex:
                    errors.append((rec, label, to_unicode(ex)))
            # rule B - after the fields are written, count the rows really
            # filled in the room and switch its identifier tags to that type
            for rec, tg in tag_todo:
                tag = tg["tag"]
                try:
                    n_now = rows_in_room(rec["room"], self.targets)
                    if not n_now:
                        continue
                    target, name, tmsg = self.tag_index.target(tag, n_now)
                    if target is None:
                        errors.append((rec, u"Identificador", tmsg))
                    elif eid_int(target.Id) != eid_int(tag.GetTypeId()):
                        try:
                            tag.ChangeTypeId(target.Id)
                        except Exception:
                            tag.RoomTagType = target          # same change, RoomTag API
                        tags_done.append((rec, name))
                except Exception as ex:
                    errors.append((rec, u"Identificador", to_unicode(ex)))
            status = tx.Commit()
            if status != DB.TransactionStatus.Committed:
                raise Exception(u"A transação terminou com status {}".format(status))
        except Exception as ex:
            if tx.HasStarted() and not tx.HasEnded():
                tx.RollBack()
            output.print_md(u"**ERRO:** atualização desfeita - nada foi gravado. `{}`".format(to_unicode(ex)))
            forms.alert(u"A atualização falhou e foi desfeita. Nada foi gravado.\n\n{}".format(to_unicode(ex)))
            return
        finally:
            tx.Dispose()

        for rec, key in written + tags_done:
            rec["updated"] = True
        for rec, key, err in errors:
            rec["write_error"] = (rec.get("write_error") or u"") + u"{}: {}. ".format(key, err)
        output.print_md(u"**[6/6] Ambientes atualizados:** {} campo(s) gravados em {} ambiente(s), {} identificador(es) "
                        u"com tipo trocado, {} erro(s).".format(
                            len(written), len(set(eid_int(r["room"].Id) for r, _ in written)), len(tags_done), len(errors)))
        for rec, key, err in errors[:40]:
            output.print_md(u"- **ERRO** Ambiente {} `{}` {}: {}".format(
                log_link({"id": rec["info"]["id"], "link": u""}), rec["info"]["number"], key, err))

        compute_changes(self.records, self.targets, self.tag_index)   # read back what is in Revit now
        for r in self.records:
            r["status"] = room_status(r)
        self._write_report()
        self.show_preview(log=False)                     # what is left (unticked / blocked rows)
        self.tb_status.Text = u"{} campo(s) gravados e {} identificador(es) trocados{}. Abra o relatório HTML para conferir.".format(
            len(written), len(tags_done), u" - {} erro(s), veja o log".format(len(errors)) if errors else u"")

    # ---------------- report / close ----------------
    def on_report(self, sender, args):
        if not os.path.exists(REPORT_PATH):
            self._write_report()
        if not open_in_browser(REPORT_PATH):
            forms.alert(u"O relatório foi salvo, mas não abriu automaticamente:\n\n{}".format(REPORT_PATH))

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
    output.print_md(u"**[1/6] Script iniciado.** Documento: `{}`".format(to_unicode(doc.Title)))
    if doc.IsFamilyDocument:
        forms.alert(u"Abra um projeto, não uma família.", exitscript=True)
    rooms = collect_rooms()
    if not rooms:
        forms.alert(u"Nenhum ambiente encontrado neste projeto.", exitscript=True)
    win = RoomFinishWindow(rooms)
    for msg in win.problems:
        output.print_md(u"**ERRO:** {}".format(msg))
    if win.problems:
        forms.alert(u"\n".join(win.problems) +
                    u"\n\nEsses parâmetros aparecem no relatório, mas não serão gravados. "
                    u"Nada é criado automaticamente.")
    win.ShowDialog()


main()
