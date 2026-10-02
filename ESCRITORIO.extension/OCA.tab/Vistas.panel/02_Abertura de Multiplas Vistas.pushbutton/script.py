# -*- coding: utf-8 -*-
"""Abertura de Múltiplas Vistas - extensão do 01_Abertura de Vistas para vários ambientes.

Mesma lógica do 01 (geometria, recorte orientado, nomenclatura, tipos de
vista, diálogo), aplicada a CADA ambiente da seleção:
  1. Usa a seleção atual (ou pede para selecionar) Ambientes (Rooms).
  2. Pergunta a margem (m) e o que fazer - UMA vista POR AMBIENTE:
       - criar nova vista de planta (Piso, Forro, Estrutural, Área), no
         nível de cada ambiente;
       - criar vista de chamada (Callout) na vista ativa.
     ("Recortar a vista ativa" do 01 não se aplica: uma vista só não pode
     ter o recorte de vários ambientes separados.)
  3. VIEW NAME escolhido uma vez e composto para cada ambiente; as opções
     "Nº - AMBIENTE - ..." usam o número/nome de cada um.
  4. Recorte orientado + margem calculado por ambiente (igual ao 01).
  5. Ordem: nível (de baixo para cima) e número do ambiente. As vistas
     criadas são abertas nessa ordem; a primeira fica ativa no final.
  6. Ambiente sem vista (não colocado, sem geometria ou falha na criação)
     é ignorado; ao final, um resumo lista o que precisa ser revisado.
"""

__title__ = "Abertura de\nMúltiplas Vistas"
__doc__ = ("Cria uma planta ou chamada recortada para cada ambiente "
           "selecionado, com margem em metros, e abre todas as vistas.")

import re
import math

from pyrevit import revit, DB, script, forms
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

BIP = DB.BuiltInParameter
VT = DB.ViewType
INVALID_ID = DB.ElementId.InvalidElementId


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


def m_to_ft(meters):
    return DB.UnitUtils.ConvertToInternalUnits(meters, DB.UnitTypeId.Meters)


HALF_PI = math.pi / 2.0
QUARTER_PI = math.pi / 4.0
ANGLE_BIN = math.radians(0.5)             # resolução do histograma de direções
N_BINS = int(round(HALF_PI / ANGLE_BIN))  # 180 bins cobrindo 0..90°
ANGLE_SNAP = math.radians(0.01)           # abaixo disso o elemento está "alinhado"
MIN_SEG = m_to_ft(0.001)                  # ignora arestas < 1 mm na projeção
MIN_CROP = m_to_ft(0.01)                  # recorte mínimo de 1 cm por lado
LOCATION_WEIGHT = 10.0                    # peso extra do eixo de paredes retas


# ------------------------------------------------------------------
# 1. Categorias aceitas + seleção
# ------------------------------------------------------------------
def _bic_int(name):
    bic = getattr(DB.BuiltInCategory, name, None)   # OST_Toposolid: Revit 2024+
    if bic is None:
        return None
    try:
        return int(bic)
    except Exception:
        return None


# comando 02: só ambientes (o 01 aceita também linhas, terreno, pisos e paredes)
ALLOWED_CATS = {}
for _name, _label in (("OST_Rooms", u"Ambiente"),):
    _v = _bic_int(_name)
    if _v is not None:
        ALLOWED_CATS[_v] = _label


def elem_cat_int(el):
    try:
        cat = el.Category
        return eid_int(cat.Id) if cat is not None else None
    except Exception:
        return None


class AllowedFilter(ISelectionFilter):
    def AllowElement(self, el):
        return elem_cat_int(el) in ALLOWED_CATS

    def AllowReference(self, ref, point):
        return False


def _unique(elements):
    seen, out = set(), []
    for el in elements:
        key = eid_int(el.Id)
        if key not in seen:
            seen.add(key)
            out.append(el)
    return out


def get_target_elements():
    """Seleção atual filtrada; se vazia, pede para o usuário selecionar."""
    current = list(revit.get_selection().elements)
    picked = [el for el in current if elem_cat_int(el) in ALLOWED_CATS]
    ignored = len(current) - len(picked)
    if picked:
        return _unique(picked), ignored
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, AllowedFilter(),
            u"Selecione os ambientes e clique em Concluir")
    except Exception:
        script.exit()   # Esc
    return _unique([doc.GetElement(r) for r in refs]), ignored


def elem_label(el):
    if isinstance(el, Room):
        num = el.get_Parameter(BIP.ROOM_NUMBER)
        name = el.get_Parameter(BIP.ROOM_NAME)
        parts = [to_unicode(p.AsString()) for p in (num, name)
                 if p is not None and p.AsString()]
        return u" ".join(parts) or u"Ambiente [{}]".format(eid_int(el.Id))
    kind = ALLOWED_CATS.get(elem_cat_int(el), u"Elemento")
    type_name = u""
    try:
        etype = doc.GetElement(el.GetTypeId())
        if etype is not None:
            type_name = to_unicode(DB.Element.Name.GetValue(etype))
    except Exception:
        pass
    return u"{} {} [{}]".format(kind, type_name, eid_int(el.Id)).replace(u"  ", u" ")


# ------------------------------------------------------------------
# 2. Geometria: pontos + arestas 3D de cada elemento
# ------------------------------------------------------------------
GEOM_OPT = DB.Options()
GEOM_OPT.DetailLevel = DB.ViewDetailLevel.Fine
GEOM_OPT.ComputeReferences = False
GEOM_OPT.IncludeNonVisibleObjects = False


class ElemGeom(object):
    def __init__(self, el):
        self.el = el
        self.label = elem_label(el)
        self.points = []   # XYZ
        self.segs = []     # (XYZ, XYZ, peso) - usados para achar a orientação

    def add_polyline(self, pts, weight=1.0):
        pts = list(pts)
        self.points.extend(pts)
        for a, b in zip(pts[:-1], pts[1:]):
            self.segs.append((a, b, weight))

    def add_curve(self, crv, weight=1.0):
        try:
            self.add_polyline(crv.Tessellate(), weight)
        except Exception:
            pass

    def walk(self, geom):
        if geom is None:
            return
        for obj in geom:
            if isinstance(obj, DB.Solid):
                for edge in obj.Edges:
                    self.add_polyline(edge.Tessellate())
            elif isinstance(obj, DB.GeometryInstance):
                self.walk(obj.GetInstanceGeometry())
            elif isinstance(obj, DB.Curve):
                self.add_curve(obj)
            elif isinstance(obj, DB.PolyLine):
                self.add_polyline(obj.GetCoordinates())
            elif isinstance(obj, DB.Mesh):
                # topografia antiga: a triangulação não indica orientação,
                # então entram só os vértices (orientação sai do casco convexo)
                self.points.extend(obj.Vertices)


def extract_geometry(el):
    g = ElemGeom(el)
    if isinstance(el, DB.SpatialElement):
        loops = el.GetBoundarySegments(DB.SpatialElementBoundaryOptions()) or []
        for loop in loops:
            for seg in loop:
                g.add_curve(seg.GetCurve())
        if isinstance(el, Room):
            try:
                g.walk(el.ClosedShell)   # altura do ambiente (cortes/elevações)
            except Exception:
                pass
    elif isinstance(el, DB.CurveElement):
        g.add_curve(el.GeometryCurve)
    else:
        g.walk(el.get_Geometry(GEOM_OPT))
        loc = getattr(el, "Location", None)
        if (isinstance(el, DB.Wall) and isinstance(loc, DB.LocationCurve)
                and isinstance(loc.Curve, DB.Line)):
            # o eixo da parede reta manda na orientação (juntas em ângulo não)
            c = loc.Curve
            g.segs.append((c.GetEndPoint(0), c.GetEndPoint(1), LOCATION_WEIGHT))
    return g


# ------------------------------------------------------------------
# 3. Orientação + retângulo com margem, no sistema do CropBox da vista
# ------------------------------------------------------------------
class Frame(object):
    """Coordenadas do CropBox de uma vista: u = direita, v = cima, w = profundidade."""

    def __init__(self, view):
        self.T = view.CropBox.Transform
        self.inv = self.T.Inverse
        self.w = self.inv.OfPoint(view.Origin).Z

    def to2d(self, p):
        q = self.inv.OfPoint(p)
        return (q.X, q.Y)

    def to3d(self, uv):
        return self.T.OfPoint(DB.XYZ(uv[0], uv[1], self.w))


def dominant_angle(segs2d):
    """Direção predominante (0..90°) ponderada pelo comprimento das arestas.

    Histograma com bins de 0,5° (circular em 90°), pico suavizado com os
    vizinhos e refinado pela média ponderada das arestas perto do pico.
    """
    bins = [0.0] * N_BINS
    samples = []
    for a, b, weight in segs2d:
        dx, dy = b[0] - a[0], b[1] - a[1]
        length = math.hypot(dx, dy)
        if length < MIN_SEG:
            continue
        ang = math.atan2(dy, dx) % HALF_PI
        w = length * weight
        bins[int(ang / ANGLE_BIN) % N_BINS] += w
        samples.append((ang, w))
    if not samples:
        return None
    best = max(range(N_BINS),
               key=lambda i: bins[i - 1] + bins[i] + bins[(i + 1) % N_BINS])
    center = (best + 0.5) * ANGLE_BIN
    sw = swd = 0.0
    for ang, w in samples:
        d = (ang - center + QUARTER_PI) % HALF_PI - QUARTER_PI
        if abs(d) <= 1.5 * ANGLE_BIN:
            sw += w
            swd += w * d
    return (center + swd / sw) % HALF_PI


def convex_hull(pts):
    pts = sorted(set(pts))
    if len(pts) < 3:
        return pts

    def cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    lower, upper = [], []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]


def _extent(pts, ang):
    c, s = math.cos(ang), math.sin(ang)
    xs = [u * c + v * s for u, v in pts]
    ys = [-u * s + v * c for u, v in pts]
    return min(xs), min(ys), max(xs), max(ys)


def min_area_angle(pts2d):
    """Fallback sem arestas úteis (ex.: malha de topografia): menor retângulo."""
    hull = convex_hull(pts2d)
    if len(hull) < 3:
        return 0.0
    best_ang, best_area = 0.0, None
    for i in range(len(hull)):
        a, b = hull[i], hull[(i + 1) % len(hull)]
        ang = math.atan2(b[1] - a[1], b[0] - a[0]) % HALF_PI
        x0, y0, x1, y1 = _extent(hull, ang)
        area = (x1 - x0) * (y1 - y0)
        if best_area is None or area < best_area - 1e-9:
            best_ang, best_area = ang, area
    return best_ang


def _snap(ang):
    if ang < ANGLE_SNAP or HALF_PI - ang < ANGLE_SNAP:
        return 0.0
    return ang


class CropCalc(object):
    """Recorte calculado para UMA vista (coordenadas no Frame dessa vista)."""

    def __init__(self, geoms, frame, margin_ft):
        self.frame = frame
        pts, segs = [], []
        for g in geoms:
            pts.extend(frame.to2d(p) for p in g.points)
            segs.extend((frame.to2d(a), frame.to2d(b), w) for a, b, w in g.segs)
        if not pts:
            raise ValueError(u"Elemento sem geometria utilizável.")

        ang = dominant_angle(segs)
        self.source = u"arestas do elemento"
        if ang is None:
            ang = min_area_angle(pts)
            self.source = u"menor retângulo envolvente"
        self.angle = _snap(ang)

        x0, y0, x1, y1 = _extent(pts, self.angle)
        m = margin_ft
        if (x1 - x0) + 2 * m < MIN_CROP or (y1 - y0) + 2 * m < MIN_CROP:
            raise ValueError(u"Recorte com largura/altura nula - use margem > 0 "
                             u"para elementos lineares.")
        self.corners = self._corners(x0 - m, y0 - m, x1 + m, y1 + m)
        self.width = (x1 - x0) + 2 * m
        self.height = (y1 - y0) + 2 * m

    def _corners(self, x0, y0, x1, y1):
        c, s = math.cos(self.angle), math.sin(self.angle)
        return [(x * c - y * s, x * s + y * c)
                for x, y in ((x0, y0), (x1, y0), (x1, y1), (x0, y1))]

    def envelope(self):
        us = [p[0] for p in self.corners]
        vs = [p[1] for p in self.corners]
        return min(us), min(vs), max(us), max(vs)


# ------------------------------------------------------------------
# 4. Aplicação do recorte
# ------------------------------------------------------------------
def _clear_shape(mgr):
    try:
        if getattr(mgr, "ShapeSet", True):
            mgr.RemoveCropRegionShape()
    except Exception:
        pass


def _set_box(view, u0, v0, u1, v1):
    old = view.CropBox
    bb = DB.BoundingBoxXYZ()
    bb.Transform = old.Transform
    bb.Min = DB.XYZ(u0, v0, old.Min.Z)
    bb.Max = DB.XYZ(u1, v1, old.Max.Z)
    view.CropBox = bb


def apply_crop(view, calc):
    """Ativa + mostra o recorte e aplica os limites. Devolve (método, aviso)."""
    view.CropBoxActive = True
    view.CropBoxVisible = True
    mgr = view.GetCropRegionShapeManager()

    # sempre forma de 4 lados, inclusive quando o elemento está alinhado à vista
    pts = [calc.frame.to3d(p) for p in calc.corners]
    loop = DB.CurveLoop()
    for i in range(4):
        loop.Append(DB.Line.CreateBound(pts[i], pts[(i + 1) % 4]))
    if mgr.CanHaveShape and mgr.IsCropRegionShapeValid(loop):
        mgr.SetCropShape(loop)
        return u"Forma orientada (4 quinas)", None

    # a vista recusou a forma: CropBox com o envelope das 4 quinas (idêntico
    # ao recorte orientado quando o elemento está alinhado à vista)
    _clear_shape(mgr)
    _set_box(view, *calc.envelope())
    if calc.angle == 0.0:
        return u"Retângulo (CropBox)", None
    return (u"Envelope (CropBox)",
            u"A vista não aceita forma de recorte; aplicado o retângulo que "
            u"envolve as 4 quinas orientadas.")


# ------------------------------------------------------------------
# 5. Níveis e tipos de vista
# ------------------------------------------------------------------
LEVELS = sorted(DB.FilteredElementCollector(doc).OfClass(DB.Level),
                key=lambda l: l.ProjectElevation)


def element_level_id(el):
    try:
        lid = el.LevelId
        if lid is not None and lid != INVALID_ID:
            return lid
    except Exception:
        pass
    if el.ViewSpecific:   # linhas de detalhe: nível da vista dona
        owner = doc.GetElement(el.OwnerViewId)
        gl = getattr(owner, "GenLevel", None)
        if gl is not None:
            return gl.Id
    bb = el.get_BoundingBox(None)   # terreno etc.: nível logo abaixo do elemento
    if bb is not None and LEVELS:
        below = [l for l in LEVELS if l.ProjectElevation <= bb.Min.Z + 1e-6]
        return (below[-1] if below else LEVELS[0]).Id
    return None


def pick_plan_level(geoms):
    """Uma única planta: nível da vista ativa se algum elemento estiver nele,
    senão o nível mais baixo entre os elementos. Devolve (Level, outros nomes)."""
    found = {}
    for g in geoms:
        lid = element_level_id(g.el)
        if lid is not None:
            found[eid_int(lid)] = doc.GetElement(lid)
    if not found:
        return None, []
    gl = getattr(active_view, "GenLevel", None)
    if gl is not None and eid_int(gl.Id) in found:
        chosen = found[eid_int(gl.Id)]
    else:
        chosen = min(found.values(), key=lambda l: l.ProjectElevation)
    others = [to_unicode(l.Name) for k, l in found.items() if k != eid_int(chosen.Id)]
    return chosen, others


PLAN_FAMILIES = (
    (DB.ViewFamily.FloorPlan, u"Planta de Piso"),
    (DB.ViewFamily.CeilingPlan, u"Planta de Forro"),
    (DB.ViewFamily.StructuralPlan, u"Planta Estrutural"),
)
PLAN_VIEWTYPES = (VT.FloorPlan, VT.CeilingPlan, VT.EngineeringPlan)
CROPPABLE_VIEWTYPES = PLAN_VIEWTYPES + (VT.AreaPlan, VT.Section, VT.Elevation, VT.Detail)
CALLOUT_PARENTS = PLAN_VIEWTYPES + (VT.Section, VT.Elevation, VT.Detail)

VIEW_FAMILY_TYPES = list(DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType))


def _vft_name(vft):
    return to_unicode(DB.Element.Name.GetValue(vft))


def plan_type_items():
    items = []
    for fam, label in PLAN_FAMILIES:
        for vft in VIEW_FAMILY_TYPES:
            if vft.ViewFamily == fam:
                items.append({"label": u"{}: {}".format(label, _vft_name(vft)),
                              "kind": "plan", "id": vft.Id})
    if any(v.ViewFamily == DB.ViewFamily.AreaPlan for v in VIEW_FAMILY_TYPES):
        for scheme in DB.FilteredElementCollector(doc).OfClass(DB.AreaScheme):
            items.append({"label": u"Planta de Área: {}".format(to_unicode(scheme.Name)),
                          "kind": "area", "id": scheme.Id})
    return items


def callout_type_items(parent):
    if parent.IsTemplate or parent.ViewType not in CALLOUT_PARENTS:
        return []
    fams = [DB.ViewFamily.Detail]
    if parent.ViewType in PLAN_VIEWTYPES:
        parent_type = doc.GetElement(parent.GetTypeId())
        if parent_type is not None:
            fams.insert(0, parent_type.ViewFamily)
    items = []
    for fam in fams:
        for vft in VIEW_FAMILY_TYPES:
            if vft.ViewFamily == fam:
                kind = u"Detalhe" if fam == DB.ViewFamily.Detail else u"Planta"
                items.append({"label": u"Chamada ({}): {}".format(kind, _vft_name(vft)),
                              "kind": "callout", "id": vft.Id})
    return items


def view_type_label(view):
    try:
        return to_unicode(view.ViewType.ToString())
    except Exception:
        return u""


# ------------------------------------------------------------------
# 6. Nomenclatura: VIEW NAME (PT) + Title on Sheet - English (automático)
# ------------------------------------------------------------------
EN_TITLE_PARAM = u"Title on Sheet - English"
ROOM_EN_PARAM = u"Room Name English"
EN_DASH = u" – "

# (chave, VIEW NAME, Title on Sheet - English, leva complemento PAV./AMPLIAÇÃO)
PLAN_TITLES = (
    ("cobertura", u"PLANTA DE COBERTURA", u"ROOF PLAN", False),
    ("civil", u"PLANTA CIVIL", u"FLOOR PLAN", True),
    ("layout", u"PLANTA DE LAYOUT", u"LAYOUT PLAN", True),
    ("piso", u"PLANTA DE PISO", u"FLOOR FINISH PLAN", True),
    ("forro", u"PLANTA DE FORRO", u"REFLECTED CEILING PLAN", True),
)
VERTICAL_TITLES = (
    ("corte", u"CORTE", u"SECTION", False),
    ("elevacao", u"ELEVAÇÃO", u"ELEVATION", False),
)
# complemento das plantas 2..6: (chave, texto PT, texto EN) + nome do nível
SUFFIXES = (("pav", u"- PAV.", u"LEVEL"),
            ("amp", u"- AMPLIAÇÃO", u"ENLARGED"))
MANUAL_KEY = "manual"


def room_info(room):
    """(número, nome, nome em inglês) - inglês cai no nome PT se vazio."""
    def text(p):
        return to_unicode(p.AsString()).strip() if p is not None and p.HasValue else u""
    number = text(room.get_Parameter(BIP.ROOM_NUMBER))
    name = text(room.get_Parameter(BIP.ROOM_NAME))
    name_en = text(room.LookupParameter(ROOM_EN_PARAM))
    return number, name, name_en


# vários ambientes: o rótulo mostra o marcador; o nome real sai de cada ambiente
ROOM_PREFIX_LABEL = u"<Nº> - <AMBIENTE> - "


def name_options(target):
    """VIEW NAME na ordem do padrão do escritório (1..11) para o tipo de vista."""
    if target == "plan":
        base, room_base = PLAN_TITLES, PLAN_TITLES[1:]   # cobertura não tem versão por ambiente
    else:
        base, room_base = VERTICAL_TITLES, VERTICAL_TITLES
    opts = [{"key": key, "label": pt, "pt": pt, "en": en, "suffix": suf, "room": False}
            for key, pt, en, suf in base]
    # depois dos simples, as versões "Nº - AMBIENTE - ..."
    opts.extend({"key": "room_" + key, "label": ROOM_PREFIX_LABEL + pt, "pt": pt, "en": en,
                 "suffix": suf, "room": True} for key, pt, en, suf in room_base)
    opts.append({"key": MANUAL_KEY, "label": u"Sem nome (manual ou sequência)",
                 "pt": u"", "en": u"", "suffix": False, "room": False})
    return opts


def level_text(level):
    if level is None:
        return u""
    name = to_unicode(level.Name).strip().upper()
    return re.sub(u"^PAV\\.?\\s*", u"", name)   # evita "PAV. PAV. TÉRREO"


def suffix_labels(level):
    """Itens do combo Complemento, já com o nome do nível."""
    lv = level_text(level)
    return [(pt + u" " + lv).strip() for _k, pt, _en in SUFFIXES]


def compose_names(opt, suffix_key, level, room, manual):
    """Devolve (VIEW NAME, Title on Sheet - English)."""
    if opt["key"] == MANUAL_KEY:
        return manual.strip(), u""
    pt, en = opt["pt"], opt["en"]
    if opt["suffix"]:
        lv = level_text(level)
        _k, suf_pt, suf_en = next(s for s in SUFFIXES if s[0] == suffix_key)
        pt = u" ".join(x for x in (pt, suf_pt, lv) if x)
        en = en + EN_DASH + u" ".join(x for x in (suf_en, lv) if x)
    if opt["room"] and room is not None:
        number, name, name_en = room_info(room)
        pt = u"{} - {} - {}".format(number, name, pt)
        en = u"{} - {} - {}".format(number, name_en or name, en)
    return pt, en


VIEW_NAMES = set(to_unicode(v.Name) for v in
                 DB.FilteredElementCollector(doc).OfClass(DB.View))
_BAD_NAME_CHARS = re.compile(u"[\\\\:{}\\[\\]|;<>?`~]")


def clean_name(name):
    return _BAD_NAME_CHARS.sub(u"-", name).strip()


def sequence_name(pattern):
    """'AMPLIAÇÃO #' -> 'AMPLIAÇÃO 01', 02... (primeiro livre)."""
    pattern = clean_name(pattern)
    if u"#" not in pattern:
        return pattern
    n = 1
    while pattern.replace(u"#", u"{:02d}".format(n)) in VIEW_NAMES:
        n += 1
    return pattern.replace(u"#", u"{:02d}".format(n))


def unique_view_name(base):
    base = clean_name(base)
    name, n = base, 2
    while name in VIEW_NAMES:
        name = u"{} ({})".format(base, n)
        n += 1
    VIEW_NAMES.add(name)
    return name


def en_param_bound():
    """O parâmetro compartilhado está vinculado à categoria Vistas?"""
    views_cat = eid_int(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Views).Id)
    it = doc.ParameterBindings.ForwardIterator()
    while it.MoveNext():
        if to_unicode(it.Key.Name) == EN_TITLE_PARAM:
            binding = it.Current
            if isinstance(binding, DB.InstanceBinding):
                return any(eid_int(c.Id) == views_cat for c in binding.Categories)
    return False


def apply_names(view, pt, en, notes):
    if pt and pt != to_unicode(view.Name):
        final = unique_view_name(pt)
        if final != clean_name(pt):
            notes.append(u"Já existia uma vista '{}'; usado '{}'.".format(clean_name(pt), final))
        view.Name = final
    if en:
        p = view.LookupParameter(EN_TITLE_PARAM)
        if p is None or p.IsReadOnly:
            notes.append(u"Parâmetro '{}' indisponível na vista; título em inglês não gravado."
                         .format(EN_TITLE_PARAM))
        else:
            p.Set(en)


# ------------------------------------------------------------------
# 7. Validação prévia dos modos (antes de abrir o diálogo)
# ------------------------------------------------------------------
MODES = ("plans", "callout")   # "current" do 01 não se aplica a vários ambientes
MODE_LABELS = {"plans": u"Nova vista de planta",
               "callout": u"Nova vista de chamada (Callout)"}


def target_kind(mode):
    """Tipo da vista resultante: define quais nomes fazem sentido."""
    if mode == "plans":
        return "plan"
    return "plan" if active_view.ViewType in PLAN_VIEWTYPES + (VT.AreaPlan,) else "vertical"


def check_modes(plan_items, callout_items, plan_level):
    reasons = {}
    av_name = to_unicode(active_view.Name)
    if not plan_items:
        reasons["plans"] = u"Nenhum tipo de vista de planta/área no modelo."
    elif plan_level is None:
        reasons["plans"] = u"Não foi possível identificar o nível dos elementos."
    if active_view.IsTemplate or active_view.ViewType not in CALLOUT_PARENTS:
        reasons["callout"] = (u"A vista ativa '{}' não aceita chamadas "
                              u"(use planta, corte, elevação ou detalhe).".format(av_name))
    elif not callout_items:
        reasons["callout"] = u"Nenhum tipo de vista compatível para a chamada."
    return reasons


def parse_margin(text):
    t = to_unicode(text).strip().lower().replace(u",", u".")
    if t.endswith(u"m"):
        t = t[:-1].strip()
    val = float(t)   # ValueError tratado no diálogo
    if val < 0 or val > 100:
        raise ValueError(u"fora do intervalo")
    return val


def natural_key(text):
    """'2', '10', 'A1' em ordem humana (2 antes de 10)."""
    return [(0, int(t), u"") if t.isdigit() else (1, 0, t.lower())
            for t in re.split(u"(\\d+)", to_unicode(text)) if t]


def room_sort_key(g):
    """Ordem de criação/abertura: nível (de baixo para cima), número, nome."""
    lid = element_level_id(g.el)
    level = doc.GetElement(lid) if lid is not None else None
    number, name, _en = room_info(g.el)
    return (level.ProjectElevation if level is not None else 0.0,
            natural_key(number), natural_key(name))


# ------------------------------------------------------------------
# 8. Diálogo (mesmo tema dos outros botões)
# ------------------------------------------------------------------
CROP_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Crop View Offset" Height="Auto" Width="480"
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
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="26"/>
      <Setter Property="Padding" Value="4,3,4,2"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,5,0,0"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#CFE3FF"/>
      <Setter Property="FontFamily" Value="Segoe UI"/>
      <Setter Property="Margin" Value="0,10,0,0"/>
    </Style>
  </Window.Resources>
  <StackPanel Margin="18">
    <TextBlock Text="RECORTE POR AMBIENTE  ·  MÚLTIPLAS VISTAS" FontSize="15" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock x:Name="info" TextWrapping="Wrap" FontSize="11"
               Foreground="#7A8FA9" Margin="0,0,0,6"/>

    <TextBlock Text="Margem ao redor de cada ambiente (m)"/>
    <TextBox x:Name="margin"/>

    <TextBlock Text="O que fazer (uma vista por ambiente)" Margin="0,14,0,0"/>
    <RadioButton x:Name="rb_plans" GroupName="mode"
                 Content="Criar nova vista de planta (Piso, Forro, Estrutural, Área)"/>
    <RadioButton x:Name="rb_callout" GroupName="mode"
                 Content="Criar vista de chamada (Callout) na vista ativa"/>
    <TextBlock x:Name="reasons" TextWrapping="Wrap" FontSize="11"
               Foreground="#FFB454" Margin="0,8,0,0"/>

    <TextBlock x:Name="vtype_lbl" Text="Tipo de vista"/>
    <ComboBox x:Name="vtype"/>

    <TextBlock Text="VIEW NAME" FontSize="12" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,18,0,3"/>
    <ComboBox x:Name="vname"/>
    <TextBlock x:Name="suffix_lbl" Text="Complemento"/>
    <ComboBox x:Name="vsuffix"/>
    <TextBlock x:Name="manual_lbl"
               Text="Nome manual (vazio = mantém o nome do Revit;  # = sequência 01, 02...)"/>
    <TextBox x:Name="manual"/>

    <Border Background="#131D33" BorderBrush="#26405F" BorderThickness="1"
            Padding="10,6,10,8" Margin="0,12,0,0">
      <StackPanel>
        <TextBlock x:Name="pv_lbl" Text="VIEW NAME" FontSize="10" Foreground="#7A8FA9"
                   TextWrapping="Wrap" Margin="0"/>
        <TextBlock x:Name="pv_pt" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,8"/>
        <TextBlock Text="TITLE ON SHEET - ENGLISH  (automático)" FontSize="10"
                   Foreground="#7A8FA9" Margin="0"/>
        <TextBlock x:Name="pv_en" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,0"/>
      </StackPanel>
    </Border>
    <TextBlock x:Name="name_hint" TextWrapping="Wrap" FontSize="11"
               Foreground="#FFB454" Margin="0,6,0,0"/>

    <CheckBox x:Name="open_views" Content="Abrir as vistas criadas ao final"/>

    <TextBlock x:Name="error" TextWrapping="Wrap" FontSize="11"
               Foreground="#FF4F9A" Margin="0,10,0,0"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="cancel" Content="Cancelar" Width="80" Height="28" Margin="0,0,10,0"/>
      <Button x:Name="ok" Content="Aplicar" Width="110" Height="28"/>
    </StackPanel>
  </StackPanel>
</Window>
"""


class CropWindow(forms.WPFWindow):
    def __init__(self, xaml, info, ctx, reasons, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self.ctx = ctx
        self._saved = cfg
        self._radios = {"plans": self.rb_plans, "callout": self.rb_callout}

        self.info.Text = info
        self.margin.Text = u"{:.2f}".format(cfg["margin"])
        self.reasons.Text = u"\n".join(
            u"• {}: {}".format(MODE_LABELS[k], reasons[k]) for k in MODES if k in reasons)
        self.name_hint.Text = u"\n".join(ctx["hints"])
        self.pv_lbl.Text = u"VIEW NAME  ·  exemplo (1º ambiente): {}".format(ctx["example"])
        self.open_views.IsChecked = bool(cfg["open"])
        self._suffix_idx = next(
            (i for i, s in enumerate(SUFFIXES) if s[0] == cfg["suffix"]), 0)

        enabled = [k for k in MODES if k not in reasons]
        for key, rb in self._radios.items():
            rb.IsEnabled = key in enabled
            rb.Checked += self._on_mode
        self.vname.SelectionChanged += self._on_name
        self.vsuffix.SelectionChanged += self._on_name
        self.manual.TextChanged += self._on_name
        start = cfg["mode"] if cfg["mode"] in enabled else enabled[0]
        self._radios[start].IsChecked = True
        self._refresh()

        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    @property
    def mode(self):
        for key in MODES:
            if self._radios[key].IsChecked:
                return key
        return None

    def _on_mode(self, sender, args):
        self._refresh()

    def _on_name(self, sender, args):
        self._update_names()

    def _refresh(self):
        mode = self.mode
        types = self.ctx["types"].get(mode, [])
        labels = [i["label"] for i in types]
        self.vtype.ItemsSource = labels
        if labels:
            self.vtype.SelectedItem = (self._saved["vtype"] if self._saved["vtype"] in labels
                                       else labels[0])
        self.vtype.IsEnabled = bool(labels)
        self.vtype_lbl.Opacity = 1.0 if labels else 0.4

        # nomes dependem do tipo da vista resultante (planta x corte/elevação)
        self._names = self.ctx["names"][mode]
        keys = [o["key"] for o in self._names]
        self.vname.ItemsSource = [o["label"] for o in self._names]
        if self._saved["vname"] in keys:
            self.vname.SelectedIndex = keys.index(self._saved["vname"])
        else:   # vários ambientes: por padrão, a 1ª opção "Nº - AMBIENTE - ..."
            self.vname.SelectedIndex = next(
                (i for i, o in enumerate(self._names) if o["room"]), 0)

        # complemento traz o nome do nível da vista resultante
        if self.vsuffix.SelectedIndex >= 0:
            self._suffix_idx = self.vsuffix.SelectedIndex
        self.vsuffix.ItemsSource = suffix_labels(self.ctx["level"][mode])
        self.vsuffix.SelectedIndex = self._suffix_idx
        self._update_names()

    @property
    def name_opt(self):
        idx = self.vname.SelectedIndex
        return self._names[idx] if 0 <= idx < len(self._names) else self._names[-1]

    @property
    def suffix_key(self):
        idx = self.vsuffix.SelectedIndex
        return SUFFIXES[idx if idx >= 0 else 0][0]

    def names(self):
        """Prévia com o 1º ambiente; na execução o nome é composto por ambiente."""
        opt = self.name_opt
        manual = to_unicode(self.manual.Text)
        pt, en = compose_names(opt, self.suffix_key, self.ctx["level"][self.mode],
                               self.ctx["room"], manual)
        if opt["key"] == MANUAL_KEY:
            pt = sequence_name(pt) if pt else u""
        return pt, en

    def _update_names(self):
        opt = self.name_opt
        is_manual = opt["key"] == MANUAL_KEY
        self.vsuffix.IsEnabled = opt["suffix"]
        self.suffix_lbl.Opacity = 1.0 if opt["suffix"] else 0.4
        self.manual.IsEnabled = is_manual
        self.manual_lbl.Opacity = 1.0 if is_manual else 0.4
        pt, en = self.names()
        self.pv_pt.Text = clean_name(pt) if pt else u"(nome padrão do Revit)"
        self.pv_en.Text = en or u"—"

    def _ok(self, sender, args):
        try:
            self.margin_m = parse_margin(self.margin.Text)
        except Exception:
            self.error.Text = u"Margem inválida: informe um número em metros entre 0 e 100 (ex.: 0,50)."
            return
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.Close()

    @property
    def result(self):
        types = self.ctx["types"].get(self.mode, [])
        sel = self.vtype.SelectedItem
        return {"mode": self.mode, "margin": self.margin_m,
                "type": next((i for i in types if i["label"] == sel), None),
                "vname": self.name_opt["key"], "suffix": self.suffix_key,
                "name_opt": self.name_opt, "manual": to_unicode(self.manual.Text),
                "open": bool(self.open_views.IsChecked)}


# ------------------------------------------------------------------
# 9. Execução
# ------------------------------------------------------------------
warnings = []
skipped = []   # ambientes sem vista: (rótulo, motivo)

elements, ignored = get_target_elements()
if ignored:
    warnings.append(u"{} elemento(s) da seleção ignorado(s): não são ambientes.".format(ignored))

geoms = []
for el in elements:
    if isinstance(el, DB.SpatialElement) and el.Area <= 0:
        skipped.append((elem_label(el), u"ambiente não colocado ou não delimitado"))
        continue
    g = extract_geometry(el)
    if not g.points:
        skipped.append((g.label, u"sem geometria utilizável"))
        continue
    geoms.append(g)


def skipped_lines(items):
    return [u"- {}: {}".format(label, why) for label, why in items]


if not geoms:
    forms.alert(u"Nenhum ambiente válido selecionado.\n\n"
                + u"\n".join(skipped_lines(skipped) + warnings), exitscript=True)

geoms.sort(key=room_sort_key)
room_levels = [pick_plan_level([g])[0] for g in geoms]   # cada planta no nível do seu ambiente

plan_level = next((l for l in room_levels if l is not None), None)
plan_items = plan_type_items()
callout_items = callout_type_items(active_view)
reasons = check_modes(plan_items, callout_items, plan_level)
if len(reasons) == len(MODES):
    forms.alert(u"Nenhuma ação disponível:\n\n" + u"\n".join(reasons.values()), exitscript=True)

name_hints = []
level_ids = set(eid_int(l.Id) for l in room_levels if l is not None)
if len(level_ids) > 1:
    name_hints.append(u"Ambientes em {} níveis: cada planta é criada no nível do seu ambiente "
                      u"e o complemento usa esse nível.".format(len(level_ids)))
no_en = [g for g in geoms if not room_info(g.el)[2]]
if no_en:
    name_hints.append(u"'{}' vazio em {} ambiente(s): o título em inglês usará o nome em português."
                      .format(ROOM_EN_PARAM, len(no_en)))
if len(geoms) > 1:
    name_hints.append(u"Nomes sem <Nº> - <AMBIENTE> se repetem entre as vistas e recebem "
                      u"(2), (3)...; no nome manual, use # para numerar.")
if not en_param_bound():
    name_hints.append(u"'{}' não está vinculado à categoria Vistas: o título em inglês "
                      u"não será gravado.".format(EN_TITLE_PARAM))

active_level = getattr(active_view, "GenLevel", None)
dialog_ctx = {
    "types": {"plans": plan_items, "callout": callout_items},
    "names": dict((m, name_options(target_kind(m))) for m in MODES),
    "level": {"plans": room_levels[0], "callout": active_level},
    "room": geoms[0].el,
    "example": geoms[0].label,
    "hints": name_hints,
}

config = script.get_config()
cfg = {"margin": config.get_option("margin_m", 0.5),
       "mode": config.get_option("mode", "plans"),
       "vtype": config.get_option("vtype", u""),
       "vname": config.get_option("vname", "room_civil"),
       "suffix": config.get_option("vsuffix", "amp"),
       "open": config.get_option("open_views", True)}

info = u"{} ambiente(s) → 1 vista por ambiente · vista ativa: {} ({})".format(
    len(geoms), to_unicode(active_view.Name), view_type_label(active_view))
win = CropWindow(CROP_XAML, info, dialog_ctx, reasons, cfg)
win.ShowDialog()
if not win.confirmed:
    script.exit()
opts = win.result

config.margin_m = opts["margin"]
config.mode = opts["mode"]
config.vtype = opts["type"]["label"] if opts["type"] else cfg["vtype"]
config.vname = opts["vname"]
config.vsuffix = opts["suffix"]
config.open_views = opts["open"]
script.save_config()

mode = opts["mode"]
margin_ft = m_to_ft(opts["margin"])

# --- verificações específicas do modo, antes de qualquer alteração ---------
if mode == "callout":
    visible = set(eid_int(i) for i in
                  DB.FilteredElementCollector(doc, active_view.Id)
                  .WhereElementIsNotElementType().ToElementIds())
    for g in geoms:
        if eid_int(g.el.Id) not in visible:
            warnings.append(u"{}: não está visível na vista ativa (recorte calculado mesmo assim).".format(g.label))


vtype = opts["type"]
name_opt = opts["name_opt"]
notes = []
created = []   # (ElemGeom, ElementId da vista), na ordem dos ambientes
failed = []    # (rótulo, motivo)


def create_room_view(g, level):
    """Mesma sequência do 01 para UM ambiente. Devolve (vista, avisos)."""
    room_notes = []
    if mode == "callout":
        parent_calc = CropCalc([g], Frame(active_view), margin_ft)
        u0, v0, u1, v1 = parent_calc.envelope()
        f = parent_calc.frame
        view = DB.ViewSection.CreateCallout(doc, active_view.Id, vtype["id"],
                                            f.to3d((u0, v0)), f.to3d((u1, v1)))
        name_level = active_level
    else:
        if level is None:
            raise ValueError(u"nível do ambiente não identificado.")
        if vtype["kind"] == "area":
            view = DB.ViewPlan.CreateAreaPlan(doc, vtype["id"], level.Id)
        else:
            view = DB.ViewPlan.Create(doc, vtype["id"], level.Id)
        name_level = level

    # o tipo de vista pode ter template padrão para vistas novas: remover
    view.ViewTemplateId = INVALID_ID

    pt, en = compose_names(name_opt, opts["suffix"], name_level, g.el, opts["manual"])
    if name_opt["key"] == MANUAL_KEY:
        pt = sequence_name(pt) if pt else u""
    apply_names(view, pt, en, room_notes)
    doc.Regenerate()

    # recalcula no sistema da vista-alvo (igual ao 01)
    calc = CropCalc([g], Frame(view), margin_ft)
    _method, warn = apply_crop(view, calc)
    if warn:
        room_notes.append(warn)
    return view, room_notes


# um grupo = um único Desfazer; uma transação por ambiente = falha isolada
tgroup = DB.TransactionGroup(doc, u"OCA - Abertura de Múltiplas Vistas")
tgroup.Start()
try:
    for g, level in zip(geoms, room_levels):
        names_before = set(VIEW_NAMES)
        t = DB.Transaction(doc, u"OCA - Vista {}".format(g.label))
        t.Start()
        try:
            view, room_notes = create_room_view(g, level)
            status = t.Commit()
            if status != DB.TransactionStatus.Committed:
                raise Exception(u"o Revit desfez a criação da vista ({}).".format(status))
            created.append((g, view.Id))
            notes.extend(u"{}: {}".format(g.label, n) for n in room_notes)
        except Exception as exc:
            if t.GetStatus() == DB.TransactionStatus.Started:
                t.RollBack()
            VIEW_NAMES.clear()   # libera nomes reservados pela vista desfeita
            VIEW_NAMES.update(names_before)
            failed.append((g.label, to_unicode(exc)))
finally:
    if tgroup.GetStatus() == DB.TransactionStatus.Started:
        tgroup.Assimilate()

# abre na mesma ordem dos ambientes; a primeira fica ativa ao final
if opts["open"] and created:
    for g, vid in created + created[:1]:
        try:
            uidoc.ActiveView = doc.GetElement(vid)
        except Exception as exc:
            notes.append(u"{}: vista criada, mas não foi possível abri-la ({}).".format(
                g.label, to_unicode(exc)))


# ------------------------------------------------------------------
# 10. Resumo: só aparece se houver falha ou algo a revisar
# ------------------------------------------------------------------
report = []
if skipped:
    report.append(u"Ambientes ignorados (nenhuma vista criada):")
    report.extend(skipped_lines(skipped))
if failed:
    report.append(u"Não foi possível criar a vista:")
    report.extend(skipped_lines(failed))
if warnings or notes:
    report.append(u"Avisos:")
    report.extend(u"- " + w for w in warnings + notes)

if not created:
    forms.alert(u"Nenhuma vista criada.", sub_msg=u"\n".join(report) or None)
elif report:
    forms.alert(u"{} de {} ambiente(s) com vista criada, com avisos:".format(
                    len(created), len(elements)),
                sub_msg=u"\n".join(report))
