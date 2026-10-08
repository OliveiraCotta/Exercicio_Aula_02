# -*- coding: utf-8 -*-
# Abertura de Múltiplas Vistas - extensão do 01_Abertura de Vistas para vários ambientes.
#
# Mesma lógica do 01 (geometria, recorte orientado, nomenclatura, tipos de
# vista, diálogo), aplicada a CADA ambiente da seleção:
#   1. Usa a seleção atual (ou pede para selecionar) Ambientes (Rooms).
#   2. Tipo de vista: Plantas, Elevações ou Plantas + Elevações.
#      Plantas - margem (m) e o que fazer, UMA vista POR AMBIENTE:
#        - criar nova vista de planta (Piso, Forro, Estrutural, Área), no
#          nível de cada ambiente;
#        - criar vista de chamada (Callout) na vista ativa.
#      ("Recortar a vista ativa" do 01 não se aplica: uma vista só não pode
#      ter o recorte de vários ambientes separados.)
#   3. Destino no Navegador de Projeto: pasta existente ou nova, gravada no
#      parâmetro de projeto "Pasta" (pelo qual o navegador agrupa as vistas).
#      O texto antes do 1º "_" da pasta (D000_...) vira o prefixo do View Name.
#   4. View Template opcional (só os compatíveis com o tipo de vista).
#   5. Nomenclatura escolhida uma vez e composta para cada ambiente:
#        Title on Sheet           = nome (Principal + Complemento)
#        View Name                = PREFIXO_ + nome
#        Title on Sheet - English = nome em inglês (sem prefixo)
#   6. Recorte orientado + margem calculado por ambiente (igual ao 01).
#   7. Ordem: nível (de baixo para cima) e número do ambiente. As vistas
#      criadas são abertas nessa ordem; a primeira fica ativa no final.
#   8. Ambiente sem vista (não colocado, sem geometria ou falha na criação)
#      é ignorado; ao final, um resumo lista o que precisa ser revisado.
#   9. Elevações (seção 6d): UM ElevationMarker nativo NOVO por ambiente,
#      num ponto dentro do Room, e as elevações Norte/Leste/Sul/Oeste criadas
#      a partir dele (como Vista > Elevação). Elevações já existentes são
#      ignoradas. Margem, View Template e pasta próprios das elevações.
#      View Name: PREFIXO_Nº - AMBIENTE - ELEVAÇÃO NORTE; títulos: ... 01.
#
# A descrição fica em comentário (e não em docstring) de propósito: o
# pyRevit 5.2 lê a docstring do módulo com .decode('utf-8'), que falha com
# acentos no IronPython e deixa o botão sem tooltip. O tooltip é só __doc__.

__title__ = u"Abertura de\nMúltiplas Vistas"
__doc__ = u"""Cria plantas e/ou elevações (marcador nativo do Revit) recortadas para cada ambiente selecionado, já com nome, pasta no Navegador de Projeto e View Template.

Selecione os ambientes e clique no botão."""

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


def natural_key(text):
    """'2', '10', 'A1' em ordem humana (2 antes de 10)."""
    return [(0, int(t), u"") if t.isdigit() else (1, 0, t.lower())
            for t in re.split(u"(\\d+)", to_unicode(text)) if t]


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

    def add_bbox(self):
        """Fallback (igual ao 01): 8 cantos da caixa envolvente do elemento."""
        try:
            bb = self.el.get_BoundingBox(None)
        except Exception:
            bb = None
        if bb is None:
            return
        t, lo, hi = bb.Transform, bb.Min, bb.Max
        for x in (lo.X, hi.X):
            for y in (lo.Y, hi.Y):
                for z in (lo.Z, hi.Z):
                    self.points.append(t.OfPoint(DB.XYZ(x, y, z)))

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
    try:   # leitura com erro não derruba o comando (igual ao 01)
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
    except Exception:
        pass
    if not g.points:
        g.add_bbox()
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


def family_view_type(fam):
    """ViewFamily do tipo -> ViewType da vista criada (filtra os View Templates)."""
    for f, vt in ((DB.ViewFamily.FloorPlan, VT.FloorPlan),
                  (DB.ViewFamily.CeilingPlan, VT.CeilingPlan),
                  (DB.ViewFamily.StructuralPlan, VT.EngineeringPlan),
                  (DB.ViewFamily.AreaPlan, VT.AreaPlan),
                  (DB.ViewFamily.Section, VT.Section),
                  (DB.ViewFamily.Elevation, VT.Elevation),
                  (DB.ViewFamily.Detail, VT.Detail)):
        if fam == f:
            return vt
    return None


def plan_type_items():
    items = []
    for fam, label in PLAN_FAMILIES:
        for vft in VIEW_FAMILY_TYPES:
            if vft.ViewFamily == fam:
                items.append({"label": u"{}: {}".format(label, _vft_name(vft)),
                              "kind": "plan", "id": vft.Id, "vt": family_view_type(fam)})
    if any(v.ViewFamily == DB.ViewFamily.AreaPlan for v in VIEW_FAMILY_TYPES):
        for scheme in DB.FilteredElementCollector(doc).OfClass(DB.AreaScheme):
            items.append({"label": u"Planta de Área: {}".format(to_unicode(scheme.Name)),
                          "kind": "area", "id": scheme.Id, "vt": VT.AreaPlan})
    return items


CALLOUT_KIND = ((DB.ViewFamily.FloorPlan, u"Planta"),
                (DB.ViewFamily.CeilingPlan, u"Forro"),
                (DB.ViewFamily.StructuralPlan, u"Estrutural"),
                (DB.ViewFamily.Section, u"Corte"),
                (DB.ViewFamily.Elevation, u"Elevação"),
                (DB.ViewFamily.Detail, u"Detalhe"))


def callout_type_items(parent):
    """Tipos aceitos numa chamada (igual ao 01): o mesmo tipo da vista-mãe
    (planta -> planta, corte -> corte, elevação -> elevação, para
    ampliações) e Detalhe."""
    if parent.IsTemplate or parent.ViewType not in CALLOUT_PARENTS:
        return []
    fams = [DB.ViewFamily.Detail]
    parent_type = doc.GetElement(parent.GetTypeId())
    if parent_type is not None and parent_type.ViewFamily != DB.ViewFamily.Detail:
        fams.insert(0, parent_type.ViewFamily)
    items = []
    for fam in fams:
        kind = next((k for f, k in CALLOUT_KIND if f == fam), u"Vista")
        for vft in VIEW_FAMILY_TYPES:
            if vft.ViewFamily == fam:
                items.append({"label": u"Chamada ({}): {}".format(kind, _vft_name(vft)),
                              "kind": "callout", "id": vft.Id, "vt": family_view_type(fam)})
    return items


def view_type_label(view):
    try:
        return to_unicode(view.ViewType.ToString())
    except Exception:
        return u""


# ------------------------------------------------------------------
# 6. Nomenclatura: Title on Sheet (PT) + Title on Sheet - English (automático)
#    + View Name (= PREFIXO DA PASTA + "_" + Title on Sheet)
# ------------------------------------------------------------------
EN_TITLE_PARAM = u"Title on Sheet - English"
ROOM_EN_PARAM = u"Room Name English"
EN_SEP = u" - "   # separador do complemento no título em inglês (hífen, como no PT)

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
# complemento: (chave, texto PT, texto EN). FREE_SUFFIX = texto digitado
# (ou nada); PAV./AMPLIAÇÃO + nível. Todas as opções valem para qualquer
# A - Principal (igual ao 01). O complemento sempre entra depois de " - ".
FREE_SUFFIX = "none"
SUFFIXES = ((FREE_SUFFIX, u"", u""),
            ("pav", u"PAV.", u"LEVEL"),
            ("amp", u"AMPLIAÇÃO", u"ENLARGED"))
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


# chamada numa elevação/corte: só o próprio título (+ versão por ambiente)
SIMPLE_KINDS = {"elevation": "elevacao", "section": "corte"}
MANUAL_OPT = {"key": MANUAL_KEY, "label": u"Sem nome, manual ou sequência",
              "pt": u"", "en": u"", "suffix": False, "room": False}


def name_options(target):
    """Nome principal na ordem do padrão do escritório para o tipo de vista;
    'Sem nome, manual ou sequência' no topo (igual ao 01)."""
    if target in SIMPLE_KINDS:
        base = [t for t in VERTICAL_TITLES if t[0] == SIMPLE_KINDS[target]]
        room_base = base
    elif target == "plan":
        base, room_base = PLAN_TITLES, PLAN_TITLES[1:]   # cobertura não tem versão por ambiente
    else:
        base, room_base = VERTICAL_TITLES, VERTICAL_TITLES
    opts = [dict(MANUAL_OPT)]
    opts.extend({"key": key, "label": pt, "pt": pt, "en": en, "suffix": suf, "room": False}
                for key, pt, en, suf in base)
    # depois dos simples, as versões "Nº - AMBIENTE - ..."
    opts.extend({"key": "room_" + key, "label": ROOM_PREFIX_LABEL + pt, "pt": pt, "en": en,
                 "suffix": suf, "room": True} for key, pt, en, suf in room_base)
    return opts


def level_text(level):
    if level is None:
        return u""
    name = to_unicode(level.Name).strip().upper()
    return re.sub(u"^PAV\\.?\\s*", u"", name)   # evita "PAV. PAV. TÉRREO"


def suffix_labels(level):
    """Itens do combo Complemento, já com o nome do nível."""
    lv = level_text(level)
    return [u"Sem complemento, manual ou sequência" if key == FREE_SUFFIX
            else u" ".join(x for x in (u"-", pt, lv) if x)
            for key, pt, _en in SUFFIXES]


def free_suffix(text):
    """Complemento digitado, sem hífens/travessões iniciais (o ' - ' é automático)."""
    return re.sub(u"^[\\s\\-–]+", u"", to_unicode(text)).strip()


def _join(base, extra, sep):
    """base + sep + extra; sem base, só o complemento."""
    if not extra:
        return base
    return base + sep + extra if base else extra


def compose_names(opt, suffix_key, level, room, manual, suffix_text=u"", suffix_text_en=u"",
                  manual_en=u""):
    """Devolve (Title on Sheet, Title on Sheet - English) - sem prefixo de pasta.
    Igual ao 01: nome manual -> manual (View Name/Title on Sheet) e manual_en
    (English); complemento manual -> suffix_text e suffix_text_en (cada campo
    pode ficar vazio). Todo complemento vale para qualquer opção."""
    if opt["key"] == MANUAL_KEY:
        pt, en = manual.strip(), manual_en.strip()
    else:
        pt, en = opt["pt"], opt["en"]
    if suffix_key == FREE_SUFFIX:
        pt = _join(pt, free_suffix(suffix_text), u" - ")
        en = _join(en, free_suffix(suffix_text_en), EN_SEP)
    else:
        lv = level_text(level)
        _k, suf_pt, suf_en = next(s for s in SUFFIXES if s[0] == suffix_key)
        pt += u" - " + u" ".join(x for x in (suf_pt, lv) if x)
        en += EN_SEP + u" ".join(x for x in (suf_en, lv) if x)
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


def with_prefix(prefix, title):
    """View Name = PREFIXO + '_' + nome (sem prefixo: só o nome)."""
    return u"{}_{}".format(prefix, title) if prefix and title else title


def resolve_names(prefix, title, en):
    """(View Name, Title on Sheet, English). '#' vira 01, 02... - o primeiro
    número com View Name livre - e o mesmo número vai para os dois títulos."""
    if u"#" not in title and u"#" not in en:
        return clean_name(with_prefix(prefix, title)), title, en
    n = 1   # '#' só no inglês: usa 01 (sem laço infinito - igual ao 01)
    while u"#" in title and clean_name(with_prefix(
            prefix, title.replace(u"#", u"{:02d}".format(n)))) in VIEW_NAMES:
        n += 1
    num = u"{:02d}".format(n)
    title, en = title.replace(u"#", num), en.replace(u"#", num)
    return clean_name(with_prefix(prefix, title)), title, en


def seq_letters():
    """A, B, ..., Z, AA, AB, ..."""
    n = 0
    while True:
        n += 1
        k, out = n, u""
        while k:
            k, r = divmod(k - 1, 26)
            out = unichr(65 + r) + out
        yield out


def letter_names(prefix, title, en):
    """Chamada em elevação/corte sem complemento (igual ao 01): '<nome> - A',
    ou a próxima letra cujo View Name esteja livre."""
    for letter in seq_letters():
        t = u"{} - {}".format(title, letter)
        vname = clean_name(with_prefix(prefix, t))
        if vname not in VIEW_NAMES:
            return vname, t, (en + EN_SEP + letter) if en else en


def plan_names(opt, suffix_key, level, room, prefix, texts, simple_kind):
    """(View Name, Title on Sheet, English) de uma planta/chamada. texts =
    (manual, manual_en, complemento, complemento_en). Chamada em elevação/
    corte com o título simples e sem complemento recebe a próxima letra livre;
    a versão por ambiente já é única e fica sem letra."""
    manual, manual_en, suf, suf_en = texts
    title, en = compose_names(opt, suffix_key, level, room, manual, suf, suf_en, manual_en)
    if (simple_kind and suffix_key == FREE_SUFFIX and opt["key"] != MANUAL_KEY
            and not opt["room"] and not free_suffix(suf) and not free_suffix(suf_en)):
        return letter_names(prefix, title, en)
    return resolve_names(prefix, title, en)


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


def apply_title_on_sheet(view, title, notes):
    """Title on Sheet nativo (VIEW_DESCRIPTION): o nome sem o prefixo da pasta."""
    if not title:
        return
    p = view.get_Parameter(BIP.VIEW_DESCRIPTION)
    if p is None or p.IsReadOnly:
        notes.append(u"Parâmetro 'Title on Sheet' indisponível na vista; título não gravado.")
    else:
        p.Set(title)


# ------------------------------------------------------------------
# 6b. Pasta do Navegador de Projeto = parâmetro de projeto "Pasta"
# ------------------------------------------------------------------
# A organização do navegador do escritório agrupa as vistas pelo parâmetro
# de projeto "Pasta" (texto, instância, categoria Vistas). As pastas
# existentes são os valores já usados nas vistas; "mandar a vista para a
# pasta" = gravar o valor nesse parâmetro. Nova pasta = valor novo.
FOLDER_PARAM = u"Pasta"
_FOLDER_PREFIX = re.compile(u"^([^_]+)_")


def folder_prefix(name):
    """'D000_IMPLANTAÇÃO' -> 'D000' (texto antes do 1º '_'); sem '_' -> ''."""
    m = _FOLDER_PREFIX.match(to_unicode(name).strip())
    return m.group(1).strip() if m else u""


def name_prefix(folder):
    """Prefixo do View Name: da pasta escolhida (o View Template não interfere);
    'Não definir pasta' = nenhum."""
    return folder_prefix(folder["name"]) if folder["mode"] != "none" else u""


class ProjectFolders(object):
    """Pastas existentes: valores do parâmetro 'Pasta' nas vistas do modelo."""

    def __init__(self):
        self.param_name = FOLDER_PARAM
        self.param_id = None
        self.reason = None
        self.entries = []   # [{"label", "name"}]
        try:
            self._load()
        except Exception as exc:
            self.reason = u"Não foi possível ler o parâmetro '{}' ({}).".format(
                FOLDER_PARAM, to_unicode(exc))
            self.entries = []

    @property
    def available(self):
        return self.reason is None

    def _load(self):
        values = set()
        for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
            # folhas e tabelas têm organização própria no navegador
            if v.IsTemplate or isinstance(v, (DB.ViewSheet, DB.ViewSchedule)):
                continue
            p = v.LookupParameter(FOLDER_PARAM)
            if p is None:
                continue
            if p.StorageType != DB.StorageType.String:
                self.reason = u"O parâmetro '{}' não é de texto.".format(FOLDER_PARAM)
                return
            if self.param_id is None:
                self.param_id = p.Id
            value = to_unicode(p.AsString()).strip()
            if value:
                values.add(value)
        if self.param_id is None:
            self.reason = (u"Parâmetro de projeto '{}' não encontrado nas vistas (precisa ser "
                           u"de texto, por instância, na categoria Vistas).".format(FOLDER_PARAM))
            return
        self.entries = [{"label": v, "name": v} for v in sorted(values, key=natural_key)]

    def template_value(self, template_id):
        """Valor de 'Pasta' imposto pelo View Template; None se ele não controla a pasta."""
        if self.param_id is None or template_id is None:
            return None
        tpl = doc.GetElement(template_id)
        try:
            controlled = (set(eid_int(i) for i in tpl.GetTemplateParameterIds())
                          - set(eid_int(i) for i in tpl.GetNonControlledTemplateParameterIds()))
        except Exception:
            return None
        if eid_int(self.param_id) not in controlled:
            return None
        p = tpl.LookupParameter(FOLDER_PARAM)
        return to_unicode(p.AsString()).strip() if p is not None and p.HasValue else u""

    @staticmethod
    def apply(view, name, notes):
        if not name:
            return
        p = view.LookupParameter(FOLDER_PARAM)
        if p is None or p.IsReadOnly:
            notes.append(u"Parâmetro '{}' indisponível na vista; pasta não definida."
                         .format(FOLDER_PARAM))
        else:
            p.Set(name)

    @staticmethod
    def check(view, name, notes):
        """Depois do View Template: avisa se a vista não ficou na pasta esperada."""
        if not name:
            return
        p = view.LookupParameter(FOLDER_PARAM)
        got = to_unicode(p.AsString()).strip() if p is not None and p.HasValue else u""
        if got != name:
            notes.append(u"Pasta: ficou '{}' em vez de '{}' (controlada pelo View Template)."
                         .format(got, name))


# ------------------------------------------------------------------
# 6c. View Template
# ------------------------------------------------------------------
NO_TEMPLATE = u"Sem View Template"
VIEW_TEMPLATES = sorted((v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
                         if v.IsTemplate),
                        key=lambda v: natural_key(v.Name))


# O ViewType de um template é o da vista que o originou: um template de
# elevação criado a partir de um corte tem ViewType = Section, e o Revit
# aceita aplicá-lo a elevações. Filtrar por ViewType exato escondia esses
# templates. Plano B (modelo sem nenhuma vista do tipo para consultar):
TEMPLATE_GROUPS = ((VT.FloorPlan, VT.EngineeringPlan, VT.AreaPlan),
                   (VT.CeilingPlan,),
                   (VT.Elevation, VT.Section, VT.Detail))
_PROBES = {}


def _probe_view(vt):
    """Uma vista existente (não-template) do tipo, para perguntar ao Revit
    quais templates ela aceita (View.IsValidViewTemplate)."""
    key = int(vt)
    if key not in _PROBES:
        _PROBES[key] = next((v for v in DB.FilteredElementCollector(doc).OfClass(DB.View)
                             if not v.IsTemplate and v.ViewType == vt), None)
    return _PROBES[key]


def template_items(vt):
    """Templates que o Revit aceita para o tipo de vista criado: a mesma
    verificação nativa usada ao aplicar (IsValidViewTemplate), feita numa
    vista existente desse tipo; sem vista para consultar, o grupo do tipo."""
    if vt is None:
        return []
    probe = _probe_view(vt)
    if probe is not None:
        def ok(t):
            try:
                return probe.IsValidViewTemplate(t.Id)
            except Exception:
                return False
    else:
        group = next((grp for grp in TEMPLATE_GROUPS if vt in grp), (vt,))

        def ok(t):
            return t.ViewType in group
    return [{"label": to_unicode(t.Name), "id": t.Id} for t in VIEW_TEMPLATES if ok(t)]


def apply_template(view, template, notes):
    if not view.IsValidViewTemplate(template["id"]):
        notes.append(u"View Template '{}' não é compatível com esta vista; criada sem template."
                     .format(template["label"]))
        return
    view.ViewTemplateId = template["id"]


# ------------------------------------------------------------------
# 6d. Elevações por ambiente (ElevationMarker nativo do Revit)
# ------------------------------------------------------------------
# Igual a Vista > Elevação: UM ElevationMarker por ambiente
# (ElevationMarker.CreateElevationMarker) e as vistas criadas a partir dele
# (marker.CreateElevation, índices 0..MaximumViewCount-1). A API diz que a
# direção vem do índice mas não fixa qual; por isso a direção de cada vista
# é lida da própria vista (View.ViewDirection, como recomenda a API), em
# relação ao Norte do Projeto (+Y interno = "para cima" nas plantas).
# (chave, palavra no View Name, nº nos títulos): 01 N, 02 L, 03 S, 04 O
ELEV_DIRS = (("N", u"NORTE", 1), ("L", u"LESTE", 2), ("S", u"SUL", 3), ("O", u"OESTE", 4))
ELEV_WORD = dict((k, w) for k, w, _n in ELEV_DIRS)
ELEV_NUM = dict((k, n) for k, _w, n in ELEV_DIRS)
ELEV_TYPES = [v for v in VIEW_FAMILY_TYPES if v.ViewFamily == DB.ViewFamily.Elevation]
GRID = 16   # malha de busca de ponto livre dentro do ambiente


def _type_name(el):
    try:
        return to_unicode(DB.Element.Name.GetValue(el))
    except Exception:
        return u""


ELEV_MARK_CAT = _bic_int("OST_ElevationMarks")


def _is_elev_mark(el):
    return isinstance(el, DB.FamilySymbol) and elem_cat_int(el) == ELEV_MARK_CAT


def _tag_link(vft):
    """(parâmetro do tipo de elevação que aponta para o Elevation Tag, tipo do
    Elevation Tag, parâmetro do tag que aponta para o corpo, símbolo do corpo).
    Achado pelos valores (ElementId -> família da categoria Elevation Marks) e
    não pelo nome do parâmetro, que muda com o idioma do Revit; assim o
    Callout Tag, que o tipo de elevação também tem, não é confundido."""
    try:
        for p in vft.Parameters:
            if p.StorageType != DB.StorageType.ElementId:
                continue
            tag = doc.GetElement(p.AsElementId())
            if not isinstance(tag, DB.ElementType) or isinstance(
                    tag, (DB.ViewFamilyType, DB.FamilySymbol)):
                continue
            for q in tag.Parameters:
                if q.StorageType == DB.StorageType.ElementId and \
                        _is_elev_mark(doc.GetElement(q.AsElementId())):
                    return p, tag, q, doc.GetElement(q.AsElementId())
    except Exception:
        pass
    return None, None, None, None


def _family_label(sym):
    return u"{} : {}".format(to_unicode(sym.FamilyName), _type_name(sym))


def elev_type_items():
    """Combo 'Tipo de elevação': as famílias de marcador carregadas no modelo.
    Família já usada por um tipo de elevação -> esse tipo; família carregada
    sem tipo -> um tipo novo é criado ao aplicar (ensure_elev_type)."""
    items, used = [], set()
    for vft in ELEV_TYPES:
        sym = _tag_link(vft)[3]
        if sym is not None:
            used.add(eid_int(sym.Id))
            items.append({"label": u"{}   (tipo: {})".format(_family_label(sym), _vft_name(vft)),
                          "id": vft.Id, "family": _family_label(sym), "symbol": None})
        else:
            items.append({"label": u"{}   (marcador não identificado)".format(_vft_name(vft)),
                          "id": vft.Id, "family": u"", "symbol": None})
    if any(_tag_link(v)[2] is not None for v in ELEV_TYPES):   # há tipo para duplicar
        for sym in (DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
                    .OfCategory(DB.BuiltInCategory.OST_ElevationMarks)):
            if _is_elev_mark(sym) and eid_int(sym.Id) not in used:
                items.append({"label": u"{}   (novo tipo de elevação)".format(_family_label(sym)),
                              "id": None, "family": _family_label(sym), "symbol": sym.Id})
    return sorted(items, key=lambda i: natural_key(i["label"]))


def _duplicate(etype, base_name):
    for n in range(1, 30):
        name = base_name if n == 1 else u"{} ({})".format(base_name, n)
        try:
            return etype.Duplicate(name)
        except Exception:   # nome já usado: tenta o próximo
            continue
    raise ValueError(u"não foi possível duplicar o tipo '{}'.".format(base_name))


def ensure_elev_type(item):
    """Id do tipo de elevação a usar. Família sem tipo: duplica um tipo de
    elevação existente e o seu Elevation Tag e aponta o corpo para a família
    escolhida. Só cria tipos novos; nenhum tipo existente é alterado."""
    if item["id"] is not None:
        return item["id"]
    base = next(v for v in ELEV_TYPES if _tag_link(v)[2] is not None)
    p, tag, q, _sym = _tag_link(base)
    sym = doc.GetElement(item["symbol"])
    name = u"OCA - " + to_unicode(sym.FamilyName)
    new_tag = _duplicate(tag, name)
    new_tag.get_Parameter(q.Definition).Set(item["symbol"])
    new_vft = _duplicate(base, name)
    new_vft.get_Parameter(p.Definition).Set(new_tag.Id)
    return new_vft.Id


def elevation_names(room, key, prefix):
    """(View Name, Title on Sheet, English) de uma elevação do ambiente.
    View Name leva a direção (NORTE...); os títulos, o número (01..04).
    Inglês: mesma regra das plantas (Room Name English, senão o nome PT)."""
    number, name, name_en = room_info(room)
    num = u"{:02d}".format(ELEV_NUM[key])
    vname = with_prefix(prefix, u"{} - {} - ELEVAÇÃO {}".format(number, name, ELEV_WORD[key]))
    title = u"{} - {} - ELEVAÇÃO {}".format(number, name, num)
    en = u"{} - {} - ELEVATION {}".format(number, name_en or name, num)
    return clean_name(vname), title, en


def look_direction(view):
    """N/L/S/O para onde a elevação olha (ViewDirection aponta para o observador)."""
    d = view.ViewDirection
    x, y = -d.X, -d.Y
    if abs(x) >= abs(y):
        return "L" if x > 0 else "O"
    return "N" if y > 0 else "S"


def _area(pts):
    return 0.5 * sum(a[0] * b[1] - b[0] * a[1] for a, b in zip(pts, pts[1:] + pts[:1]))


def _centroid(pts):
    a = _area(pts)
    if abs(a) < 1e-9:
        return None
    pairs = list(zip(pts, pts[1:] + pts[:1]))
    cx = sum((p[0] + q[0]) * (p[0] * q[1] - q[0] * p[1]) for p, q in pairs) / (6.0 * a)
    cy = sum((p[1] + q[1]) * (p[0] * q[1] - q[0] * p[1]) for p, q in pairs) / (6.0 * a)
    return cx, cy


def _seg_dist(p, a, b):
    dx, dy = b[0] - a[0], b[1] - a[1]
    ll = dx * dx + dy * dy
    t = 0.0 if ll < 1e-12 else max(0.0, min(1.0, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / ll))
    return math.hypot(p[0] - a[0] - t * dx, p[1] - a[1] - t * dy)


def room_z(g):
    zs = [p.Z for p in g.points]
    return min(zs), max(zs)


def marker_origin(g):
    """Ponto DENTRO do ambiente para o marcador (Room.IsPointInRoom):
    1. centroide do contorno externo; 2. ponto de localização do ambiente;
    3. ponto de uma malha sobre o ambiente mais afastado das paredes
    (ambientes em L, U...). None se nenhum ponto estiver dentro."""
    room = g.el
    z0, z1 = room_z(g)
    zt = (z0 + z1) / 2.0

    def inside(x, y):
        return room.IsPointInRoom(DB.XYZ(x, y, zt))

    polys = []
    for loop in room.GetBoundarySegments(DB.SpatialElementBoundaryOptions()) or []:
        pts = []
        for seg in loop:
            pts.extend((p.X, p.Y) for p in seg.GetCurve().Tessellate())
        if len(pts) >= 3:
            polys.append(pts)
    candidates = []
    if polys:
        c = _centroid(max(polys, key=lambda pts: abs(_area(pts))))
        if c:
            candidates.append(c)
    loc = getattr(room, "Location", None)
    if isinstance(loc, DB.LocationPoint):
        candidates.append((loc.Point.X, loc.Point.Y))
    for x, y in candidates:
        if inside(x, y):
            return DB.XYZ(x, y, z0)

    segs = [(a, b) for pts in polys for a, b in zip(pts, pts[1:] + pts[:1])]
    if not segs:
        return None
    xs = [a[0] for a, _b in segs]
    ys = [a[1] for a, _b in segs]
    best, best_d = None, -1.0
    for i in range(1, GRID):
        for j in range(1, GRID):
            x = min(xs) + (max(xs) - min(xs)) * i / float(GRID)
            y = min(ys) + (max(ys) - min(ys)) * j / float(GRID)
            if inside(x, y):
                d = min(_seg_dist((x, y), a, b) for a, b in segs)
                if d > best_d:
                    best, best_d = (x, y), d
    return DB.XYZ(best[0], best[1], z0) if best else None


def apply_elevation_crop(view, g, margin_ft):
    """Recorte = largura do ambiente vista nessa direção x altura do ambiente
    (piso ao limite superior do Room) + margem. Só o CropBox da vista muda;
    nenhuma geometria do modelo é tocada e a profundidade fica como o Revit criou."""
    frame = Frame(view)
    uv = [frame.to2d(p) for p in g.points]
    us = [p[0] for p in uv]
    vs = [p[1] for p in uv]
    u0, u1, v0, v1 = min(us), max(us), min(vs), max(vs)
    if v1 - v0 < MIN_CROP:   # ambiente sem volume: usa a altura do Room
        v1 = v0 + max(g.el.UnboundedHeight, m_to_ft(2.5))
    m = margin_ft
    view.CropBoxActive = True
    view.CropBoxVisible = True
    _clear_shape(view.GetCropRegionShapeManager())
    _set_box(view, u0 - m, v0 - m, u1 + m, v1 + m)


def host_plan(level, extra=None):
    """ViewPlan onde o marcador é visível (pedido por CreateElevation): a vista
    ativa, se for planta desse nível; senão a planta criada agora para o
    ambiente; senão a primeira planta de piso desse nível."""
    if level is None:
        return None
    lid = eid_int(level.Id)

    def ok(v):
        gl = getattr(v, "GenLevel", None)
        return (isinstance(v, DB.ViewPlan) and not v.IsTemplate
                and v.ViewType in PLAN_VIEWTYPES and gl is not None and eid_int(gl.Id) == lid)

    for v in [active_view, extra] + HOST_PLANS:
        if v is not None and ok(v):
            return v
    return None


HOST_PLANS = sorted((v for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan)
                     if not v.IsTemplate and v.ViewType in PLAN_VIEWTYPES),
                    key=lambda v: (v.ViewType != VT.FloorPlan, natural_key(v.Name)))


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
    if active_view.ViewType in PLAN_VIEWTYPES + (VT.AreaPlan,):
        return "plan"
    # chamada numa elevação/corte gera elevação/corte (ou detalhe dele)
    return {VT.Elevation: "elevation", VT.Section: "section"}.get(
        active_view.ViewType, "vertical")


MODE_HINTS = {
    "plans": u"Uma planta nova por ambiente, no nível de cada ambiente, recortada ao "
             u"redor dele.",
    "callout": u"Uma chamada por ambiente dentro da vista ativa. Numa elevação ou corte, "
               u"escolha o mesmo tipo para abrir outra elevação/corte (ampliação) ou Detalhe.",
}


def active_view_text(callout_items):
    """Vista ativa e o que a chamada gera a partir dela (igual ao 01)."""
    kinds = []
    for it in callout_items:
        k = it["label"].split(u")")[0].replace(u"Chamada (", u"")
        if k not in kinds:
            kinds.append(k)
    lines = [u"Vista ativa: {} ({})".format(to_unicode(active_view.Name),
                                            view_type_label(active_view))]
    if kinds:
        lines.append(u"Chamada gera: {}".format(u" ou ".join(kinds)))
    return u"\n".join(lines)


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


INITIAL_MARGIN_M = 0.50   # valor inicial das margens (plantas e elevações) ao abrir


def parse_margin(text):
    t = to_unicode(text).strip().lower().replace(u",", u".")
    if t.endswith(u"m"):
        t = t[:-1].strip()
    val = float(t)   # ValueError tratado no diálogo
    if val < 0 or val > 100:
        raise ValueError(u"fora do intervalo")
    return val


def room_sort_key(g):
    """Ordem de criação/abertura: nível (de baixo para cima), número, nome."""
    lid = element_level_id(g.el)
    level = doc.GetElement(lid) if lid is not None else None
    number, name, _en = room_info(g.el)
    return (level.ProjectElevation if level is not None else 0.0,
            natural_key(number), natural_key(name))


# ------------------------------------------------------------------
# 8. Diálogo (mesmo tema dos outros botões; três colunas para caber na tela)
# ------------------------------------------------------------------
# bloco de destino (pasta) repetido para plantas e elevações: @K@ = prefixo
# dos nomes dos controles, @T@ = título do bloco
FOLDER_BLOCK = u"""
            <TextBlock Text="@T@" FontWeight="SemiBold" Margin="0,10,0,0"/>
            <StackPanel x:Name="@K@_box">
              <RadioButton x:Name="@K@_f_exist" GroupName="@K@_folder" Content="Pasta existente"/>
              <ComboBox x:Name="@K@_list" Margin="20,4,0,0"/>
              <RadioButton x:Name="@K@_f_new" GroupName="@K@_folder" Content="Nova pasta"
                           Margin="0,8,0,0"/>
              <TextBox x:Name="@K@_new" Margin="20,4,0,0"
                       ToolTip="Ex.: D003_DETALHAMENTO - o texto antes do 1º '_' vira o prefixo do View Name"/>
              <RadioButton x:Name="@K@_f_none" GroupName="@K@_folder"
                           Content="Não definir pasta (View Name sem prefixo)" Margin="0,8,0,0"/>
            </StackPanel>
            <TextBlock x:Name="@K@_lock" TextWrapping="Wrap" FontSize="11"
                       Foreground="#FFB454" Margin="0,6,0,0"/>
            <TextBlock x:Name="@K@_prefix" FontSize="11" Margin="0,6,0,0"/>
"""

CROP_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Abertura de Múltiplas Vistas" Height="Auto" Width="1240"
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
      <Setter Property="Margin" Value="0,6,0,0"/>
    </Style>
    <Style TargetType="Border">
      <Setter Property="Background" Value="#131D33"/>
      <Setter Property="BorderBrush" Value="#26405F"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12,4,12,10"/>
      <Setter Property="Margin" Value="0,12,0,0"/>
    </Style>
  </Window.Resources>
  <ScrollViewer VerticalScrollBarVisibility="Auto">
  <StackPanel Margin="18">
    <TextBlock Text="ABERTURA DE MÚLTIPLAS VISTAS" FontSize="15" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock x:Name="info" TextWrapping="Wrap" FontSize="11"
               Foreground="#7A8FA9" Margin="0,0,0,2"/>

    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- coluna 1: premissas, plantas, elevações -->
      <StackPanel Grid.Column="0">
        <Border>
          <StackPanel>
            <TextBlock Text="PREMISSAS" FontSize="12" FontWeight="SemiBold" Foreground="#65E3FF"/>
            <TextBlock x:Name="rooms_txt" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="0,0,0,2"/>
            <TextBlock x:Name="view_txt" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="0,4,0,2"/>
            <TextBlock x:Name="margin_lbl" Text="Plantas: margem ao redor de cada ambiente (m)"/>
            <TextBox x:Name="margin"/>
            <TextBlock x:Name="e_margin_lbl" Text="Elevações: margem ao redor de cada ambiente (m)"/>
            <TextBox x:Name="e_margin"/>
            <TextBlock Text="Tipo de vista" Margin="0,14,0,0"/>
            <RadioButton x:Name="rb_c_plans" GroupName="create" Content="Plantas"/>
            <RadioButton x:Name="rb_c_elev" GroupName="create" Content="Elevações"/>
            <RadioButton x:Name="rb_c_both" GroupName="create" Content="Plantas + Elevações"/>
          </StackPanel>
        </Border>

        <Border>
          <StackPanel x:Name="plan_panel">
            <TextBlock Text="PLANTAS" FontSize="12" FontWeight="SemiBold" Foreground="#65E3FF"/>
            <TextBlock Text="O que fazer (uma vista por ambiente)"/>
            <RadioButton x:Name="rb_plans" GroupName="mode"
                         Content="Criar nova vista de planta (Piso, Forro, Estrutural, Área)"/>
            <RadioButton x:Name="rb_callout" GroupName="mode"
                         Content="Criar vista de chamada de detalhe (Callout) na vista ativa"/>
            <TextBlock x:Name="mode_hint" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="20,4,0,0"/>
            <TextBlock x:Name="reasons" TextWrapping="Wrap" FontSize="11"
                       Foreground="#FFB454" Margin="0,8,0,0"/>
            <TextBlock x:Name="vtype_lbl" Text="Tipo de vista da planta"/>
            <ComboBox x:Name="vtype"/>
          </StackPanel>
        </Border>

        <Border>
          <StackPanel x:Name="elev_panel">
            <TextBlock Text="ELEVAÇÕES  ·  MARCADOR NATIVO DO REVIT" FontSize="12"
                       FontWeight="SemiBold" Foreground="#65E3FF"/>
            <TextBlock Text="Tipo de elevação (família do marcador carregada no modelo)"/>
            <ComboBox x:Name="etype"/>
            <TextBlock Text="Direções (um marcador novo por ambiente)"/>
            <CheckBox x:Name="cb_n" Content="1. Norte   →  ELEVAÇÃO 01"/>
            <CheckBox x:Name="cb_s" Content="2. Sul   →  ELEVAÇÃO 03"/>
            <CheckBox x:Name="cb_l" Content="3. Leste   →  ELEVAÇÃO 02"/>
            <CheckBox x:Name="cb_o" Content="4. Oeste   →  ELEVAÇÃO 04"/>
            <TextBlock Text="Elevações que já existirem no ambiente são ignoradas: o marcador e as vistas são criados de novo."
                       TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,8,0,0"/>
            <TextBlock x:Name="elev_reason" TextWrapping="Wrap" FontSize="11"
                       Foreground="#FFB454" Margin="0,6,0,0"/>
          </StackPanel>
        </Border>
      </StackPanel>

      <!-- coluna 2: view template + destino -->
      <StackPanel Grid.Column="2">
        <Border>
          <StackPanel>
            <TextBlock Text="MODELO DE VISTA (VIEW TEMPLATE)" FontSize="12" FontWeight="SemiBold"
                       Foreground="#65E3FF"/>
            <TextBlock x:Name="template_lbl" Text="Plantas"/>
            <ComboBox x:Name="template"/>
            <TextBlock x:Name="etemplate_lbl" Text="Elevações"/>
            <ComboBox x:Name="etemplate"/>
            <TextBlock Text="Só aparecem os templates que o Revit aceita para cada tipo de vista."
                       TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,6,0,0"/>
          </StackPanel>
        </Border>

        <Border>
          <StackPanel>
            <TextBlock Text="DESTINO DAS VISTAS  ·  NAVEGADOR DE PROJETO" FontSize="12"
                       FontWeight="SemiBold" Foreground="#65E3FF"/>
            <TextBlock x:Name="folder_src" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="0,0,0,2"/>
            <TextBlock x:Name="folder_reason" TextWrapping="Wrap" FontSize="11"
                       Foreground="#FFB454" Margin="0,0,0,2"/>
            <StackPanel x:Name="p_panel">""" + FOLDER_BLOCK.replace(u"@K@", u"p").replace(u"@T@", u"Plantas") + u"""
            </StackPanel>
            <StackPanel x:Name="e_panel">""" + FOLDER_BLOCK.replace(u"@K@", u"e").replace(u"@T@", u"Elevações") + u"""
            </StackPanel>
          </StackPanel>
        </Border>
      </StackPanel>

      <!-- coluna 3: nomenclatura + pré-visualização -->
      <StackPanel Grid.Column="4">
        <Border>
          <StackPanel x:Name="name_panel">
            <TextBlock Text="NOMENCLATURA" FontSize="12" FontWeight="SemiBold"
                       Foreground="#65E3FF"/>
            <TextBlock Text="Plantas: vale para View Name, Title on Sheet e Title on Sheet - English. Elevações: Nº - AMBIENTE - ELEVAÇÃO (automático)."
                       TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,0,0,2"/>
            <TextBlock Text="A - Principal"/>
            <ComboBox x:Name="vname"/>
            <TextBlock x:Name="manual_lbl" TextWrapping="Wrap"
                       Text="Nome manual - View Name e Title on Sheet (vazio = mantém o nome do Revit;  # = sequência 01, 02...)"/>
            <TextBox x:Name="manual"/>
            <TextBlock x:Name="manual_en_lbl" TextWrapping="Wrap"
                       Text="Nome manual - Title on Sheet - English (vazio = sem título em inglês)"/>
            <TextBox x:Name="manual_en"/>
            <TextBlock x:Name="suffix_lbl" Text="B - Complemento"/>
            <ComboBox x:Name="vsuffix"/>
            <TextBlock x:Name="suffix_text_lbl" TextWrapping="Wrap"
                       Text="Complemento manual - View Name e Title on Sheet (vazio = sem complemento;  # = sequência 01, 02...)"/>
            <TextBox x:Name="suffix_text"/>
            <TextBlock x:Name="suffix_text_en_lbl" TextWrapping="Wrap"
                       Text="Complemento manual - Title on Sheet - English (vazio = sem complemento)"/>
            <TextBox x:Name="suffix_text_en"/>
          </StackPanel>
        </Border>

        <Border>
          <StackPanel>
            <TextBlock x:Name="pv_lbl" Text="PRÉ-VISUALIZAÇÃO" FontSize="10" Foreground="#7A8FA9"
                       TextWrapping="Wrap" Margin="0,6,0,6"/>
            <StackPanel x:Name="pv_plan">
              <TextBlock Text="PLANTA  ·  VIEW NAME  (prefixo da pasta + _ + nome)" FontSize="10"
                         Foreground="#7A8FA9" Margin="0"/>
              <TextBlock x:Name="pv_name" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,6"/>
              <TextBlock Text="TITLE ON SHEET" FontSize="10" Foreground="#7A8FA9" Margin="0"/>
              <TextBlock x:Name="pv_title" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,6"/>
              <TextBlock Text="TITLE ON SHEET - ENGLISH  (automático)" FontSize="10"
                         Foreground="#7A8FA9" Margin="0"/>
              <TextBlock x:Name="pv_en" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,8"/>
            </StackPanel>
            <StackPanel x:Name="pv_elev">
              <TextBlock x:Name="pv_e_lbl" Text="ELEVAÇÃO  ·  VIEW NAME" FontSize="10"
                         Foreground="#7A8FA9" Margin="0"/>
              <TextBlock x:Name="pv_e_name" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,6"/>
              <TextBlock Text="TITLE ON SHEET" FontSize="10" Foreground="#7A8FA9" Margin="0"/>
              <TextBlock x:Name="pv_e_title" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,6"/>
              <TextBlock Text="TITLE ON SHEET - ENGLISH  (automático)" FontSize="10"
                         Foreground="#7A8FA9" Margin="0"/>
              <TextBlock x:Name="pv_e_en" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,8"/>
            </StackPanel>
            <TextBlock Text="PASTA DE DESTINO" FontSize="10" Foreground="#7A8FA9" Margin="0"/>
            <TextBlock x:Name="pv_folder" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,6"/>
            <TextBlock Text="VIEW TEMPLATE" FontSize="10" Foreground="#7A8FA9" Margin="0"/>
            <TextBlock x:Name="pv_template" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,0"/>
          </StackPanel>
        </Border>
        <TextBlock x:Name="live_hint" TextWrapping="Wrap" FontSize="11"
                   Foreground="#FFB454" Margin="0,6,0,0"/>
        <TextBlock x:Name="name_hint" TextWrapping="Wrap" FontSize="11"
                   Foreground="#FFB454" Margin="0,4,0,0"/>
      </StackPanel>
    </Grid>

    <CheckBox x:Name="open_views" Content="Abrir as vistas criadas ao final" Margin="0,12,0,0"/>

    <TextBlock x:Name="error" TextWrapping="Wrap" FontSize="11"
               Foreground="#FF4F9A" Margin="0,10,0,0"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
      <Button x:Name="cancel" Content="Cancelar" Width="80" Height="28" Margin="0,0,10,0"/>
      <Button x:Name="ok" Content="Aplicar" Width="110" Height="28"/>
    </StackPanel>
  </StackPanel>
  </ScrollViewer>
</Window>
"""

FOLDER_MODES = ("existing", "new", "none")
CREATE_MODES = ("plans", "elev", "both")
DIR_ORDER = ("N", "S", "L", "O")   # ordem dos checkboxes no diálogo


class FolderPicker(object):
    """Um bloco de destino (pasta existente / nova / nenhuma). Se o View
    Template escolhido controla o parâmetro 'Pasta', o bloco trava: a vista
    vai para a pasta do template e o prefixo vem dela."""

    def __init__(self, win, key, folders, saved, on_change):
        def ctl(name):
            return getattr(win, u"{}_{}".format(key, name))
        self.folders = folders
        self.rb = {"existing": ctl("f_exist"), "new": ctl("f_new"), "none": ctl("f_none")}
        self.box, self.list, self.new = ctl("box"), ctl("list"), ctl("new")
        self.lock_txt, self.prefix_txt = ctl("lock"), ctl("prefix")

        labels = [e["label"] for e in folders.entries]
        self.list.ItemsSource = labels
        if labels:
            self.list.SelectedIndex = (labels.index(saved["folder"])
                                       if saved["folder"] in labels else 0)
        self.new.Text = saved["folder_new"]
        allowed = {"existing": folders.available and bool(labels),
                   "new": folders.available, "none": True}
        for k, rb in self.rb.items():
            rb.IsEnabled = allowed[k]
            rb.Checked += on_change
        self.list.SelectionChanged += on_change
        self.new.TextChanged += on_change
        # sempre abre em "Pasta existente" (igual ao 01; a pasta escolhida da
        # última vez continua selecionada na lista); sem pastas, a próxima opção
        start = next(k for k in FOLDER_MODES if allowed[k])
        self.rb[start].IsChecked = True

    @property
    def mode(self):
        return next((k for k in FOLDER_MODES if self.rb[k].IsChecked), "none")

    def choice(self, template):
        """{"mode", "name", "locked"}: name = pasta onde a vista fica (define o prefixo)."""
        forced = self.folders.template_value(template["id"]) if template else None
        if forced is not None:
            return {"mode": "template", "name": forced, "locked": True}
        mode = self.mode
        name = u""
        if mode == "existing":
            idx = self.list.SelectedIndex
            entries = self.folders.entries
            name = entries[idx]["name"] if 0 <= idx < len(entries) else u""
        elif mode == "new":
            name = to_unicode(self.new.Text).strip()
        return {"mode": mode, "name": name, "locked": False}

    def render(self, template):
        c = self.choice(template)
        self.box.IsEnabled = not c["locked"]
        self.box.Opacity = 0.4 if c["locked"] else 1.0
        if not c["locked"]:
            self.list.IsEnabled = c["mode"] == "existing"
            self.new.IsEnabled = c["mode"] == "new"
        self.lock_txt.Text = (u"Travado: o View Template '{}' define a pasta ('{}' = '{}')."
                              .format(template["label"], FOLDER_PARAM, c["name"] or u"vazio → ???")
                              if c["locked"] else u"")
        prefix = name_prefix(c)
        self.prefix_txt.Text = u"Prefixo do View Name:  {}".format(prefix or u"(nenhum)")
        return c

    def saved_values(self):
        sel = self.list.SelectedItem
        return {"folder_mode": self.mode,
                "folder": to_unicode(sel) if sel is not None else u"",
                "folder_new": to_unicode(self.new.Text).strip()}


class CropWindow(forms.WPFWindow):
    def __init__(self, xaml, info, ctx, reasons, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        try:   # telas baixas: a janela rola em vez de passar da área útil
            from System.Windows import SystemParameters
            self.MaxHeight = SystemParameters.WorkArea.Height
        except Exception:
            pass
        self.confirmed = False
        self.ctx = ctx
        self._saved = cfg
        self._ready = False    # eventos disparados durante a montagem são ignorados
        self._busy = False
        self._names = []
        self._templates = []
        self._radios = {"plans": self.rb_plans, "callout": self.rb_callout}
        self._cradios = {"plans": self.rb_c_plans, "elev": self.rb_c_elev, "both": self.rb_c_both}
        self._dir_boxes = {"N": self.cb_n, "S": self.cb_s, "L": self.cb_l, "O": self.cb_o}

        self.info.Text = info
        self.rooms_txt.Text = ctx["rooms_text"]
        self.view_txt.Text = ctx["view_text"]
        self.margin.Text = u"{:.2f}".format(cfg["margin"])
        self.e_margin.Text = u"{:.2f}".format(cfg["e_margin"])
        self.reasons.Text = u"\n".join(
            u"• {}: {}".format(MODE_LABELS[k], reasons[k]) for k in MODES if k in reasons)
        self.name_hint.Text = u"\n".join(ctx["hints"])
        self.pv_lbl.Text = u"PRÉ-VISUALIZAÇÃO  ·  exemplo (1º ambiente): {}".format(ctx["example"])
        self.open_views.IsChecked = bool(cfg["open"])
        self._suffix_idx = next(
            (i for i, s in enumerate(SUFFIXES) if s[0] == cfg["suffix"]), 0)

        enabled = [k for k in MODES if k not in reasons]
        for key, rb in self._radios.items():
            rb.IsEnabled = key in enabled
            rb.Opacity = 1.0 if key in enabled else 0.4   # bloqueada: tom mais claro
            rb.Checked += self._on_mode
        self._init_elevations()
        self._init_create(bool(enabled))

        f = ctx["folders"]
        if f.available:
            self.folder_src.Text = u"Parâmetro de projeto '{}'  ·  {} pasta(s) no modelo".format(
                f.param_name, len(f.entries))
        else:
            self.folder_reason.Text = f.reason
        self.p_folder = FolderPicker(self, "p", f, cfg["p_folder"], self._on_name)
        self.e_folder = FolderPicker(self, "e", f, cfg["e_folder"], self._on_name)

        self.vtype.SelectionChanged += self._on_vtype
        self.template.SelectionChanged += self._on_name
        self.etemplate.SelectionChanged += self._on_name
        self.vname.SelectionChanged += self._on_name
        self.vsuffix.SelectionChanged += self._on_name
        self.manual.TextChanged += self._on_name
        self.manual_en.TextChanged += self._on_name
        self.suffix_text.TextChanged += self._on_name
        self.suffix_text_en.TextChanged += self._on_name

        self._ready = True
        start = cfg["mode"] if cfg["mode"] in enabled else (enabled[0] if enabled else "plans")
        self._radios[start].IsChecked = True
        self._refresh()

        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    # --- plantas / elevações / ambas ------------------------------------
    def _init_create(self, plans_ok):
        elev_ok = bool(self.ctx["elev_types"])
        allowed = {"plans": plans_ok, "elev": elev_ok, "both": plans_ok and elev_ok}
        for key, rb in self._cradios.items():
            rb.IsEnabled = allowed[key]
            rb.Opacity = 1.0 if allowed[key] else 0.4   # bloqueada: tom mais claro
            rb.Checked += self._on_name
        start = self._saved["create"]
        if not allowed.get(start):
            start = next(k for k in CREATE_MODES if allowed[k])
        self._cradios[start].IsChecked = True

    @property
    def create(self):
        return next((k for k in CREATE_MODES if self._cradios[k].IsChecked), "plans")

    @property
    def do_plans(self):
        return self.create in ("plans", "both")

    @property
    def do_elev(self):
        return self.create in ("elev", "both")

    def _init_elevations(self):
        items = self.ctx["elev_types"]
        labels = [i["label"] for i in items]
        self.etype.ItemsSource = labels
        if labels:
            saved = self._saved["elev_type"]
            # padrão: a família OCA de elevação interna, se estiver carregada
            oca = next((i for i, it in enumerate(items)
                        if u"OCA_SIMB_ELEVA" in it["family"].upper()), 0)
            self.etype.SelectedIndex = labels.index(saved) if saved in labels else oca
        else:
            self.elev_reason.Text = (u"Nenhum tipo de vista de Elevação no modelo: crie um em "
                                     u"Vista > Elevação > Editar tipo.")
        for key, box in self._dir_boxes.items():
            box.IsChecked = key in self._saved["elev_dirs"]
            box.Checked += self._on_name
            box.Unchecked += self._on_name
        if not self.ctx["elev_templates"]:
            self.etemplate_lbl.Text = u"Elevações  (nenhum template de elevação/corte no modelo)"
        etpl = [NO_TEMPLATE] + [t["label"] for t in self.ctx["elev_templates"]]
        self.etemplate.ItemsSource = etpl
        saved_t = self._saved["etemplate"]
        self.etemplate.SelectedIndex = etpl.index(saved_t) if saved_t in etpl else 0

    @property
    def elev_type(self):
        idx = self.etype.SelectedIndex
        items = self.ctx["elev_types"]
        return items[idx] if 0 <= idx < len(items) else None

    @property
    def elev_dirs(self):
        return [k for k in DIR_ORDER if self._dir_boxes[k].IsChecked]

    @property
    def elev_template(self):
        idx = self.etemplate.SelectedIndex
        items = self.ctx["elev_templates"]
        return items[idx - 1] if 1 <= idx <= len(items) else None

    # --- tipo de vista da planta + template ------------------------------
    @property
    def mode(self):
        for key in MODES:
            if self._radios[key].IsChecked:
                return key
        return None

    @property
    def vtype_item(self):
        types = self.ctx["types"].get(self.mode, [])
        sel = self.vtype.SelectedItem
        return next((i for i in types if i["label"] == sel), None)

    def _refresh_templates(self):
        current = self.template.SelectedItem
        wanted = to_unicode(current) if current is not None else self._saved["template"]
        item = self.vtype_item
        self._templates = template_items(item["vt"] if item else None)
        labels = [NO_TEMPLATE] + [t["label"] for t in self._templates]
        self.template.ItemsSource = labels
        self.template.SelectedIndex = labels.index(wanted) if wanted in labels else 0

    @property
    def template_choice(self):
        idx = self.template.SelectedIndex
        return self._templates[idx - 1] if 1 <= idx <= len(self._templates) else None

    def _on_vtype(self, sender, args):
        if self._ready and self._names:
            self._refresh_templates()
            self._update_names()

    def _on_mode(self, sender, args):
        if self._ready:
            self._refresh()

    def _on_name(self, sender, args):
        if self._ready:
            self._update_names()

    def _refresh(self):
        mode = self.mode
        types = self.ctx["types"].get(mode, [])
        labels = [i["label"] for i in types]
        self._names = []   # evita _on_vtype com os nomes do modo anterior
        self.vtype.ItemsSource = labels
        if labels:
            self.vtype.SelectedItem = (self._saved["vtype"] if self._saved["vtype"] in labels
                                       else labels[0])
        self.vtype.IsEnabled = bool(labels)
        self.vtype_lbl.Opacity = 1.0 if labels else 0.4
        self.mode_hint.Text = MODE_HINTS.get(mode, u"")

        # nomes dependem do tipo da vista resultante (planta x corte/elevação)
        self._names = self.ctx["names"][mode]
        keys = [o["key"] for o in self._names]
        self.vname.ItemsSource = [o["label"] for o in self._names]
        if self._saved["vname"] in keys:
            self.vname.SelectedIndex = keys.index(self._saved["vname"])
        else:   # vários ambientes: por padrão, a 1ª opção "Nº - AMBIENTE - ..."
            self.vname.SelectedIndex = next(
                (i for i, o in enumerate(self._names) if o["room"]), 0)

        # complemento traz o nome do nível da vista resultante; com eventos
        # travados, senão a seleção vazia (-1) da troca de itens vira a escolha
        if self.vsuffix.SelectedIndex >= 0:
            self._suffix_idx = self.vsuffix.SelectedIndex
        self._busy = True
        try:
            self.vsuffix.ItemsSource = suffix_labels(self.ctx["level"][mode])
            # chamada em elevação/corte abre em "Sem complemento" (letra A, B...)
            if self.simple_kind:
                self._suffix_idx = 0
            self.vsuffix.SelectedIndex = self._suffix_idx
        finally:
            self._busy = False
        self._refresh_templates()
        self._update_names()

    # --- nomenclatura --------------------------------------------------
    @property
    def simple_kind(self):
        """Chamada numa elevação/corte: nome simples; sem complemento = letra."""
        return self.ctx["kinds"].get(self.mode) in SIMPLE_KINDS

    @property
    def name_opt(self):
        idx = self.vname.SelectedIndex
        if 0 <= idx < len(self._names):
            return self._names[idx]
        return next((o for o in self._names if o["key"] == MANUAL_KEY), self._names[-1])

    def name_texts(self):
        """(manual, manual_en, complemento, complemento_en) digitados; os
        complementos só valem com o combo em 'Sem complemento, manual ou sequência'."""
        on = self.suffix_key == FREE_SUFFIX
        return (to_unicode(self.manual.Text), to_unicode(self.manual_en.Text),
                to_unicode(self.suffix_text.Text) if on else u"",
                to_unicode(self.suffix_text_en.Text) if on else u"")

    @property
    def suffix_key(self):
        idx = self.vsuffix.SelectedIndex
        return SUFFIXES[idx if idx >= 0 else 0][0]

    def names(self, folder):
        """Prévia com o 1º ambiente: (View Name, Title on Sheet, English)."""
        return plan_names(self.name_opt, self.suffix_key, self.ctx["level"][self.mode],
                          self.ctx["room"], name_prefix(folder), self.name_texts(),
                          self.simple_kind)

    def _sync_suffix(self, opt):
        """Todas as opções do complemento valem para qualquer A - Principal
        (igual ao 01): só guarda a escolha atual - o combo não trava mais."""
        if self.vsuffix.SelectedIndex < 0 and self.vsuffix.Items.Count:
            self.vsuffix.SelectedIndex = self._suffix_idx
        if self.vsuffix.SelectedIndex >= 0:
            self._suffix_idx = self.vsuffix.SelectedIndex

    def _update_names(self):
        if self._busy or not self._names:   # mudar o combo aqui dispara SelectionChanged
            return
        self._busy = True
        try:
            self._render()
        finally:
            self._busy = False

    def _render(self):
        plans, elev = self.do_plans, self.do_elev
        # grupos que não se aplicam ao tipo escolhido ficam apagados
        for panel, on in ((self.plan_panel, plans), (self.name_panel, plans),
                          (self.pv_plan, plans), (self.elev_panel, elev), (self.pv_elev, elev),
                          (self.margin, plans), (self.margin_lbl, plans),
                          (self.e_margin, elev), (self.e_margin_lbl, elev),
                          (self.template, plans), (self.template_lbl, plans),
                          (self.etemplate, elev), (self.etemplate_lbl, elev),
                          (self.p_panel, plans), (self.e_panel, elev)):
            panel.IsEnabled = on
            panel.Opacity = 1.0 if on else 0.4

        opt = self.name_opt
        is_manual = opt["key"] == MANUAL_KEY
        self._sync_suffix(opt)
        free = self.suffix_key == FREE_SUFFIX
        for ctl in (self.suffix_text, self.suffix_text_en):
            ctl.IsEnabled = free
        for lbl in (self.suffix_text_lbl, self.suffix_text_en_lbl):
            lbl.Opacity = 1.0 if free else 0.4
        for ctl in (self.manual, self.manual_en):
            ctl.IsEnabled = is_manual
        for lbl in (self.manual_lbl, self.manual_en_lbl):
            lbl.Opacity = 1.0 if is_manual else 0.4

        p_tpl, e_tpl = self.template_choice, self.elev_template
        p_fold = self.p_folder.render(p_tpl)
        e_fold = self.e_folder.render(e_tpl)

        vname, title, en = self.names(p_fold)
        self.pv_name.Text = vname or u"(nome padrão do Revit)"
        self.pv_title.Text = title or u"—"
        self.pv_en.Text = en or u"—"

        dirs = sorted(self.elev_dirs, key=lambda k: ELEV_NUM[k])
        if dirs:
            e_name, e_title, e_en = elevation_names(self.ctx["room"], dirs[0], name_prefix(e_fold))
            others = [u"{} ({:02d})".format(ELEV_WORD[k], ELEV_NUM[k]) for k in dirs[1:]]
            self.pv_e_lbl.Text = (u"ELEVAÇÃO {}  ·  VIEW NAME".format(ELEV_WORD[dirs[0]])
                                  + (u"   (+ {})".format(u", ".join(others)) if others else u""))
        else:
            e_name = e_title = e_en = u"—"
            self.pv_e_lbl.Text = u"ELEVAÇÃO  ·  VIEW NAME"
        self.pv_e_name.Text, self.pv_e_title.Text, self.pv_e_en.Text = e_name, e_title, e_en

        folders, tpls = [], []
        for on, label, fold, tpl in ((plans, u"Plantas", p_fold, p_tpl),
                                     (elev, u"Elevações", e_fold, e_tpl)):
            if on:
                folders.append(u"{}: {}".format(label, fold["name"] or (
                    u"??? (template)" if fold["locked"] else u"(não definida pelo comando)")))
                tpls.append(u"{}: {}".format(label, tpl["label"] if tpl else NO_TEMPLATE))
        self.pv_folder.Text = u"\n".join(folders)
        self.pv_template.Text = u"\n".join(tpls)
        self.live_hint.Text = u"\n".join(self._live_hints(p_fold, e_fold))

    def _live_hints(self, p_fold, e_fold):
        hints = []
        entries = [e["name"] for e in self.ctx["folders"].entries]
        for on, label, fold in ((self.do_plans, u"Plantas", p_fold),
                                (self.do_elev, u"Elevações", e_fold)):
            if not on or not fold["name"]:
                continue
            if fold["mode"] != "none" and not folder_prefix(fold["name"]):
                hints.append(u"{}: a pasta '{}' não tem '_' - View Name sem prefixo."
                             .format(label, fold["name"]))
            if fold["mode"] == "new" and fold["name"] in entries:
                hints.append(u"{}: a pasta '{}' já existe - as vistas vão para ela."
                             .format(label, fold["name"]))
        opt = self.name_opt
        letters = self.simple_kind and opt["key"] != MANUAL_KEY   # recebem A, B...
        if self.do_plans and self.ctx["count"] > 1 and not opt["room"] and not letters:
            hints.append(u"Nomes sem <Nº> - <AMBIENTE> se repetem entre as vistas e recebem "
                         u"(2), (3)...; no nome manual, use # para numerar.")
        return hints

    def _ok(self, sender, args):
        try:
            self.margin_m = parse_margin(self.margin.Text)
            self.e_margin_m = parse_margin(self.e_margin.Text)
        except Exception:
            self.error.Text = u"Margem inválida: informe um número em metros entre 0 e 100 (ex.: 0,50)."
            return
        for on, label, picker, tpl in ((self.do_plans, u"plantas", self.p_folder, self.template_choice),
                                       (self.do_elev, u"elevações", self.e_folder, self.elev_template)):
            c = picker.choice(tpl)
            if on and c["mode"] == "new" and not c["name"]:
                self.error.Text = u"Informe o nome da nova pasta das {} (ex.: D003_DETALHAMENTO).".format(label)
                return
            if on and c["mode"] == "existing" and not c["name"]:
                self.error.Text = u"Selecione a pasta de destino das {}.".format(label)
                return
        if self.do_plans and self.vtype_item is None:
            self.error.Text = u"Escolha o tipo de vista da planta."
            return
        if self.do_elev and self.elev_type is None:
            self.error.Text = u"Escolha o tipo de elevação (família do marcador)."
            return
        if self.do_elev and not self.elev_dirs:
            self.error.Text = u"Marque pelo menos uma direção de elevação."
            return
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.Close()

    @property
    def result(self):
        return {"mode": self.mode, "margin": self.margin_m, "e_margin": self.e_margin_m,
                "create": self.create,
                "type": self.vtype_item,
                "vname": self.name_opt["key"], "suffix": self.suffix_key,
                "name_opt": self.name_opt, "texts": self.name_texts(),
                "simple_kind": self.simple_kind,
                "open": bool(self.open_views.IsChecked),
                "folder": self.p_folder.choice(self.template_choice),
                "e_folder": self.e_folder.choice(self.elev_template),
                "p_saved": self.p_folder.saved_values(),
                "e_saved": self.e_folder.saved_values(),
                "template": self.template_choice,
                "elev_type": self.elev_type,
                "elev_dirs": self.elev_dirs,
                "elev_template": self.elev_template}


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
elev_items = elev_type_items()
if len(reasons) == len(MODES) and not elev_items:
    forms.alert(u"Nenhuma ação disponível:\n\n" + u"\n".join(reasons.values())
                + u"\nNenhum tipo de vista de Elevação no modelo.", exitscript=True)

folders = ProjectFolders()

name_hints = []
level_ids = set(eid_int(l.Id) for l in room_levels if l is not None)
if len(level_ids) > 1:
    name_hints.append(u"Ambientes em {} níveis: cada planta é criada no nível do seu ambiente "
                      u"e o complemento usa esse nível.".format(len(level_ids)))
no_en = [g for g in geoms if not room_info(g.el)[2]]
if no_en:
    name_hints.append(u"'{}' vazio em {} ambiente(s): o título em inglês usará o nome em português."
                      .format(ROOM_EN_PARAM, len(no_en)))
if not en_param_bound():
    name_hints.append(u"'{}' não está vinculado à categoria Vistas: o título em inglês "
                      u"não será gravado.".format(EN_TITLE_PARAM))
if folders.available and not folders.entries:
    name_hints.append(u"Nenhuma vista tem o parâmetro '{}' preenchido ainda: use 'Nova pasta'."
                      .format(folders.param_name))

active_level = getattr(active_view, "GenLevel", None)
rooms_preview = u", ".join(g.label for g in geoms[:6]) + (u"..." if len(geoms) > 6 else u"")
dialog_ctx = {
    "types": {"plans": plan_items, "callout": callout_items},
    "names": dict((m, name_options(target_kind(m))) for m in MODES),
    "level": {"plans": room_levels[0], "callout": active_level},
    "room": geoms[0].el,
    "example": geoms[0].label,
    "count": len(geoms),
    "kinds": dict((m, target_kind(m)) for m in MODES),
    "view_text": active_view_text(callout_items),
    "rooms_text": u"{} ambiente(s) selecionado(s): {}".format(len(geoms), rooms_preview),
    "hints": name_hints,
    "folders": folders,
    "elev_types": elev_items,
    "elev_templates": template_items(VT.Elevation),
}

config = script.get_config()
# margens sempre abrem com o valor inicial (o usuário altera no diálogo)
cfg = {"margin": INITIAL_MARGIN_M,
       "e_margin": INITIAL_MARGIN_M,
       "mode": config.get_option("mode", "plans"),
       "create": config.get_option("create", "plans"),
       "vtype": config.get_option("vtype", u""),
       "vname": config.get_option("vname", "room_civil"),
       "suffix": config.get_option("vsuffix", "amp"),
       "open": config.get_option("open_views", True),
       "p_folder": {"folder_mode": config.get_option("folder_mode", "existing"),
                    "folder": config.get_option("folder", u""),
                    "folder_new": config.get_option("folder_new", u"")},
       "e_folder": {"folder_mode": config.get_option("e_folder_mode", "existing"),
                    "folder": config.get_option("e_folder", u""),
                    "folder_new": config.get_option("e_folder_new", u"")},
       "template": config.get_option("template", NO_TEMPLATE),
       "elev_type": config.get_option("elev_type", u""),
       "elev_dirs": config.get_option("elev_dirs", u"NLSO"),
       "etemplate": config.get_option("etemplate", NO_TEMPLATE)}

info = u"{} ambiente(s) → uma vista por ambiente".format(len(geoms))
win = CropWindow(CROP_XAML, info, dialog_ctx, reasons, cfg)
win.ShowDialog()
if not win.confirmed:
    script.exit()
opts = win.result

config.mode = opts["mode"]
config.create = opts["create"]
config.vtype = opts["type"]["label"] if opts["type"] else cfg["vtype"]
config.vname = opts["vname"]
config.vsuffix = opts["suffix"]
config.open_views = opts["open"]
config.folder_mode = opts["p_saved"]["folder_mode"]
config.folder = opts["p_saved"]["folder"]
config.folder_new = opts["p_saved"]["folder_new"]
config.e_folder_mode = opts["e_saved"]["folder_mode"]
config.e_folder = opts["e_saved"]["folder"]
config.e_folder_new = opts["e_saved"]["folder_new"]
config.template = opts["template"]["label"] if opts["template"] else NO_TEMPLATE
config.elev_type = opts["elev_type"]["label"] if opts["elev_type"] else cfg["elev_type"]
config.elev_dirs = u"".join(opts["elev_dirs"])
config.etemplate = opts["elev_template"]["label"] if opts["elev_template"] else NO_TEMPLATE
script.save_config()

mode = opts["mode"]
do_plans = opts["create"] in ("plans", "both")
do_elev = opts["create"] in ("elev", "both")
margin_ft = m_to_ft(opts["margin"])        # plantas
e_margin_ft = m_to_ft(opts["e_margin"])    # elevações

# --- verificações específicas do modo, antes de qualquer alteração ---------
if do_plans and mode == "callout":
    visible = set(eid_int(i) for i in
                  DB.FilteredElementCollector(doc, active_view.Id)
                  .WhereElementIsNotElementType().ToElementIds())
    for g in geoms:
        if eid_int(g.el.Id) not in visible:
            warnings.append(u"{}: não está visível na vista ativa (recorte calculado mesmo assim).".format(g.label))


vtype = opts["type"]
name_opt = opts["name_opt"]
folder = opts["folder"]          # plantas
prefix = name_prefix(folder)
e_folder = opts["e_folder"]      # elevações
e_prefix = name_prefix(e_folder)


def folder_to_write(fold):
    """Pasta gravada pelo comando; travada = o View Template já define."""
    return u"" if fold["locked"] else fold["name"]
template = opts["template"]
elev_type = opts["elev_type"]
elev_dirs = opts["elev_dirs"]
elev_template = opts["elev_template"]
notes = []
created = []   # (ElemGeom, ElementId da vista), na ordem dos ambientes
failed = []    # (rótulo, motivo)


def create_room_view(g, level):
    """Mesma sequência do 01 para UM ambiente + pasta e template. Devolve (vista, avisos)."""
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
    # (o template escolhido no diálogo entra no fim, depois do recorte)
    view.ViewTemplateId = INVALID_ID

    vname, title, en = plan_names(name_opt, opts["suffix"], name_level, g.el, prefix,
                                  opts["texts"], opts["simple_kind"])
    apply_names(view, vname, en, room_notes)      # View Name (com prefixo) + English
    apply_title_on_sheet(view, title, room_notes)  # Title on Sheet (sem prefixo)
    ProjectFolders.apply(view, folder_to_write(folder), room_notes)
    doc.Regenerate()

    # recalcula no sistema da vista-alvo (igual ao 01)
    calc = CropCalc([g], Frame(view), margin_ft)
    _method, warn = apply_crop(view, calc)
    if warn:
        room_notes.append(warn)

    if template is not None:
        apply_template(view, template, room_notes)
        doc.Regenerate()
        ProjectFolders.check(view, folder["name"], room_notes)
    return view, room_notes


def setup_elevation(view, g, key, notes_out):
    """Nome, pasta, recorte e template de UMA elevação (mesmas regras das plantas)."""
    view.ViewTemplateId = INVALID_ID   # tipo de elevação pode trazer template padrão
    vname, title, en = elevation_names(g.el, key, e_prefix)
    apply_names(view, vname, en, notes_out)
    apply_title_on_sheet(view, title, notes_out)
    ProjectFolders.apply(view, folder_to_write(e_folder), notes_out)
    doc.Regenerate()
    apply_elevation_crop(view, g, e_margin_ft)
    if elev_template is not None:
        apply_template(view, elev_template, notes_out)
        doc.Regenerate()
        ProjectFolders.check(view, e_folder["name"], notes_out)


def create_room_elevations(g, host, etype_id):
    """UM marcador NOVO por ambiente, no ponto seguro dentro do Room. Elevações
    que já existirem no ambiente são ignoradas (não é erro): tudo é criado de
    novo. Cada direção em uma SubTransaction: uma falha não desfaz as outras.
    Devolve ([(direção, vista)], [(direção, motivo)], avisos)."""
    made, room_notes, errors_by_key = [], [], {}
    origin = marker_origin(g)
    if origin is None:
        raise ValueError(u"nenhum ponto dentro do ambiente para colocar o marcador.")
    marker = DB.ElevationMarker.CreateElevationMarker(doc, etype_id, origin, host.Scale)

    last_error = u"o marcador não tem posição livre para essa direção."
    for i in range(marker.MaximumViewCount):
        if all(k in dict(made) for k in elev_dirs):
            break
        if not marker.IsAvailableIndex(i):
            continue
        names_before = set(VIEW_NAMES)
        key = None
        st = DB.SubTransaction(doc)
        st.Start()
        try:
            view = marker.CreateElevation(doc, host.Id, i)
            doc.Regenerate()
            key = look_direction(view)
            if key not in elev_dirs or key in dict(made):
                st.RollBack()   # direção não marcada no diálogo: desfaz
                continue
            view_notes = []
            setup_elevation(view, g, key, view_notes)
            st.Commit()
            made.append((key, view))
            room_notes.extend(u"ELEVAÇÃO {}: {}".format(ELEV_WORD[key], n) for n in view_notes)
        except Exception as exc:
            if st.GetStatus() == DB.TransactionStatus.Started:
                st.RollBack()
            VIEW_NAMES.clear()
            VIEW_NAMES.update(names_before)
            if key:
                errors_by_key[key] = to_unicode(exc)
            else:
                last_error = to_unicode(exc)

    done = set(k for k, _v in made)
    errors = [(k, errors_by_key.get(k, last_error)) for k in elev_dirs if k not in done]
    if not made:
        doc.Delete(marker.Id)   # não deixa marcador vazio no modelo
    made.sort(key=lambda kv: ELEV_NUM[kv[0]])
    return made, errors, room_notes


def elev_label(g, key):
    return u"{} · ELEVAÇÃO {}".format(g.label, ELEV_WORD[key])


# um grupo = um único Desfazer; uma transação por ambiente (planta e
# elevações separadas) = falha isolada
tgroup = DB.TransactionGroup(doc, u"OCA - Abertura de Múltiplas Vistas")
tgroup.Start()
try:
    # família do marcador sem tipo de elevação: cria o tipo uma vez, antes
    etype_id, etype_error = None, None
    if do_elev:
        t = DB.Transaction(doc, u"OCA - Tipo de elevação")
        t.Start()
        try:
            etype_id = ensure_elev_type(elev_type)
            if t.Commit() != DB.TransactionStatus.Committed:
                raise Exception(u"o Revit desfez a criação do tipo de elevação.")
            if elev_type["id"] is None:
                notes.append(u"Criado o tipo de elevação '{}' para a família {}.".format(
                    _type_name(doc.GetElement(etype_id)), elev_type["family"]))
        except Exception as exc:
            if t.GetStatus() == DB.TransactionStatus.Started:
                t.RollBack()
            etype_id, etype_error = None, u"tipo de elevação: {}".format(to_unicode(exc))

    for g, level in zip(geoms, room_levels):
        plan_view = None
        if do_plans:
            names_before = set(VIEW_NAMES)
            t = DB.Transaction(doc, u"OCA - Vista {}".format(g.label))
            t.Start()
            try:
                view, room_notes = create_room_view(g, level)
                status = t.Commit()
                if status != DB.TransactionStatus.Committed:
                    raise Exception(u"o Revit desfez a criação da vista ({}).".format(status))
                plan_view = view
                created.append((g, view.Id))
                notes.extend(u"{}: {}".format(g.label, n) for n in room_notes)
            except Exception as exc:
                if t.GetStatus() == DB.TransactionStatus.Started:
                    t.RollBack()
                VIEW_NAMES.clear()   # libera nomes reservados pela vista desfeita
                VIEW_NAMES.update(names_before)
                failed.append((g.label, to_unicode(exc)))

        if do_elev and etype_error:
            failed.extend((elev_label(g, k), etype_error) for k in elev_dirs)
        elif do_elev:
            host = host_plan(level, plan_view)
            if host is None:
                why = (u"nenhuma planta do nível '{}' para colocar o marcador (abra uma planta "
                       u"desse nível).".format(to_unicode(level.Name)) if level is not None
                       else u"nível do ambiente não identificado.")
                failed.extend((elev_label(g, k), why) for k in elev_dirs)
                continue
            names_before = set(VIEW_NAMES)
            t = DB.Transaction(doc, u"OCA - Elevações {}".format(g.label))
            t.Start()
            try:
                made, errors, room_notes = create_room_elevations(g, host, etype_id)
                status = t.Commit()
                if status != DB.TransactionStatus.Committed:
                    raise Exception(u"o Revit desfez a criação das elevações ({}).".format(status))
                created.extend((g, v.Id) for _k, v in made)
                failed.extend((elev_label(g, k), why) for k, why in errors)
                notes.extend(u"{}: {}".format(g.label, n) for n in room_notes)
            except Exception as exc:
                if t.GetStatus() == DB.TransactionStatus.Started:
                    t.RollBack()
                VIEW_NAMES.clear()
                VIEW_NAMES.update(names_before)
                failed.extend((elev_label(g, k), to_unicode(exc)) for k in elev_dirs)
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
    report.append(u"Não foi possível criar (ambiente · vista: motivo):")
    report.extend(skipped_lines(failed))
if warnings or notes:
    report.append(u"Avisos:")
    report.extend(u"- " + w for w in warnings + notes)

if not created:
    forms.alert(u"Nenhuma vista criada.", sub_msg=u"\n".join(report) or None)
elif report:
    forms.alert(u"{} vista(s) criada(s) para {} ambiente(s), com avisos:".format(
                    len(created), len(geoms)),
                sub_msg=u"\n".join(report))
