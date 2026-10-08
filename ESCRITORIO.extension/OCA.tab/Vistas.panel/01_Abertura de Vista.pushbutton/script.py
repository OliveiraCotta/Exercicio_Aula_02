# -*- coding: utf-8 -*-
# Abertura de Vistas - ajusta o recorte (CropView) de vistas a partir de elementos.
#
# Fluxo (UMA vista para toda a seleção; para uma vista por ambiente, use o 02):
#   1. Usa a seleção atual (ou pede para selecionar) - qualquer elemento do
#      modelo ou da vista: paredes, pisos, ambientes, colunas, property lines,
#      linhas, terreno, famílias, grupos, anotações... Elementos sem geometria
#      legível entram pela caixa envolvente (bounding box).
#   2. Pergunta a margem (m) e o que fazer:
#        - criar nova vista de planta (Piso, Forro, Estrutural, Área);
#        - criar vista de chamada de detalhe (Callout) na vista ativa - numa
#          elevação ou corte, do mesmo tipo (outra elevação/corte, para
#          ampliações) ou Detalhe;
#        - recortar a vista ativa (planta, corte ou elevação).
#      Antes de abrir o diálogo, cada modo é validado contra a vista ativa e o
#      modelo; modos inválidos ficam desabilitados com o motivo exibido.
#   3. Mesma interface e regras do 02_Abertura de Múltiplas Vistas:
#        Title on Sheet           = nome (Principal + Complemento)
#        View Name                = PREFIXO_ + nome (prefixo = texto antes do
#                                   1º "_" da pasta de destino)
#        Title on Sheet - English = nome em inglês (sem prefixo)
#      Complemento manual: um campo para View Name/Title on Sheet e outro para
#      o Title on Sheet - English (cada um pode ficar vazio). Em elevação e
#      corte o nome é só ELEVAÇÃO/CORTE (ou a versão por ambiente); com os
#      dois complementos vazios entra uma letra sequencial livre (CORTE - A,
#      CORTE - B...). "Manter nomenclatura atual da vista": a vista ativa não é
#      renomeada; a chamada recebe os nomes da vista ativa + " - CHAMADA n"
#      (EN: " - CALLOUT n", n = 1º livre). Em corte/elevação não há planta nova.
#      Destino no Navegador de Projeto = parâmetro de projeto "Pasta" (pasta
#      existente, nova ou nenhuma). View Template opcional, só os compatíveis;
#      na vista ativa a 1ª opção mantém o template que ela já tem.
#   4. Calcula o retângulo ORIENTADO ao elemento (as 4 quinas, paralelo e
#      perpendicular à direção predominante do elemento) + margem. A vista
#      nunca é girada: o recorte é sempre aplicado como forma de 4 lados
#      (crop region shape), inclinada quando o elemento está girado.
#   5. Opcionalmente abre a vista criada; ao final, avisa apenas se houve
#      algo a revisar. Se algo falhar, nada é alterado no modelo.
#
# A descrição fica em comentário (e não em docstring) de propósito: o
# pyRevit 5.2 lê a docstring do módulo com .decode('utf-8'), que falha com
# acentos no IronPython e deixa o botão sem tooltip. O tooltip é só __doc__.

__title__ = u"Abertura\nde Vista"
__doc__ = u"""Cria UMA vista recortada a partir dos elementos selecionados: nova planta, chamada (callout) ou recorte da vista ativa. Já sai com nome, pasta no Navegador de Projeto e View Template.

Selecione um ou mais elementos e clique no botão."""

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
LOCATION_WEIGHT = 10.0                    # peso extra do eixo de elementos lineares retos


# ------------------------------------------------------------------
# 1. Seleção: qualquer elemento com categoria (exceto vistas e tipos)
# ------------------------------------------------------------------
def is_target(el):
    try:
        if el is None or el.Category is None:
            return False
    except Exception:
        return False
    return not isinstance(el, (DB.View, DB.Viewport, DB.ElementType))


class TargetFilter(ISelectionFilter):
    def AllowElement(self, el):
        return is_target(el)

    def AllowReference(self, ref, point):
        return False


def get_target_elements():
    """Seleção atual filtrada; se vazia, pede para o usuário selecionar."""
    current = list(revit.get_selection().elements)
    picked = [el for el in current if is_target(el)]
    ignored = len(current) - len(picked)
    if picked:
        return picked, ignored
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, TargetFilter(),
            u"Selecione os elementos da vista e clique em Concluir")
    except Exception:
        script.exit()   # Esc
    return [doc.GetElement(r) for r in refs], ignored


def elem_label(el):
    if isinstance(el, Room):
        num = el.get_Parameter(BIP.ROOM_NUMBER)
        name = el.get_Parameter(BIP.ROOM_NAME)
        parts = [to_unicode(p.AsString()) for p in (num, name)
                 if p is not None and p.AsString()]
        return u" ".join(parts) or u"Ambiente [{}]".format(eid_int(el.Id))
    kind = to_unicode(el.Category.Name) if el.Category is not None else u"Elemento"
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


def owner_view(el):
    """Vista dona de elementos específicos de vista (detalhes, anotações)."""
    try:
        if el.ViewSpecific:
            return doc.GetElement(el.OwnerViewId)
    except Exception:
        pass
    return None


def geom_options(el):
    view = owner_view(el)
    if view is None:
        return GEOM_OPT
    opt = DB.Options()   # com View definida o DetailLevel vem da própria vista
    opt.View = view
    opt.ComputeReferences = False
    return opt


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
        """Fallback: 8 cantos da caixa envolvente (orientação pelo casco convexo)."""
        try:
            bb = self.el.get_BoundingBox(owner_view(self.el))
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


# vínculos e importações: a geometria inteira seria pesada demais - só a caixa
BBOX_ONLY = (DB.RevitLinkInstance, DB.ImportInstance)


def extract_geometry(el):
    g = ElemGeom(el)
    try:
        if isinstance(el, BBOX_ONLY):
            pass
        elif isinstance(el, DB.SpatialElement):
            loops = el.GetBoundarySegments(DB.SpatialElementBoundaryOptions()) or []
            for loop in loops:
                for seg in loop:
                    g.add_curve(seg.GetCurve())
            if isinstance(el, Room):
                g.walk(el.ClosedShell)   # altura do ambiente (cortes/elevações)
        elif isinstance(el, DB.CurveElement):
            g.add_curve(el.GeometryCurve)
        elif isinstance(el, DB.Grid):
            g.add_curve(el.Curve)
        elif isinstance(el, DB.ReferencePlane):
            g.add_polyline([el.BubbleEnd, el.FreeEnd])
        else:
            g.walk(el.get_Geometry(geom_options(el)))
            loc = getattr(el, "Location", None)
            if isinstance(loc, DB.LocationCurve) and isinstance(loc.Curve, DB.Line):
                # eixo de paredes/vigas/tubos retos manda na orientação
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
    """Tipos aceitos numa chamada: o mesmo tipo da vista-mãe (planta -> planta,
    corte -> corte, elevação -> elevação, para ampliações) e Detalhe."""
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
#    + View Name (= PREFIXO DA PASTA + "_" + Title on Sheet) - igual ao 02
# ------------------------------------------------------------------
EN_TITLE_PARAM = u"Title on Sheet - English"
ROOM_EN_PARAM = u"Room Name English"
EN_SEP = u" - "   # separador do complemento no título em inglês (hífen, como no PT)

# (chave, nome PT, Title on Sheet - English, leva complemento PAV./AMPLIAÇÃO)
PLAN_TITLES = (
    ("cobertura", u"PLANTA DE COBERTURA", u"ROOF PLAN", False),
    ("civil", u"PLANTA CIVIL", u"FLOOR PLAN", True),
    ("layout", u"PLANTA DE LAYOUT", u"LAYOUT PLAN", True),
    ("piso", u"PLANTA DE PISO", u"FLOOR FINISH PLAN", True),
    ("forro", u"PLANTA DE FORRO", u"REFLECTED CEILING PLAN", True),
)
# outros títulos principais de planta (sem versão por ambiente)
EXTRA_PLAN_TITLES = (
    ("memoria", u"MEMÓRIA DE CÁLCULO", u"CALCULATION REPORT", False),
    ("implantacao", u"IMPLANTAÇÃO", u"SITE PLAN", False),
    ("pavimentacao", u"PAVIMENTAÇÃO", u"PAVING PLAN", False),
)
VERTICAL_TITLES = (
    ("corte", u"CORTE", u"SECTION", False),
    ("elevacao", u"ELEVAÇÃO", u"ELEVATION", False),
)
# complemento: (chave, texto PT, texto EN). FREE_SUFFIX = texto digitado
# (ou nada) e vale para todos os títulos; PAV./AMPLIAÇÃO + nível só para
# as plantas 2..6. O complemento sempre entra depois de " - ".
FREE_SUFFIX = "none"
SUFFIXES = ((FREE_SUFFIX, u"", u""),
            ("pav", u"PAV.", u"LEVEL"),
            ("amp", u"AMPLIAÇÃO", u"ENLARGED"))
MANUAL_KEY = "manual"
KEEP_NAME_KEY = "keep"   # chamada / vista ativa: não renomeia
KEEP_NAME_OPT = {"key": KEEP_NAME_KEY, "label": u"Manter nomenclatura atual da vista",
                 "pt": u"", "en": u"", "suffix": False, "room": False}


def room_info(room):
    """(número, nome, nome em inglês) - inglês cai no nome PT se vazio."""
    def text(p):
        return to_unicode(p.AsString()).strip() if p is not None and p.HasValue else u""
    number = text(room.get_Parameter(BIP.ROOM_NUMBER))
    name = text(room.get_Parameter(BIP.ROOM_NAME))
    name_en = text(room.LookupParameter(ROOM_EN_PARAM))
    return number, name, name_en


def room_prefix(room):
    number, name, _en = room_info(room)
    return u"{} - {} - ".format(number, name)


# elevação / corte: só o próprio título (+ versão por ambiente), sem nome manual
SIMPLE_KINDS = {"elevation": "elevacao", "section": "corte"}


def name_options(target, room):
    """Nome principal na ordem do padrão do escritório para o tipo de vista.
    As versões 'Nº - AMBIENTE - ...' só aparecem com 1 ambiente na seleção."""
    if target in SIMPLE_KINDS:
        only = [t for t in VERTICAL_TITLES if t[0] == SIMPLE_KINDS[target]]
        opts = [{"key": key, "label": pt, "pt": pt, "en": en, "suffix": suf, "room": False}
                for key, pt, en, suf in only]
        if room is not None:
            prefix = room_prefix(room)
            opts.extend({"key": "room_" + key, "label": prefix + pt, "pt": pt, "en": en,
                         "suffix": suf, "room": True} for key, pt, en, suf in only)
        opts.append({"key": MANUAL_KEY, "label": u"Sem nome (manual ou sequência)",
                     "pt": u"", "en": u"", "suffix": False, "room": False})
        return opts
    if target == "plan":
        # cobertura e os títulos extras não têm versão por ambiente
        base, room_base = PLAN_TITLES + EXTRA_PLAN_TITLES, PLAN_TITLES[1:]
    else:
        base, room_base = VERTICAL_TITLES, VERTICAL_TITLES
    opts = [{"key": key, "label": pt, "pt": pt, "en": en, "suffix": suf, "room": False}
            for key, pt, en, suf in base]
    if room is not None:   # depois dos simples, as versões "Nº - AMBIENTE - ..."
        prefix = room_prefix(room)
        opts.extend({"key": "room_" + key, "label": prefix + pt, "pt": pt, "en": en,
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
    Nome manual: manual vai para View Name/Title on Sheet e manual_en para o
    Title on Sheet - English. Complemento manual: idem com suffix_text e
    suffix_text_en (cada campo pode ficar vazio)."""
    if opt["key"] == KEEP_NAME_KEY:
        return u"", u""   # nada a gravar: nomes ficam como estão
    if opt["key"] == MANUAL_KEY:
        pt, en = manual.strip(), manual_en.strip()
    else:
        pt, en = opt["pt"], opt["en"]
    if not opt["suffix"]:
        suffix_key = FREE_SUFFIX   # PAV./AMPLIAÇÃO só nas plantas 2..6
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
    n = 1
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


def view_names(view):
    """(View Name, Title on Sheet, Title on Sheet - English) atuais da vista."""
    def text(p):
        return to_unicode(p.AsString()).strip() if p is not None and p.HasValue else u""
    return (to_unicode(view.Name), text(view.get_Parameter(BIP.VIEW_DESCRIPTION)),
            text(view.LookupParameter(EN_TITLE_PARAM)))


def keep_callout_names(parent_names):
    """Chamada com 'Manter nomenclatura atual': nomes da vista-mãe + ' - CHAMADA n'
    (EN ' - CALLOUT n'); n = primeiro com View Name livre. Títulos vazios ficam vazios."""
    vname, title, en = parent_names
    n = 1
    while clean_name(u"{} - CHAMADA {}".format(vname, n)) in VIEW_NAMES:
        n += 1
    return (clean_name(u"{} - CHAMADA {}".format(vname, n)),
            u"{} - CHAMADA {}".format(title, n) if title else u"",
            u"{} - CALLOUT {}".format(en, n) if en else u"")


def letter_names(prefix, title, en, own=None):
    """Elevação/corte sem complemento: '<nome> - A', ou a próxima letra cujo
    View Name esteja livre ('own' = nome atual da vista ativa, aceito)."""
    for letter in seq_letters():
        t = u"{} - {}".format(title, letter)
        vname = clean_name(with_prefix(prefix, t))
        if vname not in VIEW_NAMES or vname == own:
            return vname, t, (en + EN_SEP + letter) if en else en


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
# 6b. Pasta do Navegador de Projeto = parâmetro de projeto "Pasta" (igual ao 02)
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
    """Prefixo do View Name: da pasta escolhida; 'Não definir pasta' = nenhum."""
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
            notes.append(u"Parâmetro '{}' indisponível na vista (ou controlado pelo View "
                         u"Template); pasta não definida.".format(FOLDER_PARAM))
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
# 6c. View Template (igual ao 02: validado pelo próprio Revit; em "Recortar a
#     vista ativa" a 1ª opção mantém o template que a vista já tem)
# ------------------------------------------------------------------
NO_TEMPLATE = u"Sem View Template"
KEEP_TEMPLATE = u"Manter o View Template atual"
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


def template_items(vt, probe=None):
    """Templates que o Revit aceita para o tipo de vista: a mesma verificação
    nativa usada ao aplicar (IsValidViewTemplate), feita numa vista existente
    desse tipo (na vista ativa, ela mesma); sem vista para consultar, o grupo."""
    if vt is None:
        return []
    probe = probe or _probe_view(vt)
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


def current_template(view):
    """Template que a vista já usa, no mesmo formato dos itens; None se não tem."""
    tid = view.ViewTemplateId
    if tid is None or tid == INVALID_ID:
        return None
    tpl = doc.GetElement(tid)
    return {"label": to_unicode(tpl.Name), "id": tid} if tpl is not None else None


def apply_template(view, template, notes):
    if not view.IsValidViewTemplate(template["id"]):
        notes.append(u"View Template '{}' não é compatível com esta vista; não aplicado."
                     .format(template["label"]))
        return
    view.ViewTemplateId = template["id"]


# ------------------------------------------------------------------
# 7. Validação prévia dos modos (antes de abrir o diálogo)
# ------------------------------------------------------------------
MODES = ("plans", "callout", "current")
MODE_LABELS = {"plans": u"Nova vista de planta",
               "callout": u"Nova vista de chamada (Callout)",
               "current": u"Recortar a vista ativa"}
VERTICAL_VIEWTYPES = (VT.Section, VT.Elevation, VT.Detail)


def target_kind(mode):
    """Tipo da vista resultante: define quais nomes fazem sentido."""
    if mode == "plans":
        return "plan"
    if active_view.ViewType in PLAN_VIEWTYPES + (VT.AreaPlan,):
        return "plan"
    # chamada numa elevação/corte gera elevação/corte (ou detalhe dele)
    return {VT.Elevation: "elevation", VT.Section: "section"}.get(
        active_view.ViewType, "vertical")


def is_vertical_view(view):
    """Corte, elevação ou detalhe gerado de um deles (direção de vista horizontal)."""
    if view.ViewType in (VT.Section, VT.Elevation):
        return True
    if view.ViewType == VT.Detail:
        try:
            return abs(view.ViewDirection.Z) < 0.5
        except Exception:
            return True
    return False


def check_modes(plan_items, callout_items, plan_level):
    reasons = {}
    av_name = to_unicode(active_view.Name)
    if is_vertical_view(active_view):
        reasons["plans"] = (u"A vista ativa '{}' é corte/elevação: use a chamada ou o "
                            u"recorte da vista ativa.".format(av_name))
    elif not plan_items:
        reasons["plans"] = u"Nenhum tipo de vista de planta/área no modelo."
    elif plan_level is None:
        reasons["plans"] = u"Não foi possível identificar o nível dos elementos."
    if active_view.IsTemplate or active_view.ViewType not in CALLOUT_PARENTS:
        reasons["callout"] = (u"A vista ativa '{}' não aceita chamadas "
                              u"(use planta, corte, elevação ou detalhe).".format(av_name))
    elif not callout_items:
        reasons["callout"] = u"Nenhum tipo de vista compatível para a chamada."
    if active_view.IsTemplate or active_view.ViewType not in CROPPABLE_VIEWTYPES:
        reasons["current"] = u"A vista ativa '{}' não suporta recorte.".format(av_name)
    return reasons


INITIAL_MARGIN_M = 0.50   # valor inicial da margem ao abrir (igual ao 02)


def parse_margin(text):
    t = to_unicode(text).strip().lower().replace(u",", u".")
    if t.endswith(u"m"):
        t = t[:-1].strip()
    val = float(t)   # ValueError tratado no diálogo
    if val < 0 or val > 100:
        raise ValueError(u"fora do intervalo")
    return val


def active_view_text(callout_items):
    """O que cada modo faz a partir da vista ativa (planta, corte ou elevação)."""
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


# ------------------------------------------------------------------
# 8. Diálogo (mesma interface do 02: três colunas)
# ------------------------------------------------------------------
FOLDER_BLOCK = u"""
            <StackPanel x:Name="p_box">
              <RadioButton x:Name="p_f_exist" GroupName="p_folder" Content="Pasta existente"/>
              <ComboBox x:Name="p_list" Margin="20,4,0,0"/>
              <RadioButton x:Name="p_f_new" GroupName="p_folder" Content="Nova pasta"
                           Margin="0,8,0,0"/>
              <TextBox x:Name="p_new" Margin="20,4,0,0"
                       ToolTip="Ex.: D003_DETALHAMENTO - o texto antes do 1º '_' vira o prefixo do View Name"/>
              <RadioButton x:Name="p_f_none" GroupName="p_folder"
                           Content="Não definir pasta (View Name sem prefixo)" Margin="0,8,0,0"/>
            </StackPanel>
            <TextBlock x:Name="p_lock" TextWrapping="Wrap" FontSize="11"
                       Foreground="#FFB454" Margin="0,6,0,0"/>
            <TextBlock x:Name="p_prefix" FontSize="11" Margin="0,6,0,0"/>
"""

CROP_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Abertura de Vista" Height="Auto" Width="1240"
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
    <TextBlock Text="ABERTURA/RECORTE DE VISTA" FontSize="15" FontWeight="SemiBold"
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

      <!-- coluna 1: premissas + o que fazer -->
      <StackPanel Grid.Column="0">
        <Border>
          <StackPanel>
            <TextBlock Text="PREMISSAS" FontSize="12" FontWeight="SemiBold" Foreground="#65E3FF"/>
            <TextBlock x:Name="sel_txt" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="0,0,0,2"/>
            <TextBlock x:Name="view_txt" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="0,4,0,2"/>
            <TextBlock Text="Margem ao redor do elemento (m)"/>
            <TextBox x:Name="margin"/>
          </StackPanel>
        </Border>

        <Border>
          <StackPanel>
            <TextBlock Text="O QUE FAZER  ·  UMA VISTA PARA TODA A SELEÇÃO" FontSize="12"
                       FontWeight="SemiBold" Foreground="#65E3FF"/>
            <RadioButton x:Name="rb_plans" GroupName="mode"
                         Content="Criar nova vista de planta (Piso, Forro, Estrutural, Área)"/>
            <RadioButton x:Name="rb_callout" GroupName="mode"
                         Content="Criar vista de chamada de detalhe (Callout) na vista ativa"/>
            <RadioButton x:Name="rb_current" GroupName="mode"
                         Content="Recortar a vista ativa (planta, corte ou elevação)"/>
            <TextBlock x:Name="mode_hint" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="20,4,0,0"/>
            <TextBlock x:Name="reasons" TextWrapping="Wrap" FontSize="11"
                       Foreground="#FFB454" Margin="0,8,0,0"/>
            <TextBlock x:Name="vtype_lbl" Text="Tipo de vista"/>
            <ComboBox x:Name="vtype"/>
          </StackPanel>
        </Border>
      </StackPanel>

      <!-- coluna 2: view template + destino -->
      <StackPanel Grid.Column="2">
        <Border>
          <StackPanel>
            <TextBlock Text="MODELO DE VISTA (VIEW TEMPLATE)" FontSize="12" FontWeight="SemiBold"
                       Foreground="#65E3FF"/>
            <ComboBox x:Name="template" Margin="0,8,0,0"/>
            <TextBlock Text="Só aparecem os templates que o Revit aceita para o tipo de vista."
                       TextWrapping="Wrap" FontSize="11" Foreground="#7A8FA9" Margin="0,6,0,0"/>
          </StackPanel>
        </Border>

        <Border>
          <StackPanel>
            <TextBlock Text="DESTINO DA VISTA  ·  NAVEGADOR DE PROJETO" FontSize="12"
                       FontWeight="SemiBold" Foreground="#65E3FF"/>
            <TextBlock x:Name="folder_src" TextWrapping="Wrap" FontSize="11"
                       Foreground="#7A8FA9" Margin="0,0,0,2"/>
            <TextBlock x:Name="folder_reason" TextWrapping="Wrap" FontSize="11"
                       Foreground="#FFB454" Margin="0,0,0,2"/>""" + FOLDER_BLOCK + u"""
          </StackPanel>
        </Border>
      </StackPanel>

      <!-- coluna 3: nomenclatura + pré-visualização -->
      <StackPanel Grid.Column="4">
        <Border>
          <StackPanel>
            <TextBlock Text="NOMENCLATURA" FontSize="12" FontWeight="SemiBold"
                       Foreground="#65E3FF"/>
            <TextBlock Text="Vale para View Name, Title on Sheet e Title on Sheet - English."
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
            <TextBlock Text="PRÉ-VISUALIZAÇÃO" FontSize="10" Foreground="#7A8FA9"
                       Margin="0,6,0,6"/>
            <TextBlock Text="VIEW NAME  (prefixo da pasta + _ + nome)" FontSize="10"
                       Foreground="#7A8FA9" Margin="0"/>
            <TextBlock x:Name="pv_name" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,6"/>
            <TextBlock Text="TITLE ON SHEET" FontSize="10" Foreground="#7A8FA9" Margin="0"/>
            <TextBlock x:Name="pv_title" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,6"/>
            <TextBlock Text="TITLE ON SHEET - ENGLISH  (automático)" FontSize="10"
                       Foreground="#7A8FA9" Margin="0"/>
            <TextBlock x:Name="pv_en" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,8"/>
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

    <CheckBox x:Name="open_views" Content="Abrir a vista criada ao final" Margin="0,12,0,0"/>

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
MODE_HINTS = {
    "plans": u"Planta nova no nível dos elementos, recortada ao redor deles.",
    "callout": u"Chamada dentro da vista ativa. Numa elevação ou corte, escolha o mesmo "
               u"tipo para abrir outra elevação/corte (ampliação) ou Detalhe.",
    "current": u"Ajusta o recorte da própria vista ativa, sem criar vista nova.",
}


class FolderPicker(object):
    """Bloco de destino (pasta existente / nova / nenhuma), igual ao 02. Se o
    View Template controla o parâmetro 'Pasta', o bloco trava: a vista vai
    para a pasta do template e o prefixo vem dela."""

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
        # sempre abre em "Pasta existente" (a pasta escolhida da última vez
        # continua selecionada na lista); sem pastas, a próxima opção possível
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
    def __init__(self, xaml, ctx, reasons, cfg):
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
        self._forced = False
        self._names = []
        self._templates = []
        self._radios = {"plans": self.rb_plans, "callout": self.rb_callout,
                        "current": self.rb_current}

        self.info.Text = ctx["info"]
        self.sel_txt.Text = ctx["sel_text"]
        self.view_txt.Text = ctx["view_text"]
        self.margin.Text = u"{:.2f}".format(cfg["margin"])
        # corte/elevação: planta nova só fica apagada, sem aviso (ctx["silent"])
        self.reasons.Text = u"\n".join(
            u"• {}: {}".format(MODE_LABELS[k], reasons[k]) for k in MODES
            if k in reasons and k not in ctx["silent"])
        self.name_hint.Text = u"\n".join(ctx["hints"])
        self.open_views.IsChecked = bool(cfg["open"])
        self._suffix_idx = next(
            (i for i, s in enumerate(SUFFIXES) if s[0] == cfg["suffix"]), 0)

        enabled = [k for k in MODES if k not in reasons]
        for key, rb in self._radios.items():
            rb.IsEnabled = key in enabled
            rb.Opacity = 1.0 if key in enabled else 0.4   # bloqueada: tom mais claro
            rb.Checked += self._on_mode

        f = ctx["folders"]
        if f.available:
            self.folder_src.Text = u"Parâmetro de projeto '{}'  ·  {} pasta(s) no modelo".format(
                f.param_name, len(f.entries))
        else:
            self.folder_reason.Text = f.reason
        self.p_folder = FolderPicker(self, "p", f, cfg["p_folder"], self._on_name)

        self.vtype.SelectionChanged += self._on_vtype
        self.template.SelectionChanged += self._on_name
        self.vname.SelectionChanged += self._on_name
        self.vsuffix.SelectionChanged += self._on_name
        self.manual.TextChanged += self._on_name
        self.manual_en.TextChanged += self._on_name
        self.suffix_text.TextChanged += self._on_name
        self.suffix_text_en.TextChanged += self._on_name

        self._ready = True
        start = cfg["mode"] if cfg["mode"] in enabled else enabled[0]
        self._radios[start].IsChecked = True
        self._refresh()

        self.ok.Click += self._ok
        self.cancel.Click += self._cancel

    # --- modo, tipo de vista e template ----------------------------------
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

    def _first_template_label(self):
        return KEEP_TEMPLATE if self.mode == "current" else NO_TEMPLATE

    def _refresh_templates(self):
        current = self.template.SelectedItem
        wanted = to_unicode(current) if current is not None else self._saved["template"]
        if self.mode == "current":   # a própria vista ativa responde quais aceita
            self._templates = template_items(active_view.ViewType, active_view)
        else:
            item = self.vtype_item
            self._templates = template_items(item["vt"] if item else None)
        labels = [self._first_template_label()] + [t["label"] for t in self._templates]
        self.template.ItemsSource = labels
        self.template.SelectedIndex = labels.index(wanted) if wanted in labels else 0

    @property
    def template_choice(self):
        """Template escolhido no combo (None = 1ª opção)."""
        idx = self.template.SelectedIndex
        return self._templates[idx - 1] if 1 <= idx <= len(self._templates) else None

    def effective_template(self):
        """Template que a vista terá: o escolhido; em 'Recortar a vista ativa'
        com 'Manter', o que ela já usa (pode travar a pasta)."""
        tpl = self.template_choice
        if tpl is None and self.mode == "current":
            return current_template(active_view)
        return tpl

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

    @property
    def simple_kind(self):
        """Elevação ou corte: nome simples; complemento vazio = letra sequencial."""
        return self.ctx["kinds"][self.mode] in SIMPLE_KINDS

    def _default_name_index(self):
        if self.simple_kind:   # padrão exclusivo: ELEVAÇÃO / CORTE
            return 0
        keys = [o["key"] for o in self._names]
        if self._saved["vname"] in keys:
            return keys.index(self._saved["vname"])
        # sem escolha salva compatível: corte/elevação pelo tipo da vista ativa
        guess = {VT.Section: "corte", VT.Elevation: "elevacao"}.get(active_view.ViewType)
        return keys.index(guess) if guess in keys else 0

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
        self.mode_hint.Text = MODE_HINTS[mode]
        # recortar a vista ativa: ela já está aberta
        self.open_views.IsEnabled = mode != "current"
        self.open_views.Opacity = 1.0 if mode != "current" else 0.4

        # nomes dependem do tipo da vista resultante (planta x corte/elevação)
        self._names = self.ctx["names"][mode]
        self.vname.ItemsSource = [o["label"] for o in self._names]
        self.vname.SelectedIndex = self._default_name_index()

        # complemento traz o nome do nível da vista resultante; com eventos
        # travados, senão a seleção vazia (-1) da troca de itens vira a escolha
        if self.vsuffix.SelectedIndex >= 0 and not self._forced:
            self._suffix_idx = self.vsuffix.SelectedIndex
        self._busy = True
        try:
            self.vsuffix.ItemsSource = suffix_labels(self.ctx["level"][mode])
            self.vsuffix.SelectedIndex = self._suffix_idx
        finally:
            self._busy = False
        self._refresh_templates()
        self._update_names()

    # --- nomenclatura --------------------------------------------------
    @property
    def name_opt(self):
        idx = self.vname.SelectedIndex
        return self._names[idx] if 0 <= idx < len(self._names) else self._names[-1]

    @property
    def suffix_key(self):
        idx = self.vsuffix.SelectedIndex
        return SUFFIXES[idx if idx >= 0 else 0][0]

    def manual_suffix_on(self):
        """Os campos de complemento manual valem? Sim com o combo em 'Sem
        complemento, manual ou sequência', exceto em 'Manter nomenclatura atual'."""
        if self.name_opt["key"] == KEEP_NAME_KEY:
            return False
        return self.suffix_key == FREE_SUFFIX

    def names(self, folder):
        """(View Name, Title on Sheet, English) que serão gravados."""
        if self.name_opt["key"] == KEEP_NAME_KEY:
            if self.mode == "callout":
                return keep_callout_names(self.ctx["active_names"])
            return u"", u"", u""   # vista ativa: nada é renomeado
        on = self.manual_suffix_on()
        pt_txt = to_unicode(self.suffix_text.Text) if on else u""
        en_txt = to_unicode(self.suffix_text_en.Text) if on else u""
        is_manual = self.name_opt["key"] == MANUAL_KEY
        title, en = compose_names(self.name_opt, self.suffix_key, self.ctx["level"][self.mode],
                                  self.ctx["room"], to_unicode(self.manual.Text),
                                  pt_txt, en_txt, to_unicode(self.manual_en.Text))
        if (on and self.simple_kind and not is_manual
                and not free_suffix(pt_txt) and not free_suffix(en_txt)):
            # corte/elevação sem complemento: próxima letra livre (CORTE - A, B...)
            own = to_unicode(active_view.Name) if self.mode == "current" else None
            return letter_names(name_prefix(folder), title, en, own)
        return resolve_names(name_prefix(folder), title, en)

    def _sync_suffix(self, opt):
        """PAV./AMPLIAÇÃO só nas plantas 2..6: nas demais o combo fica em
        'Sem nomenclatura' e a escolha anterior volta depois."""
        if opt["suffix"]:
            if self._forced or self.vsuffix.SelectedIndex < 0:
                if self.vsuffix.Items.Count:
                    self.vsuffix.SelectedIndex = self._suffix_idx
                self._forced = False
            if self.vsuffix.SelectedIndex >= 0:   # combo ainda vazio: mantém a escolha salva
                self._suffix_idx = self.vsuffix.SelectedIndex
        else:
            self._forced = True
            self.vsuffix.SelectedIndex = 0

    def _update_names(self):
        if self._busy or not self._names:   # mudar o combo aqui dispara SelectionChanged
            return
        self._busy = True
        try:
            self._render()
        finally:
            self._busy = False

    def _render(self):
        opt = self.name_opt
        is_manual = opt["key"] == MANUAL_KEY
        keep_name = opt["key"] == KEEP_NAME_KEY
        self._sync_suffix(opt)
        self.vsuffix.IsEnabled = opt["suffix"] and not keep_name
        self.suffix_lbl.Opacity = 1.0 if not keep_name else 0.4
        on = self.manual_suffix_on()
        for ctl in (self.suffix_text, self.suffix_text_en):
            ctl.IsEnabled = on
        for lbl in (self.suffix_text_lbl, self.suffix_text_en_lbl):
            lbl.Opacity = 1.0 if on else 0.4
        for ctl in (self.manual, self.manual_en):
            ctl.IsEnabled = is_manual
        for lbl in (self.manual_lbl, self.manual_en_lbl):
            lbl.Opacity = 1.0 if is_manual else 0.4

        template = self.effective_template()
        folder = self.p_folder.render(template)
        vname, title, en = self.names(folder)
        if self.mode == "current":
            # campo não gravado = continua o valor que a vista ativa já tem
            cur = self.ctx["active_names"]
            vname, title, en = vname or cur[0], title or cur[1], en or cur[2]
        self.pv_name.Text = vname or u"(nome padrão do Revit)"
        self.pv_title.Text = title or u"—"
        self.pv_en.Text = en or u"—"
        self.pv_folder.Text = folder["name"] or (u"??? (template)" if folder["locked"]
                                                 else u"(não definida pelo comando)")
        chosen = self.template_choice
        self.pv_template.Text = (chosen["label"] if chosen else
                                 (u"{} ({})".format(KEEP_TEMPLATE, template["label"])
                                  if template else self._first_template_label()))
        self.live_hint.Text = u"\n".join(self._live_hints(folder, chosen))

    def _live_hints(self, folder, chosen):
        hints = []
        entries = [e["name"] for e in self.ctx["folders"].entries]
        if folder["name"] and folder["mode"] not in ("none", "template") \
                and not folder_prefix(folder["name"]):
            hints.append(u"A pasta '{}' não tem '_': o View Name fica sem prefixo."
                         .format(folder["name"]))
        if folder["mode"] == "new" and folder["name"] in entries:
            hints.append(u"A pasta '{}' já existe: a vista vai para ela.".format(folder["name"]))
        if self.mode == "current" and chosen is not None:
            hints.append(u"O View Template '{}' substitui o template atual da vista ativa."
                         .format(chosen["label"]))
        return hints

    def _ok(self, sender, args):
        try:
            self.margin_m = parse_margin(self.margin.Text)
        except Exception:
            self.error.Text = u"Margem inválida: informe um número em metros entre 0 e 100 (ex.: 0,50)."
            return
        c = self.p_folder.choice(self.effective_template())
        if c["mode"] == "new" and not c["name"]:
            self.error.Text = u"Informe o nome da nova pasta (ex.: D003_DETALHAMENTO)."
            return
        if c["mode"] == "existing" and not c["name"]:
            self.error.Text = u"Selecione a pasta de destino."
            return
        if self.mode != "current" and self.vtype_item is None:
            self.error.Text = u"Escolha o tipo de vista."
            return
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, args):
        self.Close()

    @property
    def result(self):
        folder = self.p_folder.choice(self.effective_template())
        vname, title, en = self.names(folder)
        sel_tpl = self.template.SelectedItem
        return {"mode": self.mode, "margin": self.margin_m,
                "type": self.vtype_item,
                "vname": self.name_opt["key"], "suffix": self.suffix_key,
                "view_name": vname, "title": title, "name_en": en,
                "open": bool(self.open_views.IsChecked) and self.mode != "current",
                "folder": folder,
                "p_saved": self.p_folder.saved_values(),
                "template": self.template_choice,
                "template_label": to_unicode(sel_tpl) if sel_tpl is not None else NO_TEMPLATE}


# ------------------------------------------------------------------
# 9. Execução
# ------------------------------------------------------------------
warnings = []

elements, ignored = get_target_elements()
if ignored:
    warnings.append(u"{} elemento(s) da seleção ignorado(s): vistas, viewports e tipos "
                    u"não definem recorte.".format(ignored))

geoms = []
for el in elements:
    if isinstance(el, DB.SpatialElement) and el.Area <= 0:
        warnings.append(u"{}: ambiente não colocado ou não delimitado - ignorado.".format(elem_label(el)))
        continue
    g = extract_geometry(el)
    if not g.points:
        warnings.append(u"{}: sem geometria utilizável - ignorado.".format(g.label))
        continue
    geoms.append(g)

if not geoms:
    forms.alert(u"Nenhum elemento válido selecionado.\n\n" + u"\n".join(warnings), exitscript=True)

plan_level, other_levels = pick_plan_level(geoms)
plan_items = plan_type_items()
callout_items = callout_type_items(active_view)
reasons = check_modes(plan_items, callout_items, plan_level)
if len(reasons) == len(MODES):
    forms.alert(u"Nenhuma ação disponível:\n\n" + u"\n".join(reasons.values()), exitscript=True)

folders = ProjectFolders()

# ambiente usado nas opções "Nº - AMBIENTE - ...": exatamente 1 na seleção
rooms = [g.el for g in geoms if isinstance(g.el, Room)]
the_room = rooms[0] if len(rooms) == 1 else None
name_hints = []
if len(rooms) > 1:
    name_hints.append(u"Opções com número/nome do ambiente exigem exatamente 1 ambiente "
                      u"na seleção (há {}). Para uma vista por ambiente, use o "
                      u"02_Abertura de Múltiplas Vistas.".format(len(rooms)))
if the_room is not None and not room_info(the_room)[2]:
    name_hints.append(u"'{}' vazio neste ambiente: o título em inglês usará o nome em português."
                      .format(ROOM_EN_PARAM))
if not en_param_bound():
    name_hints.append(u"'{}' não está vinculado à categoria Vistas: o título em inglês "
                      u"não será gravado.".format(EN_TITLE_PARAM))
if folders.available and not folders.entries:
    name_hints.append(u"Nenhuma vista tem o parâmetro '{}' preenchido ainda: use 'Nova pasta'."
                      .format(folders.param_name))

active_level = getattr(active_view, "GenLevel", None)
sel_preview = u", ".join(g.label for g in geoms[:6]) + (u"..." if len(geoms) > 6 else u"")
dialog_ctx = {
    "types": {"plans": plan_items, "callout": callout_items, "current": []},
    "names": dict((m, name_options(target_kind(m), the_room)
                   + ([KEEP_NAME_OPT] if m in ("callout", "current") else []))
                  for m in MODES),
    "kinds": dict((m, target_kind(m)) for m in MODES),
    "active_names": view_names(active_view),
    "silent": set(["plans"]) if is_vertical_view(active_view) else set(),
    "level": {"plans": plan_level, "callout": active_level, "current": active_level},
    "room": the_room,
    "hints": name_hints,
    "folders": folders,
    "info": u"{} elemento(s) → 1 vista".format(len(geoms)),
    "sel_text": u"{} elemento(s) selecionado(s): {}".format(len(geoms), sel_preview),
    "view_text": active_view_text(callout_items),
}

config = script.get_config()
# margem sempre abre com o valor inicial (o usuário altera no diálogo)
cfg = {"margin": INITIAL_MARGIN_M,
       "mode": config.get_option("mode", "current"),
       "vtype": config.get_option("vtype", u""),
       "vname": config.get_option("vname", MANUAL_KEY),
       "suffix": config.get_option("vsuffix", "amp"),
       "open": config.get_option("open_views", True),
       "p_folder": {"folder_mode": config.get_option("folder_mode", "none"),
                    "folder": config.get_option("folder", u""),
                    "folder_new": config.get_option("folder_new", u"")},
       "template": config.get_option("template", NO_TEMPLATE)}

win = CropWindow(CROP_XAML, dialog_ctx, reasons, cfg)
win.ShowDialog()
if not win.confirmed:
    script.exit()
opts = win.result

config.mode = opts["mode"]
config.vtype = opts["type"]["label"] if opts["type"] else cfg["vtype"]
config.vname = opts["vname"]
config.vsuffix = opts["suffix"]
if opts["mode"] != "current":   # na vista ativa a opção fica desabilitada
    config.open_views = opts["open"]
config.folder_mode = opts["p_saved"]["folder_mode"]
config.folder = opts["p_saved"]["folder"]
config.folder_new = opts["p_saved"]["folder_new"]
config.template = opts["template_label"]
script.save_config()

mode = opts["mode"]
margin_ft = m_to_ft(opts["margin"])

# --- verificações específicas do modo, antes de qualquer alteração ---------
remove_scope_box = False
if mode == "current":
    sb = active_view.get_Parameter(BIP.VIEWER_VOLUME_OF_INTEREST_CROP)
    if sb is not None and sb.AsElementId() != INVALID_ID:
        sb_name = to_unicode(doc.GetElement(sb.AsElementId()).Name)
        if not forms.alert(u"O recorte da vista ativa está vinculado à Scope Box '{}'.\n\n"
                           u"Remover o vínculo e aplicar o recorte do elemento?".format(sb_name),
                           yes=True, no=True):
            script.exit()
        remove_scope_box = True

if mode in ("current", "callout"):
    visible = set(eid_int(i) for i in
                  DB.FilteredElementCollector(doc, active_view.Id)
                  .WhereElementIsNotElementType().ToElementIds())
    for g in geoms:
        if eid_int(g.el.Id) not in visible:
            warnings.append(u"{}: não está visível na vista ativa (recorte calculado mesmo assim).".format(g.label))

if mode == "plans" and other_levels:
    warnings.append(u"Elementos em mais de um nível: planta criada no nível '{}' "
                    u"(também havia: {}).".format(to_unicode(plan_level.Name),
                                                  u", ".join(other_levels)))


vtype = opts["type"]
folder = opts["folder"]
# pasta travada = o View Template já define; o comando não grava
folder_write = u"" if folder["locked"] else folder["name"]
template = opts["template"]
notes = []
error = None
view = None

t = DB.Transaction(doc, u"OCA - Abertura de Vistas")
t.Start()
try:
    if remove_scope_box:
        active_view.get_Parameter(BIP.VIEWER_VOLUME_OF_INTEREST_CROP).Set(INVALID_ID)
        doc.Regenerate()

    if mode == "current":
        view = active_view
    elif mode == "callout":
        parent_calc = CropCalc(geoms, Frame(active_view), margin_ft)
        u0, v0, u1, v1 = parent_calc.envelope()
        f = parent_calc.frame
        view = DB.ViewSection.CreateCallout(doc, active_view.Id, vtype["id"],
                                            f.to3d((u0, v0)), f.to3d((u1, v1)))
    elif vtype["kind"] == "area":
        view = DB.ViewPlan.CreateAreaPlan(doc, vtype["id"], plan_level.Id)
    else:
        view = DB.ViewPlan.Create(doc, vtype["id"], plan_level.Id)

    if mode != "current":
        # o tipo de vista pode ter template padrão para vistas novas: remover
        # (o template escolhido no diálogo entra no fim, depois do recorte)
        view.ViewTemplateId = INVALID_ID

    apply_names(view, opts["view_name"], opts["name_en"], notes)  # View Name + English
    apply_title_on_sheet(view, opts["title"], notes)              # Title on Sheet (sem prefixo)
    ProjectFolders.apply(view, folder_write, notes)
    doc.Regenerate()

    # recalcula no sistema da vista-alvo: a orientação relativa pode
    # mudar (ex.: planta nova sem a rotação de recorte da vista ativa)
    calc = CropCalc(geoms, Frame(view), margin_ft)
    _method, warn = apply_crop(view, calc)
    if warn:
        notes.append(warn)

    if template is not None:
        apply_template(view, template, notes)
        doc.Regenerate()
        ProjectFolders.check(view, folder_write, notes)

    status = t.Commit()
    if status != DB.TransactionStatus.Committed:
        raise Exception(u"o Revit desfez a alteração ({}).".format(status))
except Exception as exc:
    if t.GetStatus() == DB.TransactionStatus.Started:
        t.RollBack()   # nada fica pela metade (vista criada sem recorte etc.)
    error = to_unicode(exc)
    view = None

if view is not None and opts["open"]:
    try:
        uidoc.ActiveView = view
    except Exception as exc:
        notes.append(u"Vista criada, mas não foi possível abri-la ({}).".format(to_unicode(exc)))


# ------------------------------------------------------------------
# 10. Resumo: só aparece se houver falha ou algo a revisar
# ------------------------------------------------------------------
if error:
    forms.alert(u"Não foi possível concluir: {}".format(error),
                sub_msg=u"\n".join(warnings) or None)
elif warnings or notes:
    forms.alert(u"Vista '{}' ajustada, com avisos:".format(to_unicode(view.Name)),
                sub_msg=u"\n".join(u"- " + w for w in warnings + notes))
