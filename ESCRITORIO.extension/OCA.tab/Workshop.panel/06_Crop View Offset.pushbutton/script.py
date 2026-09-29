# -*- coding: utf-8 -*-
"""Crop View Offset - ajusta o recorte (CropView) de vistas a partir de elementos.

Fluxo:
  1. Usa a seleção atual (ou pede para selecionar) Linhas, Terreno, Pisos,
     Paredes ou Ambientes.
  2. Pergunta a margem (m) e o que fazer:
       - criar novas vistas de planta (Piso, Forro, Estrutural, Área);
       - criar vistas de chamada (Callout) na vista ativa;
       - recortar a vista ativa.
     Antes de abrir o diálogo, cada modo é validado contra a vista ativa e o
     modelo; modos inválidos ficam desabilitados com o motivo exibido.
  3. Calcula o retângulo ORIENTADO ao elemento (as 4 quinas, paralelo e
     perpendicular à direção predominante do elemento) + margem. A vista
     nunca é girada: o recorte é sempre aplicado como forma de 4 lados
     (crop region shape), inclinada quando o elemento está girado.
  4. Gera um relatório HTML com o desenho de cada recorte aplicado.
"""

__title__ = "Crop View\nOffset"
__doc__ = ("Recorta a vista ativa, cria plantas ou chamadas ajustadas ao "
           "elemento selecionado, com margem em metros.")

import os
import re
import json
import math
import codecs
from datetime import datetime

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


def ft_to_m(feet):
    return DB.UnitUtils.ConvertFromInternalUnits(feet, DB.UnitTypeId.Meters)


HALF_PI = math.pi / 2.0
QUARTER_PI = math.pi / 4.0
ANGLE_BIN = math.radians(0.5)             # resolução do histograma de direções
N_BINS = int(round(HALF_PI / ANGLE_BIN))  # 180 bins cobrindo 0..90°
ANGLE_SNAP = math.radians(0.01)           # abaixo disso o elemento está "alinhado"
MIN_SEG = m_to_ft(0.001)                  # ignora arestas < 1 mm na projeção
MIN_CROP = m_to_ft(0.01)                  # recorte mínimo de 1 cm por lado
LOCATION_WEIGHT = 10.0                    # peso extra do eixo de paredes retas
MAX_SVG_SEGS = 2500                       # limite de arestas desenhadas no HTML


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


ALLOWED_CATS = {}
for _name, _label in (("OST_Lines", u"Linha"),
                      ("OST_Topography", u"Terreno"),
                      ("OST_Toposolid", u"Terreno"),
                      ("OST_Floors", u"Piso"),
                      ("OST_Walls", u"Parede"),
                      ("OST_Rooms", u"Ambiente")):
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


def get_target_elements():
    """Seleção atual filtrada; se vazia, pede para o usuário selecionar."""
    current = list(revit.get_selection().elements)
    picked = [el for el in current if elem_cat_int(el) in ALLOWED_CATS]
    ignored = len(current) - len(picked)
    if picked:
        return picked, ignored
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, AllowedFilter(),
            u"Selecione linhas, terreno, pisos, paredes ou ambientes e clique em Concluir")
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
        self.tight = self._corners(x0, y0, x1, y1)
        self.corners = self._corners(x0 - m, y0 - m, x1 + m, y1 + m)
        self.width = (x1 - x0) + 2 * m
        self.height = (y1 - y0) + 2 * m
        self.segs = segs
        self.hull = convex_hull(pts) if not segs else []

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
# 5. Níveis, tipos de vista e nomes
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


PLAN_FAMILIES = (
    (DB.ViewFamily.FloorPlan, u"Planta de Piso", u"PLANTA"),
    (DB.ViewFamily.CeilingPlan, u"Planta de Forro", u"FORRO"),
    (DB.ViewFamily.StructuralPlan, u"Planta Estrutural", u"ESTRUTURAL"),
)
PLAN_VIEWTYPES = (VT.FloorPlan, VT.CeilingPlan, VT.EngineeringPlan)
CROPPABLE_VIEWTYPES = PLAN_VIEWTYPES + (VT.AreaPlan, VT.Section, VT.Elevation, VT.Detail)
CALLOUT_PARENTS = PLAN_VIEWTYPES + (VT.Section, VT.Elevation, VT.Detail)

VIEW_FAMILY_TYPES = list(DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType))


def _vft_name(vft):
    return to_unicode(DB.Element.Name.GetValue(vft))


def plan_type_items():
    items = []
    for fam, label, prefix in PLAN_FAMILIES:
        for vft in VIEW_FAMILY_TYPES:
            if vft.ViewFamily == fam:
                items.append({"label": u"{}: {}".format(label, _vft_name(vft)),
                              "kind": "plan", "id": vft.Id, "prefix": prefix})
    if any(v.ViewFamily == DB.ViewFamily.AreaPlan for v in VIEW_FAMILY_TYPES):
        for scheme in DB.FilteredElementCollector(doc).OfClass(DB.AreaScheme):
            items.append({"label": u"Planta de Área: {}".format(to_unicode(scheme.Name)),
                          "kind": "area", "id": scheme.Id, "prefix": u"ÁREA"})
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
                              "kind": "callout", "id": vft.Id, "prefix": u"CHAMADA"})
    return items


VIEW_NAMES = set(to_unicode(v.Name) for v in
                 DB.FilteredElementCollector(doc).OfClass(DB.View))
_BAD_NAME_CHARS = re.compile(u"[\\\\:{}\\[\\]|;<>?`~]")


def unique_view_name(base):
    base = _BAD_NAME_CHARS.sub(u"-", base).strip()
    name, n = base, 2
    while name in VIEW_NAMES:
        name = u"{} ({})".format(base, n)
        n += 1
    VIEW_NAMES.add(name)
    return name


def view_type_label(view):
    try:
        return to_unicode(view.ViewType.ToString())
    except Exception:
        return u""


# ------------------------------------------------------------------
# 6. Validação prévia dos modos (antes de abrir o diálogo)
# ------------------------------------------------------------------
MODES = ("plans", "callout", "current")
MODE_LABELS = {"plans": u"Novas vistas de planta",
               "callout": u"Novas vistas de chamada (Callout)",
               "current": u"Recortar a vista atual"}


def check_modes(plan_items, callout_items):
    reasons = {}
    av_name = to_unicode(active_view.Name)
    if not plan_items:
        reasons["plans"] = u"Nenhum tipo de vista de planta/área no modelo."
    elif not LEVELS:
        reasons["plans"] = u"O modelo não tem níveis."
    if active_view.IsTemplate or active_view.ViewType not in CALLOUT_PARENTS:
        reasons["callout"] = (u"A vista ativa '{}' não aceita chamadas "
                              u"(use planta, corte, elevação ou detalhe).".format(av_name))
    elif not callout_items:
        reasons["callout"] = u"Nenhum tipo de vista compatível para a chamada."
    if active_view.IsTemplate or active_view.ViewType not in CROPPABLE_VIEWTYPES:
        reasons["current"] = u"A vista ativa '{}' não suporta recorte.".format(av_name)
    return reasons


def parse_margin(text):
    t = to_unicode(text).strip().lower().replace(u",", u".")
    if t.endswith(u"m"):
        t = t[:-1].strip()
    val = float(t)   # ValueError tratado no diálogo
    if val < 0 or val > 100:
        raise ValueError(u"fora do intervalo")
    return val


# ------------------------------------------------------------------
# 7. Diálogo (mesmo tema dos outros botões)
# ------------------------------------------------------------------
CROP_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Crop View Offset" Height="Auto" Width="460"
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
      <Setter Property="Margin" Value="0,8,0,0"/>
    </Style>
  </Window.Resources>
  <StackPanel Margin="18">
    <TextBlock Text="RECORTE POR ELEMENTO" FontSize="15" FontWeight="SemiBold"
               Foreground="#65E3FF" Margin="0,0,0,4"/>
    <TextBlock x:Name="info" TextWrapping="Wrap" FontSize="11"
               Foreground="#7A8FA9" Margin="0,0,0,6"/>

    <TextBlock Text="Margem ao redor do elemento (m)"/>
    <TextBox x:Name="margin"/>

    <TextBlock Text="O que fazer" Margin="0,14,0,0"/>
    <RadioButton x:Name="rb_plans" GroupName="mode"
                 Content="Criar novas vistas de planta (Piso, Forro, Estrutural, Área)"/>
    <RadioButton x:Name="rb_callout" GroupName="mode"
                 Content="Criar vistas de chamada (Callout) na vista ativa"/>
    <RadioButton x:Name="rb_current" GroupName="mode"
                 Content="Recortar a vista ativa"/>
    <TextBlock x:Name="reasons" TextWrapping="Wrap" FontSize="11"
               Foreground="#FFB454" Margin="0,8,0,0"/>

    <TextBlock x:Name="vtype_lbl" Text="Tipo de vista"/>
    <ComboBox x:Name="vtype"/>
    <CheckBox x:Name="per_elem" Content="Uma vista por elemento (desmarcado: uma vista para toda a seleção)"/>
    <CheckBox x:Name="report" Content="Abrir relatório HTML ao final"/>

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
    def __init__(self, xaml, info, items, reasons, cfg):
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.confirmed = False
        self._items = items
        self._saved_type = cfg["vtype"]
        self._radios = {"plans": self.rb_plans, "callout": self.rb_callout,
                        "current": self.rb_current}

        self.info.Text = info
        self.margin.Text = u"{:.2f}".format(cfg["margin"])
        self.per_elem.IsChecked = cfg["per_element"]
        self.report.IsChecked = cfg["report"]
        self.reasons.Text = u"\n".join(
            u"• {}: {}".format(MODE_LABELS[k], reasons[k]) for k in MODES if k in reasons)

        enabled = [k for k in MODES if k not in reasons]
        for key, rb in self._radios.items():
            rb.IsEnabled = key in enabled
            rb.Checked += self._on_mode
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

    def _refresh(self):
        items = self._items.get(self.mode, [])
        labels = [i["label"] for i in items]
        self.vtype.ItemsSource = labels
        if labels:
            self.vtype.SelectedItem = (self._saved_type if self._saved_type in labels
                                       else labels[0])
        has_types = bool(labels)
        self.vtype.IsEnabled = has_types
        self.vtype_lbl.Opacity = 1.0 if has_types else 0.4
        self.per_elem.IsEnabled = self.mode != "current"

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
        items = self._items.get(self.mode, [])
        sel = self.vtype.SelectedItem
        chosen = next((i for i in items if i["label"] == sel), None)
        return {"mode": self.mode, "margin": self.margin_m, "type": chosen,
                "per_element": bool(self.per_elem.IsChecked),
                "report": bool(self.report.IsChecked)}


# ------------------------------------------------------------------
# 8. Execução
# ------------------------------------------------------------------
warnings = []

elements, ignored = get_target_elements()
if ignored:
    warnings.append(u"{} elemento(s) da seleção ignorado(s): categoria não suportada.".format(ignored))

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

plan_items = plan_type_items()
callout_items = callout_type_items(active_view)
reasons = check_modes(plan_items, callout_items)
if len(reasons) == len(MODES):
    forms.alert(u"Nenhuma ação disponível:\n\n" + u"\n".join(reasons.values()), exitscript=True)

config = script.get_config()
cfg = {"margin": config.get_option("margin_m", 0.5),
       "mode": config.get_option("mode", "current"),
       "vtype": config.get_option("vtype", u""),
       "per_element": config.get_option("per_element", True),
       "report": config.get_option("open_report", True)}

info = u"{} elemento(s) · vista ativa: {} ({})".format(
    len(geoms), to_unicode(active_view.Name), view_type_label(active_view))
win = CropWindow(CROP_XAML, info,
                 {"plans": plan_items, "callout": callout_items, "current": []},
                 reasons, cfg)
win.ShowDialog()
if not win.confirmed:
    script.exit()
opts = win.result

config.margin_m = opts["margin"]
config.mode = opts["mode"]
config.vtype = opts["type"]["label"] if opts["type"] else cfg["vtype"]
config.per_element = opts["per_element"]
config.open_report = opts["report"]
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

level_of = {}
if mode == "plans":
    kept = []
    for g in geoms:
        lid = element_level_id(g.el)
        if lid is None:
            warnings.append(u"{}: nível não identificado - ignorado.".format(g.label))
            continue
        level_of[eid_int(g.el.Id)] = lid
        kept.append(g)
    geoms = kept
    if not geoms:
        forms.alert(u"Nenhum elemento com nível identificável.\n\n" + u"\n".join(warnings),
                    exitscript=True)


def build_groups():
    if mode == "current" or not opts["per_element"]:
        groups = [geoms]
    else:
        groups = [[g] for g in geoms]
    if mode != "plans":
        return groups
    split = []   # plantas: uma vista por nível
    for grp in groups:
        by_level = {}
        for g in grp:
            by_level.setdefault(eid_int(level_of[eid_int(g.el.Id)]), []).append(g)
        split.extend(by_level.values())
    return split


def group_label(grp):
    return grp[0].label if len(grp) == 1 else u"Seleção ({} elementos)".format(len(grp))


def _m2(pt):
    return [round(ft_to_m(pt[0]), 3), round(ft_to_m(pt[1]), 3)]


def drawing(calc):
    seen, segs = set(), []
    for a, b, _w in calc.segs:
        key = tuple(_m2(a) + _m2(b))
        if key in seen or (key[0] == key[2] and key[1] == key[3]):
            continue
        seen.add(key)
        segs.append(list(key))
        if len(segs) >= MAX_SVG_SEGS:
            break
    hull = [_m2(p) for p in calc.hull]
    if hull:
        segs.extend(hull[i] + hull[(i + 1) % len(hull)] for i in range(len(hull)))
    return {"segs": segs,
            "crop": [_m2(p) for p in calc.corners],
            "tight": [_m2(p) for p in calc.tight]}


def record(view, action, grp, calc, method, warn):
    rec = {"view": to_unicode(view.Name), "viewId": eid_int(view.Id),
           "viewType": view_type_label(view), "action": action,
           "elements": [g.label for g in grp],
           "angle": round(math.degrees(calc.angle), 2),
           "width": round(ft_to_m(calc.width), 3),
           "height": round(ft_to_m(calc.height), 3),
           "method": method, "source": calc.source,
           "status": "warn" if warn else "ok", "msg": warn or u""}
    rec.update(drawing(calc))
    return rec


def error_record(grp, exc):
    return {"view": u"-", "viewId": None, "viewType": u"", "action": u"Falhou",
            "elements": [g.label for g in grp], "status": "error",
            "msg": to_unicode(exc), "segs": [], "crop": [], "tight": []}


results = []
vtype = opts["type"]

with revit.Transaction(u"OCA - Crop View Offset"):
    if remove_scope_box:
        active_view.get_Parameter(BIP.VIEWER_VOLUME_OF_INTEREST_CROP).Set(INVALID_ID)
        doc.Regenerate()

    for grp in build_groups():
        try:
            if mode == "current":
                view, action = active_view, u"Recortada"
            elif mode == "callout":
                parent_calc = CropCalc(grp, Frame(active_view), margin_ft)
                u0, v0, u1, v1 = parent_calc.envelope()
                f = parent_calc.frame
                view = DB.ViewSection.CreateCallout(doc, active_view.Id, vtype["id"],
                                                    f.to3d((u0, v0)), f.to3d((u1, v1)))
                action = u"Criada"
            else:
                lid = level_of[eid_int(grp[0].el.Id)]
                if vtype["kind"] == "area":
                    view = DB.ViewPlan.CreateAreaPlan(doc, vtype["id"], lid)
                else:
                    view = DB.ViewPlan.Create(doc, vtype["id"], lid)
                action = u"Criada"

            if action == u"Criada":
                view.Name = unique_view_name(u"{} - {}".format(vtype["prefix"], group_label(grp)))
                doc.Regenerate()

            # recalcula no sistema da vista-alvo: a orientação relativa pode
            # mudar (ex.: planta nova sem a rotação de recorte da vista ativa)
            calc = CropCalc(grp, Frame(view), margin_ft)
            method, warn = apply_crop(view, calc)
            results.append(record(view, action, grp, calc, method, warn))
        except Exception as exc:
            results.append(error_record(grp, exc))


# ------------------------------------------------------------------
# 9. Relatório HTML
# ------------------------------------------------------------------
n_err = sum(1 for r in results if r["status"] == "error")

if opts["report"]:
    data = {
        "project": to_unicode(doc.Title),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "activeView": to_unicode(active_view.Name),
        "mode": MODE_LABELS[mode],
        "viewType": vtype["label"] if vtype else u"",
        "margin": opts["margin"],
        "results": results,
        "warnings": warnings,
    }

    template_path = os.path.join(os.path.dirname(__file__), "cropview.html")
    with codecs.open(template_path, "r", encoding="utf-8") as f:
        html = f.read()
    html = html.replace("__DATA__", json.dumps(data, ensure_ascii=False,
                                                default=lambda o: int(o)))

    out_path = script.get_document_data_file("crop_view_offset", "html")
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
        forms.alert(u"Relatório gerado, mas não foi possível abri-lo.\n\n{}".format(out_path))

elif n_err:
    forms.alert(u"{} de {} vista(s) falharam:\n\n{}".format(
        n_err, len(results),
        u"\n".join(u"- {}: {}".format(u", ".join(r["elements"]), r["msg"])
                   for r in results if r["status"] == "error")))
