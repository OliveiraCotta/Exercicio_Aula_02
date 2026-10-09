# -*- coding: utf-8 -*-
"""Exportar Coordenadas - lê as coordenadas do Project Base Point / Survey Point
do modelo (e, opcionalmente, dos vínculos Revit) e exporta para Excel.

Sistemas de coordenadas tratados (cada um é uma referência DIFERENTE):

  * Origem interna (Internal Origin)
      Origem fixa (0,0,0) do banco de dados do Revit. Toda geometria devolvida
      pela API (XYZ, BasePoint.Position, LocationPoint...) está neste sistema,
      em PÉS, com o eixo Y = Norte do Projeto. Não pode ser movida.
  * Project Base Point (PBP)
      Ponto de referência do projeto. BasePoint.Position = onde ele está em
      coordenadas INTERNAS. Os valores N/S, L/O e Elevação mostrados nas
      propriedades do Revit são as coordenadas COMPARTILHADAS desse ponto
      (parâmetros BASEPOINT_NORTHSOUTH/EASTWEST/ELEVATION_PARAM).
  * Survey Point (SP)
      Marca a origem do sistema de coordenadas compartilhadas (quando está
      "clipped"). Os valores exibidos também são coordenadas COMPARTILHADAS.
  * Coordenadas compartilhadas (Shared Coordinates)
      Sistema definido pelo sítio ativo (doc.ActiveProjectLocation). A conversão
      interna -> compartilhada de qualquer ponto é feita pela própria API em
      ProjectLocation.GetProjectPosition(ponto) - não é calculada "na mão".
  * Norte do Projeto (Project North)
      É o eixo +Y interno (orientação das vistas "Norte do projeto").
  * Norte Verdadeiro (True North)
      É o eixo Norte/Sul do sistema compartilhado. O ângulo entre os dois é o
      "Ângulo para o Norte verdadeiro" (BASEPOINT_ANGLETON_PARAM do PBP); é uma
      propriedade do SÍTIO (sistema de coordenadas), não de um ponto.

Por isso o comando NÃO pega a posição XYZ e a apresenta como se fossem as
coordenadas do Revit: N/S, L/O e Elevação vêm dos parâmetros exibidos pelo
Revit, conferidos contra BasePoint.SharedPosition e contra
ProjectLocation.GetProjectPosition(BasePoint.Position).

Compatível com Revit 2024+ (ForgeTypeId / SpecTypeId / UnitTypeId) e com o
IronPython do pyRevit. O .xlsx é gerado direto em XML (sem Excel/COM e sem
bibliotecas externas).
"""

__title__ = "Exportar\nCoordenadas"
__doc__ = ("Lê as coordenadas do Project Base Point ou do Survey Point do modelo "
           "(e, opcionalmente, dos modelos vinculados), mostra em uma tabela e "
           "exporta para Excel (.xlsx).")

import os
import re
import math
import traceback
from datetime import datetime

from pyrevit import revit, DB, script, forms
from oca_ui import build_xaml, alert_title

import clr
clr.AddReference("System.Data")
for _asm in ("System.IO.Compression",
             "System.IO.Compression.FileSystem",   # .NET Framework (Revit <= 2024)
             "System.IO.Compression.ZipFile"):     # .NET 8 (Revit 2025+)
    try:
        clr.AddReference(_asm)
    except Exception:
        pass
from System import String
from System.Data import DataTable
from System.IO import FileStream, FileMode, FileAccess, FileShare, StreamWriter
from System.IO.Compression import ZipArchive, ZipArchiveMode
from System.Text import UTF8Encoding
from System.Windows import Visibility

try:                                    # IronPython 2.7 x CPython 3
    unicode
except NameError:                       # pragma: no cover
    unicode = str

doc = revit.doc
BIP = DB.BuiltInParameter
TITLE = alert_title(__title__)          # "OCA · Exportar Coordenadas" nos alertas


# ------------------------------------------------------------------
# Configuração
# ------------------------------------------------------------------
# Parâmetro de Informações do Projeto com o nome do modelo. Se não existir ou
# estiver vazio, usa-se o nome do arquivo (e isso é informado na observação).
BLOCK_PARAM_NAMES = [u"Bloco Nome"]

REF_PBP = "pbp"
REF_SP = "sp"
REF_LABELS = {REF_PBP: u"Project Base Point", REF_SP: u"Survey Point"}

TIPO_HOST = u"Hospedeiro"
TIPO_LINK = u"Vínculo"

PRECISIONS = [2, 3, 4]
DEFAULT_PRECISION = 3

NA = u"-"                    # valor para informação indisponível
TOL_FT = 1e-6                # tolerância (pés) nas verificações cruzadas
PROBE_FT = 100.0             # distância usada para medir a direção do Norte

LINK_STATUS_PT = {
    "Unloaded": u"descarregado",
    "LocallyUnloaded": u"descarregado localmente (só para este usuário)",
    "NotFound": u"arquivo não encontrado",
    "InClosedWorkset": u"em workset fechado",
    "Invalid": u"vínculo inválido",
}


# ------------------------------------------------------------------
# Helpers gerais
# ------------------------------------------------------------------
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
        return u"{}".format(value)


def elem_name(elem):
    """Element.Name protegido (em alguns tipos o getter falha no IronPython)."""
    try:
        return to_unicode(elem.Name)
    except Exception:
        try:
            return to_unicode(DB.Element.Name.GetValue(elem))
        except Exception:
            return u""


def param_double(elem, bip):
    """Valor Double de um BuiltInParameter, em unidades INTERNAS, ou None."""
    if elem is None:
        return None
    try:
        p = elem.get_Parameter(bip)
        if p is not None and p.HasValue and p.StorageType == DB.StorageType.Double:
            return p.AsDouble()
    except Exception:
        pass
    return None


def xyz_tuple(xyz):
    return None if xyz is None else (xyz.X, xyz.Y, xyz.Z)


def max_diff(a, b):
    return max(abs(a[i] - b[i]) for i in range(3))


def fmt_num(value, prec):
    """Número com 'prec' casas e ponto decimal (independe da cultura do Windows)."""
    if value is None:
        return NA
    return u"{0:.{1}f}".format(value, prec)


def fmt_deg(value, prec):
    return NA if value is None else fmt_num(value, prec) + u"°"


def direction_text(azimuth, prec):
    """Azimute horário (0-360°) -> '15.000° Leste' / '15.000° Oeste'."""
    if azimuth is None:
        return NA
    az = azimuth % 360.0
    half_step = 0.5 * 10 ** (-prec)
    if az < half_step or 360.0 - az < half_step:
        return u"{}° (alinhado)".format(fmt_num(0.0, prec))
    if az <= 180.0:
        return u"{}° Leste".format(fmt_num(az, prec))
    return u"{}° Oeste".format(fmt_num(360.0 - az, prec))


def pn_text(rec, prec):
    """Direção do Norte do Projeto em relação ao Norte Verdadeiro."""
    return direction_text(rec["pn_azimuth"], prec)


def tn_text(rec, prec):
    """Direção do Norte Verdadeiro em relação ao Norte do Projeto (inverso)."""
    az = rec["pn_azimuth"]
    return NA if az is None else direction_text((360.0 - az) % 360.0, prec)


def doc_file_name(document):
    """Nome do arquivo (.rvt). Modelos na nuvem podem não ter PathName local."""
    try:
        path = to_unicode(document.PathName)
        if path:
            name = re.split(r"[\\/]", path)[-1]
            if name:
                return name
    except Exception:
        pass
    try:
        return to_unicode(document.Title)
    except Exception:
        return u"(sem nome)"


def friendly_error(message, ex=None):
    """Mensagem amigável + detalhe técnico no console do pyRevit."""
    detail = to_unicode(ex) if ex is not None else u""
    if ex is not None:
        try:
            out = script.get_output()
            out.print_md(u"**{}**".format(message))
            out.print_code(to_unicode(traceback.format_exc()))
        except Exception:
            pass
    forms.alert(message, sub_msg=detail or None, title=TITLE)


# ------------------------------------------------------------------
# Unidades
# ------------------------------------------------------------------
class UnitContext(object):
    """Unidade de comprimento do projeto HOSPEDEIRO (usada em todas as linhas,
    inclusive vínculos, para que a tabela seja comparável).

    A API sempre trabalha em pés (comprimento) e radianos (ângulo). A conversão
    para a unidade do projeto é feita com UnitUtils.ConvertFromInternalUnits.
    Unidades compostas (pés-polegadas, metros-centímetros) são apresentadas em
    decimal da unidade principal.
    """

    def __init__(self, document):
        self.unit_id = None
        self.label = u"pés (unidade interna do Revit)"
        self.notes = []
        try:
            fo = document.GetUnits().GetFormatOptions(DB.SpecTypeId.Length)
            uid = fo.GetUnitTypeId()
            DB.UnitUtils.ConvertFromInternalUnits(1.0, uid)   # testa a conversão
            self.unit_id = uid
            try:
                self.label = to_unicode(DB.LabelUtils.GetLabelForUnit(uid))
            except Exception:
                self.label = to_unicode(uid.TypeId)
            tid = to_unicode(uid.TypeId)
            if "feetFractionalInches" in tid or "feetInches" in tid:
                self.notes.append(u"Unidade do projeto em pés-polegadas: valores "
                                  u"apresentados em pés decimais.")
            elif "metersCentimeters" in tid:
                self.notes.append(u"Unidade do projeto em metros-centímetros: "
                                  u"valores apresentados em metros decimais.")
        except Exception as ex:
            self.notes.append(u"Erro ao ler/converter a unidade de comprimento do "
                              u"projeto ({}). Valores apresentados em pés (unidade "
                              u"interna).".format(to_unicode(ex)))

    def length(self, feet):
        """Pés -> unidade do projeto. None se indisponível ou se falhar."""
        if feet is None:
            return None
        if self.unit_id is None:
            return feet
        try:
            return DB.UnitUtils.ConvertFromInternalUnits(feet, self.unit_id)
        except Exception as ex:
            msg = (u"Erro na conversão de unidades ({}); valor(es) afetado(s) "
                   u"exibido(s) como '-'.".format(to_unicode(ex)))
            if msg not in self.notes:
                self.notes.append(msg)
            return None


# ------------------------------------------------------------------
# Leitura de coordenadas (API)
# ------------------------------------------------------------------
def get_base_point(document, kind):
    """PBP ou SP do documento.

    BasePoint.GetProjectBasePoint / GetSurveyPoint (Revit 2020+) são os
    métodos oficiais. O coletor por categoria fica como reserva.
    """
    try:
        if kind == REF_PBP:
            bp = DB.BasePoint.GetProjectBasePoint(document)
        else:
            bp = DB.BasePoint.GetSurveyPoint(document)
        if bp is not None:
            return bp
    except Exception:
        pass
    bic = (DB.BuiltInCategory.OST_ProjectBasePoint if kind == REF_PBP
           else DB.BuiltInCategory.OST_SharedBasePoint)
    try:
        return (DB.FilteredElementCollector(document).OfCategory(bic)
                .WhereElementIsNotElementType().FirstElement())
    except Exception:
        return None


def shared_of_internal(document, internal_xyz):
    """Ponto em coordenadas internas -> (L/O, N/S, Elev.) compartilhadas,
    calculado pela própria API através do sítio ativo."""
    pp = document.ActiveProjectLocation.GetProjectPosition(internal_xyz)
    return (pp.EastWest, pp.NorthSouth, pp.Elevation)


def read_site_rotation(document):
    """Rotação entre Norte do Projeto e Norte Verdadeiro do sítio ativo.

    - angle_deg: 'Ângulo para o Norte verdadeiro' como o Revit exibe
      (BASEPOINT_ANGLETON_PARAM do PBP). Reserva: ProjectPosition.Angle.
      As duas fontes podem ter convenções de SINAL diferentes, por isso a
      direção (Leste/Oeste) é informada à parte, medida geometricamente.
    - pn_azimuth: azimute horário do Norte do Projeto medido a partir do
      Norte Verdadeiro. Obtido convertendo, pela API, dois pontos sobre o
      eixo +Y interno para coordenadas compartilhadas - não depende de
      convenção de sinal.
    """
    res = {"angle_deg": None, "pn_azimuth": None, "location": None, "notes": []}
    try:
        loc = document.ActiveProjectLocation
        res["location"] = elem_name(loc)
    except Exception as ex:
        res["notes"].append(u"Sítio ativo indisponível ({}).".format(to_unicode(ex)))
        return res

    angle = param_double(get_base_point(document, REF_PBP), BIP.BASEPOINT_ANGLETON_PARAM)
    if angle is not None:
        res["angle_deg"] = math.degrees(angle)
    else:
        try:
            res["angle_deg"] = math.degrees(loc.GetProjectPosition(DB.XYZ.Zero).Angle)
            res["notes"].append(u"Ângulo lido de ProjectPosition.Angle (parâmetro do "
                                u"PBP indisponível); o sinal pode diferir do exibido.")
        except Exception as ex:
            res["notes"].append(u"Ângulo para o Norte verdadeiro indisponível "
                                u"({}).".format(to_unicode(ex)))

    try:
        p0 = shared_of_internal(document, DB.XYZ.Zero)
        p1 = shared_of_internal(document, DB.XYZ(0.0, PROBE_FT, 0.0))
        d_east, d_north = p1[0] - p0[0], p1[1] - p0[1]
        if math.hypot(d_east, d_north) > TOL_FT:
            res["pn_azimuth"] = math.degrees(math.atan2(d_east, d_north)) % 360.0
    except Exception as ex:
        res["notes"].append(u"Direção do Norte não pôde ser medida ({}).".format(
            to_unicode(ex)))
    return res


def read_internal_origin(document):
    """Coordenadas compartilhadas da Origem Interna (InternalOrigin, 2021+)."""
    try:
        io = DB.InternalOrigin.Get(document)
        if io is not None:
            return xyz_tuple(io.SharedPosition)
    except Exception:
        pass
    try:                                    # reserva: conversão pela API
        return shared_of_internal(document, DB.XYZ.Zero)
    except Exception:
        return None


def get_model_name(document):
    """Nome do modelo = parâmetro 'Bloco Nome' das Informações do Projeto."""
    try:
        pinfo = document.ProjectInformation
        for pname in BLOCK_PARAM_NAMES:
            p = pinfo.LookupParameter(pname) if pinfo is not None else None
            if p is None or not p.HasValue:
                continue
            value = p.AsString() if p.StorageType == DB.StorageType.String \
                else p.AsValueString()
            value = to_unicode(value).strip()
            if value:
                return value, None
    except Exception:
        pass
    return doc_file_name(document), (
        u"Parâmetro '{}' não encontrado ou vazio em Informações do Projeto; "
        u"usado o nome do arquivo.".format(BLOCK_PARAM_NAMES[0]))


def new_record(tipo, ref_kind):
    return {
        "model": NA, "file": NA, "tipo": tipo, "ref": REF_LABELS[ref_kind],
        "instance": u"",
        "ns": None, "ew": None, "elev": None,            # pés, valores exibidos
        "angle_deg": None, "pn_azimuth": None, "location": None,
        "internal": None,       # (X, Y, Z) em relação à Origem Interna (pés)
        "shared": None,         # BasePoint.SharedPosition (L/O, N/S, Elev.)
        "origin_shared": None,  # Origem Interna em coordenadas compartilhadas
        "host_shared": None,    # vínculo: ponto nas coord. compart. do HOSPEDEIRO
        "clipped": None,
        "check": u"", "obs": [],
    }


def read_model(document, ref_kind, tipo, link_inst=None, host_doc=None):
    """Lê o ponto de referência escolhido em um documento (hospedeiro ou vínculo)."""
    rec = new_record(tipo, ref_kind)
    rec["model"], name_note = get_model_name(document)
    rec["file"] = doc_file_name(document)
    if name_note:
        rec["obs"].append(name_note)
    if link_inst is not None:
        rec["instance"] = elem_name(link_inst)

    ref_label = REF_LABELS[ref_kind]
    bp = get_base_point(document, ref_kind)
    checks = []
    if bp is None:
        rec["obs"].append(u"{} não encontrado no modelo.".format(ref_label))
    else:
        # Garante que não houve troca PBP <-> SP: o Survey Point é o único
        # BasePoint com IsShared = True.
        try:
            if bool(bp.IsShared) != (ref_kind == REF_SP):
                rec["obs"].append(u"Atenção: o ponto obtido não corresponde ao {} "
                                  u"(IsShared inesperado).".format(ref_label))
        except Exception:
            pass

        # 1) Valores exibidos pelo Revit (coordenadas compartilhadas do ponto)
        rec["ns"] = param_double(bp, BIP.BASEPOINT_NORTHSOUTH_PARAM)
        rec["ew"] = param_double(bp, BIP.BASEPOINT_EASTWEST_PARAM)
        rec["elev"] = param_double(bp, BIP.BASEPOINT_ELEVATION_PARAM)

        # 2) Posição interna e posição compartilhada fornecidas pela API
        try:
            rec["internal"] = xyz_tuple(bp.Position)
        except Exception:
            rec["obs"].append(u"BasePoint.Position indisponível.")
        try:
            rec["shared"] = xyz_tuple(bp.SharedPosition)
        except Exception:
            rec["obs"].append(u"BasePoint.SharedPosition indisponível.")
        try:
            rec["clipped"] = bool(bp.Clipped)
        except Exception:
            pass

        # Reserva: se algum parâmetro não existir, usa SharedPosition.
        if rec["shared"] is not None:
            fallback = []
            for key, idx, label in (("ew", 0, u"L/O"), ("ns", 1, u"N/S"),
                                    ("elev", 2, u"Elevação")):
                if rec[key] is None:
                    rec[key] = rec["shared"][idx]
                    fallback.append(label)
            if fallback:
                rec["obs"].append(u"{} lido(s) de BasePoint.SharedPosition (parâmetro "
                                  u"indisponível).".format(u", ".join(fallback)))

        # 3) Verificações cruzadas (todas com valores da própria API)
        if None not in (rec["ew"], rec["ns"], rec["elev"]) and rec["shared"]:
            diff = max_diff((rec["ew"], rec["ns"], rec["elev"]), rec["shared"])
            checks.append(u"parâmetros = SharedPosition" if diff <= TOL_FT else
                          u"parâmetros ≠ SharedPosition (dif. {:.6f} pés)".format(diff))
        if rec["internal"] and rec["shared"]:
            try:
                conv = shared_of_internal(document, bp.Position)
                diff = max_diff(conv, rec["shared"])
                checks.append(u"GetProjectPosition(Position) = SharedPosition"
                              if diff <= TOL_FT else
                              u"GetProjectPosition(Position) ≠ SharedPosition "
                              u"(dif. {:.6f} pés)".format(diff))
            except Exception:
                pass

        # 4) Vínculo: onde o ponto cai nas coordenadas compartilhadas do
        #    HOSPEDEIRO (link interno -> host interno -> host compartilhado).
        if link_inst is not None and host_doc is not None and rec["internal"]:
            try:
                host_internal = link_inst.GetTotalTransform().OfPoint(bp.Position)
                rec["host_shared"] = shared_of_internal(host_doc, host_internal)
                if rec["shared"]:
                    diff = max_diff(rec["host_shared"], rec["shared"])
                    checks.append(
                        u"coord. compartilhadas coincidem com o hospedeiro"
                        if diff <= TOL_FT else
                        u"coord. compartilhadas NÃO coincidem com o hospedeiro "
                        u"(dif. {:.6f} pés)".format(diff))
            except Exception as ex:
                rec["obs"].append(u"Posição no hospedeiro indisponível ({}).".format(
                    to_unicode(ex)))

    # Rotação é do sítio (sistema de coordenadas), vale para PBP e SP.
    rot = read_site_rotation(document)
    rec["angle_deg"] = rot["angle_deg"]
    rec["pn_azimuth"] = rot["pn_azimuth"]
    rec["location"] = rot["location"]
    rec["obs"].extend(rot["notes"])
    rec["origin_shared"] = read_internal_origin(document)
    rec["check"] = u"; ".join(checks) if checks else NA
    return rec


# ------------------------------------------------------------------
# Modelos vinculados
# ------------------------------------------------------------------
def link_status_text(link_type):
    try:
        raw = str(link_type.GetLinkedFileStatus())
        return raw, LINK_STATUS_PT.get(raw, raw)
    except Exception:
        return None, None


def read_links(host, ref_kind):
    """Uma linha por INSTÂNCIA de vínculo Revit carregado ou não.

    Limitações da API (informadas no relatório, nunca preenchidas com valores
    inventados):
      - vínculo descarregado / não encontrado / em workset fechado:
        GetLinkDocument() devolve None -> nenhum dado do arquivo é acessível;
      - vínculo aninhado: se a API expuser a instância, ela aparece como
        'Vínculo (aninhado)', mas sem a posição no hospedeiro, pois a
        transformação dele é relativa ao vínculo pai, não ao hospedeiro;
      - o vínculo é lido no sítio ATIVO salvo no arquivo vinculado.
    """
    records = []
    try:
        instances = list(DB.FilteredElementCollector(host)
                         .OfClass(DB.RevitLinkInstance).ToElements())
    except Exception as ex:
        rec = new_record(TIPO_LINK, ref_kind)
        rec["obs"].append(u"Não foi possível listar os vínculos ({}).".format(
            to_unicode(ex)))
        return [rec]

    for inst in instances:
        link_type = None
        try:
            link_type = host.GetElement(inst.GetTypeId())
        except Exception:
            pass
        status_raw, status_pt = link_status_text(link_type) if link_type else (None, None)

        link_doc = None
        access_error = None
        try:
            link_doc = inst.GetLinkDocument()
        except Exception as ex:
            access_error = to_unicode(ex)

        if link_doc is None:
            rec = new_record(TIPO_LINK, ref_kind)
            rec["model"] = elem_name(link_type) if link_type else elem_name(inst)
            rec["file"] = rec["model"]
            rec["instance"] = elem_name(inst)
            reason = status_pt or u"documento do vínculo inacessível"
            if access_error:
                reason += u" - " + access_error
            rec["obs"].append(u"Vínculo sem acesso ({}). A API só lê dados de "
                              u"vínculos carregados; nenhum valor foi preenchido."
                              .format(reason))
            rec["check"] = NA
            records.append(rec)
            continue

        nested = False
        try:
            nested = bool(link_type is not None and link_type.IsNestedLink)
        except Exception:
            pass
        try:
            rec = read_model(link_doc, ref_kind, TIPO_LINK,
                             link_inst=None if nested else inst, host_doc=host)
        except Exception as ex:
            rec = new_record(TIPO_LINK, ref_kind)
            rec["model"] = doc_file_name(link_doc)
            rec["file"] = rec["model"]
            rec["obs"].append(u"Erro ao ler o vínculo: {}".format(to_unicode(ex)))
        rec["instance"] = elem_name(inst)
        if nested:
            rec["tipo"] = u"Vínculo (aninhado)"
            rec["obs"].append(u"Vínculo aninhado: posição no hospedeiro não "
                              u"calculada (transformação relativa ao vínculo pai).")
        if status_raw and status_raw != "Loaded":
            rec["obs"].append(u"Status do vínculo: {}.".format(status_pt))
        records.append(rec)

    if not instances:
        rec = new_record(TIPO_LINK, ref_kind)
        rec["model"] = u"(nenhum vínculo Revit)"
        rec["obs"].append(u"O modelo atual não possui instâncias de vínculo Revit.")
        records.append(rec)
    records.sort(key=lambda r: (r["model"].lower(), r["instance"].lower()))
    return records


def collect(host, ref_kind, include_links):
    """Lê o hospedeiro e, se pedido, os vínculos. Nunca propaga exceção."""
    records = []
    try:
        records.append(read_model(host, ref_kind, TIPO_HOST))
    except Exception as ex:
        rec = new_record(TIPO_HOST, ref_kind)
        rec["model"] = doc_file_name(host)
        rec["obs"].append(u"Erro ao ler o modelo atual: {}".format(to_unicode(ex)))
        records.append(rec)
    if include_links:
        records.extend(read_links(host, ref_kind))
    return records


# ------------------------------------------------------------------
# Relatório (explicação + observações)
# ------------------------------------------------------------------
EXPLANATION = [
    u"COMO OS VALORES SÃO OBTIDOS",
    u"• Norte/Sul, Leste/Oeste e Elevação = valores exibidos pelo Revit nas "
    u"propriedades do ponto (BASEPOINT_NORTHSOUTH/EASTWEST/ELEVATION_PARAM). "
    u"São coordenadas COMPARTILHADAS, não a posição XYZ interna.",
    u"• Os valores são conferidos com BasePoint.SharedPosition e com "
    u"ProjectLocation.GetProjectPosition(BasePoint.Position) (coluna Verificação "
    u"da aba Detalhes).",
    u"• Posição interna = BasePoint.Position, relativa à Origem Interna (fixa).",
    u"• Ângulo = 'Ângulo para o Norte verdadeiro' do sítio ativo, lido do "
    u"Project Base Point. É uma propriedade do sistema de coordenadas: vale "
    u"igualmente para PBP e Survey Point.",
    u"• Project North = direção do Norte do Projeto em relação ao Norte "
    u"Verdadeiro. True North = direção do Norte Verdadeiro em relação ao Norte "
    u"do Projeto. Medidos convertendo, pela API, o eixo +Y interno para "
    u"coordenadas compartilhadas (independe da convenção de sinal do ângulo).",
    u"• Vínculos: valores no sistema compartilhado do PRÓPRIO arquivo vinculado "
    u"(sítio ativo salvo nele). A aba Detalhes mostra também onde o ponto cai "
    u"nas coordenadas compartilhadas do hospedeiro, indicando se coincidem.",
    u"• Todos os comprimentos estão na unidade do modelo atual; ângulos em graus "
    u"decimais.",
]


def build_notes(records, units):
    lines = list(EXPLANATION)
    lines.append(u"")
    lines.append(u"UNIDADE: {}".format(units.label))
    lines.extend(u"• " + n for n in units.notes)
    obs = [(r, o) for r in records for o in r["obs"]]
    if obs:
        lines.append(u"")
        lines.append(u"OBSERVAÇÕES / LIMITAÇÕES")
        for r, o in obs:
            who = r["model"] + (u" [{}]".format(r["instance"]) if r["instance"] else u"")
            lines.append(u"• {}: {}".format(who, o))
    return lines


# ------------------------------------------------------------------
# Excel (.xlsx gerado em XML puro, sem Excel/COM)
# ------------------------------------------------------------------
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_INVALID_XML = re.compile(u"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def xml_escape(text):
    text = _INVALID_XML.sub(u"", to_unicode(text))
    return (text.replace(u"&", u"&amp;").replace(u"<", u"&lt;")
            .replace(u">", u"&gt;").replace(u'"', u"&quot;"))


def col_letter(idx):
    """0 -> A, 25 -> Z, 26 -> AA."""
    letters = u""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = unicode(chr(65 + rem)) + letters
    return letters


def sheet_xml(headers, rows, widths):
    """Cabeçalho (estilo 1), números (estilo 2), texto inline. None -> '-'."""
    out = [u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           u'<worksheet xmlns="{}" xmlns:r="{}">'.format(NS_MAIN, NS_REL),
           u'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" '
           u'topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
           u'</sheetView></sheetViews>',
           u'<sheetFormatPr defaultRowHeight="15"/><cols>']
    for i, w in enumerate(widths):
        out.append(u'<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(i + 1, w))
    out.append(u'</cols><sheetData>')
    for r_idx, row in enumerate([headers] + rows):
        rn = r_idx + 1
        out.append(u'<row r="{}">'.format(rn))
        for c_idx, value in enumerate(row):
            ref = col_letter(c_idx) + unicode(rn)
            if r_idx > 0 and isinstance(value, float):
                out.append(u'<c r="{}" s="2"><v>{}</v></c>'.format(ref, repr(value)))
            else:
                text = NA if value is None else value
                style = u' s="1"' if r_idx == 0 else u""
                out.append(u'<c r="{}"{} t="inlineStr"><is><t xml:space="preserve">'
                           u'{}</t></is></c>'.format(ref, style, xml_escape(text)))
        out.append(u'</row>')
    out.append(u'</sheetData></worksheet>')
    return u"".join(out)


def build_xlsx_parts(sheets, prec):
    """sheets = [(nome, cabeçalhos, linhas, larguras)] -> {parte: xml}."""
    num_fmt = u"0." + u"0" * prec
    parts = {}
    parts[u"[Content_Types].xml"] = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        u'<Default Extension="xml" ContentType="application/xml"/>'
        u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        + u"".join(u'<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/'
                   u'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i + 1)
                   for i in range(len(sheets)))
        + u'</Types>')
    parts[u"_rels/.rels"] = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Relationships xmlns="{}"><Relationship Id="rId1" Type="{}/officeDocument" '
        u'Target="xl/workbook.xml"/></Relationships>'.format(NS_PKG, NS_REL))
    parts[u"xl/workbook.xml"] = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<workbook xmlns="{}" xmlns:r="{}"><sheets>'.format(NS_MAIN, NS_REL)
        + u"".join(u'<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(
            xml_escape(s[0]), i + 1, i + 1) for i, s in enumerate(sheets))
        + u'</sheets></workbook>')
    parts[u"xl/_rels/workbook.xml.rels"] = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Relationships xmlns="{}">'.format(NS_PKG)
        + u"".join(u'<Relationship Id="rId{0}" Type="{1}/worksheet" '
                   u'Target="worksheets/sheet{0}.xml"/>'.format(i + 1, NS_REL)
                   for i in range(len(sheets)))
        + u'<Relationship Id="rId{}" Type="{}/styles" Target="styles.xml"/>'.format(
            len(sheets) + 1, NS_REL)
        + u'</Relationships>')
    parts[u"xl/styles.xml"] = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<styleSheet xmlns="{}">'.format(NS_MAIN)
        + u'<numFmts count="1"><numFmt numFmtId="164" formatCode="{}"/></numFmts>'.format(num_fmt)
        + u'<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
          u'<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>'
          u'<fills count="3"><fill><patternFill patternType="none"/></fill>'
          u'<fill><patternFill patternType="gray125"/></fill>'
          u'<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill></fills>'
          u'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          u'<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          u'<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
          u'<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
          u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
          u'</styleSheet>')
    for i, (_, headers, rows, widths) in enumerate(sheets):
        parts[u"xl/worksheets/sheet{}.xml".format(i + 1)] = sheet_xml(headers, rows, widths)
    return parts


def write_xlsx(path, parts):
    """Grava o pacote ZIP. FileMode.Create falha ANTES de truncar se o arquivo
    estiver aberto no Excel, então um arquivo existente nunca é corrompido."""
    # getattr: 'None' é palavra reservada no Python 3 (motor CPython do pyRevit)
    stream = FileStream(path, FileMode.Create, FileAccess.Write,
                        getattr(FileShare, "None"))
    try:
        archive = ZipArchive(stream, ZipArchiveMode.Create)
        try:
            for name in sorted(parts):
                entry = archive.CreateEntry(name)
                writer = StreamWriter(entry.Open(), UTF8Encoding(False))
                try:
                    writer.Write(parts[name])
                finally:
                    writer.Dispose()
        finally:
            archive.Dispose()
    finally:
        stream.Dispose()


def rounded(value, prec):
    return None if value is None else round(float(value), prec)


def build_sheets(records, units, prec, read_time, notes):
    """Monta as 3 abas: Coordenadas (pedida), Detalhes (auditoria) e Notas."""
    date_txt = read_time.strftime("%d/%m/%Y %H:%M:%S")
    L = lambda ft: rounded(units.length(ft), prec)

    main_headers = [u"Modelo", u"Referência", u"Norte/Sul", u"Leste/Oeste", u"Elevação",
                    u"Ângulo", u"Project North", u"True North", u"Unidade",
                    u"Data da leitura", u"Tipo", u"Observação"]
    main_rows = []
    for r in records:
        main_rows.append([
            r["model"], r["ref"], L(r["ns"]), L(r["ew"]), L(r["elev"]),
            rounded(r["angle_deg"], prec),
            pn_text(r, prec), tn_text(r, prec),
            units.label, date_txt, r["tipo"], u" | ".join(r["obs"]) or None])

    def trio(t):
        return [None, None, None] if t is None else [L(t[0]), L(t[1]), L(t[2])]

    det_headers = [u"Modelo", u"Arquivo", u"Instância do vínculo", u"Tipo", u"Referência",
                   u"Sítio ativo", u"X interno", u"Y interno (Norte do Projeto)", u"Z interno",
                   u"SharedPosition L/O", u"SharedPosition N/S", u"SharedPosition Elev.",
                   u"Origem interna L/O", u"Origem interna N/S", u"Origem interna Elev.",
                   u"No hospedeiro L/O", u"No hospedeiro N/S", u"No hospedeiro Elev.",
                   u"Survey Point clipped", u"Verificação", u"Unidade"]
    det_rows = []
    for r in records:
        internal = r["internal"]
        clipped = None
        if r["clipped"] is not None and r["ref"] == REF_LABELS[REF_SP]:
            clipped = u"Sim" if r["clipped"] else u"Não"
        det_rows.append([r["model"], r["file"], r["instance"] or None, r["tipo"], r["ref"],
                         r["location"]] + trio(internal) + trio(r["shared"])
                        + trio(r["origin_shared"]) + trio(r["host_shared"])
                        + [clipped, r["check"], units.label])

    note_rows = [[line or u" "] for line in notes]
    return [
        (u"Coordenadas", main_headers, main_rows,
         [30, 20, 14, 14, 14, 12, 18, 18, 18, 20, 18, 60]),
        (u"Detalhes", det_headers, det_rows,
         [30, 30, 30, 18, 20, 20] + [16] * 12 + [12, 60, 18]),
        (u"Notas", [u"Notas"], note_rows, [140]),
    ]


# ------------------------------------------------------------------
# Interface (WPF - padrão visual OCA, lib/oca_ui)
# ------------------------------------------------------------------
BODY = u"""
  <Grid>
    <Grid.Resources>
      <Style x:Key="Num" TargetType="TextBlock">
        <Setter Property="HorizontalAlignment" Value="Right"/>
        <Setter Property="Margin" Value="6,0"/>
      </Style>
      <Style x:Key="Txt" TargetType="TextBlock">
        <Setter Property="Margin" Value="6,0"/>
      </Style>
    </Grid.Resources>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="130"/>
    </Grid.RowDefinitions>

    <Grid Grid.Row="0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="32"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="32"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <HeaderedContentControl Grid.Column="0" Header="PONTO DE REFERÊNCIA" Style="{StaticResource oca.Section}">
        <StackPanel>
          <RadioButton x:Name="rb_pbp" GroupName="ref" IsChecked="True" Content="Project Base Point"/>
          <RadioButton x:Name="rb_sp" GroupName="ref" Content="Survey Point"/>
        </StackPanel>
      </HeaderedContentControl>
      <HeaderedContentControl Grid.Column="2" Header="LER MODELOS VINCULADOS" Style="{StaticResource oca.Section}">
        <StackPanel>
          <RadioButton x:Name="rb_links_yes" GroupName="links" Content="Sim"/>
          <RadioButton x:Name="rb_links_no" GroupName="links" IsChecked="True" Content="Não"/>
        </StackPanel>
      </HeaderedContentControl>
      <HeaderedContentControl Grid.Column="4" Header="PRECISÃO" Style="{StaticResource oca.Section}">
        <ComboBox x:Name="cb_prec" Width="160" HorizontalAlignment="Left">
          <ComboBoxItem Content="2 casas decimais"/>
          <ComboBoxItem Content="3 casas decimais"/>
          <ComboBoxItem Content="4 casas decimais"/>
        </ComboBox>
      </HeaderedContentControl>
      <Button x:Name="btn_read" Grid.Column="6" Content="Ler coordenadas" VerticalAlignment="Bottom" Margin="0,0,0,14"/>
    </Grid>

    <DataGrid x:Name="grid_coords" Grid.Row="1" AutoGenerateColumns="False"
              IsReadOnly="True" CanUserAddRows="False" SelectionMode="Extended" RowHeight="24">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Modelo" Binding="{Binding model}" Width="2*"
                            ElementStyle="{StaticResource Txt}"/>
        <DataGridTextColumn Header="Tipo" Binding="{Binding tipo}" Width="Auto"
                            ElementStyle="{StaticResource Txt}"/>
        <DataGridTextColumn Header="Referência" Binding="{Binding ref}" Width="Auto"
                            ElementStyle="{StaticResource Txt}"/>
        <DataGridTextColumn Header="Norte/Sul" Binding="{Binding ns}" Width="*"
                            ElementStyle="{StaticResource Num}" FontFamily="Consolas"/>
        <DataGridTextColumn Header="Leste/Oeste" Binding="{Binding ew}" Width="*"
                            ElementStyle="{StaticResource Num}" FontFamily="Consolas"/>
        <DataGridTextColumn Header="Elevação" Binding="{Binding elev}" Width="*"
                            ElementStyle="{StaticResource Num}" FontFamily="Consolas"/>
        <DataGridTextColumn Header="Ângulo" Binding="{Binding angle}" Width="Auto"
                            ElementStyle="{StaticResource Num}" FontFamily="Consolas"/>
        <DataGridTextColumn Header="Project North" Binding="{Binding pn}" Width="Auto"
                            ElementStyle="{StaticResource Num}"/>
        <DataGridTextColumn Header="True North" Binding="{Binding tn}" Width="Auto"
                            ElementStyle="{StaticResource Num}"/>
        <DataGridTextColumn Header="Obs." Binding="{Binding obs}" Width="2*"
                            ElementStyle="{StaticResource Txt}"/>
      </DataGrid.Columns>
    </DataGrid>

    <Border Grid.Row="2" Style="{StaticResource oca.Msg.Info}" Margin="0,10,0,8"
            Visibility="{Binding Visibility, ElementName=tb_status}">
      <TextBlock x:Name="tb_status" Style="{StaticResource oca.MsgText}"
                 Text="Escolha as opções e clique em 'Ler coordenadas'."/>
    </Border>

    <TextBox x:Name="tb_notes" Grid.Row="3" IsReadOnly="True" TextWrapping="Wrap" Height="Auto"
             VerticalContentAlignment="Top" VerticalScrollBarVisibility="Auto" Padding="8,6"
             FontSize="11" Foreground="{StaticResource oca.Ink2}"/>
  </Grid>"""

FOOTER = u"""
  <Button x:Name="btn_export" Content="Exportar para Excel" Style="{StaticResource oca.Primary}"/>
  <Button x:Name="btn_close" Content="Fechar" Margin="8,0,0,0"/>"""

XAML = build_xaml(title=__title__, subtitle=__doc__, body=BODY, footer_right=FOOTER,
                  size="L", height=680, width=1040)


GRID_KEYS = ["model", "tipo", "ref", "ns", "ew", "elev", "angle", "pn", "tn", "obs"]
TIPO_COLUMN_INDEX = 1


class CoordinatesWindow(forms.WPFWindow):
    def __init__(self, document):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)
        self.document = document
        self.units = UnitContext(document)
        self.records = []
        self.read_time = None
        self.read_options = None

        # DataTable com nomes de coluna simples: "Norte/Sul" não pode ser usado
        # como caminho de Binding (a barra tem significado no WPF).
        self.table = DataTable("coordenadas")
        for key in GRID_KEYS:
            self.table.Columns.Add(key, String)
        self.grid_coords.ItemsSource = self.table.DefaultView
        self.grid_coords.Columns[TIPO_COLUMN_INDEX].Visibility = Visibility.Collapsed

        self.cb_prec.SelectedIndex = PRECISIONS.index(DEFAULT_PRECISION)
        self.tb_notes.Text = u"\n".join(EXPLANATION)

        self.btn_read.Click += self._on_read
        self.btn_export.Click += self._on_export
        self.btn_close.Click += self._on_close
        self.cb_prec.SelectionChanged += self._on_precision

    # -- opções ------------------------------------------------------
    @property
    def precision(self):
        idx = self.cb_prec.SelectedIndex
        return PRECISIONS[idx] if 0 <= idx < len(PRECISIONS) else DEFAULT_PRECISION

    def _options(self):
        return (REF_SP if self.rb_sp.IsChecked else REF_PBP,
                bool(self.rb_links_yes.IsChecked))

    # -- leitura -----------------------------------------------------
    def _read(self):
        ref_kind, include_links = self._options()
        self.records = collect(self.document, ref_kind, include_links)
        self.read_time = datetime.now()
        self.read_options = (ref_kind, include_links)
        self.grid_coords.Columns[TIPO_COLUMN_INDEX].Visibility = (
            Visibility.Visible if include_links else Visibility.Collapsed)
        self._render()

    def _render(self):
        prec = self.precision
        units = self.units
        self.table.Rows.Clear()
        for r in self.records:
            row = self.table.NewRow()
            row["model"] = r["model"]
            row["tipo"] = r["tipo"]
            row["ref"] = r["ref"]
            row["ns"] = fmt_num(units.length(r["ns"]), prec)
            row["ew"] = fmt_num(units.length(r["ew"]), prec)
            row["elev"] = fmt_num(units.length(r["elev"]), prec)
            row["angle"] = fmt_deg(r["angle_deg"], prec)
            row["pn"] = pn_text(r, prec)
            row["tn"] = tn_text(r, prec)
            row["obs"] = u" | ".join(r["obs"]) or u""
            self.table.Rows.Add(row)
        self.tb_notes.Text = u"\n".join(build_notes(self.records, self.units))
        if self.read_time is not None:
            n_obs = sum(1 for r in self.records if r["obs"])
            self.tb_status.Text = (
                u"{} linha(s) lida(s) em {} | Unidade: {} | {} casas decimais{}".format(
                    len(self.records), self.read_time.strftime("%d/%m/%Y %H:%M:%S"),
                    units.label, prec,
                    u" | {} linha(s) com observações - veja abaixo".format(n_obs)
                    if n_obs else u""))

    # -- eventos -----------------------------------------------------
    def _on_read(self, sender, args):
        try:
            self._read()
        except Exception as ex:
            friendly_error(u"Não foi possível ler as coordenadas.", ex)

    def _on_precision(self, sender, args):
        if self.records:
            try:
                self._render()
            except Exception as ex:
                friendly_error(u"Erro ao atualizar a precisão.", ex)

    def _on_export(self, sender, args):
        try:
            # Lê de novo se ainda não leu ou se as opções mudaram desde a leitura.
            if not self.records or self.read_options != self._options():
                self._read()
            if not self.records:
                forms.alert(u"Nenhum dado para exportar.", title=TITLE)
                return
            path = ask_save_path(self.document)
            if not path:
                forms.alert(u"Exportação cancelada pelo usuário.", title=TITLE)
                return
            prec = self.precision
            sheets = build_sheets(self.records, self.units, prec, self.read_time,
                                  build_notes(self.records, self.units))
            parts = build_xlsx_parts(sheets, prec)
        except Exception as ex:
            friendly_error(u"Erro ao preparar a exportação.", ex)
            return
        try:
            write_xlsx(path, parts)
        except Exception as ex:
            friendly_error(u"Não foi possível gravar o arquivo Excel.\n\nSe ele estiver "
                           u"aberto no Excel, feche-o e tente novamente.", ex)
            return
        if forms.alert(u"Arquivo exportado com sucesso. Abrir agora?", sub_msg=path,
                       title=TITLE, yes=True, no=True):
            try:
                os.startfile(path)
            except Exception as ex:
                friendly_error(u"O arquivo foi gravado, mas não pôde ser aberto.", ex)

    def _on_close(self, sender, args):
        self.Close()


def ask_save_path(document):
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    base = re.sub(r'[\\/:*?"<>|]+', u"_", doc_file_name(document))
    base = re.sub(r"\.rvt$", u"", base, flags=re.IGNORECASE)
    kw = {"file_ext": "xlsx", "default_name": u"Coordenadas_{}_{}".format(base, stamp)}
    try:
        path = forms.save_file(title=u"Salvar coordenadas em Excel", **kw)
    except TypeError:                           # pyRevit antigo sem 'title'
        path = forms.save_file(**kw)
    if path and not to_unicode(path).lower().endswith(".xlsx"):
        path = to_unicode(path) + u".xlsx"
    return path


# ------------------------------------------------------------------
# Main
# ------------------------------------------------------------------
def main():
    if doc is None:
        forms.alert(u"Abra um modelo do Revit antes de executar o comando.", title=TITLE)
        return
    if doc.IsFamilyDocument:
        forms.alert(u"Este comando funciona apenas em modelos de projeto.",
                    sub_msg=u"Famílias não têm Project Base Point/Survey Point.", title=TITLE)
        return
    try:
        CoordinatesWindow(doc).ShowDialog()
    except Exception as error:
        friendly_error(u"Erro inesperado ao abrir a janela de coordenadas.", error)


main()
