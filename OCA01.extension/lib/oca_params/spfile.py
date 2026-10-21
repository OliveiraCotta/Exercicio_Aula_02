# -*- coding: utf-8 -*-
"""Leitura do arquivo TXT de parâmetros compartilhados do Revit.

O arquivo é SOMENTE LIDO: este módulo nunca grava, renomeia nem cria arquivos.
Ele serve para mostrar o conteúdo na tabela e para apontar problemas antes de
o Revit abrir o arquivo. As definições usadas na vinculação vêm sempre do
próprio Revit (Application.OpenSharedParameterFile), pelo GUID.

Formato gravado pelo Revit (colunas separadas por TAB; '#' = comentário):

    # This is a Revit shared parameter file.
    # Do not edit manually.
    *META   VERSION MINVERSION
    META    2       1
    *GROUP  ID      NAME
    GROUP   1       Identificação
    *PARAM  GUID    NAME    DATATYPE  DATACATEGORY  GROUP  VISIBLE  DESCRIPTION  USERMODIFIABLE  HIDEWHENNOVALUE
    PARAM   <guid>  Nome    TEXT                    1      1        texto        1               0

- As colunas de PARAM são lidas pelo cabeçalho *PARAM, não pela posição:
  arquivos de versões antigas têm menos colunas (sem DESCRIPTION,
  USERMODIFIABLE ou HIDEWHENNOVALUE).
- DATATYPE: nome legado (TEXT, LENGTH, YESNO...) ou, em especificações que
  só existem nas versões novas, o identificador da especificação
  ("autodesk.spec.aec:...-1.0.0"). Os dois são aceitos.
- Codificação: o Revit grava em UTF-16 LE com BOM. Arquivos editados à mão
  podem estar em UTF-8 (com ou sem BOM) ou ANSI (cp1252). A codificação é
  detectada e informada.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3, sem API do Revit
(testável fora do Revit).
"""

import re
import codecs

try:
    unicode
except NameError:           # IronPython 3 / CPython
    unicode = str

# colunas padrão quando o arquivo não traz o cabeçalho *PARAM / *GROUP
DEFAULT_PARAM_COLUMNS = [u"GUID", u"NAME", u"DATATYPE", u"DATACATEGORY", u"GROUP",
                         u"VISIBLE", u"DESCRIPTION", u"USERMODIFIABLE", u"HIDEWHENNOVALUE"]
DEFAULT_GROUP_COLUMNS = [u"ID", u"NAME"]
DEFAULT_META_COLUMNS = [u"VERSION", u"MINVERSION"]

_GUID_RE = re.compile(u"^\\{?([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
                      u"[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\\}?$")

# rótulos em português dos tipos de dado legados (quando o Revit não informar o seu)
DATATYPE_LABELS = {
    u"TEXT": u"Texto",
    u"MULTILINETEXT": u"Texto multilinha",
    u"INTEGER": u"Inteiro",
    u"NUMBER": u"Número",
    u"LENGTH": u"Comprimento",
    u"AREA": u"Área",
    u"VOLUME": u"Volume",
    u"ANGLE": u"Ângulo",
    u"SLOPE": u"Inclinação",
    u"CURRENCY": u"Moeda",
    u"MASS_DENSITY": u"Densidade de massa",
    u"URL": u"URL",
    u"MATERIAL": u"Material",
    u"YESNO": u"Sim/Não",
    u"IMAGE": u"Imagem",
    u"FAMILYTYPE": u"Tipo de família",
    u"LOADCLASSIFICATION": u"Classificação de carga",
    u"NUMBER_OF_POLES": u"Número de polos",
}


class SPFileError(Exception):
    """O arquivo não pôde ser lido (inexistente, vazio ou sem nenhuma linha válida)."""


class ParamDef(object):
    """Uma linha PARAM do arquivo, exatamente como está escrita (nada é corrigido)."""

    def __init__(self, line_no):
        self.line_no = line_no
        self.guid = u""             # como está no arquivo
        self.guid_key = u""         # minúsculo, sem chaves (comparação); vazio se inválido
        self.name = u""
        self.datatype = u""
        self.datacategory = u""
        self.group_id = u""
        self.group_name = u""
        self.visible = True
        self.description = u""
        self.user_modifiable = True
        self.hide_when_no_value = False
        self.errors = []            # impedem o uso (bloqueiam)
        self.warnings = []          # informativos

    @property
    def usable(self):
        return not self.errors

    @property
    def datatype_label(self):
        return datatype_label(self.datatype, self.datacategory)

    def __repr__(self):
        return u"<ParamDef {} {}>".format(self.name, self.guid)


class SPFile(object):
    def __init__(self, path):
        self.path = path
        self.encoding = u""
        self.version = u""
        self.min_version = u""
        self.groups = {}            # id (texto) -> nome
        self.group_order = []       # ids na ordem do arquivo
        self.params = []            # [ParamDef] na ordem do arquivo
        self.issues = []            # problemas gerais do arquivo (texto)

    @property
    def usable_count(self):
        return len([p for p in self.params if p.usable])

    def by_guid(self):
        return dict((p.guid_key, p) for p in self.params if p.guid_key)


def datatype_label(datatype, datacategory=u""):
    """'LENGTH' -> 'Comprimento'; 'autodesk.spec.aec:length-2.0.0' -> 'length';
    tipo de família com categoria -> 'Tipo de família (-2000011)'."""
    raw = (datatype or u"").strip()
    if not raw:
        return u""
    up = raw.upper()
    if up in DATATYPE_LABELS:
        label = DATATYPE_LABELS[up]
        if up == u"FAMILYTYPE" and datacategory:
            label = u"{} ({})".format(label, datacategory)
        return label
    if u":" in raw and raw.lower().startswith(u"autodesk."):
        tail = raw.split(u":", 1)[1]
        return re.sub(u"-\\d+\\.\\d+\\.\\d+$", u"", tail)
    return raw


def normalize_guid(text):
    """GUID minúsculo sem chaves, ou u'' se o texto não for um GUID válido."""
    m = _GUID_RE.match((text or u"").strip())
    return m.group(1).lower() if m else u""


def decode_bytes(raw):
    """(texto, nome da codificação). Ordem: BOM > UTF-16 sem BOM > UTF-8 > cp1252."""
    for bom, enc, label in ((codecs.BOM_UTF8, "utf-8", u"UTF-8 com BOM"),
                            (codecs.BOM_UTF16_LE, "utf-16-le", u"UTF-16 LE (Unicode)"),
                            (codecs.BOM_UTF16_BE, "utf-16-be", u"UTF-16 BE")):
        if raw.startswith(bom):
            return raw[len(bom):].decode(enc), label
    # UTF-16 sem BOM: um byte nulo a cada dois caracteres ASCII
    if len(raw) >= 4 and raw[1:2] == b"\x00" and raw[3:4] == b"\x00":
        return raw.decode("utf-16-le"), u"UTF-16 LE sem BOM"
    if len(raw) >= 4 and raw[0:1] == b"\x00" and raw[2:3] == b"\x00":
        return raw.decode("utf-16-be"), u"UTF-16 BE sem BOM"
    try:
        return raw.decode("utf-8"), u"UTF-8"
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode("cp1252"), u"ANSI (Windows-1252)"
    except (UnicodeDecodeError, LookupError):
        return raw.decode("latin-1"), u"ANSI (Latin-1)"


def _flag(value, default):
    """'1'/'0' do arquivo -> bool; vazio ou outro valor -> default."""
    value = (value or u"").strip()
    if value == u"1":
        return True
    if value == u"0":
        return False
    return default


def _row(columns, cells):
    """Associa as células às colunas do cabeçalho (células a mais são ignoradas)."""
    out = {}
    for idx, col in enumerate(columns):
        out[col] = cells[idx] if idx < len(cells) else u""
    return out


def read_file(path):
    """Lê o arquivo do disco e devolve SPFile. Lança SPFileError se não der para ler."""
    try:
        # open() nativo: str no IronPython 2.7 / bytes no Python 3, igual aos BOMs de codecs
        with open(path, "rb") as fh:
            raw = fh.read()
    except (IOError, OSError) as ex:
        raise SPFileError(u"Não foi possível abrir o arquivo: {}".format(ex))
    if not raw:
        raise SPFileError(u"O arquivo está vazio.")
    text, encoding = decode_bytes(raw)
    spf = parse_text(text, path)
    spf.encoding = encoding
    return spf


def parse_text(text, path=u""):
    """Interpreta o conteúdo do TXT. Nada é corrigido: problemas viram avisos
    (warnings) ou erros (errors) na própria definição ou em SPFile.issues."""
    spf = SPFile(path)
    columns = {u"PARAM": list(DEFAULT_PARAM_COLUMNS),
               u"GROUP": list(DEFAULT_GROUP_COLUMNS),
               u"META": list(DEFAULT_META_COLUMNS)}
    seen_header = set()
    unknown_lines = []
    group_rows = []

    for line_no, line in enumerate(text.splitlines(), 1):
        line = line.rstrip(u"\r\n").replace(u"\x00", u"")
        if not line.strip() or line.lstrip().startswith(u"#"):
            continue
        cells = line.split(u"\t")
        tag = cells[0].strip()
        if tag.startswith(u"*"):
            key = tag[1:].upper()
            columns[key] = [c.strip().upper() for c in cells[1:]]
            seen_header.add(key)
            continue
        tag = tag.upper()
        data = cells[1:]
        if tag == u"META":
            row = _row(columns[u"META"], data)
            spf.version = row.get(u"VERSION", u"").strip()
            spf.min_version = row.get(u"MINVERSION", u"").strip()
        elif tag == u"GROUP":
            row = _row(columns[u"GROUP"], data)
            group_rows.append((line_no, row.get(u"ID", u"").strip(), row.get(u"NAME", u"")))
        elif tag == u"PARAM":
            spf.params.append(_param(line_no, _row(columns[u"PARAM"], data)))
        else:
            unknown_lines.append(line_no)

    for line_no, gid, gname in group_rows:
        if not gid:
            spf.issues.append(u"Linha {}: grupo sem ID.".format(line_no))
            continue
        if gid in spf.groups:
            spf.issues.append(u"Linha {}: ID de grupo {} repetido ('{}' e '{}').".format(
                line_no, gid, spf.groups[gid], gname))
            continue
        spf.groups[gid] = gname
        spf.group_order.append(gid)

    _validate(spf)

    if not spf.params and not spf.groups:
        raise SPFileError(u"Nenhum grupo ou parâmetro foi encontrado. Confira se o arquivo é "
                          u"um TXT de parâmetros compartilhados do Revit (colunas separadas por TAB).")
    if u"PARAM" not in seen_header and spf.params:
        spf.issues.append(u"Sem o cabeçalho *PARAM: as colunas foram lidas na ordem padrão do Revit.")
    if spf.version and spf.version not in (u"1", u"2"):
        spf.issues.append(u"Versão do formato ({}) diferente da esperada (2). "
                          u"Confira se o Revit em uso abre o arquivo.".format(spf.version))
    if unknown_lines:
        spf.issues.append(u"{} linha(s) não reconhecida(s) foram ignoradas (linhas {}).".format(
            len(unknown_lines), u", ".join(unicode(n) for n in unknown_lines[:10])
            + (u"..." if len(unknown_lines) > 10 else u"")))
    return spf


def _param(line_no, row):
    p = ParamDef(line_no)
    p.guid = row.get(u"GUID", u"").strip()
    p.guid_key = normalize_guid(p.guid)
    p.name = row.get(u"NAME", u"")
    p.datatype = row.get(u"DATATYPE", u"").strip()
    p.datacategory = row.get(u"DATACATEGORY", u"").strip()
    p.group_id = row.get(u"GROUP", u"").strip()
    p.visible = _flag(row.get(u"VISIBLE"), True)
    p.description = row.get(u"DESCRIPTION", u"")
    p.user_modifiable = _flag(row.get(u"USERMODIFIABLE"), True)
    p.hide_when_no_value = _flag(row.get(u"HIDEWHENNOVALUE"), False)
    return p


def _validate(spf):
    """Erros (bloqueiam o uso) e avisos por definição. Nada é alterado."""
    by_guid = {}
    by_group_name = {}
    for p in spf.params:
        if not p.guid:
            p.errors.append(u"GUID vazio.")
        elif not p.guid_key:
            p.errors.append(u"GUID inválido: '{}'.".format(p.guid))
        else:
            by_guid.setdefault(p.guid_key, []).append(p)
        if not p.name.strip():
            p.errors.append(u"Nome vazio.")
        elif p.name != p.name.strip():
            p.warnings.append(u"O nome tem espaços no início ou no fim.")
        if not p.datatype:
            p.errors.append(u"Tipo de dado vazio.")
        if p.group_id in spf.groups:
            p.group_name = spf.groups[p.group_id]
        else:
            p.group_name = u"(grupo {} inexistente)".format(p.group_id or u"?")
            p.errors.append(u"O grupo {} não existe no arquivo.".format(p.group_id or u"(vazio)"))
        if not p.visible:
            p.warnings.append(u"Definido como oculto (VISIBLE = 0): não aparece na paleta Propriedades.")
        if not p.user_modifiable:
            p.warnings.append(u"Definido como não editável pelo usuário (USERMODIFIABLE = 0).")
        if p.name.strip():
            by_group_name.setdefault((p.group_id, p.name.strip()), []).append(p)

    for key, items in by_guid.items():
        if len(items) > 1:
            lines = u", ".join(unicode(i.line_no) for i in items)
            for p in items:
                p.errors.append(u"GUID repetido no arquivo (linhas {}).".format(lines))
    for key, items in by_group_name.items():
        if len(items) > 1:
            lines = u", ".join(unicode(i.line_no) for i in items)
            for p in items:
                p.warnings.append(u"Nome repetido no mesmo grupo (linhas {}).".format(lines))
