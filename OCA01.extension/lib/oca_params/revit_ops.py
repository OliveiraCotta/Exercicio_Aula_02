# -*- coding: utf-8 -*-
"""Tudo do Gerenciador de Parâmetros Compartilhados que usa a API do Revit.

Diferenças de versão tratadas aqui (detecção por recurso, não por número):

- Grupo de propriedades
    Revit 2022+  : ForgeTypeId (GroupTypeId), ParameterUtils.GetAllBuiltInGroups(),
                   LabelUtils.GetLabelForGroup(), BindingMap.Insert(def, binding, ForgeTypeId),
                   InternalDefinition.GetGroupTypeId().
    Revit <= 2023: BuiltInParameterGroup (removido no 2024+),
                   LabelUtils.GetLabelFor(BuiltInParameterGroup), Insert(def, binding, BIPG),
                   InternalDefinition.ParameterGroup.
    No 2022/2023 existem os dois: usa-se ForgeTypeId.
    "Outros" = ForgeTypeId vazio (2022+) ou BuiltInParameterGroup.INVALID.
- Tipo de dado: ExternalDefinition.GetDataType() + LabelUtils.GetLabelForSpec()
  (2022+) ou ExternalDefinition.ParameterType (versões antigas). Só exibição.
- ElementId: .Value (2024+) ou .IntegerValue.

Arquivo de parâmetros compartilhados: a API só abre o arquivo configurado em
Application.SharedParametersFilename (configuração do usuário no Revit). O
comando troca esse caminho pelo do arquivo escolhido SÓ enquanto lê as
definições e enquanto vincula, e devolve o caminho original logo depois
(também em caso de erro). O arquivo é só aberto para leitura: nenhuma
definição é criada nele (Definitions.Create nunca é chamado), então nenhum
GUID novo é gerado.

Vinculação (uma única transação, tudo ou nada):
- create: BindingMap.Insert(ExternalDefinition, InstanceBinding|TypeBinding, grupo).
- extend: BindingMap.ReInsert(definição interna, mesmo tipo de vínculo com as
  categorias atuais + as novas, grupo atual). Antes de confirmar, valores já
  preenchidos de uma amostra de elementos são relidos; se algum mudar ou sumir,
  tudo é desfeito (RollBack).
- Qualquer recusa do Revit (Insert/ReInsert = False), exceção ou Commit que
  não termine em Committed -> RollBack: o projeto fica como estava.
- Nunca chama BindingMap.Remove, Document.Delete nem Definitions.Create.

Engine: IronPython 2.7 (pyRevit).
"""

import System
from Autodesk.Revit import DB

from oca_params import plan as P

try:
    unicode
except NameError:           # IronPython 3
    unicode = str

TRANSACTION_NAME = u"OCA · Gerenciador de Parâmetros Compartilhados"
VERIFY_LIMIT = 300          # elementos conferidos por parâmetro ampliado (extend)

# categorias sem tipos: um parâmetro de Tipo vinculado a elas não aparece
_NO_TYPE_BICS = ("OST_Rooms", "OST_Areas", "OST_MEPSpaces", "OST_HVAC_Zones",
                 "OST_ProjectInformation", "OST_Sheets", "OST_Materials")

CATEGORY_TYPE_LABELS = (("Model", u"Modelo"), ("Annotation", u"Anotação"),
                        ("AnalyticalModel", u"Analítico"))


class ApplyError(Exception):
    """Falha que desfez a transação inteira (mensagem já em português)."""


# ------------------------------------------------------------------
# utilidades
# ------------------------------------------------------------------
def to_unicode(value):
    if value is None:
        return u""
    if isinstance(value, unicode):
        return value
    try:
        return unicode(value)
    except Exception:
        return u"{!r}".format(value)


def eid_int(eid):
    """ElementId -> int (Revit 2024+ usa .Value, versões antigas .IntegerValue)."""
    raw = eid.Value if hasattr(eid, "Value") else eid.IntegerValue
    return int(raw)


def guid_key(guid):
    return to_unicode(guid.ToString()).lower()


def revit_version(app):
    try:
        return int(app.VersionNumber)
    except Exception:
        return 0


def uses_forge_groups():
    """True no Revit 2022+ (grupos por ForgeTypeId)."""
    return (hasattr(DB, "GroupTypeId") and hasattr(DB, "ParameterUtils")
            and hasattr(DB.ParameterUtils, "GetAllBuiltInGroups"))


# ------------------------------------------------------------------
# arquivo de parâmetros compartilhados (troca temporária do caminho)
# ------------------------------------------------------------------
class SharedParamFile(object):
    """with SharedParamFile(app, path) as dfile: ...

    Aponta o Revit para o arquivo escolhido, abre (DefinitionFile) e, na
    saída, devolve o caminho original do usuário, com ou sem erro.
    """

    def __init__(self, app, path):
        self.app = app
        self.path = path
        self.original = None
        self.changed = False

    def __enter__(self):
        self.original = self.app.SharedParametersFilename
        if (self.original or u"") != self.path:
            self.app.SharedParametersFilename = self.path
            self.changed = True
        try:
            dfile = self.app.OpenSharedParameterFile()
        except Exception as ex:
            self._restore()
            raise ApplyError(u"O Revit não conseguiu abrir o arquivo de parâmetros "
                             u"compartilhados: {}".format(to_unicode(ex)))
        if dfile is None:
            self._restore()
            raise ApplyError(u"O Revit não conseguiu abrir o arquivo de parâmetros compartilhados "
                             u"(arquivo inexistente ou em formato que esta versão não reconhece).")
        return dfile

    def _restore(self):
        if self.changed:
            try:
                self.app.SharedParametersFilename = self.original or u""
            except Exception:
                pass
            self.changed = False

    def __exit__(self, exc_type, exc, tb):
        self._restore()
        return False


def _external_definitions(dfile):
    """guid -> ExternalDefinition de todos os grupos do DefinitionFile."""
    out = {}
    for group in dfile.Groups:
        for definition in group.Definitions:
            try:
                out[guid_key(definition.GUID)] = definition
            except Exception:
                continue
    return out


def spec_label(definition):
    """Rótulo do tipo de dado no idioma do Revit, ou u'' se não der para saber."""
    try:
        spec = definition.GetDataType()                     # Revit 2022+
        try:
            return to_unicode(DB.LabelUtils.GetLabelForSpec(spec))
        except Exception:
            return to_unicode(spec.TypeId)
    except Exception:
        pass
    try:
        return to_unicode(DB.LabelUtils.GetLabelFor(definition.ParameterType))
    except Exception:
        return u""


def read_definitions(app, path):
    """Lê o arquivo pelo Revit. Devolve (guid -> info, erro ou None).

    info = {"usable", "reason", "revit_name", "visible", "modifiable", "datatype"}
    Nenhum objeto do Revit sai daqui: o arquivo é fechado e o caminho do usuário
    é restaurado antes do retorno.
    """
    out = {}
    try:
        with SharedParamFile(app, path) as dfile:
            for key, definition in _external_definitions(dfile).items():
                info = {"usable": True, "reason": u"",
                        "revit_name": to_unicode(definition.Name),
                        "datatype": spec_label(definition)}
                try:
                    info["visible"] = bool(definition.Visible)
                except Exception:
                    info["visible"] = None
                try:
                    info["modifiable"] = bool(definition.UserModifiable)
                except Exception:
                    info["modifiable"] = None
                out[key] = info
    except ApplyError as ex:
        return out, to_unicode(ex)
    except Exception as ex:
        return out, u"Erro ao ler o arquivo pelo Revit: {}".format(to_unicode(ex))
    return out, None


# ------------------------------------------------------------------
# grupos de propriedades (paleta Propriedades)
# ------------------------------------------------------------------
def list_property_groups():
    """[(chave, rótulo)] em ordem alfabética, válidos para a versão em execução.

    Chave: TypeId do ForgeTypeId (2022+; u"" = Outros) ou nome do
    BuiltInParameterGroup (versões antigas).
    """
    items = []
    if uses_forge_groups():
        seen = set()
        for gid in DB.ParameterUtils.GetAllBuiltInGroups():
            key = to_unicode(gid.TypeId)
            if key in seen:
                continue
            seen.add(key)
            try:
                label = to_unicode(DB.LabelUtils.GetLabelForGroup(gid))
            except Exception:
                continue
            if label.strip():
                items.append((key, label))
        if u"" not in seen:
            try:
                other = to_unicode(DB.LabelUtils.GetLabelForGroup(DB.ForgeTypeId(u"")))
            except Exception:
                other = u""
            items.append((u"", other.strip() or u"Outros"))
    else:
        for name in System.Enum.GetNames(DB.BuiltInParameterGroup):
            value = getattr(DB.BuiltInParameterGroup, name)
            try:
                label = to_unicode(DB.LabelUtils.GetLabelFor(value))
            except Exception:
                label = u""
            if name == "INVALID":
                label = label.strip() or u"Outros"
            if label.strip():
                items.append((to_unicode(name), label))
    # rótulos repetidos (grupos diferentes com o mesmo nome traduzido): mostra a chave
    counts = {}
    for _, label in items:
        counts[label] = counts.get(label, 0) + 1
    items = [(k, label if counts[label] == 1 else u"{} ({})".format(label, k.split(u":")[-1] or k))
             for k, label in items]
    items.sort(key=lambda kv: kv[1].lower())
    return items


def default_group_key(groups):
    """Dados de identidade, se existir; senão Outros; senão o primeiro."""
    keys = [k for k, _ in groups]
    try:
        if uses_forge_groups():
            ident = to_unicode(DB.GroupTypeId.IdentityData.TypeId)
        else:
            ident = u"PG_IDENTITY_DATA"
        if ident in keys:
            return ident
    except Exception:
        pass
    for k in (u"", u"INVALID"):
        if k in keys:
            return k
    return keys[0] if keys else None


def group_object(key):
    """Chave -> objeto aceito por BindingMap.Insert na versão em execução."""
    if uses_forge_groups():
        return DB.ForgeTypeId(key or u"")
    return getattr(DB.BuiltInParameterGroup, key or "INVALID")


def _definition_group(definition):
    """(chave, objeto) do grupo de uma InternalDefinition já vinculada."""
    try:
        gid = definition.GetGroupTypeId()                   # Revit 2022+
        return to_unicode(gid.TypeId), gid
    except Exception:
        pass
    try:
        bipg = definition.ParameterGroup
        return to_unicode(bipg), bipg
    except Exception:
        return None, None


def group_label_of(key, groups):
    return dict(groups).get(key, key or u"Outros")


# ------------------------------------------------------------------
# categorias
# ------------------------------------------------------------------
def _category_type_label(cat):
    try:
        ctype = to_unicode(cat.CategoryType)
    except Exception:
        return None
    return dict(CATEGORY_TYPE_LABELS).get(ctype)


def list_bindable_categories(doc):
    """[{"id", "name", "type", "cat"}] das categorias que aceitam parâmetros de
    projeto (Category.AllowsBoundParameters), sem as internas."""
    out = []
    for cat in doc.Settings.Categories:
        try:
            if not cat.AllowsBoundParameters:
                continue
        except Exception:
            continue
        type_label = _category_type_label(cat)
        if type_label is None:                              # Internal / Invalid
            continue
        out.append({"id": eid_int(cat.Id), "name": to_unicode(cat.Name),
                    "type": type_label, "cat": cat})
    out.sort(key=lambda c: c["name"].lower())
    return out


def no_type_category_ids():
    out = set()
    for name in _NO_TYPE_BICS:
        bic = getattr(DB.BuiltInCategory, name, None)
        if bic is not None:
            try:
                out.add(int(bic))
            except Exception:
                pass
    return out


# ------------------------------------------------------------------
# estado atual do projeto
# ------------------------------------------------------------------
def _binding_kind(binding):
    if isinstance(binding, DB.InstanceBinding):
        return P.INSTANCE
    if isinstance(binding, DB.TypeBinding):
        return P.TYPE
    return None


def _iter_bindings(doc):
    """(InternalDefinition, Binding) de todos os parâmetros de projeto."""
    it = doc.ParameterBindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        yield it.Key, it.Current


def scan_project(doc, groups):
    """Estado dos parâmetros do projeto (ver plan.empty_project)."""
    state = P.empty_project()
    for definition, binding in _iter_bindings(doc):
        name = to_unicode(definition.Name)
        guid = None
        try:
            element = doc.GetElement(definition.Id)
            if isinstance(element, DB.SharedParameterElement):
                guid = guid_key(element.GuidValue)
        except Exception:
            pass
        state["bound_names"].setdefault(name.strip(), set()).add(guid)
        if guid is None:
            continue
        cats = set()
        try:
            for cat in binding.Categories:
                cats.add(eid_int(cat.Id))
        except Exception:
            pass
        gkey, _ = _definition_group(definition)
        state["bound"][guid] = {"kind": _binding_kind(binding), "cats": cats,
                                "group_key": gkey, "group_label": group_label_of(gkey, groups),
                                "name": name}
    for element in DB.FilteredElementCollector(doc).OfClass(DB.SharedParameterElement):
        try:
            state["shared"][guid_key(element.GuidValue)] = to_unicode(element.Name)
        except Exception:
            continue
    return state


# ------------------------------------------------------------------
# execução
# ------------------------------------------------------------------
def _category_set(app, categories):
    cset = app.Create.NewCategorySet()
    for cat in categories:
        cset.Insert(cat)
    return cset


def _new_binding(app, kind, categories):
    cset = _category_set(app, categories)
    if kind == P.TYPE:
        return app.Create.NewTypeBinding(cset)
    return app.Create.NewInstanceBinding(cset)


def _bound_definition(doc, key):
    """(InternalDefinition, Binding) do parâmetro já vinculado com este GUID."""
    for definition, binding in _iter_bindings(doc):
        try:
            element = doc.GetElement(definition.Id)
            if isinstance(element, DB.SharedParameterElement) and guid_key(element.GuidValue) == key:
                return definition, binding
        except Exception:
            continue
    return None, None


def _param_value(param):
    st = param.StorageType
    if st == DB.StorageType.String:
        return param.AsString()
    if st == DB.StorageType.Integer:
        return param.AsInteger()
    if st == DB.StorageType.Double:
        return round(param.AsDouble(), 9)
    if st == DB.StorageType.ElementId:
        return eid_int(param.AsElementId())
    return None


def _snapshot(doc, guid, kind, categories, limit=VERIFY_LIMIT):
    """[(ElementId, valor)] de até `limit` elementos que já têm valor no parâmetro."""
    out = []
    sys_guid = System.Guid(guid)
    for cat in categories:
        collector = DB.FilteredElementCollector(doc).OfCategoryId(cat.Id)
        collector = (collector.WhereElementIsElementType() if kind == P.TYPE
                     else collector.WhereElementIsNotElementType())
        for element in collector:
            try:
                param = element.get_Parameter(sys_guid)
                if param is not None and param.HasValue:
                    out.append((element.Id, _param_value(param)))
            except Exception:
                continue
            if len(out) >= limit:
                return out
    return out


def _verify_snapshot(doc, guid, snapshot, name):
    sys_guid = System.Guid(guid)
    for eid, before in snapshot:
        element = doc.GetElement(eid)
        param = element.get_Parameter(sys_guid) if element is not None else None
        if param is None or not param.HasValue or _param_value(param) != before:
            raise ApplyError(u"'{}': ao adicionar categorias, um valor já preenchido mudou "
                             u"(elemento {}). Nada foi gravado.".format(name, eid_int(eid)))


def apply_plan(doc, path, rows, categories_by_id, groups):
    """Executa as linhas create/extend numa única transação. Tudo ou nada.

    rows             [plan.PlanRow] (só create/extend são executadas)
    categories_by_id id -> Category (de list_bindable_categories)
    Lança ApplyError (projeto inalterado) em qualquer falha.
    """
    todo = [r for r in rows if r.actionable]
    if not todo:
        raise ApplyError(u"Nada a aplicar.")
    if doc.IsReadOnly:
        raise ApplyError(u"O modelo está aberto como somente leitura.")
    app = doc.Application
    bmap = doc.ParameterBindings

    with SharedParamFile(app, path) as dfile:
        external = _external_definitions(dfile)
        tx = DB.Transaction(doc, TRANSACTION_NAME)
        try:
            tx.Start()
            for row in todo:
                cats = []
                for cid in (row.cats if row.action == P.CREATE else row.add_cats):
                    cat = categories_by_id.get(cid)
                    if cat is None:
                        raise ApplyError(u"'{}': categoria {} não encontrada no projeto.".format(
                            row.name, cid))
                    cats.append(cat)

                if row.action == P.CREATE:
                    definition = external.get(row.guid)
                    if definition is None:
                        raise ApplyError(u"'{}': a definição (GUID {}) não está no arquivo aberto "
                                         u"pelo Revit.".format(row.name, row.guid))
                    current, _ = _bound_definition(doc, row.guid)
                    if current is not None:
                        raise ApplyError(u"'{}': o parâmetro passou a estar vinculado. Revise "
                                         u"de novo.".format(row.name))
                    ok = bmap.Insert(definition, _new_binding(app, row.kind, cats),
                                     group_object(row.group_key))
                    if not ok:
                        raise ApplyError(u"'{}': o Revit recusou o vínculo.".format(row.name))
                    _check_binding(doc, row, set(row.cats))
                else:
                    definition, binding = _bound_definition(doc, row.guid)
                    if definition is None:
                        raise ApplyError(u"'{}': o vínculo atual não foi encontrado.".format(row.name))
                    if _binding_kind(binding) != row.kind:
                        raise ApplyError(u"'{}': o tipo de vínculo atual mudou. Revise de novo.".format(
                            row.name))
                    existing = list(binding.Categories)
                    snapshot = _snapshot(doc, row.guid, row.kind, existing)
                    _, group = _definition_group(definition)
                    if group is None:
                        raise ApplyError(u"'{}': não foi possível ler o grupo atual.".format(row.name))
                    new_binding = _new_binding(app, row.kind, existing + cats)
                    ok = bmap.ReInsert(definition, new_binding, group)
                    if not ok:
                        raise ApplyError(u"'{}': o Revit recusou a ampliação das categorias.".format(
                            row.name))
                    doc.Regenerate()
                    _check_binding(doc, row, set(eid_int(c.Id) for c in existing) | row.add_cats)
                    _verify_snapshot(doc, row.guid, snapshot, row.name)
            status = tx.Commit()
            if status != DB.TransactionStatus.Committed:
                raise ApplyError(u"O Revit não confirmou a transação (status {}).".format(status))
        except ApplyError:
            _rollback(tx)
            raise
        except Exception as ex:
            _rollback(tx)
            raise ApplyError(u"Erro do Revit: {}".format(to_unicode(ex)))
        finally:
            tx.Dispose()
    return len(todo)


def _check_binding(doc, row, expected):
    """Depois do ReInsert: mesmo tipo de vínculo e todas as categorias esperadas."""
    definition, binding = _bound_definition(doc, row.guid)
    if definition is None or _binding_kind(binding) != row.kind:
        raise ApplyError(u"'{}': o vínculo não ficou como esperado.".format(row.name))
    now = set(eid_int(c.Id) for c in binding.Categories)
    if not expected.issubset(now):
        raise ApplyError(u"'{}': {} categoria(s) não ficaram vinculadas.".format(
            row.name, len(expected - now)))


def _rollback(tx):
    try:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
    except Exception:
        pass
