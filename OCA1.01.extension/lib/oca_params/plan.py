# -*- coding: utf-8 -*-
"""Plano de execução: decide, para cada parâmetro configurado, o que será feito.

Trabalha só com dados simples (dicionários, conjuntos de ids), sem a API do
Revit, para que as regras possam ser testadas fora do Revit e para que a
revisão mostrada ao usuário seja exatamente o que a transação executa.

Ações:
    create      vínculo novo (o parâmetro ainda não está vinculado ao projeto)
    extend      o parâmetro já está vinculado com o MESMO tipo de vínculo:
                só as categorias que faltam são acrescentadas; o grupo de
                propriedades e o tipo de vínculo atuais são mantidos
    none        nada a fazer (já vinculado a todas as categorias pedidas)
    blocked     não será aplicado (conflito ou definição inutilizável)
    incomplete  falta informação (sem categoria): impede a execução

Regras de segurança (nunca destrutivas):
- Instância <-> Tipo de um parâmetro já vinculado NÃO é trocado: exigiria
  excluir o parâmetro e os valores dele. Fica bloqueado, com o motivo.
- Categorias já vinculadas nunca são retiradas.
- O grupo de propriedades de um parâmetro já vinculado não é alterado.
- Nome igual ao de outro parâmetro de projeto (outro GUID ou não
  compartilhado) bloqueia: o Revit teria dois parâmetros com o mesmo nome.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3, sem API do Revit.
"""

try:
    unicode
except NameError:           # IronPython 3 / CPython
    unicode = str

INSTANCE = "inst"
TYPE = "type"
KIND_LABELS = {INSTANCE: u"Instância", TYPE: u"Tipo"}

CREATE = "create"
EXTEND = "extend"
NONE = "none"
BLOCKED = "blocked"
INCOMPLETE = "incomplete"
ACTION_LABELS = {
    CREATE: u"Criar vínculo",
    EXTEND: u"Adicionar categorias",
    NONE: u"Sem alteração",
    BLOCKED: u"Não será aplicado",
    INCOMPLETE: u"Incompleto",
}
# nível visual da linha (cores da tabela): ok, warn, err, off
ACTION_LEVEL = {CREATE: "ok", EXTEND: "ok", NONE: "off", BLOCKED: "warn", INCOMPLETE: "err"}


class PlanRow(object):
    def __init__(self, item):
        self.guid = item["guid"]
        self.name = item["name"]
        self.kind = item.get("kind")
        self.group_key = item.get("group_key")
        self.cats = set(item.get("cats") or ())
        self.action = NONE
        self.add_cats = set()       # extend: categorias acrescentadas
        self.notes = []
        self.group_label = u""
        self.kind_label = u""
        self.cats_text = u""

    @property
    def level(self):
        if self.action in (CREATE, EXTEND) and self.notes and any(
                n.startswith(u"Atenção") for n in self.notes):
            return "warn"
        return ACTION_LEVEL[self.action]

    @property
    def action_label(self):
        return ACTION_LABELS[self.action]

    @property
    def actionable(self):
        return self.action in (CREATE, EXTEND)


def empty_project():
    """Estrutura do estado do projeto (preenchida por revit_ops.scan_project)."""
    return {"bound": {},        # guid -> {"kind", "cats": set, "group_key", "group_label", "name"}
            "bound_names": {},  # nome -> set(guid ou None p/ parâmetro de projeto não compartilhado)
            "shared": {}}       # guid -> nome (SharedParameterElement no projeto, vinculado ou não)


def name_conflict(guid, name, project):
    """Texto do conflito se já existe parâmetro de projeto com este nome e outro GUID."""
    others = [g for g in project["bound_names"].get(name.strip(), set()) if g != guid]
    if not others:
        return None
    if None in others:
        return (u"Já existe um parâmetro de projeto (não compartilhado) chamado '{}'. "
                u"Vincular criaria dois parâmetros com o mesmo nome.".format(name.strip()))
    return (u"Já existe um parâmetro de projeto chamado '{}' com outro GUID ({}). "
            u"Vincular criaria dois parâmetros com o mesmo nome.".format(name.strip(), others[0]))


def project_status(guid, name, project, cat_names=None):
    """(situação curta, detalhe, utilizável) para a tabela de parâmetros do arquivo."""
    bound = project["bound"].get(guid)
    if bound:
        n = len(bound["cats"])
        return (u"No projeto · {} · {} categoria(s)".format(KIND_LABELS.get(bound["kind"], u"?"), n),
                u"Já vinculado. Pode receber categorias novas; tipo de vínculo e grupo atuais são mantidos.",
                True)
    conflict = name_conflict(guid, name, project)
    if conflict:
        return (u"Conflito de nome", conflict, False)
    if guid in project["shared"]:
        pname = project["shared"][guid]
        detail = u"Existe no projeto só por famílias carregadas (sem vínculo de projeto)."
        if pname != name:
            detail += u" No projeto o nome é '{}'.".format(pname)
        return (u"Só em famílias", detail, True)
    return (u"Novo", u"", True)


def _cats_text(ids, cat_names, limit=None):
    names = sorted((cat_names.get(i, u"Categoria {}".format(i)) for i in ids),
                   key=lambda s: s.lower())
    if limit and len(names) > limit:
        return u"{} (+{})".format(u", ".join(names[:limit]), len(names) - limit)
    return u", ".join(names)


def cats_summary(ids, cat_names, limit=3):
    if not ids:
        return u"(nenhuma)"
    return u"{} · {}".format(len(ids), _cats_text(ids, cat_names, limit))


def build_plan(items, defs, project, cat_names, group_labels, no_type_cats=(),
               allow_extend=True):
    """Lista de PlanRow na ordem de items.

    items        [{"guid", "name", "group_key", "kind", "cats": set(ids)}]
    defs         guid -> {"usable": bool, "reason": texto, "revit_name": texto,
                          "visible": bool, "modifiable": bool}
    project      ver empty_project()
    cat_names    id -> nome da categoria
    group_labels chave do grupo -> rótulo
    no_type_cats ids de categorias sem tipos (vínculo de Tipo não aparece nelas)
    allow_extend False = parâmetros já vinculados não recebem categorias novas
    """
    no_type_cats = set(no_type_cats)
    rows = [PlanRow(item) for item in items]

    # nomes repetidos dentro do próprio lote (GUIDs diferentes)
    names = {}
    for r in rows:
        names.setdefault(r.name.strip(), set()).add(r.guid)
    dup_names = set(n for n, guids in names.items() if len(guids) > 1)

    for r in rows:
        info = defs.get(r.guid) or {"usable": False,
                                    "reason": u"Definição não encontrada no arquivo carregado."}
        bound = project["bound"].get(r.guid)
        r.kind_label = KIND_LABELS.get(r.kind, u"")
        r.group_label = group_labels.get(r.group_key, u"")
        r.cats_text = _cats_text(r.cats, cat_names)

        if not info.get("usable"):
            r.action = BLOCKED
            r.notes.append(info.get("reason") or u"Definição inutilizável.")
            continue
        if r.name.strip() in dup_names:
            r.action = BLOCKED
            r.notes.append(u"Outro parâmetro do lote tem o mesmo nome '{}'. "
                           u"Deixe só um deles na configuração.".format(r.name.strip()))
            continue
        if not bound:
            conflict = name_conflict(r.guid, r.name, project)
            if conflict:
                r.action = BLOCKED
                r.notes.append(conflict)
                continue
        if r.kind not in KIND_LABELS:
            r.action = INCOMPLETE
            r.notes.append(u"Escolha Instância ou Tipo.")
            continue
        if not r.cats:
            r.action = INCOMPLETE
            r.notes.append(u"Nenhuma categoria escolhida.")
            continue

        if bound:
            r.group_label = bound.get("group_label") or r.group_label
            if bound["kind"] != r.kind:
                r.action = BLOCKED
                r.notes.append(
                    u"Já está no projeto como {}. Mudar para {} exigiria excluir o parâmetro e os "
                    u"valores preenchidos: não é feito por este comando.".format(
                        KIND_LABELS.get(bound["kind"], u"?"), KIND_LABELS[r.kind]))
                continue
            missing = r.cats - bound["cats"]
            if not missing:
                r.action = NONE
                r.notes.append(u"Já vinculado a todas as categorias escolhidas.")
                continue
            if not allow_extend:
                r.action = NONE
                r.notes.append(u"Já vinculado; adicionar categorias está desligado na revisão "
                               u"({} categoria(s) ficariam de fora).".format(len(missing)))
                continue
            r.action = EXTEND
            r.add_cats = missing
            r.cats_text = u"+ " + _cats_text(missing, cat_names)
            r.notes.append(u"Mantém as {} categoria(s) atuais e o grupo '{}'.".format(
                len(bound["cats"]), bound.get("group_label") or u"atual"))
            if r.group_key != bound.get("group_key") and r.group_key is not None:
                r.notes.append(u"Grupo de propriedades pedido ('{}') ignorado: o parâmetro já "
                               u"existe e o grupo dele não é alterado.".format(
                                   group_labels.get(r.group_key, u"?")))
        else:
            if r.group_key is None or r.group_key not in group_labels:
                r.action = INCOMPLETE
                r.notes.append(u"Escolha o grupo de propriedades.")
                continue
            r.action = CREATE
            if r.guid in project["shared"]:
                r.notes.append(u"O parâmetro já existe no projeto por famílias; o vínculo usa o mesmo GUID.")

        # avisos comuns a create/extend
        target = r.add_cats if r.action == EXTEND else r.cats
        if r.kind == TYPE:
            notype = sorted(cat_names.get(i, unicode(i)) for i in target if i in no_type_cats)
            if notype:
                r.notes.append(u"Atenção: {} não tem tipos; um parâmetro de Tipo não aparece "
                               u"nessa(s) categoria(s).".format(u", ".join(notype)))
        revit_name = info.get("revit_name")
        if revit_name and revit_name != r.name:
            r.notes.append(u"Atenção: o Revit leu o nome como '{}' (codificação do arquivo?). "
                           u"É esse o nome que vai para o projeto.".format(revit_name))
        if info.get("visible") is False:
            r.notes.append(u"Atenção: definição oculta (não aparece na paleta Propriedades).")
        if info.get("modifiable") is False:
            r.notes.append(u"Atenção: definição não editável pelo usuário.")
    return rows


def summary(rows):
    """Contagens por ação."""
    out = dict((a, 0) for a in ACTION_LABELS)
    for r in rows:
        out[r.action] += 1
    return out
