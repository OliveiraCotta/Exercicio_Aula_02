# -*- coding: utf-8 -*-
"""Mensagens no padrão OCA.

Os diálogos continuam sendo os nativos (forms.alert do pyRevit = TaskDialog
do Revit): o padrão define título e texto, não as cores.

- Título: "OCA · <nome do comando>" (alert_title).
- Frase principal (msg): curta. Na confirmação é a própria pergunta
  ("Gravar 221 campos em 48 ambientes?"); no sucesso, o resultado; no
  alerta, o que impede; no erro, o que falhou.
- Detalhe (sub_msg): o que muda, o caminho do arquivo, como resolver ou a
  mensagem técnica do Revit.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

try:
    unicode
except NameError:           # IronPython 3 / CPython
    unicode = str

BRAND = u"OCA"


def command_name(title):
    """__title__ do botão ("Abertura\\nde Vistas") -> "Abertura de Vistas"."""
    return u" ".join(unicode(title or u"").split())


def alert_title(title):
    """Título padrão dos diálogos: "OCA · <nome do comando>"."""
    return u"{} · {}".format(BRAND, command_name(title))


def xml_escape(text):
    """Escapa texto para atributo XAML/HTML."""
    text = unicode(text or u"")
    for raw, esc in ((u"&", u"&amp;"), (u"<", u"&lt;"), (u">", u"&gt;"),
                     (u'"', u"&quot;"), (u"\n", u"&#10;")):
        text = text.replace(raw, esc)
    return text


def alert(msg, sub_msg=None, title=u"", **kwargs):
    """forms.alert com o título padrão. title = __title__ do comando.
    Os demais argumentos (yes, no, ok, cancel, options, exitscript...) passam
    direto para o pyRevit, com o mesmo retorno."""
    from pyrevit import forms
    if sub_msg is not None:
        kwargs["sub_msg"] = sub_msg
    return forms.alert(msg, title=alert_title(title), **kwargs)
