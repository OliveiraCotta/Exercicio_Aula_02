# -*- coding: utf-8 -*-
"""Relatórios HTML e janela de saída do pyRevit no padrão OCA.

Relatório: no modelo .html, /*__OCA_TOKENS__*/ dentro do :root recebe as cores
do tema atual do Revit (--oca-Bg, --oca-Ink, --oca-Accent...), __OCA_LOGO__
recebe o logo (base64, porque o arquivo é gravado fora da extensão),
__OCA_THEME__ recebe "light" ou "dark" e /*__OCA_XLSX__*/null recebe os
estilos das planilhas geradas no navegador (excel.js_payload).

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

import base64
import os

from oca_ds import tokens, theme, excel
from oca_ds.messages import command_name, xml_escape

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo", "logo.png")


def logo_data_uri():
    """Logo em base64 (data URI). Vazio se o arquivo não existir."""
    try:
        with open(LOGO_PATH, "rb") as f:
            data = f.read()
    except (IOError, OSError):
        return u""
    return u"data:image/png;base64," + base64.b64encode(data).decode("ascii")


def css_tokens(dark=None):
    colors = theme.palette(dark)
    return u"".join(u"--oca-{}:{};".format(k, v) for k, v in sorted(colors.items()))


def brand_report(html, dark=None):
    """Aplica a marca e o tema (o do Revit, se dark=None) a um modelo HTML."""
    is_dark = theme.is_dark(dark)
    return (html.replace(u"/*__OCA_TOKENS__*/", css_tokens(is_dark))
            .replace(u"__OCA_THEME__", u"dark" if is_dark else u"light")
            .replace(u"/*__OCA_XLSX__*/null", excel.js_payload())
            .replace(u"__OCA_LOGO__", logo_data_uri()))


def output_header(output, title, subtitle=u""):
    """Cabeçalho com logo, nome do comando e descrição na janela de saída do
    pyRevit. A janela de saída tem fundo claro próprio: usa a paleta clara."""
    c = tokens.LIGHT
    logo = logo_data_uri()
    img = (u'<img src="{}" style="height:28px;width:auto;margin-right:12px;vertical-align:middle">'.format(logo)
           if logo else u"")
    sub = (u'<div style="font-size:11px;color:{}">{}</div>'.format(c["Ink2"], xml_escape(subtitle))
           if subtitle else u"")
    output.print_html(
        u'<div style="display:flex;align-items:center;padding:0 0 10px;margin:0 0 10px;'
        u'border-bottom:2px solid {brand};font-family:Segoe UI,sans-serif">{img}<div>'
        u'<div style="font-size:15px;font-weight:600;color:{ink}">{title}</div>{sub}</div></div>'.format(
            brand=c["Brand"], img=img, ink=c["Ink"], title=xml_escape(command_name(title)), sub=sub))
