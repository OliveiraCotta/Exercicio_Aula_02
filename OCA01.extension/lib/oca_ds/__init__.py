# -*- coding: utf-8 -*-
"""Design System OCA para os comandos pyRevit.

Uso num script (a lógica da janela não muda; só o XAML):

    from oca_ds import build_xaml, alert_title

    ALERT_TITLE = alert_title(__title__)          # "OCA · Nome do Comando"
    XAML = build_xaml(title=__title__, subtitle=__doc__,
                      body=BODY, footer_right=FOOTER, size="S")

    class MinhaJanela(forms.WPFWindow):
        def __init__(self, xaml):
            forms.WPFWindow.__init__(self, xaml, literal_string=True)

    forms.alert(u"Gravar 12 campos?", sub_msg=u"Ctrl+Z desfaz.", title=ALERT_TITLE,
                yes=True, no=True)

Módulos:
    tokens    cores claro/escuro, fontes, raios, medidas, paleta do Excel
    theme     tema automático (UIThemeManager do Revit)
    wpf       estilos WPF e build_xaml (cabeçalho, corpo, rodapé)
    messages  título e texto dos diálogos
    report    relatórios HTML e cabeçalho da janela de saída
    excel     planilhas .xlsx no padrão

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

from oca_ds.messages import command_name, alert_title, alert, xml_escape
from oca_ds.wpf import build_xaml, logo_uri
from oca_ds.report import brand_report, output_header, logo_data_uri
from oca_ds.theme import is_dark
