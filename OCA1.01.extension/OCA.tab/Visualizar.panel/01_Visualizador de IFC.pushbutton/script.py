# -*- coding: utf-8 -*-
# Visualizador de IFC - abre o visualizador web IFC//HUD no navegador, pronto para carregar um arquivo .ifc.
#
# O visualizador (ifc-hud-viewer.html, na mesma pasta deste script) lê o IFC
# localmente no navegador - nada é enviado para a internet -, mas carrega as
# bibliotecas three.js e web-ifc de cdn.jsdelivr.net, então precisa de acesso à internet.
#
# A descrição fica em comentário (e não em docstring) de propósito: o
# pyRevit 5.2 lê a docstring do módulo com .decode('utf-8'), que falha com
# acentos no IronPython e deixa o botão sem tooltip. O tooltip é só __doc__.
__title__ = u"Visualizador\nde IFC"
__doc__ = u"Abre o visualizador de IFC (modelo 3D, ambientes, áreas e dados IFC) no navegador. Depois arraste ou selecione um arquivo .ifc para carregar."

import os
import codecs
from pyrevit import forms, script
from oca_ds import alert_title, brand_report

ALERT_TITLE = alert_title(__title__)          # "OCA · Visualizador de IFC"

# Padrão visual OCA (lib/oca_ds): o navegador abre uma cópia do visualizador com o
# logo e as cores do tema do Revit no momento do clique (brand_report). Se não der
# para gravar a cópia, abre o arquivo original, que usa as cores padrão (tema claro).
caminho_original = os.path.join(os.path.dirname(__file__), "ifc-hud-viewer.html")


def copia_com_marca():
    """Grava a cópia com logo e tema na pasta de dados do pyRevit e devolve o caminho.
    Se não der para ler ou gravar, devolve o caminho do original."""
    try:
        with codecs.open(caminho_original, "r", encoding="utf-8") as f:
            html = brand_report(f.read())
        destino = script.get_universal_data_file("ifc_hud_viewer", "html")
        with codecs.open(destino, "w", encoding="utf-8") as f:
            f.write(html)
        return destino
    except Exception:
        return caminho_original


caminho_html = copia_com_marca() if os.path.exists(caminho_original) else caminho_original
url = "file:///" + caminho_html.replace("\\", "/")


def abrir_no_navegador():
    # 1) Execução pelo shell do .NET (o método mais confiável dentro do Revit)
    try:
        import clr
        clr.AddReference("System")
        from System.Diagnostics import Process
        Process.Start(caminho_html)
        return True
    except Exception:
        pass
    # 2) Pelo Explorador de Arquivos (explorer.exe)
    try:
        from System.Diagnostics import Process
        Process.Start("explorer.exe", '"{}"'.format(caminho_html))
        return True
    except Exception:
        pass
    # 3) Pelo módulo webbrowser do Python
    try:
        import webbrowser
        return webbrowser.open(url)
    except Exception:
        return False


# Sem janela de saída: em caso de sucesso apenas o navegador abre.
if not os.path.exists(caminho_html):
    forms.alert(u"Arquivo do visualizador não encontrado.",
                sub_msg=caminho_html, title=ALERT_TITLE)
elif not abrir_no_navegador():
    forms.alert(u"Não foi possível abrir o visualizador automaticamente.",
                sub_msg=u"Abra por este endereço:\n{}".format(url), title=ALERT_TITLE)
