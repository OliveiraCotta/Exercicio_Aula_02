# -*- coding: utf-8 -*-
u"""Visualizador de IFC - abre o visualizador web IFC//HUD no navegador, pronto para carregar um arquivo .ifc.

O visualizador (ifc-hud-viewer.html, na mesma pasta deste script) lê o IFC
localmente no navegador - nada é enviado para a internet -, mas carrega as
bibliotecas three.js e web-ifc de cdn.jsdelivr.net, então precisa de acesso à internet.
"""
__title__ = u"Visualizador\nde IFC"
__doc__ = u"Abre o visualizador de IFC (modelo 3D, ambientes, áreas e dados IFC) no navegador. Depois arraste ou selecione um arquivo .ifc para carregar."

import os
from pyrevit import forms

caminho_html = os.path.join(os.path.dirname(__file__), "ifc-hud-viewer.html")
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
    forms.alert(u"Arquivo do visualizador não encontrado:\n{}".format(caminho_html),
                title=u"Visualizador de IFC")
elif not abrir_no_navegador():
    forms.alert(u"Não foi possível abrir o visualizador automaticamente.\n\nAbra por este endereço:\n{}".format(url),
                title=u"Visualizador de IFC")
