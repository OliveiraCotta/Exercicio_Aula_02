# -*- coding: utf-8 -*-
"""Visualizador de IFC - abre o visualizador web IFC//HUD no navegador, pronto para carregar um arquivo .ifc.

O visualizador (ifc-hud-viewer.html, ao lado deste script) le o IFC localmente
no navegador - nada e enviado para a internet - mas carrega three.js e web-ifc
de cdn.jsdelivr.net, entao precisa de acesso a internet.
"""
__title__ = "Visualizador\nde IFC"
__doc__ = "Abre o visualizador de IFC (modelo 3D, ambientes, areas e dados IFC) no navegador. Depois arraste ou selecione um arquivo .ifc para carregar."

import os
from pyrevit import script

output = script.get_output()
html_path = os.path.join(os.path.dirname(__file__), "ifc-hud-viewer.html")
url = "file:///" + html_path.replace("\\", "/")


def try_open():
    # 1) .NET shell execute (most reliable inside Revit)
    try:
        import clr
        clr.AddReference("System")
        from System.Diagnostics import Process
        Process.Start(html_path)
        return True
    except Exception:
        pass
    # 2) explorer.exe
    try:
        from System.Diagnostics import Process
        Process.Start("explorer.exe", '"{}"'.format(html_path))
        return True
    except Exception:
        pass
    # 3) python webbrowser
    try:
        import webbrowser
        return webbrowser.open(url)
    except Exception:
        return False


if os.path.exists(html_path) and try_open():
    output.print_md("## Visualizador de IFC")
    output.print_md("**Visualizador aberto** no navegador - arraste um arquivo `.ifc` para a pagina ou clique em **Browse...** para carrega-lo.")
    output.print_md("- O arquivo e lido localmente; nada e enviado.")
    output.print_md("- Precisa de acesso a internet (`cdn.jsdelivr.net`) para carregar three.js e web-ifc.")
    output.print_md("- Use Chrome, Edge ou Firefox atualizado como navegador padrao.")
else:
    output.print_md("Nao foi possivel abrir o visualizador automaticamente. Abra aqui: {}".format(url))
