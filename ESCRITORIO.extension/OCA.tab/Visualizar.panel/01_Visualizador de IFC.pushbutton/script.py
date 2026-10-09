# -*- coding: utf-8 -*-
"""Visualizador de IFC - abre o visualizador web IFC//HUD no navegador, pronto para carregar um arquivo .ifc.

O visualizador (ifc-hud-viewer.html, ao lado deste script) le o IFC localmente
no navegador - nada e enviado para a internet - mas carrega three.js e web-ifc
de cdn.jsdelivr.net, entao precisa de acesso a internet.
"""
__title__ = "Visualizador\nde IFC"
__doc__ = "Abre o visualizador de IFC (modelo 3D, ambientes, areas e dados IFC) no navegador. Depois arraste ou selecione um arquivo .ifc para carregar."

import os
import codecs
from pyrevit import script
from oca_ui import brand_report, output_header

output = script.get_output()
viewer_path = os.path.join(os.path.dirname(__file__), "ifc-hud-viewer.html")


def branded_copy():
    """Cópia do visualizador com as cores e o logo OCA (pasta de dados do pyRevit).
    Se não der para gravar, abre o arquivo original, que usa as mesmas cores sem o logo."""
    try:
        with codecs.open(viewer_path, "r", encoding="utf-8") as f:
            html = brand_report(f.read())
        out = script.get_universal_data_file("ifc_hud_viewer", "html")
        with codecs.open(out, "w", encoding="utf-8") as f:
            f.write(html)
        return out
    except Exception:
        return viewer_path


html_path = branded_copy() if os.path.exists(viewer_path) else viewer_path
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


output_header(output, __title__, u"Visualizador web de arquivos IFC")
if os.path.exists(html_path) and try_open():
    output.print_md("## Visualizador de IFC")
    output.print_md("**Visualizador aberto** no navegador - arraste um arquivo `.ifc` para a pagina ou clique em **Browse...** para carrega-lo.")
    output.print_md("- O arquivo e lido localmente; nada e enviado.")
    output.print_md("- Precisa de acesso a internet (`cdn.jsdelivr.net`) para carregar three.js e web-ifc.")
    output.print_md("- Use Chrome, Edge ou Firefox atualizado como navegador padrao.")
else:
    output.print_md("Nao foi possivel abrir o visualizador automaticamente. Abra aqui: {}".format(url))
