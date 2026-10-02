# -*- coding: utf-8 -*-
"""Workshop Run Sheet - opens the step-by-step facilitator run sheet in the browser."""
__title__ = "Run\nSheet"
__doc__ = "Opens the PA Academy workshop run sheet (every step, in order) in your browser."

import os
from pyrevit import script

output = script.get_output()
html_path = os.path.join(os.path.dirname(__file__), "runsheet.html")
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
    output.print_md("**Run sheet opened** in your browser.")
else:
    output.print_md("Could not open automatically. Open it here: {}".format(url))
