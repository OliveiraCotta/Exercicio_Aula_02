# -*- coding: utf-8 -*-
"""Tema automático: claro ou escuro conforme o tema do Revit.

Regra:
- Revit 2024 ou posterior: Autodesk.Revit.UI.UIThemeManager.CurrentTheme
  (API pública). UITheme.Dark -> escuro; qualquer outro valor -> claro.
- Revit 2023 ou anterior (a classe não existe) ou qualquer falha: claro.

O tema é lido quando cada janela é montada e quando cada relatório é gerado.
Uma janela já aberta mantém o tema com que abriu (o WPF não acompanha a troca).

Com "usar a configuração do sistema" ligado no Revit
(UIThemeManager.FollowSystemColorTheme), a documentação não garante que
CurrentTheme devolva o tema efetivamente aplicado. Isso é teste obrigatório
no Revit; o registro do Windows não é consultado.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

from oca_ds import tokens


def revit_is_dark():
    """True se o Revit está no tema escuro; False no claro ou se não der para saber."""
    try:
        from Autodesk.Revit.UI import UIThemeManager, UITheme
    except Exception:           # Revit < 2024 ou fora do Revit
        return False
    try:
        return UIThemeManager.CurrentTheme == UITheme.Dark
    except Exception:
        return False


def is_dark(dark=None):
    """dark explícito (True/False) ou, se None, o tema atual do Revit."""
    return revit_is_dark() if dark is None else bool(dark)


def palette(dark=None):
    return tokens.DARK if is_dark(dark) else tokens.LIGHT


def name(dark=None):
    return u"escuro" if is_dark(dark) else u"claro"
