# -*- coding: utf-8 -*-
"""Cor do rodape (barra de titulo) dos paineis da aba OCA conforme o tema.

O pyRevit roda este arquivo ao carregar a extensao (abrir o Revit ou
pyRevit > Reload). As cores sao aplicadas:
  - na primeira vez que o Revit fica ocioso (o ribbon ja existe);
  - quando o tema muda em Opcoes > Interface (evento ThemeChanged);
  - quando o Windows muda de claro/escuro com "Use system setting" ativo
    (verificado no Idling, que so troca as cores se o tema mudou).

No tema escuro o texto do rodape e claro, por isso cada cor tem uma
variante escura: os tons pastel deixariam o titulo ilegivel.
Revit < 2024 nao tem tema escuro: usa sempre as cores claras.
"""

import clr
clr.AddReference("AdWindows")
clr.AddReference("PresentationCore")

from System import AppDomain, EventHandler
from System.Windows.Media import SolidColorBrush, ColorConverter
from Microsoft.Win32 import Registry
from Autodesk.Windows import ComponentManager
from Autodesk.Revit.UI.Events import IdlingEventArgs

from pyrevit import HOST_APP

try:
    from Autodesk.Revit.UI import UIThemeManager, UITheme
except ImportError:
    UIThemeManager = None

TAB = "OCA"

# painel: (cor no tema claro, cor no tema escuro)
CORES = {
    "Vistas": ("#dfd6f4", "#574887"),
    "Keynote": ("#acede6", "#2e5e58"),
    "Visualizar": ("#d9d9d9", "#5a5a5a"),
}

# chave para guardar os handlers entre recargas do pyRevit
CHAVE = "OCA_TEMA_RODAPE_HANDLERS"

_estado = {"tema": None}


def _tema_escuro():
    if UIThemeManager is None:
        return False
    # "Use system setting" (Revit 2025+): segue o modo de apps do Windows
    if getattr(UIThemeManager, "FollowSystemColorTheme", False):
        chave = Registry.CurrentUser.OpenSubKey(
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        if chave is not None:
            try:
                return chave.GetValue("AppsUseLightTheme", 1) == 0
            finally:
                chave.Close()
    return UIThemeManager.CurrentTheme == UITheme.Dark


def _pincel(hex_cor):
    pincel = SolidColorBrush(ColorConverter.ConvertFromString(hex_cor))
    pincel.Freeze()
    return pincel


def aplicar_cores(forcar=False):
    escuro = _tema_escuro()
    if not forcar and _estado["tema"] == escuro:
        return
    aplicou = False
    for aba in ComponentManager.Ribbon.Tabs:
        if aba.Title != TAB:
            continue
        for painel in aba.Panels:
            cores = CORES.get(painel.Source.Title)
            if cores:
                painel.CustomPanelTitleBarBackground = _pincel(cores[escuro])
                aplicou = True
    # so memoriza o tema se o ribbon ja existia, para tentar de novo depois
    if aplicou:
        _estado["tema"] = escuro


def _ao_ficar_ocioso(sender, args):
    try:
        aplicar_cores()
    except Exception:
        pass


def _ao_trocar_tema(sender, args):
    try:
        aplicar_cores(forcar=True)
    except Exception:
        pass


def _registrar():
    uiapp = HOST_APP.uiapp
    dominio = AppDomain.CurrentDomain

    # remove os handlers da carga anterior (pyRevit > Reload)
    antigos = dominio.GetData(CHAVE)
    if antigos is not None:
        ocioso, tema = antigos
        uiapp.Idling -= ocioso
        if tema is not None:
            uiapp.ThemeChanged -= tema

    ocioso = EventHandler[IdlingEventArgs](_ao_ficar_ocioso)
    uiapp.Idling += ocioso

    tema = None
    if UIThemeManager is not None:
        from Autodesk.Revit.UI.Events import ThemeChangedEventArgs
        tema = EventHandler[ThemeChangedEventArgs](_ao_trocar_tema)
        uiapp.ThemeChanged += tema

    dominio.SetData(CHAVE, (ocioso, tema))


_registrar()
