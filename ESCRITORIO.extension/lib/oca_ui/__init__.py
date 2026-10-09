# -*- coding: utf-8 -*-
"""Padrão visual OCA para as janelas dos comandos pyRevit.

Uso num script (a lógica da janela não muda; só o XAML):

    from oca_ui import build_xaml, alert_title

    ALERT_TITLE = alert_title(__title__)
    XAML = build_xaml(title=__title__, subtitle=__doc__,
                      body=BODY, footer_right=FOOTER, size="S")

    class MinhaJanela(forms.WPFWindow):
        def __init__(self, xaml):
            forms.WPFWindow.__init__(self, xaml, literal_string=True)

    forms.alert(u"Frase principal.", sub_msg=u"Detalhe.", title=ALERT_TITLE)

build_xaml monta a janela inteira: tema escuro OCA,
cabeçalho com logo + nome do comando + descrição, corpo e rodapé. O corpo
e o rodapé são o XAML do próprio comando, com os mesmos x:Name de antes.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

import os
import re

from oca_ui import theme

try:
    unicode
except NameError:           # IronPython 3 / CPython
    unicode = str

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(HERE, "assets", "logo", "logo.png")
BRAND = u"OCA"


def command_name(title):
    """__title__ do botão ("Abertura\\nde Vistas") -> "Abertura de Vistas"."""
    return u" ".join(unicode(title or u"").split())


def alert_title(title):
    """Título padrão dos forms.alert: "OCA · <nome do comando>"."""
    return u"{} · {}".format(BRAND, command_name(title))


def _xml(text):
    """Escapa texto para atributo XAML."""
    text = unicode(text or u"")
    for raw, esc in ((u"&", u"&amp;"), (u"<", u"&lt;"), (u">", u"&gt;"),
                     (u'"', u"&quot;"), (u"\n", u"&#10;")):
        text = text.replace(raw, esc)
    return text


def logo_uri():
    """URI do logo, ou None se o arquivo não existir ou não puder ser lido.

    O arquivo é decodificado aqui uma vez: se estiver corrompido, a janela
    abre sem logo em vez de falhar ao carregar o XAML.
    """
    if not os.path.isfile(LOGO_PATH):
        return None
    try:
        import clr
        clr.AddReference("PresentationCore")
        from System import Uri
        from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
    except Exception:       # fora do .NET (testes): URI simples
        return u"file:///" + LOGO_PATH.replace(u"\\", u"/").lstrip(u"/")
    try:
        uri = Uri(LOGO_PATH)
        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.UriSource = uri
        bmp.CacheOption = BitmapCacheOption.OnLoad
        bmp.EndInit()
        return unicode(uri.AbsoluteUri)
    except Exception:
        return None


_LOGO_XAML = u"""
        <Image DockPanel.Dock="Left" Height="28" Margin="0,0,12,0" VerticalAlignment="Center"
               Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality">
          <Image.Source><BitmapImage UriSource="@uri@" CacheOption="OnLoad"/></Image.Source>
        </Image>
        <Border DockPanel.Dock="Left" Width="1" Margin="0,0,12,0" Background="{StaticResource oca.Line}"/>"""

_ICON_XAML = u"""
  <Window.Icon><BitmapImage UriSource="@uri@" CacheOption="OnLoad"/></Window.Icon>"""

_SUBTITLE_XAML = u"""
          <TextBlock Text="@subtitle@" FontSize="11.5" Margin="0,1,0,0" TextWrapping="Wrap"
                     Foreground="{StaticResource oca.Ink2}"/>"""

_WINDOW = u"""<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="@title@" Width="@width@" @sizing@ WindowStartupLocation="CenterScreen"
        Background="{DynamicResource oca.Bg}" Foreground="{DynamicResource oca.Ink}"
        FontFamily="Segoe UI" FontSize="12" UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">@icon@
  <Window.Resources>@resources@
  </Window.Resources>
  <DockPanel>
    <Border DockPanel.Dock="Top" Padding="16,11" BorderThickness="0,0,0,1"
            Background="{StaticResource oca.Surface}" BorderBrush="{StaticResource oca.Line}">
      <DockPanel>@logo@
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="@title@" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource oca.Ink}"/>@subtitle@
        </StackPanel>
      </DockPanel>
    </Border>
    <Border DockPanel.Dock="Bottom" Padding="16,10" BorderThickness="0,1,0,0"
            Background="{StaticResource oca.Surface}" BorderBrush="{StaticResource oca.Line}">
      <DockPanel LastChildFill="False">
        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">@footer_left@
        </StackPanel>
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">@footer_right@
        </StackPanel>
      </DockPanel>
    </Border>
    <Grid Margin="16,14,16,2">@body@
    </Grid>
  </DockPanel>
</Window>
"""


def _fill(template, values):
    # substituição em uma passada: marcadores dentro dos valores não são tocados
    return re.sub(u"@([a-z_]+)@", lambda m: values[m.group(1)], template)


def build_xaml(title, body, footer_right, subtitle=u"", footer_left=u"",
               size="S", height=None, dark=True):
    """XAML completo de uma janela no padrão OCA.

    title         nome do comando (aceita o __title__ com quebras de linha).
    subtitle      descrição curta - por padrão o __doc__ do botão.
    body          XAML do conteúdo (os mesmos x:Name que o código usa).
    footer_right  botões de confirmação: principal primeiro, Cancelar por último.
    footer_left   ações auxiliares ou "Voltar" (opcional).
    size          "S" 460 px, "M" 580 px (altura pelo conteúdo, tamanho fixo)
                  ou "L" 1160 px (redimensionável; informe height).
    dark          True (padrão OCA: sempre escuro); False só para testes.
    """
    name = command_name(title)
    if size == "L":
        sizing = u'Height="{}" MinWidth="900" MinHeight="460"'.format(height or 780)
    else:
        sizing = u'SizeToContent="Height" ResizeMode="NoResize"'
    uri = logo_uri()
    values = {
        "title": _xml(name),
        "width": u"{}".format(theme.SIZES[size]),
        "sizing": sizing,
        "icon": _ICON_XAML.replace(u"@uri@", _xml(uri)) if uri else u"",
        "logo": _LOGO_XAML.replace(u"@uri@", _xml(uri)) if uri else u"",
        "subtitle": _SUBTITLE_XAML.replace(u"@subtitle@", _xml(u" ".join(unicode(subtitle).split())))
                    if subtitle else u"",
        "resources": theme.resources(dark),
        "footer_left": footer_left,
        "footer_right": footer_right,
        "body": body,
    }
    return _fill(_WINDOW, values)
