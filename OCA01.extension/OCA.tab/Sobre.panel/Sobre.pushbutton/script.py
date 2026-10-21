# -*- coding: utf-8 -*-
# Sobre - central de informações institucionais e técnicas do plugin.
#
# Janela WPF no padrão OCA (lib/oca_ds): logo, empresa, nome do plugin, versão
# instalada, versões do Revit compatíveis (e a versão em uso), site, e-mail e
# aviso de direitos autorais. Tudo é local: a janela não acessa a internet nem
# abre o navegador. O site e o e-mail só abrem o navegador / cliente de e-mail
# se o usuário clicar neles.
#
# Para atualizar as informações, edite apenas o bloco CONFIGURAÇÃO abaixo.
#
# A descrição fica em comentário (e não em docstring) de propósito: o
# pyRevit 5.2 lê a docstring do módulo com .decode('utf-8'), que falha com
# acentos no IronPython e deixa o botão sem tooltip. O tooltip é só __doc__.

__title__ = u"Sobre"
__doc__ = u"Informações do OCA Automatizações em BIM: versão instalada, versões do Revit compatíveis, contato e direitos autorais."
__context__ = u"zero-doc"       # disponível mesmo sem modelo aberto

import os
import re

from pyrevit import forms
from oca_ds import build_xaml, alert_title, xml_escape

try:
    unicode
except NameError:           # IronPython 3 / CPython
    unicode = str

ALERT_TITLE = alert_title(__title__)          # "OCA · Sobre"


# ------------------------------------------------------------------
# CONFIGURAÇÃO
# ------------------------------------------------------------------
EMPRESA = u"Oliveira Cotta"
PLUGIN = u"OCA Automatizações em BIM"
VERSAO = u"13/10/2026"
REVIT_COMPATIVEIS = (u"2024", u"2025", u"2026")
SITE = u"https://www.oliveiracotta.com.br"
EMAIL = u"contato@oliveiracotta.com.br"
DIREITOS = (u"© 2026 Oliveira Cotta. Todos os direitos reservados. "
            u"Uso, cópia ou distribuição sem autorização expressa são proibidos.")

# Logo: arquivo de imagem local (.png, .jpg, .bmp...). Caminho absoluto ou
# relativo à pasta da extensão (OCA.extension). Se o arquivo não existir ou não
# puder ser lido, usa o logo do Design System; se esse também faltar, a janela
# abre sem logo.
LOGO = u"lib/oca_ds/assets/logo/logo.png"


# ------------------------------------------------------------------
# Caminhos e ambiente
# ------------------------------------------------------------------
HERE = os.path.dirname(os.path.abspath(__file__))


def extension_dir():
    """Pasta *.extension que contém este botão."""
    path = HERE
    while path and not path.lower().endswith(".extension"):
        parent = os.path.dirname(path)
        if parent == path:
            return os.path.normpath(os.path.join(HERE, u"..", u"..", u".."))
        path = parent
    return path


DEFAULT_LOGO = os.path.join(extension_dir(), u"lib", u"oca_ds", u"assets", u"logo", u"logo.png")


def logo_candidates():
    paths = []
    if LOGO:
        path = os.path.expandvars(os.path.expanduser(LOGO))
        if not os.path.isabs(path):
            path = os.path.join(extension_dir(), path)
        paths.append(os.path.normpath(path))
    paths.append(DEFAULT_LOGO)
    return paths


def load_bitmap(path):
    """BitmapImage do arquivo, já decodificado, ou None se não der para ler."""
    if not os.path.isfile(path):
        return None
    try:
        from System import Uri, UriKind
        from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.UriSource = Uri(path, UriKind.Absolute)
        bmp.CacheOption = BitmapCacheOption.OnLoad     # não prende o arquivo
        bmp.EndInit()
        bmp.Freeze()
        return bmp
    except Exception:
        return None


def revit_in_use():
    """(versão, build) do Revit em execução; (u"", u"") se não der para ler."""
    try:
        from pyrevit import HOST_APP
        return unicode(HOST_APP.version), unicode(HOST_APP.build)
    except Exception:
        return u"", u""


def pyrevit_in_use():
    try:
        from pyrevit.versionmgr import get_pyrevit_version
        return unicode(get_pyrevit_version().get_formatted())
    except Exception:
        return u""


def open_external(target):
    """Abre URL ou mailto: no programa padrão do Windows.

    UseShellExecute explícito: no .NET 8 (Revit 2025+) o padrão é False e
    Process.Start(url) falharia.
    """
    from System.Diagnostics import Process, ProcessStartInfo
    info = ProcessStartInfo(target)
    info.UseShellExecute = True
    Process.Start(info)


# ------------------------------------------------------------------
# Janela
# ------------------------------------------------------------------
# Ícones 16 x 16 em traço fino, desenhados em XAML (sem arquivos de imagem).
ICONS = {
    "version": u'<Path Data="M2.5,2.5 H8.2 L13.5,7.8 L7.8,13.5 L2.5,8.2 Z"/>'
               u'<Ellipse Canvas.Left="4.6" Canvas.Top="4.6" Width="2.2" Height="2.2"/>',
    "revit": u'<Path Data="M8,1.8 L13.8,5 V11 L8,14.2 L2.2,11 V5 Z M2.2,5 L8,8.2 L13.8,5 M8,8.2 V14.2"/>',
    "host": u'<Path Data="M1.8,2.8 H14.2 V10.8 H1.8 Z M5.5,13.6 H10.5 M8,10.8 V13.6"/>',
    "site": u'<Ellipse Canvas.Left="1.8" Canvas.Top="1.8" Width="12.4" Height="12.4"/>'
            u'<Ellipse Canvas.Left="5.3" Canvas.Top="1.8" Width="5.4" Height="12.4"/>'
            u'<Path Data="M1.8,8 H14.2"/>',
    "email": u'<Path Data="M1.8,3.5 H14.2 V12.5 H1.8 Z M1.8,3.5 L8,8.8 L14.2,3.5"/>',
    "copyright": u'<Ellipse Canvas.Left="1.8" Canvas.Top="1.8" Width="12.4" Height="12.4"/>'
                 u'<Path Data="M10.2,6.2 A2.8,2.8 0 1 0 10.2,9.8"/>',
}


def icon(name, margin=u"0,0,10,0"):
    return (u'<Canvas Width="16" Height="16" Margin="{}" VerticalAlignment="Center" '
            u'Style="{{StaticResource sobre.Icon}}">{}</Canvas>').format(margin, ICONS[name])


def row(icon_name, label, value_xaml, last=False):
    """Linha: ícone | rótulo | valor."""
    return u"""
        <Grid Margin="0,0,0,{m}">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="26"/>
            <ColumnDefinition Width="{{StaticResource oca.LabelWidth}}"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          {icon}
          <TextBlock Grid.Column="1" Text="{label}" Style="{{StaticResource oca.Label}}" Foreground="{{StaticResource oca.Ink2}}"/>
          <ContentControl Grid.Column="2" Focusable="False" IsTabStop="False" VerticalAlignment="Center">{value}
          </ContentControl>
        </Grid>""".format(m=0 if last else 10, icon=icon(icon_name, u"0"),
                          label=xml_escape(label), value=value_xaml)


def chips(values):
    return u"".join(
        u'<Border Style="{{StaticResource sobre.Chip}}"><TextBlock Text="{}"/></Border>'.format(xml_escape(v))
        for v in values)


def link(name, text, tooltip):
    return (u'<TextBlock TextTrimming="CharacterEllipsis">'
            u'<Hyperlink x:Name="{}" ToolTip="{}" Style="{{StaticResource sobre.Link}}">{}</Hyperlink>'
            u'</TextBlock>').format(name, xml_escape(tooltip), xml_escape(text))


RESOURCES = u"""
    <Style x:Key="sobre.Icon" TargetType="Canvas">
      <Style.Resources>
        <Style TargetType="Path">
          <Setter Property="Stroke" Value="{StaticResource oca.Accent}"/>
          <Setter Property="StrokeThickness" Value="1.25"/>
          <Setter Property="StrokeLineJoin" Value="Round"/>
          <Setter Property="StrokeStartLineCap" Value="Round"/>
          <Setter Property="StrokeEndLineCap" Value="Round"/>
        </Style>
        <Style TargetType="Ellipse">
          <Setter Property="Stroke" Value="{StaticResource oca.Accent}"/>
          <Setter Property="StrokeThickness" Value="1.25"/>
        </Style>
      </Style.Resources>
    </Style>
    <Style x:Key="sobre.Chip" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource oca.Surface2}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="{StaticResource oca.RadiusControl}"/>
      <Setter Property="Padding" Value="8,2"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
      <Setter Property="TextElement.FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="sobre.Link" TargetType="Hyperlink">
      <Setter Property="Foreground" Value="{StaticResource oca.Accent}"/>
      <Setter Property="TextDecorations" Value="{x:Null}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="TextDecorations" Value="Underline"/>
        </Trigger>
      </Style.Triggers>
    </Style>"""


def site_text(url):
    """https://www.exemplo.com.br/ -> www.exemplo.com.br"""
    text = url.split(u"://", 1)[-1]
    return text[:-1] if text.endswith(u"/") else text


BODY = u"""
  <StackPanel>
    <!-- identificação -->
    <Border Background="{StaticResource oca.Surface}" BorderBrush="{StaticResource oca.Line}" BorderThickness="1"
            CornerRadius="{StaticResource oca.RadiusGroup}" Padding="20,18" Margin="0,0,0,12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid x:Name="logo_box" Width="64" Height="64" Margin="0,0,18,0" VerticalAlignment="Center">
          <Image x:Name="img_logo" Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
        </Grid>
        <StackPanel Grid.Column="1" VerticalAlignment="Center">
          <TextBlock Text="@empresa@" FontSize="12" FontWeight="SemiBold" Foreground="{StaticResource oca.Accent}"/>
          <TextBlock Text="@plugin@" FontSize="19" FontWeight="SemiBold" Margin="0,1,0,0" TextWrapping="Wrap"
                     Foreground="{StaticResource oca.Ink}"/>
        </StackPanel>
      </Grid>
    </Border>

    <HeaderedContentControl Header="INFORMAÇÕES TÉCNICAS" Style="{StaticResource oca.Section}">
      <StackPanel>@row_version@@row_revit@
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="26"/>
            <ColumnDefinition Width="{StaticResource oca.LabelWidth}"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          @icon_host@
          <TextBlock Grid.Column="1" Text="Em uso" Style="{StaticResource oca.Label}" Foreground="{StaticResource oca.Ink2}"/>
          <WrapPanel Grid.Column="2" VerticalAlignment="Center">
            <TextBlock x:Name="tb_host" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="{StaticResource oca.Ink}"/>
            <Border x:Name="bd_status" CornerRadius="{StaticResource oca.RadiusControl}" Padding="7,1" BorderThickness="1"
                    VerticalAlignment="Center" Visibility="Collapsed">
              <TextBlock x:Name="tb_status" FontSize="11" FontWeight="SemiBold"/>
            </Border>
          </WrapPanel>
        </Grid>
      </StackPanel>
    </HeaderedContentControl>

    <HeaderedContentControl Header="CONTATO" Style="{StaticResource oca.Section}">
      <StackPanel>@row_site@@row_email@
      </StackPanel>
    </HeaderedContentControl>

    <Grid Margin="2,0,2,10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      @icon_copyright@
      <TextBlock Grid.Column="1" Text="@direitos@" Style="{StaticResource oca.Caption}" FontSize="10.5"
                 VerticalAlignment="Center"/>
    </Grid>
  </StackPanel>"""

FOOTER_LEFT = u"""
  <Button x:Name="btn_copy" Content="Copiar informações" Style="{StaticResource oca.Ghost}"
          ToolTip="Copia versão, compatibilidade e contato para colar em um e-mail de suporte."/>"""

FOOTER = u"""
  <Button x:Name="btn_close" Content="Fechar" Style="{StaticResource oca.Primary}" IsDefault="True"/>"""


def body_xaml():
    values = {
        u"empresa": xml_escape(EMPRESA),
        u"plugin": xml_escape(PLUGIN),
        u"direitos": xml_escape(DIREITOS),
        u"row_version": row(u"version", u"Versão instalada",
                            u'<TextBlock Text="{}" Style="{{StaticResource oca.Value}}"/>'.format(xml_escape(VERSAO))),
        u"row_revit": row(u"revit", u"Revit compatível",
                          u'<WrapPanel>{}</WrapPanel>'.format(chips(REVIT_COMPATIVEIS))),
        u"icon_host": icon(u"host", u"0"),
        u"row_site": row(u"site", u"Site", link(u"lnk_site", site_text(SITE), u"Abrir no navegador")),
        u"row_email": row(u"email", u"E-mail", link(u"lnk_email", EMAIL, u"Escrever no programa de e-mail padrão"),
                          last=True),
        u"icon_copyright": icon(u"copyright"),
    }
    # uma passada: o "@" do e-mail não é confundido com um marcador
    return re.sub(u"@([a-z_]+)@", lambda m: values[m.group(1)], BODY)


XAML = build_xaml(title=__title__, subtitle=__doc__, body=body_xaml(), footer_left=FOOTER_LEFT,
                  footer_right=FOOTER, size="M", resources=RESOURCES, logo=False)


class AboutWindow(forms.WPFWindow):
    def __init__(self):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)
        self._copy_timer = None
        self._load_logo()
        self._show_host()
        self.lnk_site.Click += self._on_site
        self.lnk_email.Click += self._on_email
        self.btn_copy.Click += self._on_copy
        self.btn_close.Click += self._on_close

    def _load_logo(self):
        from System.Windows import Visibility
        for path in logo_candidates():
            bmp = load_bitmap(path)
            if bmp is not None:
                self.img_logo.Source = bmp
                return
        self.logo_box.Visibility = Visibility.Collapsed

    def _show_host(self):
        from System.Windows import Visibility
        version, build = revit_in_use()
        pyrevit = pyrevit_in_use()
        parts = []
        if version:
            parts.append(u"Revit {}".format(version))
        if pyrevit:
            parts.append(u"pyRevit {}".format(pyrevit))
        self.tb_host.Text = u" · ".join(parts) or u"Não identificado"
        if build:
            self.tb_host.ToolTip = u"Build do Revit: {}".format(build)
        if not version:
            return
        ok = version in REVIT_COMPATIVEIS
        tone = u"Ok" if ok else u"Warn"
        self.bd_status.Background = self.FindResource(u"oca.{}Soft".format(tone))
        self.bd_status.BorderBrush = self.FindResource(u"oca.{}".format(tone))
        self.tb_status.Foreground = self.FindResource(u"oca.{}".format(tone))
        self.tb_status.Text = u"Compatível" if ok else u"Não homologada"
        if not ok:
            self.bd_status.ToolTip = (u"Esta versão do Revit não está na lista de versões compatíveis. "
                                      u"Os comandos podem funcionar, mas não foram validados nela.")
        self.bd_status.Visibility = Visibility.Visible

    def info_text(self):
        return u"\n".join([
            PLUGIN,
            EMPRESA,
            u"Versão instalada: {}".format(VERSAO),
            u"Revit compatível: {}".format(u", ".join(REVIT_COMPATIVEIS)),
            u"Em uso: {}".format(self.tb_host.Text),
            u"Site: {}".format(SITE),
            u"E-mail: {}".format(EMAIL),
        ])

    # -- eventos -----------------------------------------------------
    def _open(self, target, what):
        try:
            open_external(target)
        except Exception as ex:
            forms.alert(u"Não foi possível abrir o {}.".format(what),
                        sub_msg=u"Endereço: {}\n\n{}".format(target, ex), title=ALERT_TITLE)

    def _on_site(self, sender, args):
        self._open(SITE, u"navegador")

    def _on_email(self, sender, args):
        self._open(u"mailto:" + EMAIL, u"programa de e-mail")

    def _on_copy(self, sender, args):
        from System import TimeSpan
        from System.Windows import Clipboard
        from System.Windows.Threading import DispatcherTimer
        try:
            Clipboard.SetText(self.info_text())
        except Exception as ex:
            forms.alert(u"Não foi possível copiar as informações.",
                        sub_msg=u"A área de transferência está em uso por outro programa. "
                                u"Tente de novo.\n\n{}".format(ex), title=ALERT_TITLE)
            return
        self.btn_copy.Content = u"Copiado ✓"
        if self._copy_timer is None:
            self._copy_timer = DispatcherTimer()
            self._copy_timer.Interval = TimeSpan.FromSeconds(2)
            self._copy_timer.Tick += self._on_copy_reset
        self._copy_timer.Stop()
        self._copy_timer.Start()

    def _on_copy_reset(self, sender, args):
        self._copy_timer.Stop()
        self.btn_copy.Content = u"Copiar informações"

    def _on_close(self, sender, args):
        if self._copy_timer is not None:
            self._copy_timer.Stop()
        self.Close()


try:
    AboutWindow().ShowDialog()
except Exception as error:
    forms.alert(u"Não foi possível abrir a janela Sobre.", sub_msg=unicode(error), title=ALERT_TITLE)
