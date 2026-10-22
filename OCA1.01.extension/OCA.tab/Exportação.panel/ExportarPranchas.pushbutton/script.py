# -*- coding: utf-8 -*-
__title__ = u"Exportar\nPranchas"
__doc__ = u"Exporta pranchas em PDF e/ou DWG com nomenclatura por parâmetros e perfis por projeto."

# Compatível com IronPython 2.7 e CPython 3 do pyRevit. Revit 2022.1 ou superior (PDFExportOptions).

import os
import re
import json
import codecs

from pyrevit import revit, forms
from System.Collections.Generic import List
from System import Action
from System.Windows.Threading import DispatcherPriority
from System.Windows.Input import Keyboard, ModifierKeys
from Autodesk.Revit.UI.Events import (DialogBoxShowingEventArgs, TaskDialogShowingEventArgs,
                                      MessageBoxShowingEventArgs)
from System.Windows import Thickness, GridLength, Visibility
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Controls import Grid, ColumnDefinition, TextBox, ComboBox, Button, CheckBox
from oca_ds import build_xaml, alert_title
from Autodesk.Revit.DB import (
    ElementId, DWGExportOptions, PDFExportOptions, BuiltInParameter,
    FilteredElementCollector, ViewSheet, ViewSheetSet,
    ScheduleSheetInstance, SectionType, TextNote, Transaction, TransactionGroup,
    ColorDepthType, RasterQualityType, ZoomType, PaperPlacementType,
    PrintRange, PageOrientationType, HiddenLineViewsType, MarginType, ViewSet,
)

doc = revit.doc
ALERT_TITLE = alert_title(__title__)    # "OCA · Exportar Pranchas"
INVALID = re.compile(r'[\\/:*?"<>|]')

DEFAULT_DWG = u"<Padrão do Revit (sem configuração)>"
ALL_SHEETS = u"Todas as pranchas"
NATIVE_PRINTER = u"Revit nativo"
NO_PROFILE = u"(Nenhum perfil)"
NEW_PROFILE = u"+ Novo perfil..."
PRESET_FILE = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"),
                           "pyRevit", "ExportarPranchas_perfis.json")

# opções exibidas na janela -> enums da API
POSITIONING = [("Center", PaperPlacementType.Center), ("LowerLeft", PaperPlacementType.LowerLeft)]
ZOOM = [("Zoom", ZoomType.Zoom), ("FitToPage", ZoomType.FitToPage)]
RASTER = [("Low", RasterQualityType.Low), ("Medium", RasterQualityType.Medium),
          ("High", RasterQualityType.High), ("Presentation", RasterQualityType.Presentation)]
COLOR = [("Color", ColorDepthType.Color), ("GrayScale", ColorDepthType.GrayScale),
         ("BlackLine", ColorDepthType.BlackLine)]

# nomenclatura padrão (mesma ordem do diálogo de nome do Revit)
DEFAULT_RULES = [
    {"param": u"Nome do Arquivo", "prefix": u"", "suffix": u"", "sep": u"-"},
    {"param": u"Bloco", "prefix": u"", "suffix": u"", "sep": u"-"},
    {"param": u"Sheet Number", "prefix": u"", "suffix": u"", "sep": u"-"},
    {"param": u"Current Revision", "prefix": u"", "suffix": u"", "sep": u"-"},
    {"param": u"Sheet Name", "prefix": u"", "suffix": u"", "sep": u""},
]

# nomes alternativos -> parâmetros nativos (funciona em qualquer idioma do Revit)
BUILTIN_ALIASES = {
    "current revision": "SHEET_CURRENT_REVISION",
    "revisao atual": "SHEET_CURRENT_REVISION",
    u"revisão atual": "SHEET_CURRENT_REVISION",
    "current revision date": "SHEET_CURRENT_REVISION_DATE",
    "current revision description": "SHEET_CURRENT_REVISION_DESCRIPTION",
    "sheet issue date": "SHEET_ISSUE_DATE",
    "drawn by": "SHEET_DRAWN_BY",
    "checked by": "SHEET_CHECKED_BY",
    "approved by": "SHEET_APPROVED_BY",
}

# ----------------------------------------------------------------------
# Interface (WPF - Design System OCA, lib/oca_ds)

# combo editável no padrão OCA (o estilo de ComboBox do oca_ds não tem campo de texto)
EDIT_COMBO = u"""
    <Style x:Key="exp.EditCombo" TargetType="ComboBox" BasedOn="{StaticResource {x:Type ComboBox}}">
      <Setter Property="IsEditable" Value="True"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton Template="{StaticResource oca.ComboToggle}" Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
              <TextBox x:Name="PART_EditableTextBox" Margin="1,1,24,1" Height="Auto" VerticalAlignment="Center"
                       IsReadOnly="{Binding IsReadOnly, RelativeSource={RelativeSource TemplatedParent}}">
                <TextBox.Template>
                  <ControlTemplate TargetType="TextBox">
                    <ScrollViewer x:Name="PART_ContentHost" Focusable="False" Margin="6,0,0,0" VerticalAlignment="Center"/>
                  </ControlTemplate>
                </TextBox.Template>
              </TextBox>
              <Popup x:Name="PART_Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}"
                     AllowsTransparency="True" Focusable="False" PopupAnimation="Slide" VerticalOffset="2">
                <Border Background="{StaticResource oca.Surface}" BorderBrush="{StaticResource oca.FieldLine}" BorderThickness="1"
                        CornerRadius="{StaticResource oca.RadiusPanel}" Padding="3"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}"
                        MaxHeight="{TemplateBinding MaxDropDownHeight}">
                  <ScrollViewer SnapsToDevicePixels="True">
                    <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.7"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>"""

BODY = u"""
  <ScrollViewer VerticalScrollBarVisibility="Auto" MaxHeight="660" Padding="0,0,4,0">
    <StackPanel>

      <HeaderedContentControl Header="PERFIL (CONFIGURAÇÃO SALVA POR PROJETO)" Style="{StaticResource oca.Section}">
        <StackPanel>
          <ComboBox x:Name="cmb_profile" Margin="0,0,0,8"/>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="btn_prof_save" Content="Salvar" Style="{StaticResource oca.Compact}" Margin="0,0,6,0"/>
            <Button x:Name="btn_prof_saveas" Content="Salvar como..." Style="{StaticResource oca.Compact}" Margin="0,0,6,0"/>
            <Button x:Name="btn_prof_del" Content="Excluir" Style="{StaticResource oca.Compact}" Margin="0,0,16,0"/>
            <Button x:Name="btn_prof_import" Content="Importar..." Style="{StaticResource oca.Compact}" Margin="0,0,6,0"/>
            <Button x:Name="btn_prof_export" Content="Exportar..." Style="{StaticResource oca.Compact}"/>
          </StackPanel>
        </StackPanel>
      </HeaderedContentControl>

      <HeaderedContentControl Header="GERAL" Style="{StaticResource oca.Section}">
        <StackPanel>
          <TextBlock x:Name="tb_count" Style="{StaticResource oca.Value}" Margin="0,0,0,8"/>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="150"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="01 - Pasta para salvar" Style="{StaticResource oca.Label}"/>
            <TextBox x:Name="tb_folder" Grid.Column="1" Margin="0,0,6,0"/>
            <Button x:Name="btn_browse" Grid.Column="2" Content="Procurar..."/>
          </Grid>
          <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
            <CheckBox x:Name="cb_pdf" Content="09 - Exportar PDF" Margin="0,0,24,0"/>
            <CheckBox x:Name="cb_dwg" Content="10 - Exportar DWG" Margin="0"/>
          </StackPanel>
        </StackPanel>
      </HeaderedContentControl>

      <HeaderedContentControl Header="04 - NOMENCLATURA DO ARQUIVO" Style="{StaticResource oca.Section}">
        <StackPanel>
          <Grid Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="210"/><ColumnDefinition Width="85"/><ColumnDefinition Width="85"/><ColumnDefinition Width="50"/><ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Parâmetro" Style="{StaticResource oca.Caption}"/>
            <TextBlock Grid.Column="1" Text="Prefixo" Style="{StaticResource oca.Caption}"/>
            <TextBlock Grid.Column="2" Text="Sufixo" Style="{StaticResource oca.Caption}"/>
            <TextBlock Grid.Column="3" Text="Sep." Style="{StaticResource oca.Caption}" ToolTip="Separador colocado depois deste item"/>
          </Grid>
          <StackPanel x:Name="sp_rules"/>
          <Button x:Name="btn_addrule" Content="+ Adicionar parâmetro" Style="{StaticResource oca.Ghost}"
                  HorizontalAlignment="Left" Margin="0,4,0,0"/>
          <TextBlock Style="{StaticResource oca.Hint}" Margin="0,4,0,0"
                     Text="Digite qualquer nome de parâmetro ou escolha na lista (prancha ou Informações do Projeto). Deixe o parâmetro em branco para inserir texto fixo usando o Prefixo. Itens vazios são ignorados, e o separador só aparece entre itens preenchidos."/>
          <Border Style="{StaticResource oca.Card}" Margin="0,8,0,0">
            <StackPanel>
              <TextBlock Text="Exemplo (1ª prancha)" Style="{StaticResource oca.Caption}"/>
              <TextBlock x:Name="tb_preview" Style="{StaticResource oca.Mono}" TextWrapping="Wrap" Margin="0,2,0,0"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </HeaderedContentControl>

      <HeaderedContentControl x:Name="gp_pdf" Header="OPÇÕES DE PDF" Style="{StaticResource oca.Section}">
        <StackPanel>
          <Grid Margin="0,0,0,2">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock Text="Printer" Style="{StaticResource oca.Label}"/>
            <ComboBox x:Name="cmb_printer" Grid.Column="1"/>
          </Grid>
          <TextBlock x:Name="tb_printer_hint" Style="{StaticResource oca.Hint}" Margin="150,0,0,10"/>
          <CheckBox x:Name="cb_combined" Content="02 - Exportar em um único PDF"/>
          <Grid Margin="0,0,0,6">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock Text="03 - Nome combinado (opcional)" Style="{StaticResource oca.Label}"/>
            <TextBox x:Name="tb_combined" Grid.Column="1"/>
          </Grid>
          <Grid Margin="0,0,0,6">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock Text="05 - Positioning" Style="{StaticResource oca.Label}"/>
            <ComboBox x:Name="cmb_pos" Grid.Column="1"/>
          </Grid>
          <Grid Margin="0,0,0,6">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/><ColumnDefinition Width="80"/><ColumnDefinition Width="60"/></Grid.ColumnDefinitions>
            <TextBlock Text="06 - Zoom Type" Style="{StaticResource oca.Label}"/>
            <ComboBox x:Name="cmb_zoom" Grid.Column="1"/>
            <TextBlock Grid.Column="2" Text="Zoom (%)" Style="{StaticResource oca.Label}" HorizontalAlignment="Right"/>
            <TextBox x:Name="tb_zoompct" Grid.Column="3"/>
          </Grid>
          <Grid Margin="0,0,0,6">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="60"/><ColumnDefinition Width="24"/><ColumnDefinition Width="60"/></Grid.ColumnDefinitions>
            <TextBlock Text="Offset X / Y (Lower Left)" Style="{StaticResource oca.Label}"/>
            <TextBox x:Name="tb_xoff" Grid.Column="1"/>
            <TextBlock Grid.Column="2" Text="/" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            <TextBox x:Name="tb_yoff" Grid.Column="3"/>
          </Grid>
          <Grid Margin="0,0,0,6">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock Text="07 - Raster Quality" Style="{StaticResource oca.Label}"/>
            <ComboBox x:Name="cmb_raster" Grid.Column="1"/>
          </Grid>
          <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock Text="08 - Color Depth" Style="{StaticResource oca.Label}"/>
            <ComboBox x:Name="cmb_color" Grid.Column="1"/>
          </Grid>
          <CheckBox x:Name="cb_hide_ref" Content="Ocultar planos de referência"/>
          <CheckBox x:Name="cb_hide_tags" Content="Ocultar tags de vistas não referenciadas"/>
          <CheckBox x:Name="cb_hide_scope" Content="Ocultar scope boxes"/>
          <CheckBox x:Name="cb_hide_crop" Content="Ocultar limites de corte (crop boundaries)"/>
          <CheckBox x:Name="cb_use_raster" Content="Sempre usar raster" Margin="0"/>
        </StackPanel>
      </HeaderedContentControl>

      <HeaderedContentControl x:Name="gp_dwg" Header="OPÇÕES DE DWG" Style="{StaticResource oca.Section}">
        <StackPanel>
          <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock Text="11 - DWG Export Settings" Style="{StaticResource oca.Label}"/>
            <ComboBox x:Name="cmb_dwg" Grid.Column="1"/>
          </Grid>
          <CheckBox x:Name="cb_merged" Margin="0"
                    Content="Mesclar as views de cada prancha em um único DWG (sem xrefs)"/>
        </StackPanel>
      </HeaderedContentControl>

    </StackPanel>
  </ScrollViewer>"""

FOOTER = u"""
  <Button x:Name="btn_ok" Content="Exportar pranchas" Style="{StaticResource oca.Primary}" IsDefault="True"/>
  <Button x:Name="btn_cancel" Content="Cancelar" Margin="8,0,0,0" IsCancel="True"/>"""

XAML = build_xaml(title=__title__, subtitle=__doc__, body=BODY, footer_right=FOOTER,
                  size="M", width=620, resources=EDIT_COMBO)

PICK_BODY = u"""
  <StackPanel>
    <Grid Margin="0,0,0,8">
      <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <TextBlock Text="Buscar prancha" Style="{StaticResource oca.Label}"/>
      <TextBox x:Name="tb_search" Grid.Column="1"/>
    </Grid>
    <Grid Margin="0,0,0,10">
      <Grid.ColumnDefinitions><ColumnDefinition Width="150"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <TextBlock Text="Conjuntos de pranchas" Style="{StaticResource oca.Label}"/>
      <ComboBox x:Name="cmb_sets" Grid.Column="1"/>
    </Grid>
    <Border Height="420" Background="{StaticResource oca.Field}" BorderBrush="{StaticResource oca.FieldLine}"
            BorderThickness="1" CornerRadius="{StaticResource oca.RadiusControl}">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel x:Name="sp_sheets" Margin="6,4"/>
      </ScrollViewer>
    </Border>
    <TextBlock x:Name="tb_count" Style="{StaticResource oca.Value}" Margin="0,8,0,0"/>
    <TextBlock Style="{StaticResource oca.Hint}" Margin="0,2,0,8"
               Text="Dica: clique em uma prancha, segure Shift e clique em outra para marcar todo o intervalo."/>
    <StackPanel Orientation="Horizontal" Margin="0,0,0,12">
      <Button x:Name="btn_all" Content="Marcar todas" Style="{StaticResource oca.Compact}" Margin="0,0,6,0"/>
      <Button x:Name="btn_none" Content="Desmarcar todas" Style="{StaticResource oca.Compact}" Margin="0,0,6,0"/>
      <Button x:Name="btn_toggle" Content="Inverter seleção" Style="{StaticResource oca.Compact}"/>
    </StackPanel>
  </StackPanel>"""

PICK_FOOTER = u"""
  <Button x:Name="btn_ok" Content="Continuar" Style="{StaticResource oca.Primary}" IsDefault="True"/>
  <Button x:Name="btn_cancel" Content="Cancelar" Margin="8,0,0,0" IsCancel="True"/>"""

XAML_PICK = build_xaml(title=__title__ + u" · Pranchas", subtitle=u"Marque as pranchas que deseja exportar.",
                       body=PICK_BODY, footer_right=PICK_FOOTER, size="M")

NAME_BODY = u"""
  <StackPanel Margin="0,0,0,12">
    <TextBlock x:Name="tb_prompt" Style="{StaticResource oca.Label}" Margin="0,0,0,6"/>
    <TextBox x:Name="tb_name"/>
  </StackPanel>"""

NAME_FOOTER = u"""
  <Button x:Name="btn_ok" Content="Salvar perfil" Style="{StaticResource oca.Primary}" IsDefault="True"/>
  <Button x:Name="btn_cancel" Content="Cancelar" Margin="8,0,0,0" IsCancel="True"/>"""

XAML_NAME = build_xaml(title=__title__ + u" · Perfil", body=NAME_BODY, footer_right=NAME_FOOTER, size="S")

PROG_BODY = u"""
  <StackPanel Margin="0,0,0,12">
    <TextBlock x:Name="tb_status" Style="{StaticResource oca.Label}" Margin="0,0,0,8"/>
    <ProgressBar x:Name="pb" Height="10" Minimum="0">
      <ProgressBar.Template>
        <ControlTemplate TargetType="ProgressBar">
          <Border Background="{StaticResource oca.Surface2}" BorderBrush="{StaticResource oca.Line}"
                  BorderThickness="1" CornerRadius="{StaticResource oca.RadiusControl}">
            <Grid x:Name="PART_Track">
              <Border x:Name="PART_Indicator" HorizontalAlignment="Left" Background="{StaticResource oca.Accent}"
                      CornerRadius="{StaticResource oca.RadiusControl}"/>
            </Grid>
          </Border>
        </ControlTemplate>
      </ProgressBar.Template>
    </ProgressBar>
    <TextBlock x:Name="tb_count" Style="{StaticResource oca.Hint}" Margin="0,6,0,0"/>
  </StackPanel>"""

# sem botão próprio: o Cancelar da barra de status do Revit interrompe a exportação inteira
PROG_FOOTER = u"""
  <TextBlock Style="{StaticResource oca.Hint}" VerticalAlignment="Center"
             Text="Para cancelar, use Cancelar na barra de status do Revit."/>"""

XAML_PROG = build_xaml(title=__title__ + u" · Exportando",
                       subtitle=u"Aguarde. Cancelar no Revit interrompe todas as pranchas restantes.",
                       body=PROG_BODY, footer_right=PROG_FOOTER, size="S")


class ProgressWindow(forms.WPFWindow):
    def __init__(self, total):
        forms.WPFWindow.__init__(self, XAML_PROG, literal_string=True)
        self.total = max(total, 1)
        self.done = 0
        self.cancelled = False
        self._can_close = False
        self._text = u""
        # sem "sempre no topo": a janela fica presa à do Revit (fica na frente do modelo), mas os
        # avisos do Revit (ex.: "Exportação interrompida! Tem certeza...?") abrem na frente dela
        self.Topmost = False
        try:
            WindowInteropHelper(self).Owner = revit.uidoc.Application.MainWindowHandle
        except Exception:
            pass
        self.pb.Maximum = self.total
        self.Closing += self.on_closing

    def pump(self):
        try:
            self.Dispatcher.Invoke(Action(lambda: None), DispatcherPriority.Background)
        except Exception:
            pass

    def refresh(self):
        self.tb_status.Text = self._text
        self.pb.Value = self.done
        self.tb_count.Text = u"{} de {}".format(self.done, self.total)
        self.pump()

    def step(self, text):
        self._text = text
        self.refresh()

    def advance(self):
        self.done += 1
        self.refresh()

    def cancel_all(self):
        """Interrompe as pranchas restantes (Cancelar do Revit ou fechar esta janela)."""
        self.cancelled = True
        self._text = u"Cancelando as pranchas restantes..."
        self.refresh()

    def on_closing(self, sender, args):
        if not self._can_close:
            args.Cancel = True
            self.cancel_all()

    def finish(self):
        self._can_close = True
        self.Close()


class NameDialog(forms.WPFWindow):
    def __init__(self, prompt, default):
        forms.WPFWindow.__init__(self, XAML_NAME, literal_string=True)
        self.result = None
        self.tb_prompt.Text = prompt
        self.tb_name.Text = default or u""
        self.btn_ok.Click += self.ok_click
        self.btn_cancel.Click += self.cancel_click
        self.tb_name.Focus()
        self.tb_name.SelectAll()

    def ok_click(self, sender, args):
        txt = (self.tb_name.Text or u"").strip()
        if txt:
            self.result = txt
            self.Close()

    def cancel_click(self, sender, args):
        self.Close()


def ask_name(prompt, default=u""):
    dlg = NameDialog(prompt, default)
    dlg.show_dialog()
    return dlg.result


class SheetPicker(forms.WPFWindow):
    def __init__(self, sheets, sets, preselected):
        forms.WPFWindow.__init__(self, XAML_PICK, literal_string=True)
        self.result = None
        self.sets = sets
        self.items = []          # (checkbox, sheet, texto_minusculo, id)
        self.last_cb = None
        self._loading = True
        for s in sheets:
            label = u"{} - {}".format(s.SheetNumber, s.Name)
            cb = CheckBox()
            cb.Content = label
            cb.Margin = Thickness(4, 3, 4, 3)
            cb.IsChecked = s.Id.ToString() in preselected
            cb.Checked += self.count_changed
            cb.Unchecked += self.count_changed
            cb.Click += self.item_click
            self.sp_sheets.Children.Add(cb)
            self.items.append((cb, s, label.lower(), s.Id.ToString()))

        self.cmb_sets.ItemsSource = [ALL_SHEETS] + sorted(sets.keys(), key=lambda x: x.lower())
        self.cmb_sets.SelectedIndex = 0
        self.tb_search.TextChanged += self.apply_filter
        self.cmb_sets.SelectionChanged += self.apply_filter
        self.btn_all.Click += lambda s, a: self.set_visible(True)
        self.btn_none.Click += lambda s, a: self.set_visible(False)
        self.btn_toggle.Click += self.toggle_visible
        self.btn_ok.Click += self.ok_click
        self.btn_cancel.Click += self.cancel_click
        self._loading = False
        self.update_count()

    def visible_items(self):
        return [i for i in self.items if i[0].Visibility == Visibility.Visible]

    def apply_filter(self, sender, args):
        if self._loading:
            return
        text = (self.tb_search.Text or u"").strip().lower()
        sel = self.cmb_sets.SelectedItem
        allowed = None if (sel in (None, ALL_SHEETS)) else self.sets.get(sel, set())
        for cb, s, low, sid in self.items:
            ok = (not text or text in low) and (allowed is None or sid in allowed)
            cb.Visibility = Visibility.Visible if ok else Visibility.Collapsed
        self.update_count()

    def set_visible(self, value):
        self._loading = True
        for i in self.visible_items():
            i[0].IsChecked = value
        self._loading = False
        self.update_count()

    def toggle_visible(self, sender, args):
        self._loading = True
        for i in self.visible_items():
            i[0].IsChecked = not bool(i[0].IsChecked)
        self._loading = False
        self.update_count()

    def item_click(self, sender, args):
        """Shift + clique marca/desmarca o intervalo entre o último clique e este."""
        vis = [i[0] for i in self.visible_items()]
        shift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift
        if shift and self.last_cb is not None and self.last_cb in vis and sender in vis:
            a, b = vis.index(self.last_cb), vis.index(sender)
            state = bool(sender.IsChecked)
            self._loading = True
            for cb in vis[min(a, b):max(a, b) + 1]:
                cb.IsChecked = state
            self._loading = False
            self.update_count()
        self.last_cb = sender

    def count_changed(self, sender, args):
        if not self._loading:
            self.update_count()

    def update_count(self):
        n = len([i for i in self.items if i[0].IsChecked])
        self.tb_count.Text = u"{} de {} prancha(s) selecionada(s)".format(n, len(self.items))

    def ok_click(self, sender, args):
        chosen = [i[1] for i in self.items if i[0].IsChecked]
        if not chosen:
            forms.alert(u"Nenhuma prancha está marcada.", sub_msg=u"Marque pelo menos uma prancha para continuar.",
                        title=ALERT_TITLE)
            return
        self.result = chosen
        self.Close()

    def cancel_click(self, sender, args):
        self.Close()


def pick_sheets():
    sheets = []
    for s in FilteredElementCollector(doc).OfClass(ViewSheet):
        try:
            if s.IsPlaceholder:
                continue
        except Exception:
            pass
        sheets.append(s)
    sheets.sort(key=lambda x: x.SheetNumber.lower())
    if not sheets:
        forms.alert(u"Este projeto não tem pranchas.", sub_msg=u"Crie ao menos uma prancha antes de exportar.",
                    title=ALERT_TITLE)
        return None

    sets = {}
    for vs in FilteredElementCollector(doc).OfClass(ViewSheetSet):
        ids = set()
        try:
            for v in vs.Views:
                if isinstance(v, ViewSheet):
                    ids.add(v.Id.ToString())
        except Exception:
            pass
        sets[vs.Name] = ids

    pre = set()
    try:
        for eid in revit.uidoc.Selection.GetElementIds():
            if isinstance(doc.GetElement(eid), ViewSheet):
                pre.add(eid.ToString())
    except Exception:
        pass

    picker = SheetPicker(sheets, sets, pre)
    picker.show_dialog()
    return picker.result



# ----------------------------------------------------------------------
# perfis salvos
def load_store():
    """Lê os perfis salvos; se o arquivo estiver danificado, usa a cópia anterior (.bak)."""
    for path in (PRESET_FILE, PRESET_FILE + ".bak"):
        try:
            with codecs.open(path, "r", "utf-8-sig") as f:
                st = json.loads(f.read())
            if isinstance(st, dict):
                st.setdefault("profiles", {})
                return st
        except Exception:
            pass
    return {"profiles": {}, "last": None, "last_profile": None}


def xml_text(text):
    """Desfaz os escapes de XML (&amp;, &lt;, &#45;...)."""
    def ent(m):
        e = m.group(1)
        if e.startswith(u"#x"):
            return unichr_(int(e[2:], 16))
        if e.startswith(u"#"):
            return unichr_(int(e[1:]))
        return {u"amp": u"&", u"lt": u"<", u"gt": u">", u"quot": u'"', u"apos": u"'"}.get(e, m.group(0))
    return re.sub(u"&(#x[0-9a-fA-F]+|#[0-9]+|[a-z]+);", ent, text or u"")


def unichr_(n):
    try:
        return unichr(n)
    except NameError:       # CPython 3
        return chr(n)


def xml_tag(block, tag):
    """Conteúdo de <tag>...</tag> (vazio para <tag />)."""
    m = re.search(u"<{0}\\b[^>]*?(?:/>|>(.*?)</{0}>)".format(tag), block, re.S)
    return xml_text(m.group(1)) if (m and m.group(1)) else u""


def read_profile_xml(text, base):
    """Perfil .xml: lê a nomenclatura de Select Sheet Parameters.

    Cada <ParameterModel> vira uma regra: ParameterName, Prefix, Suffix e o
    separador colocado depois do item, gravado no atributo
    xml:space_x003D_preserve (ex.: Nome do Arquivo + "" + Bloco + "-" + Sheet Number)."""
    profiles = {}
    for pm in re.finditer(u"<Profile>(.*?)</Profile>", text, re.S):
        block = pm.group(1)
        name = xml_tag(block, u"Name") or base
        rules = []
        for section in (u"SelectSheetParameters", u"CustomFileNameParameters"):
            sm = re.search(u"<{0}>(.*?)</{0}>".format(section), block, re.S)
            if not sm:
                continue
            for m in re.finditer(u"<ParameterModel\\b([^>]*)>(.*?)</ParameterModel>", sm.group(1), re.S):
                sep = re.search(u'space[^=\\s]*\\s*=\\s*"([^"]*)"', m.group(1))
                rules.append({"param": xml_tag(m.group(2), u"ParameterName"),
                              "prefix": xml_tag(m.group(2), u"Prefix"),
                              "suffix": xml_tag(m.group(2), u"Suffix"),
                              "sep": xml_text(sep.group(1)) if sep else u""})
            if rules:
                break
        st = profile_settings(block)
        if rules:
            st["rules"] = rules
        if st:
            profiles[name] = st
    if not profiles:
        raise ValueError(u"nenhuma configuração reconhecida no XML")
    return profiles


# subseções do perfil .xml com tags de mesmo nome (IsCenter, RasterQuality...) que não são do PDF/DWG
XML_SKIP_SECTIONS = (u"DWF", u"NWC", u"IFC", u"IMG", u"SelectionSheets", u"SelectionViews",
                  u"SelectedProjectParameters", u"SelectSheetParameters", u"SelectViewParameters",
                  u"CustomFileNameParameters")


def profile_settings(block):
    """Opções de PDF/DWG do perfil .xml, nas chaves de estado desta janela.
    Só entram as tags presentes no XML; o resto da janela fica como está."""
    for sec in XML_SKIP_SECTIONS:
        block = re.sub(u"<{0}>.*?</{0}>|<{0}\\s*/>".format(sec), u"", block, flags=re.S)

    def has(tag):
        return re.search(u"<{}\\b".format(tag), block) is not None

    def flag(tag):
        return xml_tag(block, tag).strip().lower() == u"true"

    def number(tag):
        m = re.search(u"-?[0-9]+(?:[.,][0-9]+)?", xml_tag(block, tag))
        return m.group(0).replace(u",", u".") if m else None

    st = {}
    folder = xml_tag(block, u"CreateExportFolderPath").strip()
    if folder:
        st["folder"] = folder.rstrip(u"\\/") if len(folder) > 3 else folder
    for key, tag in (("pdf", u"IsPDFChecked"), ("dwg", u"IsDWGChecked"),
                     ("hide_ref", u"HidePlanes"), ("hide_scope", u"HideScopeBox"),
                     ("hide_tags", u"HideUnreferencedTags"), ("hide_crop", u"HideCropBoundaries"),
                     ("merged", u"DWG_MergedViews")):
        if has(tag):
            st[key] = flag(tag)
    if has(u"IsSeparateFile"):
        st["combined"] = not flag(u"IsSeparateFile")
    if has(u"IsVectorProcessing"):
        st["use_raster"] = not flag(u"IsVectorProcessing")
    if has(u"IsCenter"):
        st["positioning"] = "Center" if flag(u"IsCenter") else "LowerLeft"
    if has(u"IsFitToPage"):
        st["zoom"] = "FitToPage" if flag(u"IsFitToPage") else "Zoom"
    for key, tag in (("zoom_pct", u"Zoom"), ("xoff", u"OffsetX"), ("yoff", u"OffsetY")):
        val = number(tag)
        if val is not None:
            st[key] = val
    for key, tag, items in (("raster", u"RasterQuality", RASTER), ("color", u"Color", COLOR)):
        val = xml_tag(block, tag).strip().lower()
        for opt, _ in items:
            if opt.lower() == val:
                st[key] = opt
    if has(u"DWGSettingName"):
        dwg = xml_tag(block, u"DWGSettingName").strip()
        # só é aplicado se existir no projeto; vazio = padrão do Revit
        st["dwg_setup"] = dwg if dwg and dwg != u"No Export Setup" else DEFAULT_DWG
    if has(u"PrinterName"):        # impressora do perfil ("PDF24", "Revit Native"...)
        pr = xml_tag(block, u"PrinterName").strip()
        st["printer"] = NATIVE_PRINTER if (not pr or pr.lower().startswith(u"revit")) else pr
    if has(u"OCA_CombinedName"):   # tag própria deste botão
        st["combined_name"] = xml_tag(block, u"OCA_CombinedName")
    return st


def xml_esc(value):
    text = u"{}".format(value if value is not None else u"")
    for raw, esc in ((u"&", u"&amp;"), (u"<", u"&lt;"), (u">", u"&gt;"), (u'"', u"&quot;")):
        text = text.replace(raw, esc)
    return text


def write_profile_xml(path, name, st):
    """Grava o perfil em .xml no mesmo formato dos perfis .xml que o botão importa
    (o Importar... lê de volta exatamente o que foi gravado aqui)."""
    def b(v):
        return u"true" if v else u"false"

    params, combo = [], u""
    rules = st.get("rules") or []
    for i, r in enumerate(rules):
        sep = r.get("sep") or u""
        params.append(
            u'            <ParameterModel xml:space_x003D_preserve="{sep}">\n'
            u'              <ParameterName>{p}</ParameterName>\n'
            u'              <Prefix>{pre}</Prefix>\n'
            u'              <Suffix>{suf}</Suffix>\n'
            u'            </ParameterModel>'.format(sep=xml_esc(sep), p=xml_esc(r.get("param")),
                                                    pre=xml_esc(r.get("prefix")), suf=xml_esc(r.get("suffix"))))
        combo += (r.get("param") or u"") + (sep if i < len(rules) - 1 else u"")
    dwg = st.get("dwg_setup")
    dwg = u"" if (not dwg or dwg == DEFAULT_DWG) else dwg
    tags = [
        (u"CreateExportFolderPath", st.get("folder")),
        (u"DWGSettingName", dwg),
        (u"DWG_MergedViews", b(st.get("merged"))),
        (u"IsCenter", b(st.get("positioning") == "Center")),
        (u"OffsetX", u"{}mm".format(st.get("xoff") or 0)),
        (u"OffsetY", u"{}mm".format(st.get("yoff") or 0)),
        (u"IsFitToPage", b(st.get("zoom") == "FitToPage")),
        (u"Zoom", st.get("zoom_pct")),
        (u"IsVectorProcessing", b(not st.get("use_raster"))),
        (u"RasterQuality", st.get("raster")),
        (u"Color", st.get("color")),
        (u"HidePlanes", b(st.get("hide_ref"))),
        (u"HideScopeBox", b(st.get("hide_scope"))),
        (u"HideUnreferencedTags", b(st.get("hide_tags"))),
        (u"HideCropBoundaries", b(st.get("hide_crop"))),
        (u"IsSeparateFile", b(not st.get("combined"))),
        (u"IsPDFChecked", b(st.get("pdf"))),
        (u"IsDWGChecked", b(st.get("dwg"))),
        (u"PrinterName", u"Revit Native" if st.get("printer", NATIVE_PRINTER) == NATIVE_PRINTER
                         else st.get("printer")),
        (u"OCA_CombinedName", st.get("combined_name")),
    ]
    xml = (u'<?xml version="1.0" encoding="utf-8"?>\n'
           u'<Profiles>\n  <List>\n    <Profile>\n'
           u'      <Name>{name}</Name>\n      <IsCurrent>true</IsCurrent>\n      <Version>1</Version>\n'
           u'      <TemplateInfo>\n'
           u'        <SelectSheetParameters>\n'
           u'          <CombineParameterName>{combo}</CombineParameterName>\n'
           u'          <CombineParameters>\n{params}\n          </CombineParameters>\n'
           u'        </SelectSheetParameters>\n'
           u'{tags}\n'
           u'      </TemplateInfo>\n    </Profile>\n  </List>\n</Profiles>\n').format(
        name=xml_esc(name), combo=xml_esc(combo), params=u"\n".join(params),
        tags=u"\n".join(u"        <{0}>{1}</{0}>".format(t, xml_esc(v)) for t, v in tags))
    with codecs.open(path, "w", "utf-8") as f:
        f.write(xml)


def read_profiles_file(path):
    """Lê um perfil .xml (exportado por este botão ou outro perfil .xml no mesmo formato):
    nomenclatura, pasta e opções de PDF/DWG."""
    with codecs.open(path, "r", "utf-8-sig") as f:
        text = f.read()
    if u"<Profile>" not in text:
        raise ValueError(u"o arquivo não é um perfil de exportação (.xml)")
    return read_profile_xml(text, os.path.splitext(os.path.basename(path))[0])


def to_json(data):
    """JSON só com caracteres ASCII (acentos viram \\u00fa...).

    O json do IronPython falha com acentos quando ensure_ascii=True ("'unknown' codec
    can't decode byte 0xfa"), então os acentos são escapados aqui, sem passar por ele."""
    text = json.dumps(data, indent=2, ensure_ascii=False)
    if not isinstance(text, type(u"")):
        text = text.decode("utf-8")
    return re.sub(u"[^\\x00-\\x7f]", lambda m: u"\\u{:04x}".format(ord(m.group(0))), text)


def save_store(store):
    """Grava os perfis sem nunca deixar o arquivo cortado: monta o texto inteiro, grava num
    arquivo temporário e só então troca pelo atual (o anterior fica em .bak)."""
    tmp = PRESET_FILE + ".tmp"
    try:
        d = os.path.dirname(PRESET_FILE)
        if not os.path.isdir(d):
            os.makedirs(d)
        text = to_json(store)
        with codecs.open(tmp, "w", "utf-8") as f:
            f.write(text)
        if os.path.isfile(PRESET_FILE):
            try:
                with codecs.open(PRESET_FILE, "r", "utf-8-sig") as f:   # fecha antes de renomear
                    json.loads(f.read())
                keep = PRESET_FILE + ".bak"          # atual bom: vira a cópia anterior
            except Exception:
                keep = PRESET_FILE + ".corrompido"   # atual danificado: guardado à parte, sem virar .bak
            if os.path.isfile(keep):
                os.remove(keep)
            os.rename(PRESET_FILE, keep)
        os.rename(tmp, PRESET_FILE)
    except Exception as ex:
        try:
            if os.path.isfile(tmp):
                os.remove(tmp)
        except Exception:
            pass
        forms.alert(u"O perfil não foi salvo.", sub_msg=u"{}".format(ex), title=ALERT_TITLE)


# ----------------------------------------------------------------------
class ExportForm(forms.WPFWindow):
    def __init__(self, sheets, dwg_setups, param_names):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)
        self.result = None
        self.sheets = sheets
        self.param_names = param_names
        self.rows = []
        self._loading = True
        self.store = load_store()

        self.tb_count.Text = u"{} prancha(s) selecionada(s)".format(len(sheets))
        for combo, items in ((self.cmb_pos, POSITIONING), (self.cmb_zoom, ZOOM),
                             (self.cmb_raster, RASTER), (self.cmb_color, COLOR)):
            combo.ItemsSource = [i[0] for i in items]
        self.cmb_dwg.ItemsSource = [DEFAULT_DWG] + list(dwg_setups)
        self.dwg_setups = list(dwg_setups)
        self.printer_modes = pdf_printers()
        self.printers = [NATIVE_PRINTER] + sorted(self.printer_modes, key=lambda x: x.lower())
        self.cmb_printer.ItemsSource = self.printers
        self.cmb_printer.SelectedItem = NATIVE_PRINTER

        self.set_state(self.default_state())
        if self.store.get("last"):
            self.set_state(self.store["last"])
        self.refresh_profiles(self.store.get("last_profile"))

        self.btn_browse.Click += self.browse_click
        self.btn_ok.Click += self.ok_click
        self.btn_cancel.Click += self.cancel_click
        self.btn_addrule.Click += self.addrule_click
        self.btn_prof_save.Click += self.prof_save_click
        self.btn_prof_saveas.Click += self.prof_saveas_click
        self.btn_prof_del.Click += self.prof_del_click
        self.btn_prof_import.Click += self.prof_import_click
        self.btn_prof_export.Click += self.prof_export_click
        self.cmb_profile.SelectionChanged += self.profile_changed
        for cb in (self.cb_pdf, self.cb_dwg, self.cb_combined):
            cb.Checked += self.update_state
            cb.Unchecked += self.update_state
        self.cmb_pos.SelectionChanged += self.update_state
        self.cmb_zoom.SelectionChanged += self.update_state
        self.cmb_printer.SelectionChanged += self.printer_changed
        self._loading = False
        self.update_state(None, None)
        self.update_preview(None, None)

    # ---------------- estado ----------------
    def default_state(self):
        return {
            "folder": u"C:\\Exportacoes_Revit", "pdf": True, "dwg": True,
            "rules": DEFAULT_RULES, "combined": False, "combined_name": u"Pranchas_Combinadas",
            "positioning": "Center", "zoom": "Zoom", "zoom_pct": u"100", "xoff": u"0", "yoff": u"0",
            "raster": "High", "color": "Color",
            "hide_ref": True, "hide_tags": True, "hide_scope": True, "hide_crop": True,
            "use_raster": False, "printer": NATIVE_PRINTER,
            "dwg_setup": self.dwg_setups[0] if self.dwg_setups else DEFAULT_DWG, "merged": True,
        }

    def collect_state(self):
        rules = []
        for r in self.rows:
            rule = {"param": r["cmb"].Text or u"", "prefix": r["pre"].Text or u"",
                    "suffix": r["suf"].Text or u"", "sep": r["sep"].Text or u""}
            if rule["param"].strip() or rule["prefix"] or rule["suffix"]:
                rules.append(rule)
        return {
            "folder": (self.tb_folder.Text or u"").strip(),
            "pdf": bool(self.cb_pdf.IsChecked), "dwg": bool(self.cb_dwg.IsChecked),
            "rules": rules,
            "combined": bool(self.cb_combined.IsChecked),
            "combined_name": (self.tb_combined.Text or u"").strip(),
            "positioning": self.cmb_pos.SelectedItem, "zoom": self.cmb_zoom.SelectedItem,
            "zoom_pct": self.tb_zoompct.Text or u"", "xoff": self.tb_xoff.Text or u"",
            "yoff": self.tb_yoff.Text or u"",
            "raster": self.cmb_raster.SelectedItem, "color": self.cmb_color.SelectedItem,
            "hide_ref": bool(self.cb_hide_ref.IsChecked), "hide_tags": bool(self.cb_hide_tags.IsChecked),
            "hide_scope": bool(self.cb_hide_scope.IsChecked), "hide_crop": bool(self.cb_hide_crop.IsChecked),
            "use_raster": bool(self.cb_use_raster.IsChecked),
            "printer": self.cmb_printer.SelectedItem or NATIVE_PRINTER,
            "dwg_setup": self.cmb_dwg.SelectedItem or DEFAULT_DWG, "merged": bool(self.cb_merged.IsChecked),
        }

    def set_state(self, st):
        """Aplica só as chaves presentes (permite importar perfis parciais, ex.: só nomenclatura)."""
        was = self._loading
        self._loading = True
        if st.get("folder"):
            self.tb_folder.Text = st["folder"]
        for key, cb in (("pdf", self.cb_pdf), ("dwg", self.cb_dwg), ("combined", self.cb_combined),
                        ("hide_ref", self.cb_hide_ref), ("hide_tags", self.cb_hide_tags),
                        ("hide_scope", self.cb_hide_scope), ("hide_crop", self.cb_hide_crop),
                        ("use_raster", self.cb_use_raster), ("merged", self.cb_merged)):
            if key in st:
                cb.IsChecked = bool(st[key])
        for key, tb in (("combined_name", self.tb_combined), ("zoom_pct", self.tb_zoompct),
                        ("xoff", self.tb_xoff), ("yoff", self.tb_yoff)):
            if key in st:
                tb.Text = u"{}".format(st[key])
        for combo, items, key in ((self.cmb_pos, POSITIONING, "positioning"), (self.cmb_zoom, ZOOM, "zoom"),
                                  (self.cmb_raster, RASTER, "raster"), (self.cmb_color, COLOR, "color")):
            if st.get(key) in [i[0] for i in items]:
                combo.SelectedItem = st[key]
        if "printer" in st:   # impressora que não existe neste computador -> Revit nativo
            pr = st["printer"]
            match = [p for p in self.printers if p.lower() == (pr or u"").lower()]
            self.cmb_printer.SelectedItem = match[0] if match else NATIVE_PRINTER
        if "dwg_setup" in st:
            dw = st["dwg_setup"]
            if dw == DEFAULT_DWG or dw in self.dwg_setups:
                self.cmb_dwg.SelectedItem = dw
            elif self.cmb_dwg.SelectedItem is None:
                self.cmb_dwg.SelectedItem = DEFAULT_DWG
        if "rules" in st:
            rules = st["rules"] if st["rules"] is not None else DEFAULT_RULES
            self.sp_rules.Children.Clear()
            self.rows = []
            for r in rules:
                self.add_rule_row(r.get("param", u""), r.get("prefix", u""), r.get("suffix", u""), r.get("sep", u""))
        self._loading = was
        if not was:
            self.update_state(None, None)
            self.update_preview(None, None)

    # ---------------- linhas de nomenclatura ----------------
    def add_rule_row(self, param=u"", prefix=u"", suffix=u"", sep=u"-"):
        grid = Grid()
        grid.Margin = Thickness(0, 2, 0, 2)
        for w in (210, 85, 85, 50, 28, 28, 28):
            cd = ColumnDefinition()
            cd.Width = GridLength(w)
            grid.ColumnDefinitions.Add(cd)

        cmb = ComboBox()
        cmb.Style = self.FindResource("exp.EditCombo")
        cmb.IsEditable = True
        cmb.ItemsSource = self.param_names
        cmb.Text = param
        cmb.Margin = Thickness(0, 0, 4, 0)
        pre, suf, sp = TextBox(), TextBox(), TextBox()
        for tb, val in ((pre, prefix), (suf, suffix), (sp, sep)):
            tb.Text = val
            tb.Margin = Thickness(0, 0, 4, 0)
            tb.TextChanged += self.update_preview
        cmb.SelectionChanged += self.update_preview
        cmb.KeyUp += self.update_preview
        cmb.LostFocus += self.update_preview

        row = {"panel": grid, "cmb": cmb, "pre": pre, "suf": suf, "sep": sp}
        btns = []
        for text, tip, handler in ((u"\u25B2", u"Subir", lambda s, a, r=row: self.move_row(r, -1)),
                                   (u"\u25BC", u"Descer", lambda s, a, r=row: self.move_row(r, 1)),
                                   (u"\u2715", u"Remover", lambda s, a, r=row: self.remove_row(r))):
            b = Button()
            b.Style = self.FindResource("oca.Compact")
            b.Padding = Thickness(0)
            b.Margin = Thickness(2, 0, 0, 0)
            b.Content = text
            b.ToolTip = tip
            b.Click += handler
            btns.append(b)
        for i, ctrl in enumerate([cmb, pre, suf, sp] + btns):
            Grid.SetColumn(ctrl, i)
            grid.Children.Add(ctrl)
        self.rows.append(row)
        self.sp_rules.Children.Add(grid)

    def rebuild_rows(self):
        self.sp_rules.Children.Clear()
        for r in self.rows:
            self.sp_rules.Children.Add(r["panel"])
        self.update_preview(None, None)

    def move_row(self, row, delta):
        i = self.rows.index(row)
        j = i + delta
        if 0 <= j < len(self.rows):
            self.rows[i], self.rows[j] = self.rows[j], self.rows[i]
            self.rebuild_rows()

    def remove_row(self, row):
        self.rows.remove(row)
        self.rebuild_rows()

    def addrule_click(self, sender, args):
        self.add_rule_row(u"", u"", u"", u"-")
        self.update_preview(None, None)

    # ---------------- eventos ----------------
    def update_state(self, sender, args):
        if self._loading:
            return
        self.gp_pdf.IsEnabled = bool(self.cb_pdf.IsChecked)
        self.gp_dwg.IsEnabled = bool(self.cb_dwg.IsChecked)
        self.tb_combined.IsEnabled = bool(self.cb_combined.IsChecked)
        self.tb_zoompct.IsEnabled = (self.cmb_zoom.SelectedItem == "Zoom")
        native = (self.cmb_printer.SelectedItem or NATIVE_PRINTER) == NATIVE_PRINTER
        lower_left = (self.cmb_pos.SelectedItem == "LowerLeft")
        self.tb_xoff.IsEnabled = lower_left and native   # deslocamento só no exportador nativo
        self.tb_yoff.IsEnabled = lower_left and native
        if native:
            self.tb_printer_hint.Text = (u"Exportador de PDF nativo do Revit." if len(self.printers) > 1 else
                                         u"Exportador de PDF nativo do Revit. Instale um programa de "
                                         u"impressão em PDF para habilitar a opção de selecionar a "
                                         u"impressora desejada.")
        elif not self.printer_modes.get(self.cmb_printer.SelectedItem):
            self.tb_printer_hint.Text = (u"O PDF24 está instalado, mas falta criar a impressora \"OCA PDF24\" "
                                         u"(feita a partir do PDF24, uma vez por computador). Ao exportar, o "
                                         u"script oferece criar; o Windows pede permissão de administrador.")
        else:
            device = self.printer_modes.get(self.cmb_printer.SelectedItem) or u""
            self.tb_printer_hint.Text = (u"O Revit imprime cada folha na impressora \"{}\" e o PDF24 salva o "
                                         u"PDF sozinho, sem abrir janelas. O papel é escolhido pelo tamanho "
                                         u"de cada folha (A4 a A0).".format(device))

    def printer_changed(self, sender, args):
        self.update_state(sender, args)
        if self._loading:
            return
        if self.cmb_printer.SelectedItem == PDF24_LABEL and not self.printer_modes.get(PDF24_LABEL):
            # a pergunta abre depois que a lista fecha (não dá para trocar a seleção aqui dentro)
            self.Dispatcher.BeginInvoke(Action(self.offer_printer_install))

    def offer_printer_install(self):
        """PDF24 escolhido sem a impressora "OCA PDF24": oferece instalar agora."""
        ans = forms.alert(
            u"Falta criar a impressora \"OCA PDF24\" neste computador.",
            sub_msg=u"O PDF24 já está instalado. A impressora \"OCA PDF24\" é uma impressora do PDF24 "
                    u"configurada para salvar o PDF de cada folha sozinha, sem abrir janelas. Ela é criada "
                    u"uma vez por computador; o Windows pede permissão de administrador e abre uma janela "
                    u"do instalador (pressione Enter nela ao final). Criar agora?",
            options=[u"Instalar a impressora OCA PDF24", u"Agora não"], title=ALERT_TITLE)
        ok = False
        if ans == u"Instalar a impressora OCA PDF24":
            try:
                ok = install_oca_printer()
            except Exception as ex:
                forms.alert(u"A impressora OCA PDF24 não foi instalada.", sub_msg=u"{}".format(ex),
                            title=ALERT_TITLE)
        self.printer_modes = pdf_printers()
        self._loading = True
        self.cmb_printer.SelectedItem = PDF24_LABEL if ok else NATIVE_PRINTER
        self._loading = False
        self.update_state(None, None)
        if ans == u"Instalar a impressora OCA PDF24" and not ok:
            forms.alert(u"A impressora OCA PDF24 não ficou instalada.",
                        sub_msg=u"Confira a mensagem da janela do instalador. O PDF segue pelo Revit nativo.",
                        title=ALERT_TITLE)

    def update_preview(self, sender, args):
        if self._loading:
            return
        try:
            rules = self.collect_state()["rules"]
            ex = build_names(self.sheets[:1], rules)[0]
            self.tb_preview.Text = ex      # o título "Exemplo (1ª prancha)" já está no cartão
        except Exception as e:
            self.tb_preview.Text = u"Não foi possível gerar o exemplo: {}".format(e)

    def browse_click(self, sender, args):
        folder = forms.pick_folder(title=u"Pasta para salvar")
        if folder:
            self.tb_folder.Text = folder

    def cancel_click(self, sender, args):
        self.Close()

    # ---------------- perfis ----------------
    # a lista de perfis começa com "(Nenhum perfil)" e termina com "+ Novo perfil..."
    def current_profile(self):
        """Perfil salvo selecionado, ou None ("(Nenhum perfil)" / "+ Novo perfil...")."""
        name = self.cmb_profile.SelectedItem
        return name if (name and name in self.store["profiles"]) else None

    def refresh_profiles(self, select=None):
        was = self._loading
        self._loading = True
        names = sorted(self.store["profiles"].keys(), key=lambda x: x.lower())
        self.cmb_profile.ItemsSource = [NO_PROFILE] + names + [NEW_PROFILE]
        self.cmb_profile.SelectedItem = select if (select and select in names) else NO_PROFILE
        self._prev_profile = self.cmb_profile.SelectedItem
        self._loading = was

    def profile_changed(self, sender, args):
        if self._loading:
            return
        name = self.cmb_profile.SelectedItem
        if name == NEW_PROFILE:
            # a pergunta do nome abre depois que a lista fecha (não dá para trocar a seleção aqui dentro)
            self.Dispatcher.BeginInvoke(Action(self.create_profile))
            return
        self._prev_profile = name
        if name in self.store["profiles"]:
            self.set_state(self.store["profiles"][name])
        # "(Nenhum perfil)": mantém os campos como estão, só sem perfil associado

    def create_profile(self):
        """"+ Novo perfil...": cria um perfil com as configurações atuais da janela."""
        if not self.save_profile_as(u""):
            self.refresh_profiles(self._prev_profile)   # cancelou: volta para o que estava

    def save_profile_as(self, default):
        name = ask_name(u"Nome do novo perfil (ex.: nome do projeto):", default)
        if not name:
            return False
        if name in (NO_PROFILE, NEW_PROFILE):
            forms.alert(u"Esse nome é reservado.", sub_msg=u"Escolha outro nome para o perfil.", title=ALERT_TITLE)
            return False
        if name in self.store["profiles"] and not forms.alert(
                u"Já existe um perfil '{}'. Substituir?".format(name),
                sub_msg=u"Ele passa a ter as configurações atuais da janela.", yes=True, no=True, title=ALERT_TITLE):
            return False
        self.store["profiles"][name] = self.collect_state()
        self.store["last_profile"] = name
        save_store(self.store)
        self.refresh_profiles(name)
        return True

    def prof_saveas_click(self, sender, args):
        self.save_profile_as(self.current_profile() or u"")

    def prof_save_click(self, sender, args):
        name = self.current_profile()
        if not name:
            self.save_profile_as(u"")
            return
        self.store["profiles"][name] = self.collect_state()
        save_store(self.store)

    def prof_del_click(self, sender, args):
        name = self.current_profile()
        if name and forms.alert(u"Excluir o perfil '{}'?".format(name),
                                sub_msg=u"O perfil sai da lista deste computador. Não dá para desfazer.",
                                options=[u"Excluir", u"Cancelar"], title=ALERT_TITLE) == u"Excluir":
            self.store["profiles"].pop(name, None)
            if self.store.get("last_profile") == name:
                self.store["last_profile"] = None
            save_store(self.store)
            self.refresh_profiles()

    def prof_import_click(self, sender, args):
        path = forms.pick_file(files_filter=u"Perfil de exportação (*.xml)|*.xml", title=u"Importar perfil")
        if not path:
            return
        try:
            incoming = read_profiles_file(path)
        except Exception as ex:
            forms.alert(u"O arquivo de perfil é inválido.", sub_msg=u"{}".format(ex), title=ALERT_TITLE)
            return
        first = None
        for name, st in incoming.items():
            new = name
            if new in self.store["profiles"]:
                new = name + u" (importado)"
                n = 2
                while new in self.store["profiles"]:
                    new = u"{} (importado {})".format(name, n)
                    n += 1
            self.store["profiles"][new] = st
            first = first or new
        save_store(self.store)
        self.refresh_profiles(first)
        applied = self.store["profiles"][first]
        self.set_state(applied)
        note = u""
        dw = applied.get("dwg_setup")
        if dw and dw != DEFAULT_DWG and dw not in self.dwg_setups:
            note = (u"\n\nA configuração de DWG '{}' não existe neste projeto; "
                    u"ficou a que já estava selecionada.".format(dw))
        forms.alert(u"{} perfil(is) importado(s).".format(len(incoming)),
                    sub_msg=u"O perfil '{}' foi aplicado.{}".format(first, note), title=ALERT_TITLE)

    def prof_export_click(self, sender, args):
        name = u"{}".format(self.current_profile() or u"perfil_exportacao")
        path = forms.save_file(files_filter=u"Perfil de exportação (*.xml)|*.xml",
                               default_name=clean(name) or u"perfil_exportacao", title=u"Exportar perfil")
        if not path:
            return
        if not path.lower().endswith(".xml"):
            path += ".xml"
        # o perfil leva o nome dado ao arquivo (é com esse nome que ele aparece ao importar)
        name = os.path.splitext(os.path.basename(path))[0] or name
        try:
            write_profile_xml(path, name, self.collect_state())
            forms.alert(u"Perfil exportado.", sub_msg=path, title=ALERT_TITLE)
        except Exception as ex:
            forms.alert(u"O perfil não foi exportado.", sub_msg=u"{}".format(ex), title=ALERT_TITLE)

    # ---------------- confirmar ----------------
    def ok_click(self, sender, args):
        # PDF24 escolhido (ex.: vindo do perfil salvo) sem a impressora "OCA PDF24": oferece instalar
        if bool(self.cb_pdf.IsChecked) and self.cmb_printer.SelectedItem == PDF24_LABEL \
                and not self.printer_modes.get(PDF24_LABEL):
            self.offer_printer_install()
            if not self.printer_modes.get(PDF24_LABEL):
                return               # não instalou: a janela continua aberta, já em "Revit nativo"
        st = self.collect_state()
        if not st["folder"]:
            forms.alert(u"A pasta para salvar está vazia.", sub_msg=u"Digite a pasta ou use Procurar...",
                        title=ALERT_TITLE)
            return
        if not (st["pdf"] or st["dwg"]):
            forms.alert(u"Nenhum formato está marcado.", sub_msg=u"Marque PDF, DWG ou os dois.", title=ALERT_TITLE)
            return
        if not st["rules"]:
            forms.alert(u"A nomenclatura está vazia.", sub_msg=u"Adicione pelo menos um item na nomenclatura.",
                        title=ALERT_TITLE)
            return
        try:
            zoom_pct = int(float(st["zoom_pct"].replace(",", ".")))
            xoff = float(st["xoff"].replace(",", "."))
            yoff = float(st["yoff"].replace(",", "."))
        except Exception:
            forms.alert(u"Zoom (%) e Offsets precisam ser números.", sub_msg=u"Ex.: 100, 0 ou 0,5.",
                        title=ALERT_TITLE)
            return

        self.store["last"] = st
        self.store["last_profile"] = self.current_profile()
        save_store(self.store)

        res = dict(st)
        res.update({
            "zoom_pct": zoom_pct, "xoff": xoff, "yoff": yoff,
            "positioning": dict(POSITIONING)[st["positioning"]],
            "zoom": dict(ZOOM)[st["zoom"]],
            "raster": dict(RASTER)[st["raster"]],
            "color": dict(COLOR)[st["color"]],
            "dwg_setup": "" if st["dwg_setup"] == DEFAULT_DWG else st["dwg_setup"],
        })
        self.result = res
        self.Close()


# ----------------------------------------------------------------------
# nomes de arquivo
def clean(text):
    return INVALID.sub("_", text).strip().rstrip(".")


def part_value(sheet, token):
    """Valor de um parâmetro: prancha primeiro, depois Informações do Projeto."""
    t = token.strip()
    tl = t.lower()
    if tl in ("sheet name", "nome da folha", "nome da prancha"):
        return sheet.Name
    if tl in ("sheet number", "numero da folha", u"número da folha", "numero da prancha", u"número da prancha"):
        return sheet.SheetNumber
    cands = []
    p = sheet.LookupParameter(t)
    if p is not None:
        cands.append(p)
    bip_name = BUILTIN_ALIASES.get(tl)
    if bip_name and hasattr(BuiltInParameter, bip_name):
        bp = sheet.get_Parameter(getattr(BuiltInParameter, bip_name))
        if bp is not None:
            cands.append(bp)
    try:
        pp = doc.ProjectInformation.LookupParameter(t)
        if pp is not None:
            cands.append(pp)
    except Exception:
        pass
    for c in cands:
        v = c.AsString() or c.AsValueString()
        if v:
            return v
    return u""


def render_name(sheet, rules):
    """Cada regra: parâmetro + prefixo + sufixo + separador (usado só entre itens preenchidos).
    Parâmetro em branco = texto fixo (prefixo + sufixo)."""
    pieces = []
    for r in rules:
        param = (r.get("param") or u"").strip()
        pre = r.get("prefix") or u""
        suf = r.get("suffix") or u""
        sep = r.get("sep")
        if sep is None:
            sep = u""
        if param:
            val = (part_value(sheet, param) or u"").strip()
            if not val:
                continue
            text = pre + val + suf
        else:
            text = pre + suf
            if not text:
                continue
        pieces.append((text, sep))
    name = u""
    for i, (text, sep) in enumerate(pieces):
        name += text
        if i < len(pieces) - 1:
            name += sep
    return clean(name)


def build_names(sheets, rules):
    names, used = [], {}
    for s in sheets:
        base = render_name(s, rules) or clean(s.SheetNumber)
        n = used.get(base.lower(), 0) + 1
        used[base.lower()] = n
        names.append(base if n == 1 else u"{}_{}".format(base, n))
    return names


def sheet_param_names(sheets):
    names = set([u"Sheet Name", u"Sheet Number", u"Current Revision"])
    for s in sheets[:3]:
        for p in s.Parameters:
            try:
                names.add(p.Definition.Name)
            except Exception:
                pass
    try:
        for p in doc.ProjectInformation.Parameters:
            names.add(p.Definition.Name)
    except Exception:
        pass
    return sorted(names, key=lambda x: x.lower())


# ----------------------------------------------------------------------
# exportação
def ids_of(sheets):
    ids = List[ElementId]()
    for s in sheets:
        ids.Add(s.Id)
    return ids


def pdf_options(o, file_name, combine):
    p = PDFExportOptions()
    p.Combine = combine
    p.FileName = file_name
    p.ColorDepth = o["color"]
    p.RasterQuality = o["raster"]
    p.ZoomType = o["zoom"]
    p.ZoomPercentage = o["zoom_pct"]
    p.PaperPlacement = o["positioning"]
    p.OriginOffsetX = o["xoff"]
    p.OriginOffsetY = o["yoff"]
    p.HideReferencePlane = o["hide_ref"]
    p.HideUnreferencedViewTags = o["hide_tags"]
    p.HideScopeBoxes = o["hide_scope"]
    p.HideCropBoundaries = o["hide_crop"]
    p.AlwaysUseRaster = o["use_raster"]
    return p


# ----------------------------------------------------------------------
# fontes: texto passando para fora das tabelas em alguns computadores
#
# O PDF é gerado com as fontes instaladas no computador que exporta. Se uma fonte usada
# nas folhas não estiver instalada, o Windows usa outra (normalmente mais larga) e o texto
# das tabelas passa das células. Antes de exportar, o script confere as fontes das tabelas
# e das notas de texto das folhas escolhidas e avisa se faltar alguma.
EXPORT_ANYWAY = u"Exportar sem trocar a fonte"
CANCEL_FONTS = u"Cancelar a exportação"
# fontes que todo Windows tem: a que falta é trocada por uma delas SÓ durante a exportação do PDF
FALLBACK_FONTS = (u"Arial", u"Tahoma")
USE_FONT = u"Trocar por {} nesta exportação (recomendado)"
USE_FONT_ALT = u"Trocar por {} nesta exportação"


def net_import(module, name, assemblies):
    """Importa um tipo .NET. No Revit 2025+ (.NET 8) algumas bibliotecas não vêm carregadas
    (ex.: o registro do Windows fica em Microsoft.Win32.Registry.dll): carrega e tenta de novo."""
    try:
        return getattr(__import__(module, fromlist=[name]), name)
    except Exception:
        pass
    import clr
    for asm in assemblies:
        try:
            clr.AddReference(asm)
        except Exception:
            pass
    return getattr(__import__(module, fromlist=[name]), name)


def win_registry():
    return net_import("Microsoft.Win32", "Registry", ("Microsoft.Win32.Registry", "mscorlib"))


def installed_fonts():
    """Nomes das fontes instaladas neste computador (para todos os usuários e só para este)."""
    names = set()
    try:
        InstalledFontCollection = net_import("System.Drawing.Text", "InstalledFontCollection",
                                             ("System.Drawing.Common", "System.Drawing"))
        coll = InstalledFontCollection()
        for fam in coll.Families:
            names.add(fam.Name.lower())
        coll.Dispose()
    except Exception:
        pass
    try:  # registro do Windows: complementa a lista acima (inclui fontes instaladas só para o usuário)
        Registry = win_registry()
        for hive in (Registry.LocalMachine, Registry.CurrentUser):
            key = hive.OpenSubKey(r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
            if key is None:
                continue
            for value in key.GetValueNames():
                # "Arial Narrow Bold (TrueType)" -> "arial narrow bold" e "arial narrow"
                base = re.sub(r"\s*\(.*?\)\s*$", u"", value).strip().lower()
                for part in base.split(u"&"):
                    part = part.strip()
                    names.add(part)
                    names.add(re.sub(r"\s+(bold|italic|bold italic|regular|negrito|it[aá]lico)$", u"", part))
            key.Close()
    except Exception:
        pass
    return names


def sheet_fonts(sheets):
    """Fontes usadas nas tabelas e notas de texto das folhas -> lugares onde aparecem."""
    used = {}

    def add(font, where):
        font = (font or u"").strip()
        if font:
            places = used.setdefault(font, [])
            if where not in places:
                places.append(where)

    def type_font(type_id):
        t = doc.GetElement(type_id) if type_id and type_id != ElementId.InvalidElementId else None
        p = t.get_Parameter(BuiltInParameter.TEXT_FONT) if t is not None else None
        return p.AsString() if p is not None else None

    for sheet in sheets:
        # tabelas (quantitativos) colocadas na folha: fonte do título/cabeçalho/corpo e de cada célula
        for ssi in FilteredElementCollector(doc, sheet.Id).OfClass(ScheduleSheetInstance):
            try:
                if ssi.IsTitleblockRevisionSchedule:
                    continue
                vs = doc.GetElement(ssi.ScheduleId)
                where = u"tabela '{}'".format(vs.Name)
                for attr in ("TitleTextTypeId", "HeaderTextTypeId", "BodyTextTypeId"):
                    try:
                        add(type_font(getattr(vs, attr)), where)
                    except Exception:
                        pass
                table = vs.GetTableData()
                for st in (SectionType.Header, SectionType.Body):
                    sec = table.GetSectionData(st)
                    if sec is None:
                        continue
                    for r in range(sec.FirstRowNumber, sec.LastRowNumber + 1):
                        for c in range(sec.FirstColumnNumber, sec.LastColumnNumber + 1):
                            try:
                                add(sec.GetTableCellStyle(r, c).FontName, where)
                            except Exception:
                                pass
            except Exception:
                pass
        # notas de texto na folha e nas vistas colocadas nela
        try:
            view_ids = [sheet.Id] + list(sheet.GetAllPlacedViews())
        except Exception:
            view_ids = [sheet.Id]
        for vid in view_ids:
            try:
                type_ids = set(t.GetTypeId() for t in FilteredElementCollector(doc, vid).OfClass(TextNote))
            except Exception:
                continue
            for tid in type_ids:
                add(type_font(tid), u"notas de texto")
    return used


def check_fonts(sheets):
    """Confere as fontes das folhas. Se faltar alguma neste computador, pergunta o que fazer.

    Devolve (seguir, troca): seguir=False cancela a exportação; troca = None (não trocar) ou
    (fonte nova, conjunto das fontes que faltam, em minúsculas)."""
    try:
        used = sheet_fonts(sheets)
        installed = installed_fonts()
    except Exception:
        return True, None
    if not installed:          # não conseguiu ler as fontes do computador: não atrapalha a exportação
        return True, None
    missing = sorted((f for f in used if f.lower() not in installed), key=lambda x: x.lower())
    if not missing:
        return True, None
    choices = [f for f in FALLBACK_FONTS if f.lower() in installed]
    lines = []
    for f in missing:
        places = used[f]
        extra = u" e mais {}".format(len(places) - 3) if len(places) > 3 else u""
        lines.append(u"- {}: {}{}".format(f, u", ".join(places[:3]), extra))
    options = [(USE_FONT if i == 0 else USE_FONT_ALT).format(f) for i, f in enumerate(choices)]
    ans = forms.alert(
        u"{} fonte(s) usada(s) nas folhas não está(ão) instalada(s) neste computador.".format(len(missing)),
        sub_msg=u"\n".join(lines) +
                u"\n\nSem a fonte certa, o texto das tabelas passa das células no PDF. Trocando por uma "
                u"fonte instalada, o Revit refaz as quebras de linha das tabelas com ela e o texto fica "
                u"dentro das células (as linhas podem ficar um pouco mais altas). A troca vale só para "
                u"o PDF desta exportação: o modelo volta como estava logo depois.",
        options=options + [EXPORT_ANYWAY, CANCEL_FONTS], title=ALERT_TITLE)
    for f, opt in zip(choices, options):
        if ans == opt:
            return True, (f, set(m.lower() for m in missing))
    if ans == EXPORT_ANYWAY:
        return True, None
    return False, None


def apply_font_fallback(sheets, new_font, missing):
    """Troca as fontes que faltam por new_font nas tabelas e notas de texto das folhas.
    Precisa estar dentro de uma Transaction. Devolve quantos itens foram trocados."""
    changed = [0]

    def fix_type(type_id):
        t = doc.GetElement(type_id) if type_id and type_id != ElementId.InvalidElementId else None
        p = t.get_Parameter(BuiltInParameter.TEXT_FONT) if t is not None else None
        if p is not None and not p.IsReadOnly and (p.AsString() or u"").strip().lower() in missing:
            p.Set(new_font)
            changed[0] += 1

    def fix_style(style):
        """Estilo de célula/coluna com fonte trocada à mão -> novo estilo, ou None se não precisa."""
        if (style.FontName or u"").strip().lower() not in missing:
            return None
        style.FontName = new_font
        opts = style.GetCellStyleOverrideOptions()
        opts.Font = True
        style.SetCellStyleOverrideOptions(opts)
        return style

    done_types, done_schedules = set(), set()
    for sheet in sheets:
        for ssi in FilteredElementCollector(doc, sheet.Id).OfClass(ScheduleSheetInstance):
            try:
                if ssi.IsTitleblockRevisionSchedule or ssi.ScheduleId.ToString() in done_schedules:
                    continue
            except Exception:
                pass
            vs = doc.GetElement(ssi.ScheduleId)
            done_schedules.add(ssi.ScheduleId.ToString())
            for attr in ("TitleTextTypeId", "HeaderTextTypeId", "BodyTextTypeId"):
                try:
                    tid = getattr(vs, attr)
                    if tid.ToString() not in done_types:
                        done_types.add(tid.ToString())
                        fix_type(tid)
                except Exception:
                    pass
            # cabeçalho/título: fonte trocada célula a célula
            try:
                sec = vs.GetTableData().GetSectionData(SectionType.Header)
                for r in range(sec.FirstRowNumber, sec.LastRowNumber + 1):
                    for c in range(sec.FirstColumnNumber, sec.LastColumnNumber + 1):
                        try:
                            st = fix_style(sec.GetTableCellStyle(r, c))
                            if st is not None:
                                sec.SetCellStyle(r, c, st)
                                changed[0] += 1
                        except Exception:
                            pass
            except Exception:
                pass
            # corpo: fonte trocada por coluna (campo da tabela)
            try:
                definition = vs.Definition
                for i in range(definition.GetFieldCount()):
                    field = definition.GetField(i)
                    try:
                        st = fix_style(field.GetStyle())
                        if st is not None:
                            field.SetStyle(st)
                            changed[0] += 1
                    except Exception:
                        pass
            except Exception:
                pass
        # notas de texto na folha e nas vistas dela
        try:
            view_ids = [sheet.Id] + list(sheet.GetAllPlacedViews())
        except Exception:
            view_ids = [sheet.Id]
        for vid in view_ids:
            try:
                for t in FilteredElementCollector(doc, vid).OfClass(TextNote):
                    tid = t.GetTypeId()
                    if tid.ToString() not in done_types:
                        done_types.add(tid.ToString())
                        fix_type(tid)
            except Exception:
                pass
    return changed[0]


class TemporaryFonts(object):
    """Troca as fontes que faltam só enquanto o PDF é exportado e desfaz tudo em seguida.

    Usa um TransactionGroup: a troca é gravada numa Transaction (o Revit refaz as tabelas com a
    fonte nova), o PDF é exportado e o grupo inteiro é desfeito (RollBack), sem deixar nada no
    modelo nem no histórico de desfazer."""

    def __init__(self, sheets, fallback):
        self.sheets, self.fallback = sheets, fallback
        self.group = None
        self.count = 0
        self.error = None

    def __enter__(self):
        if not self.fallback:
            return self
        new_font, missing = self.fallback
        try:
            self.group = TransactionGroup(doc, u"Exportar Pranchas - fontes temporárias")
            self.group.Start()
            t = Transaction(doc, u"Trocar fontes que faltam por {}".format(new_font))
            t.Start()
            try:
                self.count = apply_font_fallback(self.sheets, new_font, missing)
                t.Commit()
            except Exception:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                raise
        except Exception as ex:
            self.error = ex
            self.__exit__(None, None, None)
        return self

    def __exit__(self, *args):
        if self.group is not None:
            try:
                if self.group.HasStarted() and not self.group.HasEnded():
                    self.group.RollBack()
            except Exception:
                pass
            self.group = None
        return False


REPLACE_ALL = u"Substituir os arquivos existentes"
KEEP_EXISTING = u"Manter os existentes e exportar só os novos"
CANCEL_EXPORT = u"Cancelar a exportação"


def combined_name(o):
    return clean(o["combined_name"]) or u"Pranchas_Combinadas"


def planned_files(o, names):
    """Nomes dos arquivos que a exportação vai gravar na pasta."""
    files = []
    if o["pdf"]:
        files += [n + u".pdf" for n in ([combined_name(o)] if o["combined"] else names)]
    if o["dwg"]:
        files += [n + u".dwg" for n in names]
    return files


def ask_existing(o, names):
    """Se já houver arquivos com os mesmos nomes na pasta, pergunta UMA vez o que fazer.
    Retorna "replace", "skip" ou None (cancelar). Sem arquivos repetidos, não pergunta nada."""
    existing = [f for f in planned_files(o, names) if os.path.isfile(os.path.join(o["folder"], f))]
    if not existing:
        return "replace"
    shown = existing[:12]
    more = len(existing) - len(shown)
    ans = forms.alert(
        u"{} arquivo(s) com o mesmo nome já existem na pasta. Substituir?".format(len(existing)),
        sub_msg=u"{}\n\n{}{}".format(o["folder"], u"\n".join(shown),
                                     u"\n... e mais {}".format(more) if more > 0 else u""),
        options=[REPLACE_ALL, KEEP_EXISTING, CANCEL_EXPORT], title=ALERT_TITLE)
    if ans == REPLACE_ALL:
        return "replace"
    if ans == KEEP_EXISTING:
        return "skip"
    return None


def prepare_target(ctx, path, label):
    """Antes de exportar um arquivo: aplica a escolha sobre arquivos já existentes.
    Devolve False quando esse arquivo não deve ser exportado."""
    if not os.path.isfile(path):
        return True
    if ctx["existing"] == "skip":
        ctx["kept"].append(os.path.basename(path))
        return False
    try:
        os.remove(path)          # substituir: apaga o antigo para o Revit gravar o novo no lugar
        return True
    except Exception:
        handle_failure(ctx, label, u"o arquivo existente não pôde ser substituído; "
                                   u"talvez esteja aberto em outro programa")
        return False


def is_cancel(ex):
    """Exceção causada pelo Cancelar da barra de status do Revit (OperationCanceledException)."""
    text = u"{} {}".format(type(ex).__name__, ex).lower()
    return u"cancel" in text


def stop_all(ctx, label):
    """Cancelar no Revit interrompe a prancha atual E todas as restantes, sem perguntar."""
    if not ctx["prog"].cancelled:
        ctx["stopped_at"] = label
    ctx["prog"].cancel_all()


def run_export(ctx, label, export):
    """Roda um doc.Export. Quando o usuário clica em Cancelar na barra de status do Revit, o
    Export devolve False (ou lança OperationCanceledException): a exportação inteira para.
    Outros erros seguem para handle_failure (pergunta única: cancelar tudo ou ignorar)."""
    try:
        if export():
            return True
        stop_all(ctx, label)
    except Exception as ex:
        if is_cancel(ex):
            stop_all(ctx, label)
        else:
            handle_failure(ctx, label, ex)
    return False


def handle_failure(ctx, label, ex=None):
    """Na primeira falha/cancelamento pergunta UMA vez; a escolha vale para todas as demais pranchas."""
    ctx["failed"].append(label + (u" ({})".format(ex) if ex else u""))
    diag(u"FALHA {}: {}".format(label, ex))
    prog = ctx["prog"]
    if ctx["policy"] is None and not prog.cancelled:
        watcher = ctx.get("watcher")
        if watcher is not None:
            watcher.asking = True    # esta pergunta é do script: o observador de avisos do Revit ignora
        try:
            ans = forms.alert(
                u"Não foi possível exportar {}.".format(label),
                sub_msg=u"A escolha abaixo vale para todas as pranchas restantes.{}".format(
                    u"\n\n{}".format(ex) if ex else u""),
                options=[u"Cancelar toda a exportação", u"Ignorar e continuar com as demais"], title=ALERT_TITLE)
        finally:
            if watcher is not None:
                watcher.asking = False
        ctx["policy"] = "continue" if ans == u"Ignorar e continuar com as demais" else "stop"
    if ctx["policy"] == "stop":
        prog.cancelled = True


AUTO_ALL = u"Confirmar em todas as folhas desta exportação"
SHOW_EACH = u"Mostrar este aviso em cada folha"

# códigos de resposta tentados em ordem até o Revit aceitar um:
# OK, Sim, primeiro link de comando, Fechar (TaskDialogResult / IDs do Windows)
CONFIRM_CODES = (1, 6, 1001, 8)

# confirmação do próprio Revit ao clicar em Cancelar na barra de status
# ("Exportação interrompida! Tem certeza que deseja parar esta operação?"): nunca é interceptada
CANCEL_PROMPT = (u"interrompid", u"parar esta opera", u"interrupted", u"stop this operation")


# erros de impressão do Revit (ex.: "could not start the print job ... unable to connect to the
# specified printer"): não viram pergunta; a falha aparece com o motivo na exportação
PRINT_ERRORS = (u"print job", u"specified printer", u"trabalho de impress", u"impressora especificada")


def is_cancel_prompt(msg):
    low = (msg or u"").lower()
    return any(k in low for k in CANCEL_PROMPT)


class DialogWatcher(object):
    """Fica ligado só durante a exportação (UIApplication.DialogBoxShowing).

    Nada aparece se o Revit não mostrar nenhum aviso. Quando um aviso aparece numa folha
    (ex.: keynote desatualizado), pergunta UMA vez o que fazer com ESSE aviso nas demais
    folhas da exportação:
      - confirmar em todas as folhas: o mesmo aviso é respondido sozinho (OK/Sim) nas próximas folhas;
      - mostrar em cada folha: o Revit mostra esse aviso normalmente, sem nova pergunta.
    Um aviso diferente gera uma nova pergunta. A confirmação de Cancelar do Revit não é tocada."""

    def __init__(self, prog):
        self.prog = prog
        self.policies = {}          # aviso (DialogId ou texto) -> "auto" ou "show"
        self.asking = False
        self.answered = []          # textos dos avisos respondidos automaticamente
        self.unanswered = []        # avisos que o Revit não deixou responder sozinho
        self.seen = []              # todos os avisos mostrados (para explicar falhas de impressão)

    @staticmethod
    def describe(args):
        try:
            msg = args.Message if isinstance(args, (TaskDialogShowingEventArgs, MessageBoxShowingEventArgs)) else u""
        except Exception:
            msg = u""
        try:
            did = args.DialogId
        except Exception:
            did = u""
        return (msg or u"").strip(), did or u""

    def ask(self, msg):
        topmost = self.prog.Topmost
        self.prog.Topmost = False   # a janela de progresso não pode cobrir a pergunta
        try:
            ans = forms.alert(
                u"O Revit mostrou um aviso ao exportar esta folha.",
                sub_msg=(msg or u"(aviso sem texto)") +
                        u"\n\nEscolha o que fazer com este aviso em todas as folhas desta exportação. "
                        u"Confirmando, ele é respondido sozinho (OK/Sim) nas próximas folhas, sem parar "
                        u"a exportação, e aparece no resumo ao final.",
                options=[AUTO_ALL, SHOW_EACH], title=ALERT_TITLE)
        finally:
            self.prog.Topmost = topmost
        return "auto" if ans == AUTO_ALL else "show"

    def confirm(self, args):
        for code in CONFIRM_CODES:
            try:
                ok = args.OverrideResult(code)
            except Exception:
                continue
            if ok is None or ok:     # algumas versões não devolvem nada
                return True
        return False

    def __call__(self, sender, args):
        if self.asking:              # a própria pergunta (TaskDialog do pyRevit) também passa por aqui
            return
        try:
            if not isinstance(args, DialogBoxShowingEventArgs):
                return
            msg, did = self.describe(args)
            low = msg.lower()
            if is_cancel_prompt(msg):
                return               # o usuário clicou em Cancelar: o Revit pergunta e decide sozinho
            self.seen.append(msg or did or u"(aviso sem texto)")
            if any(k in low for k in PRINT_ERRORS):
                self.confirm(args)   # erro de impressora: fecha o aviso; a falha é tratada na exportação
                return
            key = did or msg or u"(aviso sem texto)"
            if key not in self.policies:
                self.asking = True
                try:
                    self.policies[key] = self.ask(msg or did)
                finally:
                    self.asking = False
            if self.policies[key] == "auto":
                text = msg or did or u"(aviso sem texto)"
                (self.answered if self.confirm(args) else self.unanswered).append(text)
        except Exception:
            pass

    def summary(self):
        """Linhas para a mensagem final (sem janela de log)."""
        lines = []
        for title, items in ((u"Avisos do Revit respondidos automaticamente", self.answered),
                             (u"Avisos que o Revit não deixou responder automaticamente", self.unanswered)):
            if items:
                lines.append(u"{}:".format(title))
                for m in sorted(set(items)):
                    lines.append(u"- {} ({}x)".format(m.replace(u"\n", u" "), items.count(m)))
        return lines


# ----------------------------------------------------------------------
# PDF pela impressora PDF24 (opção "Printer")
#
# O Revit não deixa "imprimir em arquivo" com impressoras de PDF (na janela Imprimir do Revit essa
# opção fica desativada para elas). Por isso o PDF é feito pela impressora "OCA PDF24", criada pelo
# instalador da OCA (pasta Instaladores da extensão) e configurada para SALVAR SOZINHA numa pasta,
# sem abrir janela. Sem ela instalada, a opção PDF24 não aparece (só "Revit nativo").
# O Revit imprime cada folha nela, o PDF24 grava o PDF na pasta dela e o script move o arquivo
# para a pasta escolhida, com o nome da nomenclatura. As configurações de impressão do modelo
# são alteradas só durante a exportação (TransactionGroup desfeito no final).
PDF24_LABEL = u"PDF24"
OCA_PDF24_PRINTER = u"OCA PDF24"
ISO_PAPERS = [(u"A4", 210, 297), (u"A3", 297, 420), (u"A2", 420, 594), (u"A1", 594, 841),
              (u"A0", 841, 1189), (u"2A", 1189, 1682), (u"4A", 1682, 2378)]
PS_TIMEOUT = 180     # segundos esperando o PDF24 salvar o PDF de uma folha


def find_ghostscript():
    """Ghostscript do PDF24, do PDFCreator ou instalado à parte (usado só para juntar folhas)."""
    import glob
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramW6432"),
                 os.environ.get("ProgramFiles(x86)")):
        if not base:
            continue
        patterns = [r"PDF24\gs\bin\gswin64c.exe", r"PDF24\gs\bin\gswin32c.exe",
                    r"PDFCreator\Ghostscript\Bin\gswin64c.exe", r"PDFCreator\Ghostscript\Bin\gswin32c.exe",
                    r"gs\gs*\bin\gswin64c.exe", r"gs\gs*\bin\gswin32c.exe"]
        for pat in patterns:
            found = sorted(glob.glob(os.path.join(base, pat)), reverse=True)   # versão mais nova primeiro
            if found:
                return found[0]
    return None


def autosave_printers():
    """Impressoras PDF24 configuradas para salvar sozinhas -> pasta onde o PDF24 grava o PDF.

    Uma impressora assim usa a porta "\\\\.\\pipe\\PDFPrint - <serviço>" e o serviço do PDF24
    (HKLM\\SOFTWARE\\PDF24\\Services\\<serviço>) tem Handler = autoSave e AutoSaveDir."""
    result = {}
    try:
        from System import Environment
        Registry = win_registry()
        lm = Registry.LocalMachine
        printers = lm.OpenSubKey(r"SYSTEM\CurrentControlSet\Control\Print\Printers")
        if printers is None:
            return result
        for name in printers.GetSubKeyNames():
            key = printers.OpenSubKey(name)
            port = u"{}".format(key.GetValue("Port") or u"") if key is not None else u""
            if key is not None:
                key.Close()
            m = re.match(r"^\\\\\.\\pipe\\PDFPrint - (.+)$", port, re.I)
            if not m:
                continue
            svc = lm.OpenSubKey(u"SOFTWARE\\PDF24\\Services\\" + m.group(1))
            if svc is None:
                continue
            handler = u"{}".format(svc.GetValue("Handler") or u"").lower()
            folder = u"{}".format(svc.GetValue("AutoSaveDir") or u"")
            svc.Close()
            if handler == u"autosave" and folder:
                result[name] = Environment.ExpandEnvironmentVariables(folder)
        printers.Close()
    except Exception:
        pass
    return result


def pdf24_printer():
    """(impressora, pasta) para a opção "PDF24": só a "OCA PDF24", instalada pelo instalador da OCA."""
    for name, folder in autosave_printers().items():
        if name.lower() == OCA_PDF24_PRINTER.lower():
            return name, folder
    return None, None


def pdf24_installed():
    """O PDF24 Creator está instalado (driver de impressora "PDF24")?"""
    try:
        lm = win_registry().LocalMachine
        for env in (r"Windows x64", r"Windows NT x86"):
            for ver in ("Version-3", "Version-4"):
                key = lm.OpenSubKey(u"SYSTEM\\CurrentControlSet\\Control\\Print\\Environments\\{}\\Drivers\\{}\\PDF24"
                                    .format(env, ver))
                if key is not None:
                    key.Close()
                    return True
    except Exception:
        pass
    return False


def pdf_printers():
    """Opções de impressora da janela: {nome mostrado: impressora usada}.
    Com o PDF24 instalado mas sem a "OCA PDF24", a opção aparece com None (oferece instalar)."""
    device, _ = pdf24_printer()
    if device:
        return {PDF24_LABEL: device}
    return {PDF24_LABEL: None} if pdf24_installed() else {}


def installer_path():
    """Instalador da impressora: <extensão OCA>\\Instaladores\\Instalar impressora OCA PDF24.ps1."""
    try:
        here = os.path.dirname(os.path.abspath(__file__))       # ...\ExportarPranchas.pushbutton
    except NameError:
        from pyrevit import script as pyscript
        here = pyscript.get_script_path()
    ext = os.path.dirname(os.path.dirname(os.path.dirname(here)))  # ...\OCA.extension
    return os.path.join(ext, u"Instaladores", u"Instalar impressora OCA PDF24.ps1")


def install_oca_printer():
    """Roda o instalador da "OCA PDF24" como administrador (o Windows pede permissão) e espera
    terminar. Devolve True se a impressora ficou instalada."""
    path = installer_path()
    if not os.path.isfile(path):
        raise Exception(u"instalador não encontrado: {}".format(path))
    Process = net_import("System.Diagnostics", "Process", ("System.Diagnostics.Process", "System"))
    ProcessStartInfo = net_import("System.Diagnostics", "ProcessStartInfo",
                                  ("System.Diagnostics.Process", "System"))
    info = ProcessStartInfo(u"powershell.exe",
                            u'-NoProfile -ExecutionPolicy Bypass -File "{}"'.format(path))
    info.UseShellExecute = True
    info.Verb = u"runas"                 # "Executar como administrador"
    proc = Process.Start(info)           # recusar a permissão gera exceção
    proc.WaitForExit()
    return pdf24_printer()[0] is not None


def sheet_size_mm(sheet):
    """Tamanho da folha em mm pelo carimbo (Largura/Altura da folha do title block). O contorno da
    folha (Outline) inclui o que estiver fora do carimbo e pode dar um papel maior que o certo."""
    try:
        from Autodesk.Revit.DB import BuiltInCategory
        for tb in FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks) \
                .WhereElementIsNotElementType():
            w = tb.get_Parameter(BuiltInParameter.SHEET_WIDTH)
            h = tb.get_Parameter(BuiltInParameter.SHEET_HEIGHT)
            if w is not None and h is not None and w.AsDouble() > 0 and h.AsDouble() > 0:
                return w.AsDouble() * 304.8, h.AsDouble() * 304.8
    except Exception:
        pass
    o = sheet.Outline
    return (o.Max.U - o.Min.U) * 304.8, (o.Max.V - o.Min.V) * 304.8


def paper_key(name):
    """Nome de papel normalizado: "A1", "ISOA1", "ISO A1", "A1 (594 x 841 mm)" -> "a1"."""
    n = re.sub(r"\(.*?\)", u"", u"{}".format(name)).strip().lower().replace(u" ", u"")
    if n.startswith(u"iso") and len(n) > 3 and n[3] == u"a":
        n = n[3:]
    return n


def pick_paper(pm, sheet):
    """Menor papel ISO (A4 a 4A0) onde a folha cabe, entre os papéis da impressora."""
    w, h = sheet_size_mm(sheet)
    short, long_ = min(w, h), max(w, h)
    papers = {}
    for p in pm.PaperSizes:
        papers.setdefault(paper_key(p.Name), p)
    fallback = None
    for name, pw, ph in ISO_PAPERS:
        p = papers.get(paper_key(name))
        if p is None:
            continue
        fallback = p
        if short <= pw + 2 and long_ <= ph + 2:
            return p, w >= h, True
    return fallback, w >= h, False      # folha maior que o maior papel: ajusta à página


LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or u".", u"OCA_ExportarPranchas")
LOG_FILE = os.path.join(LOG_DIR, u"registro_impressao.txt")


def diag(msg, reset=False):
    """Registro passo a passo da impressão (para descobrir por que uma impressora falhou)."""
    try:
        import datetime
        if not os.path.isdir(LOG_DIR):
            os.makedirs(LOG_DIR)
        with codecs.open(LOG_FILE, "w" if reset else "a", "utf-8") as f:
            f.write(u"{} {}\n".format(datetime.datetime.now().strftime("%H:%M:%S"), msg))
    except Exception:
        pass


def pdf_page_mm(path):
    """Tamanho da 1ª página do PDF em mm (para o registro), ex.: "841x1189 (girada 90)"."""
    try:
        with open(path, "rb") as f:
            data = f.read()
        text = data.decode("latin-1") if not isinstance(data, type(u"")) else data
        m = re.search(r"/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)", text)
        r = re.search(r"/Rotate\s+(\d+)", text)
        w = (float(m.group(3)) - float(m.group(1))) * 25.4 / 72
        h = (float(m.group(4)) - float(m.group(2))) * 25.4 / 72
        return u"{:.0f}x{:.0f} mm{}".format(w, h, u" (girada {})".format(r.group(1)) if r else u"")
    except Exception:
        return u"?"


def pdf_snapshot(folder):
    """PDFs da pasta -> data de gravação (para reconhecer o PDF novo, mesmo se sobrescrito)."""
    snap = {}
    try:
        for f in os.listdir(folder):
            if f.lower().endswith(u".pdf"):
                p = os.path.join(folder, f)
                snap[os.path.normcase(p)] = os.path.getmtime(p)
    except Exception:
        pass
    return snap


def wait_new_pdf(folder, before, prog, timeout=PS_TIMEOUT):
    """Espera o PDF24 terminar de gravar um PDF novo na pasta. Devolve o caminho ou None."""
    from System.Threading import Thread
    sizes, stable, waited = {}, {}, 0.0
    base = prog._text
    try:
        while waited < timeout:
            if prog.cancelled:
                return None
            if waited >= 2 and int(waited * 4) % 4 == 0:     # mostra que está esperando, a cada 1 s
                prog._text = u"{}\nAguardando o PDF24 salvar a folha... {} s".format(base, int(waited))
                prog.refresh()
            for key, mtime in pdf_snapshot(folder).items():
                if key in before and before[key] == mtime:
                    continue                                   # já estava lá antes desta folha
                size = os.path.getsize(key) if os.path.isfile(key) else 0
                if size > 0 and size == sizes.get(key, -1):
                    stable[key] = stable.get(key, 0) + 1
                    if stable[key] >= 2:
                        try:
                            with open(key, "ab"):
                                return key
                        except Exception:
                            pass                               # o PDF24 ainda está gravando
                else:
                    stable[key] = 0
                sizes[key] = size
            prog.pump()
            Thread.CurrentThread.Join(250)   # espera sem travar as mensagens da janela do Revit
            waited += 0.25
        return None
    finally:
        prog._text = base


def ps_to_pdf(gs, ps_files, pdf_path):
    """Junta arquivos .pdf (uma folha cada) num PDF com o Ghostscript."""
    Process = net_import("System.Diagnostics", "Process", ("System.Diagnostics.Process", "System"))
    ProcessStartInfo = net_import("System.Diagnostics", "ProcessStartInfo",
                                  ("System.Diagnostics.Process", "System"))
    q = lambda s: u'"{}"'.format(s)
    args = [u"-dBATCH", u"-dNOPAUSE", u"-dQUIET", u"-dSAFER", u"-sDEVICE=pdfwrite",
            u"-dAutoRotatePages=/PageByPage", q(u"-sOutputFile=" + pdf_path)] + [q(f) for f in ps_files]
    info = ProcessStartInfo(gs, u" ".join(args))
    info.UseShellExecute = False
    info.CreateNoWindow = True
    proc = Process.Start(info)
    if not proc.WaitForExit(600000):
        proc.Kill()
        raise Exception(u"o Ghostscript demorou demais para juntar as folhas")
    if proc.ExitCode != 0 or not os.path.isfile(pdf_path):
        raise Exception(u"o Ghostscript não juntou as folhas (código {})".format(proc.ExitCode))
    return True


def run_gs(gs, args):
    """Roda o Ghostscript sem janela. args: lista de argumentos (aspas já tratadas aqui)."""
    Process = net_import("System.Diagnostics", "Process", ("System.Diagnostics.Process", "System"))
    ProcessStartInfo = net_import("System.Diagnostics", "ProcessStartInfo",
                                  ("System.Diagnostics.Process", "System"))
    q = lambda s: u'"{}"'.format(s) if (u" " in s or u'"' in s) else s
    info = ProcessStartInfo(gs, u" ".join(q(a) for a in args))
    info.UseShellExecute = False
    info.CreateNoWindow = True
    proc = Process.Start(info)
    if not proc.WaitForExit(600000):
        proc.Kill()
        raise Exception(u"o Ghostscript demorou demais")
    if proc.ExitCode != 0:
        raise Exception(u"o Ghostscript falhou (código {})".format(proc.ExitCode))


# imagens mantidas como estão (sem reduzir resolução) ao regravar o PDF
GS_KEEP_QUALITY = [u"-dDownsampleColorImages=false", u"-dDownsampleGrayImages=false",
                   u"-dDownsampleMonoImages=false", u"-dPassThroughJPEGImages=true"]


def crop_pdf(src, dst, paper_w, paper_h, sheet_w, sheet_h):
    """Recorta o PDF impresso (folha centralizada em 100% num papel maior) no tamanho exato da
    folha. 1) aplica a rotação da página (fica do jeito que se vê); 2) página do tamanho da folha,
    deslocando o desenho para o centro coincidir. Medidas em mm."""
    gs = find_ghostscript()
    if gs is None:
        raise Exception(u"o Ghostscript (vem com o PDF24) não foi encontrado para ajustar a página")
    pt = lambda mm: mm * 72.0 / 25.4
    if (paper_w > paper_h) != (sheet_w > sheet_h):
        sheet_w, sheet_h = sheet_h, sheet_w
    tmp = dst + u".tmp.pdf"
    base = [u"-dSAFER", u"-dNOPAUSE", u"-dBATCH", u"-dQUIET", u"-sDEVICE=pdfwrite",
            u"-dAutoRotatePages=/None"] + GS_KEEP_QUALITY
    try:
        run_gs(gs, base + [u"-sOutputFile=" + tmp, src])
        ox, oy = (pt(paper_w) - pt(sheet_w)) / 2.0, (pt(paper_h) - pt(sheet_h)) / 2.0
        run_gs(gs, base + [u"-dFIXEDMEDIA",
                           u"-dDEVICEWIDTHPOINTS={:.2f}".format(pt(sheet_w)),
                           u"-dDEVICEHEIGHTPOINTS={:.2f}".format(pt(sheet_h)),
                           u"-sOutputFile=" + dst,
                           u"-c", u"<</PageOffset [{:.2f} {:.2f}]>> setpagedevice".format(-ox, -oy),
                           u"-f", tmp])
    finally:
        try:
            if os.path.isfile(tmp):
                os.remove(tmp)
        except Exception:
            pass
    if not os.path.isfile(dst):
        raise Exception(u"o Ghostscript não gerou o PDF ajustado")


def move_file(src, dst):
    """Move o PDF salvo pelo PDF24 para o destino (pode ser outro disco ou a rede)."""
    import shutil
    if os.path.isfile(dst):
        os.remove(dst)
    shutil.move(src, dst)


class PrinterSession(object):
    """Configura o PrintManager do Revit para a impressora PDF24 e desfaz tudo no final."""

    def __init__(self, o, device, autosave_dir, watcher=None):
        self.o = o
        self.device = device
        self.autosave_dir = autosave_dir
        self.watcher = watcher
        self.group = None

    def __enter__(self):
        self.group = TransactionGroup(doc, u"Exportar Pranchas - impressão PDF")
        self.group.Start()
        return self

    def __exit__(self, *args):
        try:
            if self.group is not None and self.group.HasStarted() and not self.group.HasEnded():
                self.group.RollBack()      # devolve impressora e configurações de impressão do modelo
        except Exception:
            pass
        return False

    def setup(self, sheet):
        """Configura a impressão desta folha. Precisa estar dentro de uma Transaction.
        Devolve (PrintManager, nome do papel) para a mensagem de erro."""
        o = self.o
        pm = doc.PrintManager
        # 1) troca a impressora e aplica ANTES de ler os papéis (a lista é da impressora nova)
        pm.SelectNewPrintDriver(self.device)
        pm.Apply()
        if (pm.PrinterName or u"").lower() != self.device.lower():
            raise Exception(u"o Revit não aceitou a impressora '{}' (ficou '{}')".format(
                self.device, pm.PrinterName))
        # 2) imprime a própria folha (SubmitPrint(folha)), sem "imprimir em arquivo": o PDF24 salva
        #    sozinho. (O conjunto de folhas "na sessão" não é aceito: "There are no views/sheets selected")
        pm.PrintRange = PrintRange.Current
        try:
            pm.PrintToFile = False
        except Exception:
            pass
        setup = pm.PrintSetup
        setup.CurrentPrintSetting = setup.InSession
        pp = setup.CurrentPrintSetting.PrintParameters
        paper, landscape, fits = pick_paper(pm, sheet)
        if paper is not None:
            pp.PaperSize = paper
        pp.PageOrientation = PageOrientationType.Landscape if landscape else PageOrientationType.Portrait
        # folha no centro do papel, em 100%: depois o PDF é recortado no tamanho exato da folha
        # (o papel da impressora é o menor ISO onde a folha cabe; o PDF24 não tem margens)
        self.crop = None
        if fits:
            pp.ZoomType = ZoomType.Zoom
            pp.Zoom = 100
            dims = [(pw, ph) for n, pw, ph in ISO_PAPERS if paper is not None and paper_key(n) == paper_key(paper.Name)]
            if dims:
                sw, sh = sheet_size_mm(sheet)
                pw, ph = (max(dims[0]), min(dims[0])) if landscape else (min(dims[0]), max(dims[0]))
                self.crop = (pw, ph, sw, sh)
        else:
            pp.ZoomType = ZoomType.FitToPage        # folha maior que o maior papel: sem recorte
        pp.PaperPlacement = PaperPlacementType.Center
        pp.ColorDepth = o["color"]
        pp.RasterQuality = o["raster"]
        pp.HiddenLineViews = (HiddenLineViewsType.RasterProcessing if o["use_raster"]
                              else HiddenLineViewsType.VectorProcessing)
        pp.HideReforWorkPlanes = o["hide_ref"]
        pp.HideUnreferencedViewTags = o["hide_tags"]
        pp.HideScopeBoxes = o["hide_scope"]
        pp.HideCropBoundaries = o["hide_crop"]
        # 3) grava a configuração com um nome: alterada só "na sessão", o Revit imprimia com o papel
        #    padrão da impressora (A4). O TransactionGroup apaga essa configuração no final.
        try:
            self.count = getattr(self, "count", 0) + 1
            setup.SaveAs(u"OCA_Exportar_Pranchas_{:03d}".format(self.count))
            saved = True
        except Exception as ex:
            saved = u"não ({})".format(ex)
        pm.Apply()
        applied = setup.CurrentPrintSetting.PrintParameters
        diag(u"impressora={} folha={} tamanho={:.0f}x{:.0f} mm papel escolhido={} | gravada={} | "
             u"aplicado: papel={} orientação={}".format(
                 pm.PrinterName, sheet.SheetNumber, sheet_size_mm(sheet)[0], sheet_size_mm(sheet)[1],
                 paper.Name if paper is not None else None, saved,
                 applied.PaperSize.Name if applied.PaperSize is not None else None, applied.PageOrientation))
        return pm, (paper.Name if paper is not None else u"(nenhum papel encontrado)")

    def print_sheet(self, sheet, pdf_path, prog):
        """Imprime uma folha e move o PDF salvo pelo PDF24 para pdf_path. Devolve False só quando
        o usuário cancelou no Revit; erro vira exceção com o motivo."""
        folder = self.autosave_dir
        if not os.path.isdir(folder):
            os.makedirs(folder)
        before = pdf_snapshot(folder)
        seen = len(self.watcher.seen) if self.watcher is not None else 0
        # configura e envia na mesma Transaction (o TransactionGroup desfaz tudo no final)
        t = Transaction(doc, u"Imprimir folha em PDF")
        t.Start()
        try:
            pm, paper = self.setup(sheet)
            sent = pm.SubmitPrint(sheet)   # imprime só esta folha
            diag(u"SubmitPrint devolveu {}".format(sent))
            t.Commit()
        except Exception:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            raise
        if not sent:
            msgs = self.watcher.seen[seen:] if self.watcher is not None else []
            if prog.cancelled or not msgs or any(is_cancel_prompt(m) for m in msgs):
                return False          # Cancelar na barra de status do Revit
            raise Exception(u"o Revit não conseguiu imprimir na impressora '{}' (papel {}): {}".format(
                self.device, paper, u" | ".join(m.replace(u"\n", u" ") for m in msgs)))
        found = wait_new_pdf(folder, before, prog)
        diag(u"PDF salvo pelo PDF24: {} | página do PDF: {}".format(found, pdf_page_mm(found)))
        if not found:
            if prog.cancelled:
                return False
            raise Exception(u"o PDF24 não salvou o PDF da folha em {} s (pasta {})".format(PS_TIMEOUT, folder))
        if self.crop is not None:
            # recorta a página no tamanho da folha do Revit (sem a borda do papel maior)
            prog.step(u"Ajustando a página ao tamanho da folha...")
            crop_pdf(found, pdf_path, *self.crop)
            diag(u"recortado para {:.0f}x{:.0f} mm: {} | página: {}".format(
                self.crop[2], self.crop[3], pdf_path, pdf_page_mm(pdf_path)))
            try:
                os.remove(found)
            except Exception:
                pass
        else:
            move_file(found, pdf_path)
        return True


def export_pdf_printer(o, sheets, names, ctx):
    """PDF pela impressora "OCA PDF24", que salva o PDF sozinha."""
    prog, folder, ok = ctx["prog"], o["folder"], 0
    device, autosave = pdf24_printer()
    diag(u"início: opção='{}' impressora='{}' pasta do PDF24='{}'".format(o["printer"], device, autosave),
         reset=True)
    if not device:
        handle_failure(ctx, u"PDF", u"nenhuma impressora PDF24 que salva sozinha foi encontrada; rode "
                                    u"'Instalar impressora OCA PDF24' (pasta Instaladores da extensão OCA)")
        return 0
    with PrinterSession(o, device, autosave, ctx.get("watcher")) as session:
        if o["combined"]:
            name = combined_name(o)
            pdf = os.path.join(folder, name + u".pdf")
            if prepare_target(ctx, pdf, u"PDF combinado"):
                parts = []

                def do_combined():
                    gs = find_ghostscript()
                    if gs is None:
                        raise Exception(u"para juntar as folhas num único PDF é preciso o Ghostscript "
                                        u"(vem com o PDF24)")
                    if not os.path.isdir(LOG_DIR):
                        os.makedirs(LOG_DIR)
                    for i, s in enumerate(sheets):
                        if prog.cancelled:
                            return False
                        prog.step(u"PDF combinado: folha {} de {}".format(i + 1, len(sheets)))
                        part = os.path.join(LOG_DIR, u"folha_{:03d}.pdf".format(i + 1))
                        if not session.print_sheet(s, part, prog):
                            return False
                        parts.append(part)
                    prog.step(u"PDF combinado: gerando {}".format(name))
                    return ps_to_pdf(gs, parts, pdf)

                if run_export(ctx, u"PDF combinado", do_combined):
                    ok = 1
                for f in parts:
                    try:
                        os.remove(f)
                    except Exception:
                        pass
            prog.advance()
            return ok
        for s, n in zip(sheets, names):
            if prog.cancelled:
                break
            prog.step(u"PDF ({}): {}".format(o["printer"], n))
            pdf = os.path.join(folder, n + u".pdf")
            if prepare_target(ctx, pdf, u"PDF: {}".format(n)) and \
                    run_export(ctx, u"PDF: {}".format(n), lambda: session.print_sheet(s, pdf, prog)):
                ok += 1
            prog.advance()
    return ok


def export_pdf(o, sheets, names, ctx):
    if o.get("printer", NATIVE_PRINTER) != NATIVE_PRINTER:
        return export_pdf_printer(o, sheets, names, ctx)
    return export_pdf_native(o, sheets, names, ctx)


def export_pdf_native(o, sheets, names, ctx):
    prog, folder, ok = ctx["prog"], o["folder"], 0
    if o["combined"]:
        name = combined_name(o)
        prog.step(u"PDF combinado: {}".format(name))
        if prepare_target(ctx, os.path.join(folder, name + u".pdf"), u"PDF combinado") and \
                run_export(ctx, u"PDF combinado",
                           lambda: doc.Export(folder, ids_of(sheets), pdf_options(o, name, True))):
            ok = 1
        prog.advance()
        return ok
    for s, n in zip(sheets, names):
        if prog.cancelled:
            break
        prog.step(u"PDF: {}".format(n))
        if prepare_target(ctx, os.path.join(folder, n + u".pdf"), u"PDF: {}".format(n)) and \
                run_export(ctx, u"PDF: {}".format(n),
                           lambda: doc.Export(folder, ids_of([s]), pdf_options(o, n, True))):
            ok += 1
        prog.advance()
    return ok


def dwg_options(o):
    opts = None
    if o["dwg_setup"]:
        opts = DWGExportOptions.GetPredefinedOptions(doc, o["dwg_setup"])
    if opts is None:
        opts = DWGExportOptions()
    opts.MergedViews = o["merged"]
    return opts


def export_dwg(o, sheets, names, ctx):
    prog, ok = ctx["prog"], 0
    for s, n in zip(sheets, names):
        if prog.cancelled:
            break
        prog.step(u"DWG: {}".format(n))
        if prepare_target(ctx, os.path.join(o["folder"], n + u".dwg"), u"DWG: {}".format(n)) and \
                run_export(ctx, u"DWG: {}".format(n),
                           lambda: doc.Export(o["folder"], n, ids_of([s]), dwg_options(o))):
            ok += 1
        prog.advance()
    return ok


def main():
    sheets = pick_sheets()
    if not sheets:
        return

    setups = sorted(list(DWGExportOptions.GetPredefinedSetupNames(doc)))
    form = ExportForm(sheets, setups, sheet_param_names(sheets))
    form.show_dialog()
    o = form.result
    if not o:
        return

    folder = o["folder"]
    if not os.path.isdir(folder):
        os.makedirs(folder)

    names = build_names(sheets, o["rules"])
    total = 0
    if o["pdf"]:
        total += 1 if o["combined"] else len(sheets)
    if o["dwg"]:
        total += len(sheets)

    # fontes das tabelas/notas que não existem neste computador (texto saindo das células no PDF):
    # pergunta uma vez e, se escolhido, troca por Arial/Tahoma só durante o PDF
    fallback = None
    if o["pdf"]:
        go, fallback = check_fonts(sheets)
        if not go:
            return

    # arquivos com o mesmo nome já na pasta: pergunta uma vez antes de começar
    existing = ask_existing(o, names)
    if existing is None:
        return

    prog = ProgressWindow(total)
    prog.Show()
    prog.step(u"Preparando exportação...")
    ctx = {"prog": prog, "policy": None, "failed": [], "stopped_at": None,
           "existing": existing, "kept": []}

    # avisos do Revit: só pergunta se algum aparecer, e uma única vez para a exportação inteira
    watcher = DialogWatcher(prog)
    ctx["watcher"] = watcher
    handler = watcher.__call__
    uiapp = revit.uidoc.Application
    n_pdf = n_dwg = 0
    font_note = None
    try:
        uiapp.DialogBoxShowing += handler
        if o["pdf"]:
            # a troca de fonte vale só para o PDF; o DWG sai com as fontes originais do modelo
            with TemporaryFonts(sheets, fallback) as fonts:
                if fonts.error is not None:
                    font_note = u"As fontes não puderam ser trocadas ({}); o PDF saiu com as originais.".format(
                        fonts.error)
                elif fallback:
                    font_note = u"Fontes que faltavam trocadas por {} só no PDF ({} item(ns)); " \
                                u"o modelo não foi alterado.".format(fallback[0], fonts.count)
                n_pdf = export_pdf(o, sheets, names, ctx)
        if o["dwg"] and not prog.cancelled:
            n_dwg = export_dwg(o, sheets, names, ctx)
    finally:
        try:
            uiapp.DialogBoxShowing -= handler
        except Exception:
            pass
        was_cancelled = prog.cancelled
        prog.finish()

    for f in os.listdir(folder):  # lixo que o export de DWG às vezes deixa
        if f.lower().endswith(".pcp"):
            try:
                os.remove(os.path.join(folder, f))
            except Exception:
                pass

    # resumo só na mensagem final (sem janela de log do pyRevit)
    status = u"Exportação cancelada" if was_cancelled else u"Exportação concluída"
    details = [u"Pasta: {}".format(folder)]
    if font_note:
        details.append(font_note)
    if ctx["stopped_at"]:
        details.append(u"Cancelado no Revit em {}; as pranchas restantes não foram exportadas.".format(
            ctx["stopped_at"]))
    if ctx["kept"]:
        details.append(u"Mantidos (já existiam, não foram exportados de novo): {}".format(len(ctx["kept"])))
    if ctx["failed"]:
        details.append(u"Falhas:")
        details.extend(u"- {}".format(f) for f in ctx["failed"])
    details.extend(watcher.summary())
    forms.alert(u"{}: {} PDF, {} DWG, {} falha(s).".format(status, n_pdf, n_dwg, len(ctx["failed"])),
                sub_msg=u"\n".join(details), title=ALERT_TITLE)
    if n_pdf or n_dwg:
        os.startfile(folder)


main()
