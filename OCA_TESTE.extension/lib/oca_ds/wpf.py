# -*- coding: utf-8 -*-
"""Janelas WPF no padrão OCA: estilos de todos os controles e montagem da janela.

As cores vêm de tokens.py pelo tema atual (theme.py). Os estilos referenciam
só os pincéis "oca.*", nunca códigos hexadecimais, para que o mesmo XAML
funcione nos dois temas.

Raios: 4 px em campos, listas e botões; 6 px em seções, mensagens, tabelas,
cartões e dicas; 8 px nos menus suspensos. O contorno externo da janela é o
do Windows (o Windows 11 arredonda, o 10 não): a moldura nativa é mantida para
não perder mover, redimensionar, encaixar e acessibilidade.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

import os
import re

from oca_ds import tokens, theme
from oca_ds.messages import command_name, xml_escape

try:
    unicode
except NameError:           # IronPython 3 / CPython
    unicode = str

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(HERE, "assets", "logo", "logo.png")


# Os pincéis (oca.*) são gerados por window_resources() e vêm antes destes estilos,
# por isso StaticResource funciona aqui. @r_control@, @r_group@ e @r_panel@
# recebem os raios de tokens.RADIUS.
STYLES = u"""
    <GridLength x:Key="oca.LabelWidth">@label_width@</GridLength>
    <CornerRadius x:Key="oca.RadiusControl">@r_control@</CornerRadius>
    <CornerRadius x:Key="oca.RadiusGroup">@r_group@</CornerRadius>
    <CornerRadius x:Key="oca.RadiusPanel">@r_panel@</CornerRadius>

    <!-- textos -->
    <Style x:Key="oca.Label" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style x:Key="oca.Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="oca.Caption" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="oca.Value" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="oca.Mono" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="@font_mono@"/>
      <Setter Property="FontSize" Value="11.5"/>
    </Style>

    <!-- seção: caixa com título em caixa-alta; o conteúdo ocupa o espaço restante -->
    <Style x:Key="oca.Section" TargetType="HeaderedContentControl">
      <Setter Property="Margin" Value="0,0,0,12"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="IsTabStop" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="HeaderedContentControl">
            <Border Background="{StaticResource oca.Surface}" BorderBrush="{StaticResource oca.Line}"
                    BorderThickness="1" CornerRadius="{StaticResource oca.RadiusGroup}" Padding="12,9,12,10">
              <DockPanel>
                <ContentPresenter ContentSource="Header" DockPanel.Dock="Top" Margin="0,0,0,8"
                                  TextElement.FontSize="10.5" TextElement.FontWeight="SemiBold"
                                  TextElement.Foreground="{StaticResource oca.Ink2}"/>
                <ContentPresenter/>
              </DockPanel>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- mensagens em faixa: a faixa some quando o texto está vazio -->
    <Style x:Key="oca.MsgText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Style.Triggers>
        <Trigger Property="Text" Value="">
          <Setter Property="Visibility" Value="Collapsed"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="oca.Msg.Info" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource oca.AccentSoft}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="{StaticResource oca.RadiusGroup}"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="Margin" Value="0,8,0,0"/>
    </Style>
    <Style x:Key="oca.Msg.Ok" TargetType="Border" BasedOn="{StaticResource oca.Msg.Info}">
      <Setter Property="Background" Value="{StaticResource oca.OkSoft}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Ok}"/>
    </Style>
    <Style x:Key="oca.Msg.Warn" TargetType="Border" BasedOn="{StaticResource oca.Msg.Info}">
      <Setter Property="Background" Value="{StaticResource oca.WarnSoft}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Warn}"/>
    </Style>
    <Style x:Key="oca.Msg.Err" TargetType="Border" BasedOn="{StaticResource oca.Msg.Info}">
      <Setter Property="Background" Value="{StaticResource oca.ErrSoft}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Err}"/>
    </Style>

    <!-- cartão de pré-visualização -->
    <Style x:Key="oca.Card" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource oca.Field}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="{StaticResource oca.RadiusGroup}"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
    </Style>

    <!-- botões: secundário (padrão), principal, sem borda (Voltar) e compacto -->
    <Style TargetType="Button">
      <Setter Property="Height" Value="@button_height@"/>
      <Setter Property="MinWidth" Value="92"/>
      <Setter Property="Padding" Value="14,0"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="Background" Value="{StaticResource oca.Surface}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.FieldLine}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="{StaticResource oca.RadiusControl}" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="oca.Primary" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="{StaticResource oca.Brand}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Brand}"/>
      <Setter Property="Foreground" Value="{StaticResource oca.BrandInk}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="oca.Ghost" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Accent}"/>
      <Setter Property="MinWidth" Value="0"/>
      <Setter Property="Padding" Value="6,0"/>
    </Style>
    <Style x:Key="oca.Compact" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Height" Value="24"/>
      <Setter Property="MinWidth" Value="0"/>
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="FontSize" Value="11.5"/>
    </Style>

    <!-- campos -->
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="@field_height@"/>
      <Setter Property="Padding" Value="6,0"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="Background" Value="{StaticResource oca.Field}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.FieldLine}"/>
      <Setter Property="CaretBrush" Value="{StaticResource oca.Ink}"/>
      <Setter Property="SelectionBrush" Value="{StaticResource oca.Accent}"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="{StaticResource oca.RadiusControl}">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" Focusable="False"
                            VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                            HorizontalScrollBarVisibility="{TemplateBinding HorizontalScrollBarVisibility}"
                            VerticalScrollBarVisibility="{TemplateBinding VerticalScrollBarVisibility}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Ink2}"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <ControlTemplate x:Key="oca.ComboToggle" TargetType="ToggleButton">
      <Border x:Name="bd" Background="{StaticResource oca.Field}" BorderBrush="{StaticResource oca.FieldLine}"
              BorderThickness="1" CornerRadius="{StaticResource oca.RadiusControl}">
        <Path HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,9,0" Data="M0,0 L4,4 L8,0"
              Stroke="{StaticResource oca.Ink2}" StrokeThickness="1.5"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
        </Trigger>
        <Trigger Property="IsChecked" Value="True">
          <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
        </Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="@field_height@"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="MaxDropDownHeight" Value="320"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton Template="{StaticResource oca.ComboToggle}" Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
              <ContentPresenter IsHitTestVisible="False" Margin="8,0,26,0" HorizontalAlignment="Left" VerticalAlignment="Center"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"/>
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
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="Padding" Value="8,4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" Background="Transparent" Padding="{TemplateBinding Padding}"
                    CornerRadius="{StaticResource oca.RadiusControl}">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource oca.AccentSoft}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- dicas -->
    <Style TargetType="ToolTip">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="MaxWidth" Value="420"/>
      <Setter Property="HasDropShadow" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToolTip">
            <Border Background="{StaticResource oca.Surface}" BorderBrush="{StaticResource oca.Line}" BorderThickness="1"
                    CornerRadius="{StaticResource oca.RadiusGroup}" Padding="8,5">
              <ContentPresenter/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Setter Property="ContentTemplate">
        <Setter.Value>
          <DataTemplate>
            <TextBlock Text="{Binding}" TextWrapping="Wrap"/>
          </DataTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- barras de rolagem -->
    <Style x:Key="oca.ScrollThumb" TargetType="Thumb">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border x:Name="bd" CornerRadius="3" Background="{StaticResource oca.FieldLine}"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource oca.Ink2}"/>
              </Trigger>
              <Trigger Property="IsDragging" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource oca.Accent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="oca.ScrollPage" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="IsTabStop" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton">
            <Border Background="Transparent"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Width" Value="11"/>
      <Setter Property="MinWidth" Value="11"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Border Background="{TemplateBinding Background}">
              <Track x:Name="PART_Track" IsDirectionReversed="True" Margin="2">
                <Track.DecreaseRepeatButton>
                  <RepeatButton Style="{StaticResource oca.ScrollPage}" Command="ScrollBar.PageUpCommand"/>
                </Track.DecreaseRepeatButton>
                <Track.IncreaseRepeatButton>
                  <RepeatButton Style="{StaticResource oca.ScrollPage}" Command="ScrollBar.PageDownCommand"/>
                </Track.IncreaseRepeatButton>
                <Track.Thumb>
                  <Thumb Style="{StaticResource oca.ScrollThumb}"/>
                </Track.Thumb>
              </Track>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="11"/>
          <Setter Property="MinHeight" Value="11"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Border Background="{TemplateBinding Background}">
                  <Track x:Name="PART_Track" IsDirectionReversed="False" Margin="2">
                    <Track.DecreaseRepeatButton>
                      <RepeatButton Style="{StaticResource oca.ScrollPage}" Command="ScrollBar.PageLeftCommand"/>
                    </Track.DecreaseRepeatButton>
                    <Track.IncreaseRepeatButton>
                      <RepeatButton Style="{StaticResource oca.ScrollPage}" Command="ScrollBar.PageRightCommand"/>
                    </Track.IncreaseRepeatButton>
                    <Track.Thumb>
                      <Thumb Style="{StaticResource oca.ScrollThumb}"/>
                    </Track.Thumb>
                  </Track>
                </Border>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- abas -->
    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="0,1,0,0"/>
      <Setter Property="Padding" Value="0,12,0,0"/>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,2"
                    Padding="14,7" Margin="0,0,2,-1">
              <ContentPresenter ContentSource="Header" TextElement.Foreground="{TemplateBinding Foreground}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
                <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- tabelas: contorno arredondado; o recuo de 2 px mantém o cabeçalho e as
         linhas dentro da curva do canto -->
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{StaticResource oca.Field}"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="{StaticResource oca.Field}"/>
      <Setter Property="AlternatingRowBackground" Value="{StaticResource oca.Surface2}"/>
      <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="VerticalGridLinesBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="RowHeaderWidth" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGrid">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="{StaticResource oca.RadiusGroup}"
                    Padding="2" SnapsToDevicePixels="True">
              <ScrollViewer x:Name="DG_ScrollViewer" Focusable="False">
                <ScrollViewer.Template>
                  <ControlTemplate TargetType="ScrollViewer">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                      </Grid.RowDefinitions>
                      <DataGridColumnHeadersPresenter x:Name="PART_ColumnHeadersPresenter" Grid.Column="1"
                          Visibility="{Binding HeadersVisibility, RelativeSource={RelativeSource AncestorType=DataGrid}, Converter={x:Static DataGrid.HeadersVisibilityConverter}, ConverterParameter={x:Static DataGridHeadersVisibility.Column}}"/>
                      <ScrollContentPresenter x:Name="PART_ScrollContentPresenter" Grid.Row="1" Grid.ColumnSpan="2"
                                              CanContentScroll="{TemplateBinding CanContentScroll}"/>
                      <ScrollBar x:Name="PART_VerticalScrollBar" Grid.Row="1" Grid.Column="2" Orientation="Vertical"
                                 Maximum="{TemplateBinding ScrollableHeight}" ViewportSize="{TemplateBinding ViewportHeight}"
                                 Value="{Binding VerticalOffset, Mode=OneWay, RelativeSource={RelativeSource TemplatedParent}}"
                                 Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"/>
                      <ScrollBar x:Name="PART_HorizontalScrollBar" Grid.Row="2" Grid.Column="1" Orientation="Horizontal"
                                 Maximum="{TemplateBinding ScrollableWidth}" ViewportSize="{TemplateBinding ViewportWidth}"
                                 Value="{Binding HorizontalOffset, Mode=OneWay, RelativeSource={RelativeSource TemplatedParent}}"
                                 Visibility="{TemplateBinding ComputedHorizontalScrollBarVisibility}"/>
                    </Grid>
                  </ControlTemplate>
                </ScrollViewer.Template>
                <ItemsPresenter SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
              </ScrollViewer>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="{StaticResource oca.Surface2}"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{StaticResource oca.AccentSoft}"/>
          <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- opções -->
    <Style TargetType="RadioButton">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Grid Background="Transparent">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Grid Width="14" Height="14" VerticalAlignment="Center">
                <Ellipse x:Name="ring" Fill="{StaticResource oca.Field}" Stroke="{StaticResource oca.FieldLine}" StrokeThickness="1"/>
                <Ellipse x:Name="dot" Width="6" Height="6" Fill="{StaticResource oca.Accent}" Visibility="Collapsed"/>
              </Grid>
              <ContentPresenter Grid.Column="1" Margin="8,0,0,0" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ring" Property="Stroke" Value="{StaticResource oca.Accent}"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="ring" Property="Stroke" Value="{StaticResource oca.Accent}"/>
                <Setter TargetName="ring" Property="StrokeThickness" Value="2"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="ring" Property="Stroke" Value="{StaticResource oca.Accent}"/>
                <Setter TargetName="dot" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Grid Background="Transparent">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Border x:Name="box" Width="14" Height="14" CornerRadius="3" BorderThickness="1" VerticalAlignment="Center"
                      Background="{StaticResource oca.Field}" BorderBrush="{StaticResource oca.FieldLine}">
                <Path x:Name="mark" Data="M2,6 L5,9 L10,3" Stroke="{StaticResource oca.AccentInk}" StrokeThickness="1.8"
                      Visibility="Collapsed"/>
              </Border>
              <ContentPresenter Grid.Column="1" Margin="8,0,0,0" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
                <Setter TargetName="box" Property="BorderThickness" Value="2"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="box" Property="Background" Value="{StaticResource oca.Accent}"/>
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
                <Setter TargetName="mark" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
"""


def window_resources(dark=None):
    """Conteúdo de <Window.Resources>: pincéis oca.* do tema + estilos."""
    colors = theme.palette(dark)
    brushes = u"".join(
        u'\n    <SolidColorBrush x:Key="oca.{}" Color="{}"/>'.format(key, colors[key])
        for key in sorted(colors))
    values = {
        "label_width": unicode(tokens.LABEL_WIDTH),
        "r_control": unicode(tokens.RADIUS["control"]),
        "r_group": unicode(tokens.RADIUS["group"]),
        "r_panel": unicode(tokens.RADIUS["panel"]),
        "field_height": unicode(tokens.FIELD_HEIGHT),
        "button_height": unicode(tokens.BUTTON_HEIGHT),
        "font_mono": tokens.FONT_MONO,
    }
    return brushes + _fill(STYLES, values)


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
        FontFamily="@font@" FontSize="12" UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display" Tag="oca-@theme@">@icon@
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
               size="S", height=None, width=None, resources=u"", dark=None):
    """XAML completo de uma janela no padrão OCA.

    title         nome do comando (aceita o __title__ com quebras de linha).
    subtitle      descrição curta - por padrão o __doc__ do botão.
    body          XAML do conteúdo (os mesmos x:Name que o código usa).
    footer_right  botão principal primeiro (Style oca.Primary), depois o secundário.
    footer_left   ações auxiliares ou "Voltar" (opcional).
    size          "S" 460 px, "M" 580 px (altura pelo conteúdo, tamanho fixo)
                  ou "L" 1160 px (redimensionável; informe height).
    width         largura em px, só quando a janela original era maior que o padrão.
    resources     XAML extra para <Window.Resources> (estilos que o código busca com
                  FindResource ou que precisam valer na janela inteira).
    dark          None = segue o tema do Revit (padrão); True/False força o tema.
    """
    is_dark = theme.is_dark(dark)
    name = command_name(title)
    if size == "L":
        sizing = u'Height="{}" MinWidth="900" MinHeight="460"'.format(height or 780)
    else:
        sizing = u'SizeToContent="Height" ResizeMode="NoResize"'
    uri = logo_uri()
    values = {
        "title": xml_escape(name),
        "width": unicode(width or tokens.SIZES[size]),
        "sizing": sizing,
        "font": tokens.FONT,
        "theme": u"dark" if is_dark else u"light",
        "icon": _ICON_XAML.replace(u"@uri@", xml_escape(uri)) if uri else u"",
        "logo": _LOGO_XAML.replace(u"@uri@", xml_escape(uri)) if uri else u"",
        "subtitle": _SUBTITLE_XAML.replace(u"@subtitle@", xml_escape(u" ".join(unicode(subtitle).split())))
                    if subtitle else u"",
        "resources": window_resources(is_dark) + (resources or u""),
        "footer_left": footer_left,
        "footer_right": footer_right,
        "body": body,
    }
    return _fill(_WINDOW, values)
