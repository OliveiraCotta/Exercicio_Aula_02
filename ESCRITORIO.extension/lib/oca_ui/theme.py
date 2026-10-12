# -*- coding: utf-8 -*-
"""Paleta e estilos WPF do padrão visual OCA.

As janelas usam sempre o tema escuro (DARK), independente do tema do Revit.
LIGHT fica disponível para uso explícito (build_xaml(..., dark=False)).

As cores ficam só aqui. As janelas referenciam os pincéis pelas chaves
"oca.*" ({StaticResource oca.Ink2}, {StaticResource oca.Warn}...), nunca
por código hexadecimal, para que o mesmo XAML funcione nos dois temas.

Engine: IronPython 2.7 (pyRevit). Sintaxe neutra py2/py3.
"""

# Marca: #2FBAA7 (cor do logo), igual nos dois temas - logo e botão principal.
# Accent: tom do mesmo matiz com contraste >= 4,5:1 para texto e indicadores.
LIGHT = {
    "Bg": "#F3F5F6",
    "Surface": "#FFFFFF",
    "Ink": "#1D2428",
    "Ink2": "#5A666D",
    "Line": "#CDD4D8",
    "Field": "#FFFFFF",
    "FieldLine": "#A9B4BA",
    "Accent": "#1D7266",
    "AccentInk": "#FFFFFF",
    "AccentSoft": "#E3F4F1",
    "Brand": "#2FBAA7",
    "BrandInk": "#0E2226",
    "Ok": "#1D7438",
    "OkSoft": "#E5F3E9",
    "Warn": "#8A5100",
    "WarnSoft": "#FCF0DB",
    "Err": "#AE251D",
    "ErrSoft": "#FCE4E2",
    "Violet": "#6A55C2",
}

DARK = {
    "Bg": "#2A2F35",
    "Surface": "#33393F",
    "Ink": "#E7EBEE",
    "Ink2": "#A9B3BA",
    "Line": "#474F57",
    "Field": "#24292E",
    "FieldLine": "#5D666F",
    "Accent": "#3FC0AE",
    "AccentInk": "#0E2226",
    "AccentSoft": "#1E3B37",
    "Brand": "#2FBAA7",
    "BrandInk": "#0E2226",
    "Ok": "#73C98E",
    "OkSoft": "#1F3527",
    "Warn": "#EFAE55",
    "WarnSoft": "#3D3020",
    "Err": "#F28C84",
    "ErrSoft": "#432726",
    "Violet": "#B7A9F0",    # categoria extra dos relatórios (ex.: "atualizado")
}

# Larguras padrão das janelas (px): P = formulários, M = formulários longos,
# G = grades (redimensionável).
SIZES = {"S": 460, "M": 580, "L": 1160}
LABEL_WIDTH = 120


def palette(dark=True):
    return DARK if dark else LIGHT


# Estilos compartilhados. Os pincéis (oca.*) são gerados por resources() e
# vêm antes destes estilos, por isso StaticResource funciona aqui.
STYLES = u"""
    <GridLength x:Key="oca.LabelWidth">@label_width@</GridLength>

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

    <!-- seção: título em caixa-alta + filete -->
    <Style x:Key="oca.Section" TargetType="HeaderedContentControl">
      <Setter Property="Margin" Value="0,0,0,14"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="IsTabStop" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="HeaderedContentControl">
            <StackPanel>
              <DockPanel Margin="0,0,0,8">
                <ContentPresenter ContentSource="Header" DockPanel.Dock="Left" VerticalAlignment="Center"
                                  TextElement.FontSize="10.5" TextElement.FontWeight="SemiBold"
                                  TextElement.Foreground="{StaticResource oca.Ink2}"/>
                <Border Height="1" Margin="8,0,0,0" VerticalAlignment="Center" Background="{StaticResource oca.Line}"/>
              </DockPanel>
              <ContentPresenter/>
            </StackPanel>
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
      <Setter Property="BorderThickness" Value="3,0,0,0"/>
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
      <Setter Property="Background" Value="{StaticResource oca.Surface}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="2"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
    </Style>

    <!-- botões: secundário (padrão), principal e sem borda (Voltar) -->
    <Style TargetType="Button">
      <Setter Property="Height" Value="26"/>
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
                    BorderThickness="1" CornerRadius="3" Padding="{TemplateBinding Padding}">
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

    <!-- campos -->
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="24"/>
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
                    BorderThickness="1" CornerRadius="2">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" Focusable="False"
                            VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                            HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
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
              BorderThickness="1" CornerRadius="2">
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
      <Setter Property="Height" Value="24"/>
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
                     AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                <Border Background="{StaticResource oca.Surface}" BorderBrush="{StaticResource oca.FieldLine}" BorderThickness="1"
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
            <Border x:Name="bd" Background="Transparent" Padding="{TemplateBinding Padding}">
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
      <Setter Property="Background" Value="{StaticResource oca.Surface}"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="MaxWidth" Value="420"/>
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
      <Setter Property="Background" Value="{StaticResource oca.Bg}"/>
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

    <!-- grades -->
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{StaticResource oca.Field}"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="{StaticResource oca.Field}"/>
      <Setter Property="AlternatingRowBackground" Value="{StaticResource oca.Field}"/>
      <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="VerticalGridLinesBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="RowHeaderWidth" Value="0"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="{StaticResource oca.Surface}"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="0,0,1,1"/>
    </Style>
    <Style TargetType="DataGridRow">
      <Setter Property="Background" Value="{StaticResource oca.Field}"/>
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
              <Border x:Name="box" Width="14" Height="14" CornerRadius="2" BorderThickness="1" VerticalAlignment="Center"
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


def resources(dark=True):
    """Conteúdo de <Window.Resources>: pincéis oca.* + estilos."""
    colors = palette(dark)
    brushes = u"".join(
        u'\n    <SolidColorBrush x:Key="oca.{}" Color="{}"/>'.format(key, colors[key])
        for key in sorted(colors))
    return brushes + STYLES.replace(u"@label_width@", u"{}".format(LABEL_WIDTH))
