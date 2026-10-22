# -*- coding: utf-8 -*-
# GERENCIADOR DE PARÂMETROS COMPARTILHADOS  (pyRevit - Revit 2024/2025/2026;
# grupos de parâmetro também tratados para 2021-2023, ver lib/oca_params/revit_ops.py)
#
# (Comentário, não docstring. O tooltip do botão vem SÓ de __doc__ abaixo, que
# precisa ser UMA string literal: o pyRevit 6 lê o __doc__ do código-fonte sem
# executar o script; numa concatenação ("a" "b") ele pega só o primeiro trecho.)
#
# Janela com duas abas principais (como as abas do Explorador de Arquivos) e a
# barra do arquivo TXT logo abaixo, comum às duas:
#
# A. VINCULAR DO TXT AO PROJETO - passo a passo, com Voltar / Avançar no rodapé:
#   Passo 1 · Escolher parâmetros: tabela do TXT (nome, grupo de definição,
#     tipo de dado, GUID, situação no projeto), busca sem acento, filtro por
#     grupo do TXT e por situação (todas / novos / já no projeto), "só
#     marcados". Seleção múltipla: Shift/Ctrl + clique seleciona linhas e um
#     clique na caixa ✓ de uma delas marca (ou desmarca) todas; Shift + clique
#     numa caixa marca o intervalo desde a última caixa clicada; Espaço alterna.
#     As marcas definem o lote e continuam lá ao voltar.
#   Passo 2 · Configurar: lista do lote; categorias numa coluna própria (1/3 da
#     largura); grupo de propriedades e Instância/Tipo; "Aplicar às
#     selecionadas" grava grupo, vínculo e categorias nas linhas selecionadas.
#   Passo 3 · Revisar e aplicar: o que será feito com cada parâmetro e por quê;
#     "Aplicar no projeto" pede confirmação e grava tudo em UMA transação.
#
# B. PARÂMETROS NO PROJETO:
#   Todos os parâmetros compartilhados do modelo (vínculo, grupo, categorias,
#   origem, GUID, situação no TXT). À direita, a edição do parâmetro de
#   projeto selecionado: grupo de propriedades e categorias -> "Atualizar no
#   modelo" (confirmação; retirar categoria apaga os valores preenchidos nela
#   e o número de valores é mostrado antes). "Exportar para o TXT" grava os
#   marcados que estão fora do TXT, SEMPRE no grupo "Exported Parameters".
#
# Regras (detalhes em lib/oca_params/plan.py e revit_ops.py):
#   * O TXT só é gravado pela exportação (aba B), com confirmação e cópia de
#     segurança antes; o GUID do projeto é mantido.
#   * As definições vêm do próprio Revit (OpenSharedParameterFile), pelo GUID.
#     O caminho do arquivo de parâmetros do usuário é trocado só durante a
#     leitura e a gravação e restaurado logo depois.
#   * Vincular nunca exclui nada; parâmetro já vinculado só recebe categorias
#     novas. Instância <-> Tipo de parâmetro existente: não é feito.
#   * Qualquer falha desfaz a transação inteira (RollBack).
#   * Tabelas com ✓ têm ordem fixa (coluna Ordem, sem ordenar pelo cabeçalho):
#     marcar uma linha não a move. Antes, clicar no cabeçalho "✓" ordenava a
#     tabela por marcado/desmarcado e cada marcação mudava a linha de lugar.

__title__ = u"Gerenciador\nde Parâmetros\nCompartilhados"
__doc__ = u"Gerência, importa e exporta os parâmetros compartilhados através dos parâmetros existentes no projeto e/ou do TXT de parâmetros."

import os
import re
from collections import OrderedDict

# Referências .NET ANTES de qualquer "from System.<namespace> import": no
# IronPython um namespace .NET só vira módulo importável depois que o assembly
# que o contém foi referenciado no motor do script (clr.AddReference). O WPF
# já estar carregado no processo do Revit não basta. Sem isso: "ImportError:
# No module named Windows". Os nomes valem no .NET Framework 4.8 (Revit 2024)
# e no .NET 8 (Revit 2025/2026).
import clr
clr.AddReference("PresentationFramework")    # System.Windows.Controls (DataGrid, CheckBox...)
clr.AddReference("PresentationCore")         # System.Windows (Visibility, RoutedEventHandler), Keyboard
clr.AddReference("WindowsBase")              # System.Windows.Input (Key) e tipos base do WPF
clr.AddReference("System.Data")              # System.Data (DataTable)
try:
    clr.AddReference("System.Data.Common")   # .NET 8: onde DataTable está de fato (Revit 2025+)
except Exception:                            # .NET Framework 4.8 (Revit 2024): não existe, nem precisa
    pass

from pyrevit import revit, script, forms     # pyRevit também referencia o WPF (pyrevit.framework)

import System
from System.Data import DataTable
from System.Windows import Visibility, RoutedEventHandler
from System.Windows.Controls import CheckBox, DataGridEditingUnit
from System.Windows.Controls.Primitives import ButtonBase
from System.Windows.Input import Key, Keyboard

try:
    # pacotes da pasta lib da extensão (o pyRevit põe <extensão>\lib no caminho de importação)
    from oca_ds import build_xaml, alert_title
    from oca_params import spfile
    from oca_params import plan as P
    from oca_params import revit_ops as R
except ImportError as _missing:
    forms.alert(u"Arquivos do comando não encontrados.",
                sub_msg=u"{}\n\nCopie as pastas 'lib\\oca_ds' e 'lib\\oca_params' para a pasta 'lib' "
                        u"da mesma extensão deste botão e use Recarregar no pyRevit.".format(_missing),
                title=u"OCA · Gerenciador de Parâmetros Compartilhados")
    script.exit()

ALERT_TITLE = alert_title(__title__)   # "OCA · Gerenciador de Parâmetros Compartilhados"

doc = revit.doc
config = script.get_config()

to_unicode = R.to_unicode

ALL_GROUPS = u"Todos os grupos do arquivo"
STATUS_FILTERS = OrderedDict([
    (u"Todas as situações", u""),
    (u"Novos", u"SitKey = 'new'"),
    (u"Já no projeto", u"SitKey = 'bound'"),
])
ALL_CAT_TYPES = u"Todos os tipos"
PROJ_ORIGINS = OrderedDict([
    (u"Todas as origens", u""),
    (u"Parâmetros de projeto", u"OrigemKey = 'bound'"),
    (u"Só em famílias", u"OrigemKey = 'family'"),
])
PROJ_TXT = OrderedDict([
    (u"Todas (TXT)", u""),
    (u"Fora do TXT", u"TxtKey = 'out'"),
    (u"Já no TXT", u"TxtKey IN ('in', 'rename')"),
    (u"Conflito de nome", u"TxtKey = 'conflict'"),
])
STEP_FILE, STEP_CFG, STEP_REVIEW = 1, 2, 3
STEP_TEXT = {
    STEP_FILE: u"Passo 1 de 3 · Marque na coluna ✓ os parâmetros do TXT que vão para o projeto. "
               u"Para marcar vários de uma vez: selecione as linhas com Shift (ou Ctrl) + clique e "
               u"clique na caixa ✓ de uma delas. Depois, clique em Avançar.",
    STEP_CFG: u"Passo 2 de 3 · Selecione linhas na lista (Shift/Ctrl + clique), marque as categorias, "
              u"escolha o grupo e Instância ou Tipo, e clique em 'Aplicar às selecionadas'. "
              u"Todos precisam de ao menos uma categoria para avançar.",
    STEP_REVIEW: u"Passo 3 de 3 · Confira o que será feito com cada parâmetro e clique em "
                 u"'Aplicar no projeto'. Nada é gravado antes disso.",
}


def fold(text):
    """Texto para busca: minúsculo e sem acento ('Área' acha 'area')."""
    text = to_unicode(text).lower()
    try:
        import unicodedata
        text = u"".join(c for c in unicodedata.normalize("NFD", text)
                        if unicodedata.category(c) != "Mn")
    except Exception:
        pass
    return text


def natural_key(text):
    """'Piso 2' antes de 'Piso 10', sem diferença de acento ou maiúscula."""
    return [(0, int(t), u"") if t.isdigit() else (1, 0, t)
            for t in re.split(u"(\\d+)", fold(text)) if t]


def file_order(spf):
    """Parâmetros por grupo (ordem do arquivo) e nome. O Revit grava as linhas
    PARAM fora de ordem; a ordem da tabela não altera nada no arquivo."""
    rank = dict((gid, i) for i, gid in enumerate(spf.group_order))
    return sorted(spf.params, key=lambda p: (rank.get(p.group_id, len(rank)), natural_key(p.name)))


def like_escape(text):
    """Escapa texto para DataView.RowFilter ... LIKE '%texto%'."""
    out = []
    for ch in text:
        if ch in u"[]*%":
            out.append(u"[" + ch + u"]")
        elif ch == u"'":
            out.append(u"''")
        else:
            out.append(ch)
    return u"".join(out)


def search_filter(text, column=u"Busca"):
    terms = [t for t in fold(text).split() if t]
    return u" AND ".join(u"{} LIKE '%{}%'".format(column, like_escape(t)) for t in terms)


def eq_filter(column, value):
    return u"{} = '{}'".format(column, to_unicode(value).replace(u"'", u"''"))


def new_table(name, columns):
    """DataTable com colunas (nome, tipo .NET). Binding direto no DataGrid.

    Com coluna "Ordem", a vista fica ordenada por ela (ordem estável). As
    tabelas com ✓ não deixam ordenar pelo cabeçalho: ordenadas por "Sel", cada
    marcação mudava a linha de lugar."""
    table = DataTable(name)
    for col, clr_type in columns:
        table.Columns.Add(col, clr.GetClrType(clr_type))
    if table.Columns.Contains("Ordem"):
        table.DefaultView.Sort = u"Ordem ASC"
    return table


def shift_down():
    return Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)


def check_targets(rows, row, anchor, selected_before, shift):
    """Linhas que recebem o valor da caixa ✓ clicada (além dela mesma).

    rows             linhas visíveis, na ordem da tabela
    row              linha da caixa clicada
    anchor           linha da última caixa clicada (ou None)
    selected_before  linhas selecionadas antes do clique
    shift            Shift pressionado
    - Shift + clique: intervalo entre a última caixa clicada e esta;
    - clique numa linha que faz parte de uma seleção de 2+ linhas: todas elas;
    - senão: só a própria (lista vazia = nada a propagar)."""
    if shift and anchor is not None and anchor in rows and row in rows:
        i, j = rows.index(anchor), rows.index(row)
        return rows[min(i, j):max(i, j) + 1]
    if len(selected_before) > 1 and row in selected_before:
        return list(selected_before)
    return []


# ------------------------------------------------------------------
# Interface (WPF - Design System OCA, lib/oca_ds)
# ------------------------------------------------------------------
RESOURCES = u"""
    <Style x:Key="cell" TargetType="TextBlock">
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="Margin" Value="4,0"/>
      <Setter Property="ToolTip" Value="{Binding Text, RelativeSource={RelativeSource Self}}"/>
      <Style.Triggers>
        <Trigger Property="Text" Value=""><Setter Property="ToolTip" Value="{x:Null}"/></Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="mono" TargetType="TextBlock" BasedOn="{StaticResource cell}">
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="11.5"/>
    </Style>
    <Style x:Key="wrap" TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="4,3"/>
    </Style>
    <!-- situação / ação: cor pelo nível (ok, warn, err, off) -->
    <Style x:Key="status" TargetType="TextBlock" BasedOn="{StaticResource cell}">
      <Style.Triggers>
        <DataTrigger Binding="{Binding Nivel}" Value="ok">
          <Setter Property="Foreground" Value="{StaticResource oca.Ok}"/>
          <Setter Property="FontWeight" Value="SemiBold"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Nivel}" Value="warn">
          <Setter Property="Foreground" Value="{StaticResource oca.Warn}"/>
          <Setter Property="FontWeight" Value="SemiBold"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Nivel}" Value="err">
          <Setter Property="Foreground" Value="{StaticResource oca.Err}"/>
          <Setter Property="FontWeight" Value="SemiBold"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Nivel}" Value="off">
          <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="statusWrap" TargetType="TextBlock" BasedOn="{StaticResource status}">
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="TextTrimming" Value="None"/>
      <Setter Property="Margin" Value="4,3"/>
    </Style>
    <!-- definição inutilizável: linha inteira em cinza -->
    <Style x:Key="paramRow" TargetType="DataGridRow" BasedOn="{StaticResource {x:Type DataGridRow}}">
      <Style.Triggers>
        <DataTrigger Binding="{Binding Usavel}" Value="False">
          <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="tight" TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
      <Setter Property="Margin" Value="0"/>
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Focusable" Value="False"/>
    </Style>
    <Style x:Key="inline" TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
      <Setter Property="Margin" Value="0,0,14,0"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="fieldLabel" TargetType="TextBlock" BasedOn="{StaticResource oca.Label}">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
    </Style>
    <!-- passo a passo: caixa de cada etapa (cores trocadas em código pelos pincéis oca.*) -->
    <Style x:Key="stepBox" TargetType="Border">
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="{StaticResource oca.RadiusGroup}"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>
    <Style x:Key="stepDot" TargetType="Border">
      <Setter Property="Width" Value="26"/>
      <Setter Property="Height" Value="26"/>
      <Setter Property="CornerRadius" Value="13"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="stepArrow" TargetType="TextBlock">
      <Setter Property="Text" Value="›"/>
      <Setter Property="FontSize" Value="22"/>
      <Setter Property="Margin" Value="8,0"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
    </Style>
    <!-- abas principais (como as abas do Explorador de Arquivos): faixa mais escura,
         aba ativa da cor da página e encaixada nela -->
    <Style x:Key="tabStrip" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource oca.Line}"/>
      <Setter Property="CornerRadius" Value="8,8,0,0"/>
      <Setter Property="Padding" Value="6,6,6,0"/>
    </Style>
    <Style x:Key="tabPage" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource oca.Bg}"/>
      <Setter Property="BorderBrush" Value="{StaticResource oca.Line}"/>
      <Setter Property="BorderThickness" Value="1,0,1,1"/>
      <Setter Property="CornerRadius" Value="0,0,8,8"/>
      <Setter Property="Padding" Value="12,12,12,6"/>
    </Style>
    <Style x:Key="browserTab" TargetType="RadioButton">
      <Setter Property="Foreground" Value="{StaticResource oca.Ink2}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="0,0,4,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="1,1,1,0"
                    CornerRadius="8,8,0,0" Padding="18,8,18,8" MinWidth="200">
              <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Left"
                                TextElement.Foreground="{TemplateBinding Foreground}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource oca.Surface2}"/>
                <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource oca.Bg}"/>
                <Setter Property="Foreground" Value="{StaticResource oca.Ink}"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource oca.Accent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>"""

_GRID = u"""AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
              CanUserReorderColumns="False" CanUserResizeRows="False"
              SelectionUnit="FullRow" AlternationCount="2" """

# tabelas com caixa ✓: ordem fixa (sem ordenar pelo cabeçalho) para nada mudar de lugar
_CHECK_GRID = _GRID + u"""CanUserSortColumns="False" SelectionMode="Extended" """


def _check_col(enabled=None):
    return u"""
                  <DataGridTemplateColumn Header="✓" Width="34" CanUserResize="False">
                    <DataGridTemplateColumn.CellTemplate>
                      <DataTemplate>
                        <CheckBox Style="{StaticResource tight}"%s
                                  IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}"/>
                      </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                  </DataGridTemplateColumn>""" % (
        u' IsEnabled="{Binding %s}"' % enabled if enabled else u"")


def _step(n, title, subtitle):
    return u"""
            <Border x:Name="st%(n)d" Grid.Column="%(col)d" Style="{StaticResource stepBox}"
                    ToolTip="Ir para o passo %(n)d">
              <DockPanel>
                <Border x:Name="sd%(n)d" Style="{StaticResource stepDot}">
                  <TextBlock x:Name="sn%(n)d" Text="%(n)d" FontWeight="Bold"
                             HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                  <TextBlock x:Name="stt%(n)d" Text="%(title)s" FontWeight="SemiBold"/>
                  <TextBlock Text="%(subtitle)s" Style="{StaticResource oca.Hint}"/>
                </StackPanel>
              </DockPanel>
            </Border>""" % {"n": n, "col": (n - 1) * 2, "title": title, "subtitle": subtitle}


def _cat_picker(prefix, hint):
    """Busca + filtro por tipo + botões + lista de categorias com ✓ (passo 2 e edição)."""
    return u"""
                <DockPanel>
                  <Grid DockPanel.Dock="Top" Margin="0,0,0,6">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="f_%(p)scat_search" ToolTip="Buscar categoria pelo nome (sem diferença de acento)"/>
                    <ComboBox x:Name="cb_%(p)scat_type" Grid.Column="1" Width="120" Margin="8,0,0,0"
                              ToolTip="Modelo, Anotação ou Analítico"/>
                  </Grid>
                  <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6" LastChildFill="False">
                    <Button x:Name="b_%(p)scat_all" Content="Marcar visíveis" Style="{StaticResource oca.Compact}"
                            ToolTip="Marca as categorias que estão na lista com a busca e o filtro atuais"/>
                    <Button x:Name="b_%(p)scat_none" Content="Desmarcar todas" Style="{StaticResource oca.Compact}"
                            Margin="6,0,0,0"/>
                    <CheckBox x:Name="chk_%(p)scat_only" Content="Só marcadas" DockPanel.Dock="Right"
                              Style="{StaticResource inline}" Margin="0"/>
                  </DockPanel>
                  <TextBlock x:Name="tb_%(p)scat_count" DockPanel.Dock="Bottom" Style="{StaticResource oca.Hint}"
                             Margin="0,6,0,0"/>
                  <TextBlock DockPanel.Dock="Bottom" Style="{StaticResource oca.Hint}" Margin="0,4,0,0"
                             Text="%(hint)s"/>
                  <DataGrid x:Name="grid_%(p)scats" """ % {"p": prefix, "hint": hint} + _CHECK_GRID + u"""RowHeight="24">
                    <DataGrid.Columns>""" + _check_col() + u"""
                      <DataGridTextColumn Header="Categoria" Binding="{Binding Nome}" Width="*" IsReadOnly="True"
                                          ElementStyle="{StaticResource cell}"/>
                      <DataGridTextColumn Header="Tipo" Binding="{Binding Tipo}" Width="86" IsReadOnly="True"
                                          ElementStyle="{StaticResource cell}"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </DockPanel>"""


_SHIFT_HINT = u"Shift/Ctrl + clique seleciona várias linhas; clique na caixa ✓ de uma delas para marcar todas."

STEP1 = u"""
            <!-- ============ PASSO 1 · ESCOLHER PARÂMETROS ============ -->
            <Grid x:Name="pg1">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*" MinWidth="160"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Buscar" Style="{StaticResource oca.Label}"/>
                <TextBox x:Name="f_search" Grid.Column="1"
                         ToolTip="Nome, GUID, grupo, tipo de dado ou descrição. Sem diferença de acento ou maiúscula."/>
                <ComboBox x:Name="cb_fgroup" Grid.Column="2" Width="220" Margin="8,0,0,0"
                          ToolTip="Filtrar pelo grupo de definição do arquivo TXT"/>
                <ComboBox x:Name="cb_fstatus" Grid.Column="3" Width="170" Margin="8,0,0,0"
                          ToolTip="Filtrar pela situação do parâmetro no projeto aberto"/>
                <CheckBox x:Name="chk_only_marked" Grid.Column="4" Content="Só marcados" Margin="12,0,0,0"
                          VerticalAlignment="Center"/>
              </Grid>
              <DataGrid x:Name="grid_params" Grid.Row="1" """ + _CHECK_GRID + u"""RowHeight="26"
                        RowStyle="{StaticResource paramRow}">
                <DataGrid.Columns>""" + _check_col(u"Usavel") + u"""
                  <DataGridTextColumn Header="Nome" Binding="{Binding Nome}" Width="2*" MinWidth="140"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="Grupo de definição (TXT)" Binding="{Binding GrupoArq}" Width="1.3*"
                                      MinWidth="110" IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="Tipo de dado" Binding="{Binding Tipo}" Width="*" MinWidth="80"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="GUID" Binding="{Binding GuidTxt}" Width="275" IsReadOnly="True"
                                      ElementStyle="{StaticResource mono}"/>
                  <DataGridTextColumn Header="Situação no projeto" Binding="{Binding Situacao}" Width="1.3*"
                                      MinWidth="110" IsReadOnly="True" ElementStyle="{StaticResource status}"/>
                  <DataGridTextColumn Header="Observações" Binding="{Binding Obs}" Width="2*" MinWidth="120"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                </DataGrid.Columns>
              </DataGrid>
              <StackPanel Grid.Row="2">
                <Border Style="{StaticResource oca.Msg.Err}" Visibility="{Binding Visibility, ElementName=tb_params_msg}">
                  <TextBlock x:Name="tb_params_msg" Style="{StaticResource oca.MsgText}"/>
                </Border>
                <DockPanel Margin="0,10,0,2" LastChildFill="False">
                  <Button x:Name="b_sel_all" Content="Selecionar todos" DockPanel.Dock="Left"
                          ToolTip="Marca todos os parâmetros disponíveis que estão visíveis com a busca e os filtros atuais"/>
                  <Button x:Name="b_sel_none" Content="Limpar seleção" DockPanel.Dock="Left" Margin="8,0,0,0"
                          ToolTip="Desmarca todos os parâmetros (inclusive os escondidos pelos filtros)"/>
                  <TextBlock x:Name="tb_count" DockPanel.Dock="Left" Margin="14,0,0,0" VerticalAlignment="Center"
                             Foreground="{StaticResource oca.Ink2}"/>
                </DockPanel>
                <TextBlock Style="{StaticResource oca.Hint}"
                           Text="Linhas cinza não podem ser usadas: o motivo está em Observações. """ + _SHIFT_HINT + u""""/>
              </StackPanel>
            </Grid>"""

STEP2 = u"""
            <!-- ============ PASSO 2 · CONFIGURAR ============ -->
            <Grid x:Name="pg2" Visibility="Collapsed">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.15*" MinWidth="320"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="*" MinWidth="300"/>
                <ColumnDefinition Width="12"/>
                <ColumnDefinition Width="300"/>
              </Grid.ColumnDefinitions>

              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DockPanel Margin="0,0,0,8" LastChildFill="False">
                  <TextBlock Text="PARÂMETROS DO LOTE" Style="{StaticResource fieldLabel}" VerticalAlignment="Center"
                             Margin="0,0,12,0"/>
                  <Button x:Name="b_cfg_all" Content="Selecionar todas" Style="{StaticResource oca.Compact}"
                          ToolTip="Seleciona todas as linhas da lista"/>
                  <Button x:Name="b_remove" Content="Remover da configuração" Style="{StaticResource oca.Compact}"
                          Margin="6,0,0,0" DockPanel.Dock="Right"
                          ToolTip="Tira as linhas selecionadas do lote e desmarca no passo 1 (não altera o projeto nem o arquivo)"/>
                </DockPanel>
                <DataGrid x:Name="grid_cfg" Grid.Row="1" """ + _GRID + u"""SelectionMode="Extended" RowHeight="26"
                          IsReadOnly="True" CanUserSortColumns="False">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Parâmetro" Binding="{Binding Nome}" Width="1.4*" MinWidth="120"
                                        ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="Grupo de propriedades" Binding="{Binding Grupo}" Width="*" MinWidth="90"
                                        ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="Vínculo" Binding="{Binding Vinculo}" Width="70"
                                        ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="Categorias" Binding="{Binding Categorias}" Width="1.3*" MinWidth="100"
                                        ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="Situação" Binding="{Binding Situacao}" Width="1.2*" MinWidth="100"
                                        ElementStyle="{StaticResource status}"/>
                  </DataGrid.Columns>
                </DataGrid>
                <Border Grid.Row="2" Style="{StaticResource oca.Msg.Info}"
                        Visibility="{Binding Visibility, ElementName=tb_cfg_msg}">
                  <TextBlock x:Name="tb_cfg_msg" Style="{StaticResource oca.MsgText}"/>
                </Border>
              </Grid>

              <HeaderedContentControl Grid.Column="2" Header="CATEGORIAS" Style="{StaticResource oca.Section}" Margin="0">
""" + _cat_picker(u"", u"Marcar uma categoria não muda a ordem da lista.") + u"""
              </HeaderedContentControl>

              <StackPanel Grid.Column="4">
                <HeaderedContentControl Header="GRUPO DE PROPRIEDADES" Style="{StaticResource oca.Section}">
                  <StackPanel>
                    <ComboBox x:Name="cb_group"
                              ToolTip="Seção da paleta Propriedades em que o parâmetro aparece (lista da versão do Revit em uso)"/>
                    <TextBlock Style="{StaticResource oca.Hint}" Margin="0,6,0,0"
                               Text="Onde o parâmetro aparece na paleta Propriedades. Não é o grupo de definição do TXT."/>
                  </StackPanel>
                </HeaderedContentControl>
                <HeaderedContentControl Header="TIPO DE VÍNCULO" Style="{StaticResource oca.Section}">
                  <StackPanel>
                    <RadioButton x:Name="rb_inst" GroupName="kind" IsChecked="True"
                                 Content="Instância: cada elemento tem o seu próprio valor"/>
                    <RadioButton x:Name="rb_type" GroupName="kind"
                                 Content="Tipo: um valor por tipo, igual em todas as instâncias dele"/>
                    <Border Style="{StaticResource oca.Msg.Warn}" Margin="0,2,0,0">
                      <TextBlock Style="{StaticResource oca.MsgText}"
                                 Text="Muda onde o dado é preenchido (Propriedades x Editar tipo) e como aparece em tabelas. Depois de aplicada, trocar exige excluir o parâmetro e os valores: este comando não faz isso."/>
                    </Border>
                  </StackPanel>
                </HeaderedContentControl>
                <HeaderedContentControl Header="APLICAR NA LISTA" Style="{StaticResource oca.Section}" Margin="0">
                  <StackPanel>
                    <Button x:Name="b_apply_sel" Content="Aplicar às selecionadas" Height="32"
                            ToolTip="Grava o grupo, o vínculo e as categorias marcadas nas linhas selecionadas da lista"/>
                    <TextBlock x:Name="tb_apply_hint" Style="{StaticResource oca.Hint}" Margin="0,6,0,0"/>
                  </StackPanel>
                </HeaderedContentControl>
              </StackPanel>
            </Grid>"""

STEP3 = u"""
            <!-- ============ PASSO 3 · REVISAR E APLICAR ============ -->
            <Grid x:Name="pg3" Visibility="Collapsed">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border x:Name="bd_review" Style="{StaticResource oca.Msg.Info}" Margin="0,0,0,8"
                      Visibility="{Binding Visibility, ElementName=tb_review}">
                <TextBlock x:Name="tb_review" Style="{StaticResource oca.MsgText}"/>
              </Border>
              <StackPanel Grid.Row="1" Margin="0,0,0,8">
                <CheckBox x:Name="chk_extend" IsChecked="True"
                          Content="Adicionar as categorias novas aos parâmetros que já estão vinculados ao projeto"/>
                <TextBlock Style="{StaticResource oca.Hint}" Margin="22,0,0,0"
                           Text="As categorias atuais, o tipo de vínculo e o grupo são mantidos. Valores já preenchidos são conferidos antes de confirmar; se algum mudar, nada é gravado."/>
              </StackPanel>
              <DataGrid x:Name="grid_review" Grid.Row="2" """ + _GRID + u"""IsReadOnly="True" SelectionMode="Extended">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Parâmetro" Binding="{Binding Nome}" Width="1.3*" MinWidth="120"
                                      ElementStyle="{StaticResource wrap}"/>
                  <DataGridTextColumn Header="Ação" Binding="{Binding Acao}" Width="140"
                                      ElementStyle="{StaticResource statusWrap}"/>
                  <DataGridTextColumn Header="Grupo de propriedades" Binding="{Binding Grupo}" Width="*" MinWidth="100"
                                      ElementStyle="{StaticResource wrap}"/>
                  <DataGridTextColumn Header="Vínculo" Binding="{Binding Vinculo}" Width="72"
                                      ElementStyle="{StaticResource wrap}"/>
                  <DataGridTextColumn Header="Categorias" Binding="{Binding Categorias}" Width="2*" MinWidth="140"
                                      ElementStyle="{StaticResource wrap}"/>
                  <DataGridTextColumn Header="Observações" Binding="{Binding Obs}" Width="2*" MinWidth="140"
                                      ElementStyle="{StaticResource wrap}"/>
                </DataGrid.Columns>
              </DataGrid>
            </Grid>"""

PROJ = u"""
          <!-- ============ PARÂMETROS NO PROJETO ============ -->
          <Grid x:Name="pn_proj" Visibility="Collapsed">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="2*" MinWidth="420"/>
              <ColumnDefinition Width="14"/>
              <ColumnDefinition Width="*" MinWidth="360"/>
            </Grid.ColumnDefinitions>
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*" MinWidth="140"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Buscar" Style="{StaticResource oca.Label}"/>
                <TextBox x:Name="f_psearch" Grid.Column="1"
                         ToolTip="Nome, GUID, tipo de dado, grupo ou categoria. Sem diferença de acento ou maiúscula."/>
                <ComboBox x:Name="cb_porigin" Grid.Column="2" Width="180" Margin="8,0,0,0"
                          ToolTip="Parâmetros de projeto (vinculados) ou que existem só por famílias carregadas"/>
                <ComboBox x:Name="cb_ptxt" Grid.Column="3" Width="150" Margin="8,0,0,0"
                          ToolTip="Comparação com o TXT carregado (pelo GUID e pelo nome)"/>
                <CheckBox x:Name="chk_p_only" Grid.Column="4" Content="Só marcados" Margin="12,0,0,0"
                          VerticalAlignment="Center"/>
              </Grid>
              <DataGrid x:Name="grid_proj" Grid.Row="1" """ + _CHECK_GRID + u"""RowHeight="26">
                <DataGrid.Columns>""" + _check_col(u"Pode") + u"""
                  <DataGridTextColumn Header="Nome" Binding="{Binding Nome}" Width="1.6*" MinWidth="130"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="Tipo de dado" Binding="{Binding Tipo}" Width="*" MinWidth="80"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="Vínculo" Binding="{Binding Vinculo}" Width="72"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="Grupo de propriedades" Binding="{Binding Grupo}" Width="*" MinWidth="100"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="Categorias" Binding="{Binding Categorias}" Width="1.4*" MinWidth="110"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="Origem" Binding="{Binding Origem}" Width="*" MinWidth="90"
                                      IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  <DataGridTextColumn Header="No TXT" Binding="{Binding NoTxt}" Width="1.3*" MinWidth="110"
                                      IsReadOnly="True" ElementStyle="{StaticResource status}"/>
                  <DataGridTextColumn Header="GUID" Binding="{Binding Guid}" Width="275" IsReadOnly="True"
                                      ElementStyle="{StaticResource mono}"/>
                </DataGrid.Columns>
              </DataGrid>
              <StackPanel Grid.Row="2">
                <Border Style="{StaticResource oca.Msg.Err}" Visibility="{Binding Visibility, ElementName=tb_proj_msg}">
                  <TextBlock x:Name="tb_proj_msg" Style="{StaticResource oca.MsgText}"/>
                </Border>
                <DockPanel Margin="0,10,0,4" LastChildFill="False">
                  <Button x:Name="b_psel_all" Content="Selecionar todos" DockPanel.Dock="Left"
                          ToolTip="Marca para exportar os parâmetros visíveis que ainda não estão no TXT"/>
                  <Button x:Name="b_psel_none" Content="Limpar seleção" DockPanel.Dock="Left" Margin="8,0,0,0"
                          ToolTip="Desmarca todos (inclusive os escondidos pelos filtros)"/>
                  <TextBlock x:Name="tb_pcount" DockPanel.Dock="Left" Margin="14,0,0,0" VerticalAlignment="Center"
                             Foreground="{StaticResource oca.Ink2}"/>
                </DockPanel>
                <Border Style="{StaticResource oca.Card}">
                  <TextBlock x:Name="tb_export_info" TextWrapping="Wrap" Foreground="{StaticResource oca.Ink}"/>
                </Border>
                <TextBlock Style="{StaticResource oca.Hint}" Margin="0,4,0,0"
                           Text="Só parâmetros fora do TXT podem ser marcados (✓) para exportar. """ + _SHIFT_HINT + u""" Clique numa linha para editá-la à direita."/>
              </StackPanel>
            </Grid>

            <HeaderedContentControl Grid.Column="2" Header="EDITAR PARÂMETRO SELECIONADO" Style="{StaticResource oca.Section}" Margin="0">
              <DockPanel>
                <StackPanel DockPanel.Dock="Top" Margin="0,0,0,8">
                  <TextBlock x:Name="tb_d_name" Style="{StaticResource oca.Value}" Text="Clique numa linha da tabela."/>
                  <TextBlock x:Name="tb_d_guid" Style="{StaticResource oca.Mono}" Margin="0,2,0,0" TextWrapping="Wrap"
                             Foreground="{StaticResource oca.Ink2}"/>
                  <TextBlock x:Name="tb_d_info" Style="{StaticResource oca.Hint}" Margin="0,6,0,0"/>
                </StackPanel>
                <StackPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
                  <Border Style="{StaticResource oca.Msg.Warn}" Margin="0,0,0,8"
                          Visibility="{Binding Visibility, ElementName=tb_edit_msg}">
                    <TextBlock x:Name="tb_edit_msg" Style="{StaticResource oca.MsgText}"/>
                  </Border>
                  <Button x:Name="b_update" Content="Atualizar no modelo" Height="32"
                          ToolTip="Grava no modelo o grupo e as categorias escolhidos (pede confirmação; Ctrl+Z desfaz)"/>
                  <TextBlock Style="{StaticResource oca.Hint}" Margin="0,6,0,0"
                             Text="Instância/Tipo não muda. Tirar uma categoria apaga os valores preenchidos nela: a quantidade aparece antes de confirmar. Alterações não gravadas se perdem ao trocar de linha."/>
                </StackPanel>
                <DockPanel x:Name="pn_edit">
                  <StackPanel DockPanel.Dock="Top" Margin="0,0,0,8">
                    <TextBlock Text="Grupo de propriedades" Style="{StaticResource fieldLabel}"/>
                    <ComboBox x:Name="cb_egroup" ToolTip="Seção da paleta Propriedades em que o parâmetro aparece"/>
                  </StackPanel>
                  <TextBlock DockPanel.Dock="Top" Text="Categorias" Style="{StaticResource fieldLabel}"/>
""" + _cat_picker(u"e", u"Desmarcar uma categoria vinculada a retira do parâmetro.") + u"""
                </DockPanel>
              </DockPanel>
            </HeaderedContentControl>
          </Grid>"""

BODY = u"""
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <!-- abas principais, no estilo do Explorador de Arquivos -->
    <Border Style="{StaticResource tabStrip}">
      <StackPanel Orientation="Horizontal">
        <RadioButton x:Name="rb_mode_link" GroupName="mode" IsChecked="True" Style="{StaticResource browserTab}"
                     Content="Vincular do TXT ao projeto"
                     ToolTip="Escolher parâmetros do TXT, configurar e vincular ao projeto (passo a passo)"/>
        <RadioButton x:Name="rb_mode_proj" GroupName="mode" Style="{StaticResource browserTab}"
                     Content="Parâmetros no projeto"
                     ToolTip="Ver e editar os parâmetros compartilhados do projeto e exportar para o TXT"/>
      </StackPanel>
    </Border>

    <Border Grid.Row="1" Style="{StaticResource tabPage}">
      <DockPanel>
        <!-- barra do arquivo TXT: vale para as duas abas -->
        <StackPanel DockPanel.Dock="Top" Margin="0,0,0,10">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Arquivo TXT" Style="{StaticResource oca.Label}" FontWeight="SemiBold"/>
            <TextBox x:Name="f_path" Grid.Column="1" IsReadOnly="True"
                     ToolTip="Arquivo em uso. Só é gravado ao exportar parâmetros do projeto, com confirmação e cópia de segurança antes."/>
            <Button x:Name="b_load" Grid.Column="2" Content="Carregar arquivo…" Margin="8,0,0,0"
                    ToolTip="Escolher um arquivo .txt de parâmetros compartilhados"/>
            <Button x:Name="b_reload" Grid.Column="3" Content="Recarregar" Margin="8,0,0,0"
                    ToolTip="Ler de novo o mesmo arquivo (depois de editá-lo no Revit)"/>
          </Grid>
          <TextBlock x:Name="tb_file_info" Style="{StaticResource oca.Hint}" Margin="0,6,0,0"/>
          <Border Style="{StaticResource oca.Msg.Warn}" Visibility="{Binding Visibility, ElementName=tb_file_warn}">
            <TextBlock x:Name="tb_file_warn" Style="{StaticResource oca.MsgText}"/>
          </Border>
        </StackPanel>

        <Grid>
          <!-- ============ VINCULAR DO TXT: PASSO A PASSO ============ -->
          <DockPanel x:Name="pn_link">
            <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>""" + _step(1, u"Escolher parâmetros", u"Marque no TXT o que vai para o projeto") + u"""
              <TextBlock Grid.Column="1" Style="{StaticResource stepArrow}"/>""" + _step(2, u"Configurar", u"Grupo, Instância/Tipo e categorias") + u"""
              <TextBlock Grid.Column="3" Style="{StaticResource stepArrow}"/>""" + _step(3, u"Revisar e aplicar", u"Confira e grave no modelo") + u"""
            </Grid>
            <Border DockPanel.Dock="Top" Style="{StaticResource oca.Msg.Info}" Margin="0,0,0,10">
              <TextBlock x:Name="tb_step" Style="{StaticResource oca.MsgText}"/>
            </Border>
            <Grid>""" + STEP1 + STEP2 + STEP3 + u"""
            </Grid>
          </DockPanel>
""" + PROJ + u"""
        </Grid>
      </DockPanel>
    </Border>
  </Grid>"""

FOOTER_LEFT = u"""
  <Button x:Name="b_back" Content="←  Voltar" Style="{StaticResource oca.Ghost}" Margin="0,0,12,0"
          ToolTip="Volta ao passo anterior (as marcações e configurações são mantidas)"/>
  <TextBlock x:Name="tb_footer" VerticalAlignment="Center" Foreground="{StaticResource oca.Ink2}"/>"""

FOOTER = u"""
  <Button x:Name="b_next" Content="Avançar  →" MinWidth="150" Style="{StaticResource oca.Primary}"
          ToolTip="Vai para o próximo passo"/>
  <Button x:Name="b_apply" Content="Aplicar no projeto" MinWidth="150" Style="{StaticResource oca.Primary}"
          Visibility="Collapsed"
          ToolTip="Pede confirmação e grava tudo numa única transação"/>
  <Button x:Name="b_export" Content="Exportar para o TXT" MinWidth="150" Style="{StaticResource oca.Primary}"
          Visibility="Collapsed"
          ToolTip="Grava no TXT carregado, no grupo 'Exported Parameters', os parâmetros marcados (pede confirmação e faz uma cópia do TXT antes)"/>
  <Button x:Name="b_cancel" Content="Cancelar" Margin="8,0,0,0" ToolTip="Fecha sem alterar o projeto"/>"""

XAML = build_xaml(title=__title__, subtitle=__doc__, body=BODY,
                  footer_right=FOOTER, footer_left=FOOTER_LEFT, resources=RESOURCES,
                  size="L", width=1320, height=920, min_width=1120, min_height=720)


class ParamManagerWindow(forms.WPFWindow):
    def __init__(self, document):
        forms.WPFWindow.__init__(self, XAML, literal_string=True)
        self.doc = document
        self.app = document.Application
        self.applied = False
        self._bulk = False

        # dados do arquivo
        self.path = u""
        self.spf = None
        self.spf_by_guid = {}
        self.defs = {}              # guid -> info lida pelo Revit
        self.revit_error = None
        # dados do projeto e da versão do Revit
        self.groups = R.list_property_groups()
        self.group_keys = [k for k, _ in self.groups]
        self.group_labels = dict(self.groups)
        self.categories = R.list_bindable_categories(document)
        self.cat_by_id = dict((c["id"], c["cat"]) for c in self.categories)
        self.cat_names = dict((c["id"], c["name"]) for c in self.categories)
        self.no_type = R.no_type_category_ids()
        self.project = R.scan_project(document, self.groups)
        # lote: guid -> {"guid", "name", "group_key", "kind", "cats"}
        self.cfg = OrderedDict()
        self.plan_rows = []
        self.step = STEP_FILE
        # aba "Parâmetros no projeto"
        self.proj_items = {}        # guid -> item de R.list_project_shared
        self.txt_names = {}         # nome no TXT -> set(guid)
        self.proj_mode = False
        self.edit_guid = None       # parâmetro mostrado no painel de edição

        self._build_tables()
        self._init_controls()
        self._wire_events()
        self._refresh_cfg()
        self._fill_proj()
        self._go(STEP_FILE)

        start = to_unicode(config.get_option("last_path", u"") or u"")
        if not (start and os.path.isfile(start)):
            start = to_unicode(self.app.SharedParametersFilename or u"")
        if start and os.path.isfile(start):
            self._load(start)
        else:
            self.tb_file_info.Text = u"Nenhum arquivo carregado. Clique em 'Carregar arquivo…'."
        self._update_export_info()

    # ---------------------------------------------------------- montagem
    def _build_tables(self):
        S, B, I = System.String, System.Boolean, System.Int32
        self.t_params = new_table(u"params", [
            ("Ordem", I), ("Sel", B), ("Usavel", B), ("Guid", S), ("Nome", S), ("GrupoArq", S),
            ("Tipo", S), ("GuidTxt", S), ("Situacao", S), ("SitKey", S), ("Nivel", S), ("Obs", S),
            ("Busca", S)])
        self.t_cfg = new_table(u"cfg", [
            ("Ordem", I), ("Guid", S), ("Nome", S), ("Grupo", S), ("Vinculo", S), ("Categorias", S),
            ("Situacao", S), ("Nivel", S)])
        self.t_cats = self._category_table(u"cats")
        self.t_ecats = self._category_table(u"ecats")
        self.t_proj = new_table(u"proj", [
            ("Ordem", I), ("Sel", B), ("Pode", B), ("Guid", S), ("Nome", S), ("Tipo", S),
            ("Vinculo", S), ("Grupo", S), ("Categorias", S), ("Origem", S), ("OrigemKey", S),
            ("NoTxt", S), ("TxtKey", S), ("Nivel", S), ("Busca", S)])
        self.t_review = new_table(u"review", [
            ("Ordem", I), ("Nome", S), ("Acao", S), ("Grupo", S), ("Vinculo", S), ("Categorias", S),
            ("Obs", S), ("Nivel", S)])
        self.grid_params.ItemsSource = self.t_params.DefaultView
        self.grid_cfg.ItemsSource = self.t_cfg.DefaultView
        self.grid_cats.ItemsSource = self.t_cats.DefaultView
        self.grid_ecats.ItemsSource = self.t_ecats.DefaultView
        self.grid_review.ItemsSource = self.t_review.DefaultView
        self.grid_proj.ItemsSource = self.t_proj.DefaultView

    def _category_table(self, name):
        S, B, I = System.String, System.Boolean, System.Int32
        table = new_table(name, [("Ordem", I), ("Sel", B), ("Id", I), ("Nome", S), ("Tipo", S),
                                 ("Busca", S)])
        for i, c in enumerate(self.categories):
            row = table.NewRow()
            row["Ordem"] = i
            row["Sel"] = False
            row["Id"] = c["id"]
            row["Nome"] = c["name"]
            row["Tipo"] = c["type"]
            row["Busca"] = fold(c["name"])
            table.Rows.Add(row)
        return table

    def _init_controls(self):
        self.cb_fstatus.ItemsSource = list(STATUS_FILTERS.keys())
        self.cb_fstatus.SelectedIndex = 0
        self.cb_fgroup.ItemsSource = [ALL_GROUPS]
        self.cb_fgroup.SelectedIndex = 0
        types = [ALL_CAT_TYPES] + sorted(set(c["type"] for c in self.categories))
        for combo in (self.cb_cat_type, self.cb_ecat_type):
            combo.ItemsSource = types
            combo.SelectedIndex = 0
        labels = [label for _, label in self.groups]
        self.cb_group.ItemsSource = labels
        self.cb_egroup.ItemsSource = labels
        default = R.default_group_key(self.groups)
        self.cb_group.SelectedIndex = self.group_keys.index(default) if default in self.group_keys else -1
        self.cb_porigin.ItemsSource = list(PROJ_ORIGINS.keys())
        self.cb_porigin.SelectedIndex = 0
        self.cb_ptxt.ItemsSource = list(PROJ_TXT.keys())
        self.cb_ptxt.SelectedIndex = 0
        self._update_cat_count(u"")
        self._update_cat_count(u"e")
        self._update_count()
        self._show_edit(None)

    def _wire_events(self):
        sf = self._safe
        # arquivo
        self.b_load.Click += sf(self._on_browse)
        self.b_reload.Click += sf(self._on_reload)
        # abas principais e navegação
        self.rb_mode_link.Checked += sf(lambda s, a: self._set_mode(False))
        self.rb_mode_proj.Checked += sf(lambda s, a: self._set_mode(True))
        self.b_back.Click += sf(lambda s, a: self._go(self.step - 1))
        self.b_next.Click += sf(lambda s, a: self._go(self.step + 1))
        for n in (STEP_FILE, STEP_CFG, STEP_REVIEW):
            getattr(self, "st%d" % n).MouseLeftButtonUp += sf((lambda k: lambda s, a: self._go(k))(n))
        self.b_apply.Click += sf(self._on_apply)
        self.b_cancel.Click += sf(lambda s, a: self.Close())
        # passo 1
        self.f_search.TextChanged += sf(lambda s, a: self._filter_params())
        for combo in (self.cb_fgroup, self.cb_fstatus):
            combo.SelectionChanged += sf(lambda s, a: self._filter_params())
        self.chk_only_marked.Click += sf(lambda s, a: self._filter_params())
        self.b_sel_all.Click += sf(lambda s, a: self._set_marks(
            self.t_params, [drv.Row for drv in self.t_params.DefaultView], True, u"Usavel"))
        self.b_sel_none.Click += sf(lambda s, a: self._set_marks(
            self.t_params, list(self.t_params.Rows), False, u"Usavel"))
        self._wire_checks(self.grid_params, self.t_params, u"Usavel", self._update_count)
        self.t_params.ColumnChanged += self._on_param_changed
        # passo 2
        self.b_cfg_all.Click += sf(lambda s, a: self.grid_cfg.SelectAll())
        self.b_remove.Click += sf(self._on_remove)
        self.b_apply_sel.Click += sf(self._apply_controls)
        self.grid_cfg.SelectionChanged += sf(lambda s, a: self._update_apply_hint())
        self._wire_cat_picker(u"")
        # passo 3
        self.chk_extend.Click += sf(lambda s, a: self._refresh_cfg())
        # parâmetros no projeto
        self.f_psearch.TextChanged += sf(lambda s, a: self._filter_proj())
        for combo in (self.cb_porigin, self.cb_ptxt):
            combo.SelectionChanged += sf(lambda s, a: self._filter_proj())
        self.chk_p_only.Click += sf(lambda s, a: self._filter_proj())
        self.grid_proj.SelectionChanged += sf(self._on_proj_select)
        self.b_psel_all.Click += sf(lambda s, a: self._set_marks(
            self.t_proj, [drv.Row for drv in self.t_proj.DefaultView], True, u"Pode"))
        self.b_psel_none.Click += sf(lambda s, a: self._set_marks(
            self.t_proj, list(self.t_proj.Rows), False, u"Pode"))
        self._wire_checks(self.grid_proj, self.t_proj, u"Pode", self._update_proj_count)
        self.t_proj.ColumnChanged += self._on_proj_changed
        self._wire_cat_picker(u"e")
        self.b_update.Click += sf(self._on_update)
        self.b_export.Click += sf(self._on_export)

    def _safe(self, fn):
        """Erro num evento vira mensagem; nunca chega ao Revit."""
        def handler(sender, args):
            try:
                return fn(sender, args)
            except Exception as ex:
                forms.alert(u"Ocorreu um erro inesperado na janela.", sub_msg=to_unicode(ex),
                            title=ALERT_TITLE)
        return handler

    # ---------------------------------------------------------- seleção múltipla com ✓
    def _commit(self, grid):
        """Fecha edição pendente da tabela (evita erro ao refiltrar)."""
        try:
            grid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            try:
                grid.CancelEdit()
            except Exception:
                pass

    def _scroll_top(self, grid):
        """Depois de filtrar, volta ao início: a rolagem antiga pode ficar além do
        fim da lista e mostrar linhas vazias."""
        try:
            if grid.Items.Count:
                grid.ScrollIntoView(grid.Items[0])
        except Exception:
            pass

    def _set_marks(self, table, rows, value, enabled_col=None):
        """Marca/desmarca linhas (só as habilitadas quando marca)."""
        for grid in (self.grid_params, self.grid_cats, self.grid_ecats, self.grid_proj):
            self._commit(grid)
        self._bulk = True
        try:
            for row in rows:
                if not value or enabled_col is None or row[enabled_col]:
                    row["Sel"] = value
        finally:
            self._bulk = False
        self._after_marks(table)

    def _after_marks(self, table):
        if table is self.t_params:
            self._update_count()
        elif table is self.t_proj:
            self._update_proj_count()
        elif table is self.t_cats:
            self._update_cat_count(u"")
        elif table is self.t_ecats:
            self._update_cat_count(u"e")

    def _wire_checks(self, grid, table, enabled_col, after):
        """Seleção múltipla na coluna ✓:
        - linhas selecionadas (Shift/Ctrl + clique) e clique na caixa de uma delas:
          todas as selecionadas recebem o mesmo valor;
        - Shift + clique numa caixa: marca o intervalo desde a última caixa clicada;
        - Espaço: alterna as linhas selecionadas."""
        state = {"anchor": None, "sel": []}

        def on_down(sender, args):
            # seleção ANTES do clique (o clique pode mudar a seleção da tabela)
            state["sel"] = [d.Row for d in list(grid.SelectedItems)]

        def on_click(sender, args):
            box = args.OriginalSource
            if not isinstance(box, CheckBox):
                return
            try:
                row = box.DataContext.Row
            except Exception:
                return
            value = bool(box.IsChecked)
            targets = check_targets([d.Row for d in table.DefaultView], row, state["anchor"],
                                    state["sel"], shift_down())
            state["anchor"] = row
            if targets:
                self._set_marks(table, targets, value, enabled_col)
            else:
                after()

        def on_key(sender, args):
            if args.Key != Key.Space:
                return
            rows = [d.Row for d in list(grid.SelectedItems)]
            if rows:
                usable = [r for r in rows if enabled_col is None or r[enabled_col]]
                self._set_marks(table, rows, not all(bool(r["Sel"]) for r in usable), enabled_col)
                args.Handled = True

        grid.PreviewMouseLeftButtonDown += self._safe(on_down)
        grid.AddHandler(ButtonBase.ClickEvent, RoutedEventHandler(self._safe(on_click)))
        grid.PreviewKeyDown += self._safe(on_key)

    # ---------------------------------------------------------- passo a passo
    def _paint_steps(self):
        res = self.FindResource
        for n in (STEP_FILE, STEP_CFG, STEP_REVIEW):
            box, dot = getattr(self, "st%d" % n), getattr(self, "sd%d" % n)
            num, title = getattr(self, "sn%d" % n), getattr(self, "stt%d" % n)
            if n == self.step:
                box.Background, box.BorderBrush = res("oca.AccentSoft"), res("oca.Accent")
                dot.Background, num.Foreground = res("oca.Accent"), res("oca.AccentInk")
                num.Text, title.Foreground = unicode(n), res("oca.Ink")
            elif n < self.step:
                box.Background, box.BorderBrush = res("oca.Surface"), res("oca.Line")
                dot.Background, num.Foreground = res("oca.AccentSoft"), res("oca.Accent")
                num.Text, title.Foreground = u"✓", res("oca.Ink")
            else:
                box.Background, box.BorderBrush = res("oca.Surface"), res("oca.Line")
                dot.Background, num.Foreground = res("oca.Surface2"), res("oca.Ink2")
                num.Text, title.Foreground = unicode(n), res("oca.Ink2")

    def _go(self, step):
        """Vai para o passo pedido. Para avançar, valida o passo atual."""
        step = max(STEP_FILE, min(STEP_REVIEW, step))
        if step > self.step:
            if self.step == STEP_FILE and not self._sync_cfg():
                return
            if step == STEP_REVIEW:
                if self.step == STEP_FILE:            # clicou direto no passo 3
                    self.step = STEP_CFG
                if not self._cfg_complete():
                    step = STEP_CFG
        self.step = step
        for n in (STEP_FILE, STEP_CFG, STEP_REVIEW):
            getattr(self, "pg%d" % n).Visibility = (Visibility.Visible if n == step
                                                    else Visibility.Collapsed)
        if step == STEP_REVIEW:
            self.project = R.scan_project(self.doc, self.groups)
            self._refresh_cfg()
        self.tb_step.Text = STEP_TEXT[step]
        self._paint_steps()
        self._update_apply_hint()
        self._update_footer()

    def _sync_cfg(self):
        """Passo 1 -> 2: o lote passa a ser exatamente o que está marcado. Quem já
        estava no lote mantém a configuração; os novos recebem os controles atuais."""
        self._commit(self.grid_params)
        marked = [to_unicode(r["Guid"]) for r in self.t_params.Select(u"Sel = true", u"Ordem ASC")
                  if r["Usavel"]]
        if not marked:
            self.tb_params_msg.Text = u"Marque ao menos um parâmetro disponível (coluna ✓) para avançar."
            return False
        self.tb_params_msg.Text = u""
        group_key, kind, cats = self._controls()
        new_cfg, added = OrderedDict(), []
        for guid in marked:
            if guid in self.cfg:
                new_cfg[guid] = self.cfg[guid]
            else:
                p = self.spf_by_guid.get(guid)
                if p is None:
                    continue
                new_cfg[guid] = {"guid": guid, "name": p.name, "group_key": group_key,
                                 "kind": kind, "cats": set(cats)}
                added.append(guid)
        removed = len([g for g in self.cfg if g not in new_cfg])
        self.cfg = new_cfg
        self._refresh_cfg(select=added or None)
        parts = []
        if added:
            parts.append(u"{} parâmetro(s) novo(s) no lote (já selecionados na lista).".format(len(added)))
        if removed:
            parts.append(u"{} desmarcado(s) no passo 1 saíram do lote.".format(removed))
        self.tb_cfg_msg.Text = u" ".join(parts)
        return True

    def _cfg_complete(self):
        """Passo 2 -> 3: todos com grupo e ao menos uma categoria."""
        missing = [i["name"] for i in self.cfg.values()
                   if not i["cats"] or i["group_key"] not in self.group_labels]
        if missing:
            self.tb_cfg_msg.Text = (u"{} parâmetro(s) sem categoria ou sem grupo: {}. Selecione-os, "
                                    u"marque as categorias e clique em 'Aplicar às selecionadas'.".format(
                                        len(missing), u", ".join(missing[:5])
                                        + (u"…" if len(missing) > 5 else u"")))
            return False
        return True

    def _set_mode(self, proj):
        """Troca a aba principal; o rodapé acompanha."""
        self.proj_mode = bool(proj)
        shown, hidden = Visibility.Visible, Visibility.Collapsed
        self.pn_proj.Visibility = shown if proj else hidden
        self.pn_link.Visibility = hidden if proj else shown
        if proj:
            self.project = R.scan_project(self.doc, self.groups)
            self._fill_proj()
        self._update_footer()

    def _update_footer(self):
        shown, hidden = Visibility.Visible, Visibility.Collapsed
        if self.proj_mode:
            for b in (self.b_back, self.b_next, self.b_apply):
                b.Visibility = hidden
            self.b_export.Visibility = shown
            self.tb_footer.Text = u"{} parâmetro(s) compartilhado(s) no projeto".format(
                len(self.proj_items))
            return
        self.b_export.Visibility = hidden
        self.b_back.Visibility = shown if self.step > STEP_FILE else hidden
        self.b_next.Visibility = shown if self.step < STEP_REVIEW else hidden
        self.b_apply.Visibility = shown if self.step == STEP_REVIEW else hidden
        self.b_apply.IsEnabled = self.step == STEP_REVIEW and self._ready()
        n = len(self.cfg)
        self.tb_footer.Text = u"Lote: {} parâmetro(s)".format(n) if n else u""

    # ---------------------------------------------------------- arquivo
    def _on_browse(self, sender, args):
        folder = os.path.dirname(self.path) if self.path else None
        kw = {"file_ext": "txt", "title": u"Arquivo de parâmetros compartilhados"}
        if folder and os.path.isdir(folder):
            kw["init_dir"] = folder
        try:
            picked = forms.pick_file(**kw)
        except TypeError:                       # pyRevit sem init_dir/title
            picked = forms.pick_file(file_ext="txt")
        if picked:
            self._load(to_unicode(picked))

    def _on_reload(self, sender, args):
        if not self.path:
            self.tb_params_msg.Text = u"Nenhum arquivo carregado."
            return
        self._load(self.path)

    def _load(self, path):
        changed = path != self.path
        if self.cfg and changed:
            if not forms.alert(u"Carregar outro arquivo e limpar o lote?",
                               sub_msg=u"O lote tem {} parâmetro(s) do arquivo atual. Cada lote "
                                       u"usa um arquivo só.".format(len(self.cfg)),
                               title=ALERT_TITLE, yes=True, no=True):
                return
            self.cfg.clear()
        self.tb_params_msg.Text = u""
        try:
            spf = spfile.read_file(path)
        except spfile.SPFileError as ex:
            self.tb_params_msg.Text = u"O arquivo não pôde ser lido: {}".format(to_unicode(ex))
            return
        defs, error = R.read_definitions(self.app, path)
        self.path, self.spf, self.defs, self.revit_error = path, spf, defs, error
        self.spf_by_guid = spf.by_guid()
        self.project = R.scan_project(self.doc, self.groups)
        self.f_path.Text = path
        try:
            config.last_path = path
            script.save_config()
        except Exception:
            pass

        # parâmetros que sumiram ou ficaram inutilizáveis saem do lote
        dropped = [g for g in list(self.cfg) if not self._def_info_by_guid(g).get("usable")]
        for g in dropped:
            del self.cfg[g]
        for g, item in self.cfg.items():            # nome pode ter mudado no arquivo
            item["name"] = self.spf_by_guid[g].name

        self._fill_params(keep_marks=not changed)
        n_off = len([p for p in spf.params if not self._def_info(p).get("usable")])
        self.tb_file_info.Text = (
            u"{} parâmetro(s) em {} grupo(s) · {} indisponível(is) · codificação {} · formato {} · "
            u"Revit {}. Só é gravado ao exportar (aba Parâmetros no projeto).".format(
                len(spf.params), len(spf.groups), n_off, spf.encoding, spf.version or u"?",
                R.revit_version(self.app) or u"?"))
        warns = list(spf.issues)
        if error:
            warns.insert(0, u"{} Nenhum parâmetro pode ser usado.".format(error))
        if dropped:
            warns.append(u"{} parâmetro(s) saíram do lote: não estão mais disponíveis no "
                         u"arquivo.".format(len(dropped)))
        self.tb_file_warn.Text = u"\n".join(warns)
        if changed and self.step != STEP_FILE:
            self._go(STEP_FILE)
        self._refresh_cfg()
        self._fill_proj()
        self._update_export_info()

    def _def_info(self, p):
        """Combina a leitura do TXT com a do Revit: o parâmetro só é utilizável
        se o TXT estiver correto E o Revit tiver aberto a mesma definição."""
        if not p.usable:
            return {"usable": False, "reason": u" ".join(p.errors)}
        if self.revit_error:
            return {"usable": False, "reason": u"O Revit não abriu o arquivo."}
        info = self.defs.get(p.guid_key)
        if info is None:
            return {"usable": False,
                    "reason": u"O Revit não reconheceu esta definição (tipo de dado '{}' pode não "
                              u"existir nesta versão).".format(p.datatype)}
        return info

    def _def_info_by_guid(self, guid):
        p = self.spf_by_guid.get(guid)
        if p is None:
            return {"usable": False, "reason": u"Não está no arquivo carregado."}
        return self._def_info(p)

    # ---------------------------------------------------------- passo 1
    def _fill_params(self, keep_marks=True):
        """Remonta a tabela do TXT. As marcas (✓) são mantidas: as do lote e as
        que já estavam na tabela."""
        self._commit(self.grid_params)
        marked = set(self.cfg)
        if keep_marks:
            marked |= set(to_unicode(r["Guid"]) for r in self.t_params.Select(u"Sel = true"))
        self._bulk = True
        try:
            self.t_params.Rows.Clear()
            groups = []
            for i, p in enumerate(file_order(self.spf)):
                info = self._def_info(p)
                usable = bool(info.get("usable"))
                datatype = info.get("datatype") or p.datatype_label
                if usable:
                    sit, detail, usable = P.project_status(p.guid_key, p.name, self.project)
                    if not usable:
                        key = u"conflict"
                    elif p.guid_key in self.project["bound"]:
                        key = u"bound"
                    elif p.guid_key in self.project["shared"]:
                        key = u"family"
                    else:
                        key = u"new"
                    level = {u"new": u"ok", u"family": u"ok", u"bound": u"off"}.get(key, u"warn")
                else:
                    sit, detail, key, level = u"Indisponível", info.get("reason", u""), u"off", u"err"
                notes = [detail] if detail else []
                notes += [w for w in p.warnings]
                if p.description.strip():
                    notes.append(u"Descrição: " + p.description.strip())
                row = self.t_params.NewRow()
                row["Ordem"] = i
                row["Sel"] = usable and p.guid_key in marked
                row["Usavel"] = usable
                row["Guid"] = p.guid_key or u"#linha{}".format(p.line_no)
                row["Nome"] = p.name
                row["GrupoArq"] = p.group_name
                row["Tipo"] = datatype
                row["GuidTxt"] = p.guid
                row["Situacao"] = sit
                row["SitKey"] = key
                row["Nivel"] = level
                row["Obs"] = u" · ".join(notes)
                row["Busca"] = fold(u" ".join((p.name, p.guid, p.group_name, datatype, p.datatype,
                                               p.description)))
                self.t_params.Rows.Add(row)
                if p.group_name not in groups:
                    groups.append(p.group_name)
        finally:
            self._bulk = False
        current = self.cb_fgroup.SelectedItem
        self.cb_fgroup.ItemsSource = [ALL_GROUPS] + groups
        self.cb_fgroup.SelectedItem = current if current in groups else ALL_GROUPS
        self._filter_params()

    def _filter_params(self):
        self._commit(self.grid_params)
        parts = []
        text = search_filter(self.f_search.Text or u"")
        if text:
            parts.append(text)
        group = self.cb_fgroup.SelectedItem
        if group and group != ALL_GROUPS:
            parts.append(eq_filter(u"GrupoArq", group))
        status = STATUS_FILTERS.get(self.cb_fstatus.SelectedItem or u"", u"")
        if status:
            parts.append(status)
        if self.chk_only_marked.IsChecked:
            parts.append(u"Sel = true")
        self.t_params.DefaultView.RowFilter = u" AND ".join(u"({})".format(p) for p in parts)
        self._scroll_top(self.grid_params)
        self._update_count()

    def _on_param_changed(self, sender, args):
        if not self._bulk and args.Column.ColumnName == "Sel":
            self._update_count()

    def _update_count(self):
        total = self.t_params.Rows.Count
        if not total:
            self.tb_count.Text = u""
            return
        marked = len(self.t_params.Select(u"Sel = true"))
        off = len(self.t_params.Select(u"Usavel = false"))
        self.tb_count.Text = u"{} marcado(s) · {} de {} exibido(s) · {} indisponível(is)".format(
            marked, self.t_params.DefaultView.Count, total, off)

    # ---------------------------------------------------------- passo 2
    def _controls(self):
        """(grupo, vínculo, categorias) escolhidos nos controles do passo 2."""
        idx = self.cb_group.SelectedIndex
        group_key = self.group_keys[idx] if 0 <= idx < len(self.group_keys) else None
        kind = P.TYPE if self.rb_type.IsChecked else P.INSTANCE
        self._commit(self.grid_cats)
        cats = set(int(r["Id"]) for r in self.t_cats.Select(u"Sel = true"))
        return group_key, kind, cats

    def _selected_cfg_guids(self):
        return [to_unicode(drv["Guid"]) for drv in list(self.grid_cfg.SelectedItems)]

    def _update_apply_hint(self):
        n = len(self.grid_cfg.SelectedItems)
        self.b_apply_sel.Content = (u"Aplicar às selecionadas ({})".format(n) if n
                                    else u"Aplicar às selecionadas")
        self.tb_apply_hint.Text = (
            u"Grava o grupo, o vínculo e as categorias marcadas nas {} linha(s) selecionada(s) "
            u"da lista.".format(n) if n else
            u"Selecione linhas na lista (Shift/Ctrl + clique ou 'Selecionar todas').")

    def _apply_controls(self, sender, args):
        targets = self._selected_cfg_guids()
        if not targets:
            self.tb_cfg_msg.Text = u"Selecione ao menos uma linha da lista (Shift/Ctrl + clique ou 'Selecionar todas')."
            return
        group_key, kind, cats = self._controls()
        if group_key is None:
            self.tb_cfg_msg.Text = u"Escolha o grupo de propriedades."
            return
        if not cats:
            self.tb_cfg_msg.Text = u"Marque ao menos uma categoria na coluna Categorias."
            return
        for guid in targets:
            item = self.cfg[guid]
            item["group_key"], item["kind"], item["cats"] = group_key, kind, set(cats)
        self._refresh_cfg(select=targets)
        self.tb_cfg_msg.Text = u"Aplicado em {} parâmetro(s): {} · {} · {} categoria(s).".format(
            len(targets), self.group_labels.get(group_key, u""), P.KIND_LABELS[kind], len(cats))

    def _on_remove(self, sender, args):
        guids = self._selected_cfg_guids()
        if not guids:
            self.tb_cfg_msg.Text = u"Selecione as linhas que quer tirar do lote."
            return
        for guid in guids:
            self.cfg.pop(guid, None)
        self._set_marks(self.t_params, [r for r in self.t_params.Rows
                                        if to_unicode(r["Guid"]) in set(guids)], False)
        self._refresh_cfg()
        self.tb_cfg_msg.Text = u"{} parâmetro(s) removido(s) do lote e desmarcado(s) no passo 1.".format(
            len(guids))
        if not self.cfg:
            self._go(STEP_FILE)

    def _build_plan(self):
        items = list(self.cfg.values())
        defs = dict((g, self._def_info_by_guid(g)) for g in self.cfg)
        self.plan_rows = P.build_plan(items, defs, self.project, self.cat_names, self.group_labels,
                                      self.no_type, allow_extend=bool(self.chk_extend.IsChecked))
        return self.plan_rows

    def _refresh_cfg(self, select=None):
        """Remonta a lista do lote e a revisão a partir de self.cfg."""
        keep = set(select) if select is not None else set(self._selected_cfg_guids())
        rows = self._build_plan()
        self.t_cfg.Rows.Clear()
        for i, r in enumerate(rows):
            item = self.cfg[r.guid]
            row = self.t_cfg.NewRow()
            row["Ordem"] = i
            row["Guid"] = r.guid
            row["Nome"] = r.name
            row["Grupo"] = self.group_labels.get(item["group_key"], u"")
            row["Vinculo"] = P.KIND_LABELS.get(item["kind"], u"")
            row["Categorias"] = P.cats_summary(item["cats"], self.cat_names)
            row["Situacao"] = r.action_label + (u": " + r.notes[0] if r.notes and r.action in (
                P.BLOCKED, P.INCOMPLETE, P.NONE) else u"")
            row["Nivel"] = r.level
            self.t_cfg.Rows.Add(row)
        self.grid_cfg.SelectedItems.Clear()
        for drv in self.t_cfg.DefaultView:
            if to_unicode(drv["Guid"]) in keep:
                self.grid_cfg.SelectedItems.Add(drv)
        self._fill_review(rows)
        self._update_apply_hint()
        self._update_footer()

    # ---------------------------------------------------------- listas de categorias
    def _picker(self, prefix):
        """(tabela, grade, busca, filtro de tipo, só marcadas, contador) do seletor."""
        table = self.t_ecats if prefix else self.t_cats
        get = lambda name: getattr(self, name % prefix)
        return (table, get("grid_%scats"), get("f_%scat_search"), get("cb_%scat_type"),
                get("chk_%scat_only"), get("tb_%scat_count"))

    def _wire_cat_picker(self, prefix):
        sf = self._safe
        table, grid, search, ctype, only, _ = self._picker(prefix)
        search.TextChanged += sf(lambda s, a: self._filter_cats(prefix))
        ctype.SelectionChanged += sf(lambda s, a: self._filter_cats(prefix))
        only.Click += sf(lambda s, a: self._filter_cats(prefix))
        getattr(self, "b_%scat_all" % prefix).Click += sf(lambda s, a: self._set_marks(
            table, [drv.Row for drv in table.DefaultView], True))
        getattr(self, "b_%scat_none" % prefix).Click += sf(lambda s, a: self._set_marks(
            table, list(table.Rows), False))
        self._wire_checks(grid, table, None, lambda: self._update_cat_count(prefix))
        table.ColumnChanged += (lambda s, a: None if self._bulk or a.Column.ColumnName != "Sel"
                                else self._update_cat_count(prefix))

    def _filter_cats(self, prefix):
        table, grid, search, ctype, only, _ = self._picker(prefix)
        self._commit(grid)
        parts = []
        text = search_filter(search.Text or u"")
        if text:
            parts.append(text)
        kind = ctype.SelectedItem
        if kind and kind != ALL_CAT_TYPES:
            parts.append(eq_filter(u"Tipo", kind))
        if only.IsChecked:
            parts.append(u"Sel = true")
        table.DefaultView.RowFilter = u" AND ".join(u"({})".format(p) for p in parts)
        self._scroll_top(grid)
        self._update_cat_count(prefix)

    def _update_cat_count(self, prefix):
        table, _, _, _, _, label = self._picker(prefix)
        label.Text = u"{} marcada(s) · {} de {} na lista".format(
            len(table.Select(u"Sel = true")), table.DefaultView.Count, table.Rows.Count)

    def _set_cat_selection(self, table, ids):
        self._bulk = True
        try:
            for row in table.Rows:
                row["Sel"] = int(row["Id"]) in ids
        finally:
            self._bulk = False
        self._after_marks(table)

    # ---------------------------------------------------------- passo 3 e execução
    def _fill_review(self, rows):
        self.t_review.Rows.Clear()
        for i, r in enumerate(rows):
            row = self.t_review.NewRow()
            row["Ordem"] = i
            row["Nome"] = r.name
            row["Acao"] = r.action_label
            row["Grupo"] = r.group_label
            row["Vinculo"] = r.kind_label
            row["Categorias"] = r.cats_text or u"(nenhuma)"
            row["Obs"] = u" ".join(r.notes)
            row["Nivel"] = r.level
            self.t_review.Rows.Add(row)
        count = P.summary(rows)
        todo = count[P.CREATE] + count[P.EXTEND]
        if not rows:
            text, style = u"O lote está vazio: marque parâmetros no passo 1.", "oca.Msg.Info"
        elif count[P.INCOMPLETE]:
            text = (u"{} parâmetro(s) incompleto(s) (sem categoria ou sem grupo): volte ao passo 2."
                    .format(count[P.INCOMPLETE]))
            style = "oca.Msg.Err"
        elif not todo:
            text, style = u"Nada a aplicar: veja as observações.", "oca.Msg.Warn"
        else:
            text = u"Será feito: {} vínculo(s) novo(s) · {} parâmetro(s) com categorias adicionadas.".format(
                count[P.CREATE], count[P.EXTEND])
            extra = []
            if count[P.NONE]:
                extra.append(u"{} sem alteração".format(count[P.NONE]))
            if count[P.BLOCKED]:
                extra.append(u"{} não será(ão) aplicado(s)".format(count[P.BLOCKED]))
            if extra:
                text += u" " + u" · ".join(extra) + u" (veja Observações)."
            style = "oca.Msg.Warn" if count[P.BLOCKED] else "oca.Msg.Ok"
        self.tb_review.Text = text
        try:
            self.bd_review.Style = self.FindResource(style)
        except Exception:
            pass

    def _ready(self):
        count = P.summary(self.plan_rows)
        return (count[P.CREATE] + count[P.EXTEND]) > 0 and not count[P.INCOMPLETE]

    def _on_apply(self, sender, args):
        # estado do projeto lido de novo: a revisão confirmada é a que será executada
        self.project = R.scan_project(self.doc, self.groups)
        self._refresh_cfg()
        if not self._ready():
            forms.alert(u"Nada a aplicar.", sub_msg=self.tb_review.Text, title=ALERT_TITLE)
            return
        count = P.summary(self.plan_rows)
        n = count[P.CREATE] + count[P.EXTEND]
        details = [u"{} vínculo(s) novo(s) · {} com categorias adicionadas.".format(
            count[P.CREATE], count[P.EXTEND])]
        if count[P.BLOCKED] or count[P.NONE]:
            details.append(u"{} parâmetro(s) ficam como estão (veja a revisão).".format(
                count[P.BLOCKED] + count[P.NONE]))
        details.append(u"Uma única transação: se algo falhar, nada é gravado. Ctrl+Z desfaz depois.")
        details.append(u"Nenhum parâmetro é excluído e o arquivo TXT não é alterado.")
        if not forms.alert(u"Aplicar {} parâmetro(s) ao projeto?".format(n),
                           sub_msg=u"\n".join(details), title=ALERT_TITLE, yes=True, no=True):
            return
        try:
            R.apply_plan(self.doc, self.path, self.plan_rows, self.cat_by_id, self.groups)
        except R.ApplyError as ex:
            self.project = R.scan_project(self.doc, self.groups)
            self._refresh_cfg()
            forms.alert(u"A vinculação falhou e foi desfeita. Nada foi gravado.",
                        sub_msg=to_unicode(ex), title=ALERT_TITLE)
            return
        self.applied = True
        self.Close()

    # ---------------------------------------------------------- parâmetros no projeto
    def _txt_status(self, item):
        """(texto, chave, nível, pode exportar) do parâmetro em relação ao TXT carregado."""
        if self.spf is None:
            return u"Sem TXT carregado", u"none", u"off", False
        if self.revit_error:
            return u"O Revit não abriu o TXT", u"none", u"off", False
        p = self.spf_by_guid.get(item["guid"])
        if p is not None:
            if p.name.strip() == item["name"].strip():
                return u"No TXT · " + p.group_name, u"in", u"ok", False
            return u"No TXT com outro nome: '{}'".format(p.name.strip()), u"rename", u"warn", False
        if self.txt_names.get(item["name"].strip()):
            return u"Conflito: nome no TXT com outro GUID", u"conflict", u"err", False
        return u"Fora do TXT", u"out", u"", True

    def _fill_proj(self):
        """Remonta a tabela dos parâmetros compartilhados do projeto (marcas mantidas)."""
        self._commit(self.grid_proj)
        marked = set(to_unicode(r["Guid"]) for r in self.t_proj.Select(u"Sel = true"))
        selected = self.edit_guid
        self.txt_names = {}
        if self.spf is not None:
            for p in self.spf.params:
                if p.guid_key:
                    self.txt_names.setdefault(p.name.strip(), set()).add(p.guid_key)
        items = R.list_project_shared(self.doc, self.project, self.cat_names)
        items.sort(key=lambda i: natural_key(i["name"]))
        self.proj_items = dict((i["guid"], i) for i in items)
        self._bulk = True
        try:
            self.t_proj.Rows.Clear()
            for i, item in enumerate(items):
                text, key, level, can = self._txt_status(item)
                kind = P.KIND_LABELS.get(item["kind"], u"—") if item["bound"] else u"—"
                cats = (P.cats_summary(item["cats"], self.cat_names) if item["bound"] else u"—")
                origin = u"Parâmetro de projeto" if item["bound"] else u"Só em famílias"
                row = self.t_proj.NewRow()
                row["Ordem"] = i
                row["Sel"] = can and item["guid"] in marked
                row["Pode"] = can
                row["Guid"] = item["guid"]
                row["Nome"] = item["name"]
                row["Tipo"] = item["datatype"]
                row["Vinculo"] = kind
                row["Grupo"] = item["group_label"] or u"—"
                row["Categorias"] = cats
                row["Origem"] = origin
                row["OrigemKey"] = u"bound" if item["bound"] else u"family"
                row["NoTxt"] = text
                row["TxtKey"] = key
                row["Nivel"] = level
                row["Busca"] = fold(u" ".join([item["name"], item["guid"], item["datatype"],
                                              item["group_label"], origin, text]
                                             + item["cats_names"]))
                self.t_proj.Rows.Add(row)
        finally:
            self._bulk = False
        self.rb_mode_proj.Content = u"Parâmetros no projeto ({})".format(len(items))
        self._filter_proj()
        # volta a mostrar o parâmetro que estava em edição (se ainda existir)
        if selected in self.proj_items:
            for drv in self.t_proj.DefaultView:
                if to_unicode(drv["Guid"]) == selected:
                    self.grid_proj.SelectedItem = drv
                    break
        self._show_edit(selected if selected in self.proj_items else None)
        self._update_footer()

    def _filter_proj(self):
        self._commit(self.grid_proj)
        parts = []
        text = search_filter(self.f_psearch.Text or u"")
        if text:
            parts.append(text)
        origin = PROJ_ORIGINS.get(self.cb_porigin.SelectedItem or u"", u"")
        if origin:
            parts.append(origin)
        txt = PROJ_TXT.get(self.cb_ptxt.SelectedItem or u"", u"")
        if txt:
            parts.append(txt)
        if self.chk_p_only.IsChecked:
            parts.append(u"Sel = true")
        self.t_proj.DefaultView.RowFilter = u" AND ".join(u"({})".format(p) for p in parts)
        self._scroll_top(self.grid_proj)
        self._update_proj_count()

    def _on_proj_changed(self, sender, args):
        if not self._bulk and args.Column.ColumnName == "Sel":
            self._update_proj_count()

    def _update_proj_count(self):
        total = self.t_proj.Rows.Count
        if not total:
            self.tb_pcount.Text = u"Nenhum parâmetro compartilhado no projeto."
            return
        self.tb_pcount.Text = u"{} marcado(s) para exportar · {} de {} exibido(s) · {} fora do TXT".format(
            len(self.t_proj.Select(u"Sel = true")), self.t_proj.DefaultView.Count, total,
            len(self.t_proj.Select(u"Pode = true")))

    def _update_export_info(self):
        if self.path:
            self.tb_export_info.Text = (
                u"Exportar para o TXT grava os marcados (✓) em:  {}  →  grupo '{}'. "
                u"Uma cópia do TXT é salva antes, na mesma pasta.".format(self.path, R.EXPORT_GROUP))
        else:
            self.tb_export_info.Text = u"Carregue um arquivo TXT (barra acima) para poder exportar."

    def _on_proj_select(self, sender, args):
        if not System.Object.ReferenceEquals(args.OriginalSource, self.grid_proj):
            return
        selected = list(self.grid_proj.SelectedItems)
        self._show_edit(to_unicode(selected[0]["Guid"]) if len(selected) == 1 else None,
                        many=len(selected) > 1)

    def _show_edit(self, guid, many=False):
        """Mostra no painel à direita o parâmetro selecionado, pronto para editar."""
        item = self.proj_items.get(guid) if guid else None
        self.edit_guid = guid if item else None
        self.tb_edit_msg.Text = u""
        if item is None:
            self.tb_d_name.Text = (u"Várias linhas selecionadas: clique em uma só para editar."
                                   if many else u"Clique numa linha da tabela.")
            self.tb_d_guid.Text = self.tb_d_info.Text = u""
            self.pn_edit.IsEnabled = self.b_update.IsEnabled = False
            self.cb_egroup.SelectedIndex = -1
            self._set_cat_selection(self.t_ecats, set())
            return
        text = self._txt_status(item)[0]
        self.tb_d_name.Text = item["name"]
        self.tb_d_guid.Text = item["guid"]
        if item["bound"]:
            self.tb_d_info.Text = u"Tipo de dado: {} · Vínculo: {} · {} categoria(s) · No TXT: {}".format(
                item["datatype"] or u"?", P.KIND_LABELS.get(item["kind"], u"?"), len(item["cats"]), text)
            key = self.project["bound"][guid].get("group_key")
            self.cb_egroup.SelectedIndex = self.group_keys.index(key) if key in self.group_keys else -1
            self._set_cat_selection(self.t_ecats, item["cats"])
            self.pn_edit.IsEnabled = self.b_update.IsEnabled = True
        else:
            self.tb_d_info.Text = u"Tipo de dado: {} · No TXT: {}".format(item["datatype"] or u"?", text)
            self.cb_egroup.SelectedIndex = -1
            self._set_cat_selection(self.t_ecats, set())
            self.pn_edit.IsEnabled = self.b_update.IsEnabled = False
            self.tb_edit_msg.Text = (u"Existe só em famílias carregadas: não tem grupo nem categorias "
                                     u"de projeto. Para vinculá-lo, use a aba 'Vincular do TXT ao "
                                     u"projeto'.")

    def _on_update(self, sender, args):
        guid = self.edit_guid
        item = self.proj_items.get(guid) if guid else None
        if item is None or not item["bound"]:
            self.tb_edit_msg.Text = u"Clique num parâmetro de projeto na tabela."
            return
        self._commit(self.grid_ecats)
        bound = self.project["bound"][guid]
        idx = self.cb_egroup.SelectedIndex
        group_key = self.group_keys[idx] if 0 <= idx < len(self.group_keys) else bound.get("group_key")
        new_ids = set(int(r["Id"]) for r in self.t_ecats.Select(u"Sel = true"))
        if not new_ids:
            self.tb_edit_msg.Text = (u"Marque ao menos uma categoria. Para tirar o parâmetro do "
                                     u"projeto, use Gerenciar > Parâmetros de projeto.")
            return
        old_ids = set(bound["cats"])
        added, removed = new_ids - old_ids, old_ids - new_ids
        group_changed = group_key != bound.get("group_key")
        if not (added or removed or group_changed):
            self.tb_edit_msg.Text = u"Nada mudou: grupo e categorias são os atuais."
            return
        names = lambda ids: u", ".join(sorted((self.cat_names.get(i, unicode(i)) for i in ids),
                                              key=lambda n: n.lower()))
        lost = 0
        if removed:
            lost = R.count_values(self.doc, guid, bound["kind"],
                                  [self.cat_by_id[i] for i in removed if i in self.cat_by_id])
        details = []
        if group_changed:
            details.append(u"Grupo de propriedades: {} → {}".format(
                bound.get("group_label") or u"?", self.group_labels.get(group_key, u"?")))
        if added:
            details.append(u"Categorias adicionadas ({}): {}".format(len(added), names(added)))
        if removed:
            details.append(u"Categorias retiradas ({}): {}".format(len(removed), names(removed)))
            details.append(u"ATENÇÃO: {} valor(es) preenchido(s) nessas categorias serão apagados."
                           .format(lost) if lost else u"Nenhum valor preenchido nessas categorias.")
        details.append(u"Instância/Tipo não muda. Uma transação: se algo falhar, nada é gravado. "
                       u"Ctrl+Z desfaz depois.")
        question = (u"Retirar categorias e apagar {} valor(es) de '{}'?".format(lost, item["name"])
                    if lost else u"Atualizar '{}' no modelo?".format(item["name"]))
        if not forms.alert(question, sub_msg=u"\n".join(details), title=ALERT_TITLE, yes=True, no=True):
            return
        try:
            R.update_binding(self.doc, guid, new_ids, group_key, self.cat_by_id)
        except R.ApplyError as ex:
            forms.alert(u"A atualização falhou e foi desfeita. Nada foi gravado.",
                        sub_msg=to_unicode(ex), title=ALERT_TITLE)
            return
        # projeto relido: as duas abas mostram o estado novo
        self.project = R.scan_project(self.doc, self.groups)
        self._fill_proj()
        if self.spf is not None:
            self._fill_params()
        self._refresh_cfg()
        self.tb_edit_msg.Text = u"'{}' atualizado no modelo.".format(item["name"])

    def _on_export(self, sender, args):
        self._commit(self.grid_proj)
        self.tb_proj_msg.Text = u""
        if self.spf is None or not self.path:
            self.tb_proj_msg.Text = u"Carregue um arquivo TXT antes de exportar ('Carregar arquivo…')."
            return
        if self.revit_error:
            self.tb_proj_msg.Text = u"O Revit não abriu o TXT carregado; não é possível exportar para ele."
            return
        guids = [to_unicode(r["Guid"]) for r in self.t_proj.Select(u"Sel = true", u"Ordem ASC")
                 if r["Pode"]]
        if not guids:
            self.tb_proj_msg.Text = u"Marque (✓) ao menos um parâmetro que esteja fora do TXT."
            return
        details = [u"Arquivo: {}".format(self.path),
                   u"Grupo do TXT: {}".format(R.EXPORT_GROUP),
                   u"GUID, nome e tipo de dado são os do projeto. Nada muda no projeto.",
                   u"Antes de gravar, uma cópia do TXT é salva na mesma pasta.",
                   u"Descrição fica vazia e 'editável pelo usuário' fica sim: a API não informa "
                   u"esses dados no projeto."]
        if not forms.alert(u"Exportar {} parâmetro(s) para o TXT?".format(len(guids)),
                           sub_msg=u"\n".join(details), title=ALERT_TITLE, yes=True, no=True):
            return
        try:
            written, backup = R.export_definitions(self.doc, self.path, guids, R.EXPORT_GROUP)
        except R.ExportError as ex:
            self._load(self.path)
            info = [to_unicode(ex)]
            if ex.written:
                info.insert(0, u"Gravados antes do erro: {}.".format(u", ".join(ex.written)))
            if ex.backup:
                info.append(u"Cópia do TXT anterior: {}".format(ex.backup))
            forms.alert(u"A exportação parou no meio." if ex.written
                        else u"A exportação falhou. Nenhum parâmetro foi gravado.",
                        sub_msg=u"\n".join(info), title=ALERT_TITLE)
            return
        self._load(self.path)
        forms.alert(u"{} parâmetro(s) exportado(s) para o TXT.".format(len(written)),
                    sub_msg=u"Grupo '{}'. Cópia do arquivo anterior: {}".format(R.EXPORT_GROUP, backup),
                    title=ALERT_TITLE)


try:
    unicode
except NameError:           # IronPython 3 / CPython (testes)
    unicode = str


def main():
    if doc is None:
        forms.alert(u"Abra um modelo do Revit antes de executar o comando.", title=ALERT_TITLE)
        return
    if doc.IsFamilyDocument:
        forms.alert(u"Este comando funciona apenas em modelos de projeto.",
                    sub_msg=u"Em famílias, parâmetros compartilhados são adicionados pelo "
                            u"Gerenciador de famílias.", title=ALERT_TITLE)
        return
    if doc.IsReadOnly:
        forms.alert(u"O modelo está aberto como somente leitura.", title=ALERT_TITLE)
        return
    try:
        win = ParamManagerWindow(doc)
    except Exception as ex:
        forms.alert(u"Não foi possível abrir a janela.", sub_msg=to_unicode(ex), title=ALERT_TITLE)
        return
    win.ShowDialog()
    if win.applied:
        forms.alert(u"Parâmetros aplicados ao projeto.", sub_msg=u"Ctrl+Z desfaz.", title=ALERT_TITLE)


main()
