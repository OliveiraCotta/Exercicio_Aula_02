# -*- coding: utf-8 -*-
# GERENCIADOR DE PARÂMETROS COMPARTILHADOS  (pyRevit - Revit 2024/2025/2026;
# grupos de parâmetro também tratados para 2021-2023, ver lib/oca_params/revit_ops.py)
#
# Vincula ao projeto, em lote, parâmetros de um arquivo TXT de parâmetros
# compartilhados, sem passar pela janela Parâmetros de projeto do Revit.
# (Comentário, não docstring: o tooltip do botão vem SÓ de __doc__ abaixo.)
#
# Janela com duas abas principais (no estilo das abas do Explorador de
# Arquivos) e a barra do arquivo TXT compartilhada entre elas:
#
# A. VINCULAR DO TXT AO PROJETO (três etapas):
#   1. Parâmetros do arquivo - lê o TXT escolhido (só leitura). Tabela com
#      nome, grupo de definição (do TXT), tipo de dado, GUID e situação no
#      projeto; busca sem acento, filtros por grupo / tipo de dado / situação,
#      "só marcados", Selecionar todos, Limpar seleção, Espaço marca as linhas
#      selecionadas. Definição que não pode ser usada fica cinza, com o motivo.
#      "Configurar marcados" leva os marcados para a configuração (lote).
#   2. Configuração - o lote, uma linha por parâmetro. À direita, os controles
#      em lote: grupo de propriedades (paleta Propriedades; não é o grupo do
#      TXT), Instância ou Tipo e categorias (busca, filtro por tipo, seleção
#      múltipla). "Aplicar às selecionadas" (configuração individual) ou
#      "Aplicar a todas". "+ Adicionar parâmetros" volta à aba 1 para incluir
#      mais parâmetros no lote.
#   3. Revisão - o que será feito com cada parâmetro (criar vínculo, adicionar
#      categorias, sem alteração, não será aplicado, incompleto) e por quê.
#      "Aplicar no projeto" pede confirmação e grava tudo em UMA transação.
#
# B. PARÂMETROS NO PROJETO:
#   Todos os parâmetros compartilhados do projeto (SharedParameterElement):
#   nome, tipo de dado, Instância/Tipo, grupo de propriedades, categorias,
#   origem (parâmetro de projeto ou só em famílias), GUID e situação no TXT
#   carregado (no TXT, no TXT com outro nome, conflito de nome, fora do TXT).
#   Busca, filtros, detalhe com todas as categorias da linha selecionada.
#   "Exportar para o TXT" grava no TXT carregado os marcados que estão fora
#   dele, com o GUID do projeto, num grupo existente ou novo do TXT.
#
# Regras (detalhes em lib/oca_params/plan.py e revit_ops.py):
#   * O TXT só é gravado pela exportação (aba B), com confirmação e cópia de
#     segurança antes; o GUID do projeto é mantido. Vincular não grava no TXT.
#   * As definições vêm do próprio Revit (OpenSharedParameterFile), pelo GUID.
#     O caminho do arquivo de parâmetros do usuário é trocado só durante a
#     leitura e a gravação e restaurado logo depois.
#   * Nada é excluído. Parâmetro já vinculado: só recebe categorias novas
#     (mesmo tipo de vínculo, grupo atual mantido, valores conferidos antes de
#     confirmar). Instância <-> Tipo de parâmetro existente: não é feito.
#   * Qualquer falha desfaz a transação inteira (RollBack).
#   * Sem relatório ao final (pedido do escritório): só a confirmação.

__title__ = u"Gerenciador\nde Parâmetros\nCompartilhados"
__doc__ = (u"Vincula ao projeto, em lote, parâmetros de um arquivo TXT de parâmetros "
           u"compartilhados (GUIDs originais) e mostra os compartilhados do projeto, com "
           u"exportação para o TXT.\n\n"
           u"Vincular: escolha os parâmetros, o grupo na paleta Propriedades, Instância ou Tipo "
           u"e as categorias; revise e aplique.\n"
           u"Parâmetros no projeto: veja grupo, vínculo e categorias e exporte para o TXT "
           u"(cópia de segurança antes). Nada é excluído do projeto.")

import os
import re
from collections import OrderedDict

import clr
clr.AddReference("System.Data")
import System
from System.Data import DataTable
from System.Windows.Controls import DataGridEditingUnit
from System.Windows.Input import Key

from pyrevit import revit, script, forms
from oca_ds import build_xaml, alert_title
from oca_params import spfile
from oca_params import plan as P
from oca_params import revit_ops as R

ALERT_TITLE = alert_title(__title__)   # "OCA · Gerenciador de Parâmetros Compartilhados"

doc = revit.doc
config = script.get_config()

to_unicode = R.to_unicode

ALL_GROUPS = u"Todos os grupos do arquivo"
ALL_TYPES = u"Todos os tipos de dado"
STATUS_FILTERS = OrderedDict([
    (u"Todas as situações", u""),
    (u"Disponíveis", u"Usavel = true"),
    (u"Novos", u"SitKey = 'new'"),
    (u"Já no projeto", u"SitKey = 'bound'"),
    (u"Só em famílias", u"SitKey = 'family'"),
    (u"Indisponíveis", u"Usavel = false"),
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
NEW_GROUP = u"Novo grupo…"
DEFAULT_NEW_GROUP = u"Parâmetros exportados"
TAB_FILE, TAB_CFG, TAB_REVIEW = 0, 1, 2


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
    """DataTable com colunas (nome, tipo .NET). Binding direto no DataGrid."""
    table = DataTable(name)
    for col, clr_type in columns:
        table.Columns.Add(col, clr.GetClrType(clr_type))
    return table


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
                     ToolTip="Escolher parâmetros do TXT, configurar e vincular ao projeto"/>
        <RadioButton x:Name="rb_mode_proj" GroupName="mode" Style="{StaticResource browserTab}"
                     Content="Parâmetros no projeto"
                     ToolTip="Ver os parâmetros compartilhados do projeto (grupo, vínculo, categorias) e exportar para o TXT"/>
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
          <!-- ============ VINCULAR DO TXT ============ -->
          <Grid x:Name="pn_link">
          <TabControl x:Name="tabs">

            <!-- ============ 1. PARÂMETROS DO ARQUIVO ============ -->
            <TabItem Header="1 · Parâmetros do arquivo">
              <Grid>
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
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <TextBlock Text="Buscar" Style="{StaticResource oca.Label}"/>
                  <TextBox x:Name="f_search" Grid.Column="1"
                           ToolTip="Nome, GUID, grupo, tipo de dado ou descrição. Sem diferença de acento ou maiúscula."/>
                  <ComboBox x:Name="cb_fgroup" Grid.Column="2" Width="200" Margin="8,0,0,0"
                            ToolTip="Filtrar pelo grupo de definição do arquivo TXT"/>
                  <ComboBox x:Name="cb_ftype" Grid.Column="3" Width="170" Margin="8,0,0,0"
                            ToolTip="Filtrar pelo tipo de dado"/>
                  <ComboBox x:Name="cb_fstatus" Grid.Column="4" Width="160" Margin="8,0,0,0"
                            ToolTip="Filtrar pela situação do parâmetro no projeto aberto"/>
                  <CheckBox x:Name="chk_only_marked" Grid.Column="5" Content="Só marcados" Margin="12,0,0,0"
                            VerticalAlignment="Center"/>
                </Grid>

                <DataGrid x:Name="grid_params" Grid.Row="1" """ + _GRID + u"""
                          SelectionMode="Extended" RowHeight="26" RowStyle="{StaticResource paramRow}">
                  <DataGrid.Columns>
                    <DataGridTemplateColumn Header="✓" Width="34" CanUserResize="False" SortMemberPath="Sel">
                      <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                          <CheckBox Style="{StaticResource tight}" IsEnabled="{Binding Usavel}"
                                    IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}"/>
                        </DataTemplate>
                      </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Nome" Binding="{Binding Nome}" Width="2*" MinWidth="140"
                                        IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="Grupo de definição (TXT)" Binding="{Binding GrupoArq}" Width="1.3*"
                                        MinWidth="110" IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="Tipo de dado" Binding="{Binding Tipo}" Width="*" MinWidth="90"
                                        IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="GUID" Binding="{Binding GuidTxt}" Width="275" IsReadOnly="True"
                                        ElementStyle="{StaticResource mono}"/>
                    <DataGridTextColumn Header="Situação no projeto" Binding="{Binding Situacao}" Width="1.3*"
                                        MinWidth="110" IsReadOnly="True" ElementStyle="{StaticResource status}"/>
                    <DataGridTextColumn Header="Lote" Binding="{Binding Lote}" Width="44" IsReadOnly="True"
                                        ElementStyle="{StaticResource cell}"/>
                    <DataGridTextColumn Header="Observações" Binding="{Binding Obs}" Width="2*" MinWidth="120"
                                        IsReadOnly="True" ElementStyle="{StaticResource cell}"/>
                  </DataGrid.Columns>
                </DataGrid>

                <StackPanel Grid.Row="2">
                  <Border Style="{StaticResource oca.Msg.Err}" Visibility="{Binding Visibility, ElementName=tb_params_msg}">
                    <TextBlock x:Name="tb_params_msg" Style="{StaticResource oca.MsgText}"/>
                  </Border>
                  <DockPanel Margin="0,10,0,4" LastChildFill="False">
                    <Button x:Name="b_sel_all" Content="Selecionar todos" DockPanel.Dock="Left"
                            ToolTip="Marca todos os parâmetros disponíveis que estão visíveis com a busca e os filtros atuais"/>
                    <Button x:Name="b_sel_none" Content="Limpar seleção" DockPanel.Dock="Left" Margin="8,0,0,0"
                            ToolTip="Desmarca todos os parâmetros (inclusive os escondidos pelos filtros)"/>
                    <TextBlock x:Name="tb_count" DockPanel.Dock="Left" Margin="14,0,0,0" VerticalAlignment="Center"
                               Foreground="{StaticResource oca.Ink2}"/>
                    <Button x:Name="b_configure" Content="Configurar marcados  →" DockPanel.Dock="Right" MinWidth="170"
                            ToolTip="Leva os parâmetros marcados para a configuração (aba 2) com as opções atuais de grupo, vínculo e categorias"/>
                  </DockPanel>
                  <TextBlock Style="{StaticResource oca.Hint}"
                             Text="Marque na coluna ✓ ou selecione linhas (Ctrl/Shift + clique) e tecle Espaço para marcar ou desmarcar. Linhas cinza não podem ser usadas: o motivo está em Observações."/>
                </StackPanel>
              </Grid>
            </TabItem>

            <!-- ============ 2. CONFIGURAÇÃO ============ -->
            <TabItem Header="2 · Configuração">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*" MinWidth="360"/>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="430"/>
                </Grid.ColumnDefinitions>

                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>
                  <WrapPanel Margin="0,0,0,8">
                    <Button x:Name="b_add_more" Content="+ Adicionar parâmetros" Margin="0,0,8,0"
                            ToolTip="Volta à aba 1 para incluir mais parâmetros nesta configuração"/>
                    <Button x:Name="b_cfg_all" Content="Selecionar todas" Margin="0,0,8,0"
                            ToolTip="Seleciona todas as linhas da configuração"/>
                    <Button x:Name="b_copy_from" Content="Copiar da linha" Margin="0,0,8,0"
                            ToolTip="Mostra nos controles à direita o grupo, o vínculo e as categorias da linha selecionada"/>
                    <Button x:Name="b_remove" Content="Remover da configuração"
                            ToolTip="Tira as linhas selecionadas do lote (não altera o projeto nem o arquivo)"/>
                  </WrapPanel>
                  <DataGrid x:Name="grid_cfg" Grid.Row="1" """ + _GRID + u"""
                            SelectionMode="Extended" RowHeight="26" IsReadOnly="True">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="Parâmetro" Binding="{Binding Nome}" Width="1.4*" MinWidth="120"
                                          ElementStyle="{StaticResource cell}"/>
                      <DataGridTextColumn Header="Grupo de propriedades" Binding="{Binding Grupo}" Width="*" MinWidth="100"
                                          ElementStyle="{StaticResource cell}"/>
                      <DataGridTextColumn Header="Vínculo" Binding="{Binding Vinculo}" Width="72"
                                          ElementStyle="{StaticResource cell}"/>
                      <DataGridTextColumn Header="Categorias" Binding="{Binding Categorias}" Width="1.6*" MinWidth="120"
                                          ElementStyle="{StaticResource cell}"/>
                      <DataGridTextColumn Header="Situação" Binding="{Binding Situacao}" Width="1.4*" MinWidth="110"
                                          ElementStyle="{StaticResource status}"/>
                    </DataGrid.Columns>
                  </DataGrid>
                  <StackPanel Grid.Row="2">
                    <Border Style="{StaticResource oca.Msg.Info}" Visibility="{Binding Visibility, ElementName=tb_cfg_msg}">
                      <TextBlock x:Name="tb_cfg_msg" Style="{StaticResource oca.MsgText}"/>
                    </Border>
                    <DockPanel Margin="0,10,0,4">
                      <Button x:Name="b_review" Content="Revisar  →" DockPanel.Dock="Right" MinWidth="120" Margin="8,0,0,0"
                              ToolTip="Mostra o que será feito com cada parâmetro antes de aplicar"/>
                      <TextBlock Style="{StaticResource oca.Hint}" VerticalAlignment="Center"
                                 Text="Configuração individual: selecione linhas (Ctrl/Shift + clique), ajuste os controles à direita e clique em 'Aplicar às selecionadas'."/>
                    </DockPanel>
                  </StackPanel>
                </Grid>

                <Grid Grid.Column="2">
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*" MinHeight="180"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>

                  <HeaderedContentControl Header="GRUPO DE PROPRIEDADES (PALETA PROPRIEDADES)" Style="{StaticResource oca.Section}">
                    <StackPanel>
                      <ComboBox x:Name="cb_group"
                                ToolTip="Seção da paleta Propriedades em que o parâmetro aparece (lista da versão do Revit em uso)"/>
                      <TextBlock Style="{StaticResource oca.Hint}" Margin="0,6,0,0"
                                 Text="É onde o parâmetro aparece no Revit. Não é o grupo de definição do TXT, que só organiza o arquivo."/>
                    </StackPanel>
                  </HeaderedContentControl>

                  <HeaderedContentControl Grid.Row="1" Header="TIPO DE VÍNCULO" Style="{StaticResource oca.Section}">
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

                  <HeaderedContentControl Grid.Row="2" Header="CATEGORIAS" Style="{StaticResource oca.Section}">
                    <DockPanel>
                      <Grid DockPanel.Dock="Top" Margin="0,0,0,6">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="f_cat_search" ToolTip="Buscar categoria pelo nome (sem diferença de acento)"/>
                        <ComboBox x:Name="cb_cat_type" Grid.Column="1" Width="130" Margin="8,0,0,0"
                                  ToolTip="Modelo, Anotação ou Analítico"/>
                      </Grid>
                      <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6" LastChildFill="False">
                        <Button x:Name="b_cat_all" Content="Marcar visíveis" Style="{StaticResource oca.Compact}"
                                ToolTip="Marca as categorias que estão na lista com a busca e o filtro atuais"/>
                        <Button x:Name="b_cat_none" Content="Desmarcar todas" Style="{StaticResource oca.Compact}" Margin="6,0,0,0"/>
                        <CheckBox x:Name="chk_cat_only" Content="Só marcadas" DockPanel.Dock="Right"
                                  Style="{StaticResource inline}" Margin="0"/>
                      </DockPanel>
                      <TextBlock x:Name="tb_cat_count" DockPanel.Dock="Bottom" Style="{StaticResource oca.Hint}" Margin="0,6,0,0"/>
                      <DataGrid x:Name="grid_cats" """ + _GRID + u""" SelectionMode="Extended" RowHeight="24">
                        <DataGrid.Columns>
                          <DataGridTemplateColumn Header="✓" Width="34" CanUserResize="False" SortMemberPath="Sel">
                            <DataGridTemplateColumn.CellTemplate>
                              <DataTemplate>
                                <CheckBox Style="{StaticResource tight}"
                                          IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}"/>
                              </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                          </DataGridTemplateColumn>
                          <DataGridTextColumn Header="Categoria" Binding="{Binding Nome}" Width="*" IsReadOnly="True"
                                              ElementStyle="{StaticResource cell}"/>
                          <DataGridTextColumn Header="Tipo" Binding="{Binding Tipo}" Width="86" IsReadOnly="True"
                                              ElementStyle="{StaticResource cell}"/>
                        </DataGrid.Columns>
                      </DataGrid>
                    </DockPanel>
                  </HeaderedContentControl>

                  <HeaderedContentControl Grid.Row="3" Header="APLICAR NA CONFIGURAÇÃO" Style="{StaticResource oca.Section}" Margin="0">
                    <StackPanel>
                      <WrapPanel Margin="0,0,0,8">
                        <CheckBox x:Name="chk_set_group" Content="Grupo" IsChecked="True" Style="{StaticResource inline}"/>
                        <CheckBox x:Name="chk_set_kind" Content="Vínculo" IsChecked="True" Style="{StaticResource inline}"/>
                        <CheckBox x:Name="chk_set_cats" Content="Categorias" IsChecked="True" Style="{StaticResource inline}"/>
                      </WrapPanel>
                      <Grid>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="8"/>
                          <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name="b_apply_sel" Content="Aplicar às selecionadas"
                                ToolTip="Usa as opções marcadas acima só nas linhas selecionadas da configuração"/>
                        <Button x:Name="b_apply_all" Grid.Column="2" Content="Aplicar a todas"
                                ToolTip="Usa as opções marcadas acima em todas as linhas da configuração"/>
                      </Grid>
                    </StackPanel>
                  </HeaderedContentControl>
                </Grid>
              </Grid>
            </TabItem>

            <!-- ============ 3. REVISÃO ============ -->
            <TabItem Header="3 · Revisão">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="Auto"/>
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
                <DataGrid x:Name="grid_review" Grid.Row="2" """ + _GRID + u""" IsReadOnly="True"
                          SelectionMode="Extended">
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
                <TextBlock Grid.Row="3" Style="{StaticResource oca.Hint}" Margin="0,8,0,4"
                           Text="Nada é gravado até clicar em 'Aplicar no projeto'. Tudo vai numa única transação: se algo falhar, nada é gravado. Nenhum parâmetro é excluído e o arquivo TXT não é alterado."/>
              </Grid>
            </TabItem>
          </TabControl>
          </Grid>

          <!-- ============ PARÂMETROS NO PROJETO ============ -->
          <Grid x:Name="pn_proj" Visibility="Collapsed">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*" MinWidth="420"/>
              <ColumnDefinition Width="14"/>
              <ColumnDefinition Width="360"/>
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
                <ComboBox x:Name="cb_porigin" Grid.Column="2" Width="190" Margin="8,0,0,0"
                          ToolTip="Parâmetros de projeto (vinculados) ou que existem só por famílias carregadas"/>
                <ComboBox x:Name="cb_ptxt" Grid.Column="3" Width="170" Margin="8,0,0,0"
                          ToolTip="Comparação com o TXT carregado (pelo GUID e pelo nome)"/>
                <CheckBox x:Name="chk_p_only" Grid.Column="4" Content="Só marcados" Margin="12,0,0,0"
                          VerticalAlignment="Center"/>
              </Grid>
              <DataGrid x:Name="grid_proj" Grid.Row="1" """ + _GRID + u"""
                        SelectionMode="Extended" RowHeight="26">
                <DataGrid.Columns>
                  <DataGridTemplateColumn Header="✓" Width="34" CanUserResize="False" SortMemberPath="Sel">
                    <DataGridTemplateColumn.CellTemplate>
                      <DataTemplate>
                        <CheckBox Style="{StaticResource tight}" IsEnabled="{Binding Pode}"
                                  IsChecked="{Binding Sel, UpdateSourceTrigger=PropertyChanged}"/>
                      </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                  </DataGridTemplateColumn>
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
                          ToolTip="Marca os parâmetros visíveis que ainda não estão no TXT"/>
                  <Button x:Name="b_psel_none" Content="Limpar seleção" DockPanel.Dock="Left" Margin="8,0,0,0"
                          ToolTip="Desmarca todos (inclusive os escondidos pelos filtros)"/>
                  <TextBlock x:Name="tb_pcount" DockPanel.Dock="Left" Margin="14,0,0,0" VerticalAlignment="Center"
                             Foreground="{StaticResource oca.Ink2}"/>
                </DockPanel>
                <TextBlock Style="{StaticResource oca.Hint}"
                           Text="Só parâmetros que ainda não estão no TXT podem ser marcados para exportar. Clique numa linha para ver todas as categorias. Espaço marca as linhas selecionadas."/>
              </StackPanel>
            </Grid>

            <Grid Grid.Column="2">
              <Grid.RowDefinitions>
                <RowDefinition Height="*" MinHeight="200"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <HeaderedContentControl Header="PARÂMETRO SELECIONADO" Style="{StaticResource oca.Section}">
                <DockPanel>
                  <StackPanel DockPanel.Dock="Top">
                    <TextBlock x:Name="tb_d_name" Style="{StaticResource oca.Value}"
                               Text="Selecione uma linha da tabela."/>
                    <TextBlock x:Name="tb_d_guid" Style="{StaticResource oca.Mono}" Margin="0,2,0,0"
                               Foreground="{StaticResource oca.Ink2}" TextWrapping="Wrap"/>
                    <TextBlock x:Name="tb_d_info" Style="{StaticResource oca.Hint}" Margin="0,6,0,8"/>
                    <TextBlock Text="Categorias" Style="{StaticResource oca.Label}" Margin="0,0,0,4"/>
                  </StackPanel>
                  <TextBox x:Name="tb_d_cats" IsReadOnly="True" Height="Auto" TextWrapping="Wrap"
                           VerticalContentAlignment="Top" VerticalScrollBarVisibility="Auto" Padding="8,6"/>
                </DockPanel>
              </HeaderedContentControl>
              <HeaderedContentControl Grid.Row="1" Header="EXPORTAR PARA O TXT CARREGADO" Style="{StaticResource oca.Section}" Margin="0">
                <StackPanel>
                  <TextBlock Text="Grupo de definição no TXT" Style="{StaticResource oca.Label}" Margin="0,0,0,4"/>
                  <ComboBox x:Name="cb_xgroup"
                            ToolTip="Grupo do arquivo TXT em que as definições serão gravadas (não é o grupo da paleta Propriedades)"/>
                  <TextBox x:Name="f_xgroup" Margin="0,6,0,0" ToolTip="Nome do grupo novo no TXT"/>
                  <TextBlock Style="{StaticResource oca.Hint}" Margin="0,8,0,0"
                             Text="Grava no TXT os parâmetros marcados com o GUID, o nome e o tipo de dado do projeto. Antes, uma cópia do TXT é salva na mesma pasta (_backup_data). O Revit regrava o arquivo no formato dele. Nada é alterado no projeto."/>
                </StackPanel>
              </HeaderedContentControl>
            </Grid>
          </Grid>
        </Grid>
      </DockPanel>
    </Border>
  </Grid>"""

FOOTER_LEFT = u"""
  <TextBlock x:Name="tb_footer" VerticalAlignment="Center" Foreground="{StaticResource oca.Ink2}"/>"""

FOOTER = u"""
  <Button x:Name="b_apply" Content="Aplicar no projeto" MinWidth="150" Style="{StaticResource oca.Primary}"
          ToolTip="Disponível na aba Revisão. Pede confirmação e grava tudo numa única transação."/>
  <Button x:Name="b_export" Content="Exportar para o TXT" MinWidth="150" Style="{StaticResource oca.Primary}"
          Visibility="Collapsed"
          ToolTip="Grava no TXT carregado os parâmetros marcados (pede confirmação e faz uma cópia do TXT antes)"/>
  <Button x:Name="b_cancel" Content="Cancelar" Margin="8,0,0,0" ToolTip="Fecha sem alterar o projeto"/>"""

XAML = build_xaml(title=__title__, subtitle=__doc__.split(u"\n")[0], body=BODY,
                  footer_right=FOOTER, footer_left=FOOTER_LEFT, resources=RESOURCES,
                  size="L", width=1280, height=900, min_width=1060, min_height=700)


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
        self.group_labels = dict(self.groups)
        self.categories = R.list_bindable_categories(document)
        self.cat_by_id = dict((c["id"], c["cat"]) for c in self.categories)
        self.cat_names = dict((c["id"], c["name"]) for c in self.categories)
        self.no_type = R.no_type_category_ids()
        self.project = R.scan_project(document, self.groups)
        # lote: guid -> {"guid", "name", "group_key", "kind", "cats"}
        self.cfg = OrderedDict()
        self.plan_rows = []
        # aba "Parâmetros no projeto": guid -> item de R.list_project_shared
        self.proj_items = {}
        self.txt_names = {}         # nome no TXT -> set(guid)
        self.proj_mode = False

        self._build_tables()
        self._init_controls()
        self._wire_events()
        self._refresh_cfg()
        self._fill_proj()
        self._fill_xgroups()
        self._update_footer()

        start = to_unicode(config.get_option("last_path", u"") or u"")
        if not (start and os.path.isfile(start)):
            start = to_unicode(self.app.SharedParametersFilename or u"")
        if start and os.path.isfile(start):
            self._load(start)
        else:
            self.tb_file_info.Text = u"Nenhum arquivo carregado. Clique em 'Carregar arquivo…'."

    # ---------------------------------------------------------- montagem
    def _build_tables(self):
        S, B, I = System.String, System.Boolean, System.Int32
        self.t_params = new_table(u"params", [
            ("Sel", B), ("Usavel", B), ("Guid", S), ("Nome", S), ("GrupoArq", S), ("Tipo", S),
            ("GuidTxt", S), ("Situacao", S), ("SitKey", S), ("Nivel", S), ("Lote", S), ("Obs", S),
            ("Busca", S)])
        self.t_cfg = new_table(u"cfg", [
            ("Guid", S), ("Nome", S), ("Grupo", S), ("Vinculo", S), ("Categorias", S),
            ("Situacao", S), ("Nivel", S)])
        self.t_cats = new_table(u"cats", [("Sel", B), ("Id", I), ("Nome", S), ("Tipo", S), ("Busca", S)])
        self.t_proj = new_table(u"proj", [
            ("Sel", B), ("Pode", B), ("Guid", S), ("Nome", S), ("Tipo", S), ("Vinculo", S),
            ("Grupo", S), ("Categorias", S), ("Origem", S), ("OrigemKey", S), ("NoTxt", S),
            ("TxtKey", S), ("Nivel", S), ("Busca", S)])
        self.t_review = new_table(u"review", [
            ("Nome", S), ("Acao", S), ("Grupo", S), ("Vinculo", S), ("Categorias", S), ("Obs", S),
            ("Nivel", S)])
        for c in self.categories:
            row = self.t_cats.NewRow()
            row["Sel"] = False
            row["Id"] = c["id"]
            row["Nome"] = c["name"]
            row["Tipo"] = c["type"]
            row["Busca"] = fold(c["name"])
            self.t_cats.Rows.Add(row)
        self.grid_params.ItemsSource = self.t_params.DefaultView
        self.grid_cfg.ItemsSource = self.t_cfg.DefaultView
        self.grid_cats.ItemsSource = self.t_cats.DefaultView
        self.grid_review.ItemsSource = self.t_review.DefaultView
        self.grid_proj.ItemsSource = self.t_proj.DefaultView

    def _init_controls(self):
        self.cb_fstatus.ItemsSource = list(STATUS_FILTERS.keys())
        self.cb_fstatus.SelectedIndex = 0
        self.cb_fgroup.ItemsSource = [ALL_GROUPS]
        self.cb_fgroup.SelectedIndex = 0
        self.cb_ftype.ItemsSource = [ALL_TYPES]
        self.cb_ftype.SelectedIndex = 0
        types = sorted(set(c["type"] for c in self.categories))
        self.cb_cat_type.ItemsSource = [ALL_CAT_TYPES] + types
        self.cb_cat_type.SelectedIndex = 0
        self.cb_group.ItemsSource = [label for _, label in self.groups]
        default = R.default_group_key(self.groups)
        keys = [k for k, _ in self.groups]
        self.cb_group.SelectedIndex = keys.index(default) if default in keys else -1
        self.cb_porigin.ItemsSource = list(PROJ_ORIGINS.keys())
        self.cb_porigin.SelectedIndex = 0
        self.cb_ptxt.ItemsSource = list(PROJ_TXT.keys())
        self.cb_ptxt.SelectedIndex = 0
        self._update_cat_count()
        self._update_count()

    def _wire_events(self):
        sf = self._safe
        self.b_load.Click += sf(self._on_browse)
        self.b_reload.Click += sf(self._on_reload)
        self.f_search.TextChanged += sf(lambda s, a: self._filter_params())
        for combo in (self.cb_fgroup, self.cb_ftype, self.cb_fstatus):
            combo.SelectionChanged += sf(lambda s, a: self._filter_params())
        self.chk_only_marked.Click += sf(lambda s, a: self._filter_params())
        self.grid_params.PreviewKeyDown += sf(self._on_params_key)
        self.b_sel_all.Click += sf(self._on_select_all)
        self.b_sel_none.Click += sf(self._on_select_none)
        self.b_configure.Click += sf(self._on_configure)
        self.t_params.ColumnChanged += self._on_param_changed

        self.b_add_more.Click += sf(lambda s, a: self._go(TAB_FILE))
        self.b_cfg_all.Click += sf(lambda s, a: self.grid_cfg.SelectAll())
        self.b_copy_from.Click += sf(self._on_copy_from)
        self.b_remove.Click += sf(self._on_remove)
        self.b_review.Click += sf(lambda s, a: self._go(TAB_REVIEW))
        self.b_apply_sel.Click += sf(lambda s, a: self._apply_controls(selected_only=True))
        self.b_apply_all.Click += sf(lambda s, a: self._apply_controls(selected_only=False))
        self.f_cat_search.TextChanged += sf(lambda s, a: self._filter_cats())
        self.cb_cat_type.SelectionChanged += sf(lambda s, a: self._filter_cats())
        self.chk_cat_only.Click += sf(lambda s, a: self._filter_cats())
        self.b_cat_all.Click += sf(self._on_cat_all)
        self.b_cat_none.Click += sf(self._on_cat_none)
        self.grid_cats.PreviewKeyDown += sf(self._on_cats_key)
        self.t_cats.ColumnChanged += self._on_cat_changed

        self.chk_extend.Click += sf(lambda s, a: self._refresh_cfg())
        self.tabs.SelectionChanged += sf(self._on_tab)
        self.b_apply.Click += sf(self._on_apply)
        self.b_cancel.Click += sf(lambda s, a: self.Close())

        self.rb_mode_link.Checked += sf(lambda s, a: self._set_mode(False))
        self.rb_mode_proj.Checked += sf(lambda s, a: self._set_mode(True))
        self.f_psearch.TextChanged += sf(lambda s, a: self._filter_proj())
        for combo in (self.cb_porigin, self.cb_ptxt):
            combo.SelectionChanged += sf(lambda s, a: self._filter_proj())
        self.chk_p_only.Click += sf(lambda s, a: self._filter_proj())
        self.grid_proj.SelectionChanged += sf(self._on_proj_select)
        self.grid_proj.PreviewKeyDown += sf(self._on_proj_key)
        self.b_psel_all.Click += sf(lambda s, a: self._set_proj_marks(
            [drv.Row for drv in self.t_proj.DefaultView], True))
        self.b_psel_none.Click += sf(lambda s, a: self._set_proj_marks(list(self.t_proj.Rows), False))
        self.t_proj.ColumnChanged += self._on_proj_changed
        self.cb_xgroup.SelectionChanged += sf(self._on_xgroup)
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
        if self.cfg and path != self.path:
            if not forms.alert(u"Carregar outro arquivo e limpar a configuração?",
                               sub_msg=u"A configuração tem {} parâmetro(s) do arquivo atual. Cada lote "
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

        self._fill_params()
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
            warns.append(u"{} parâmetro(s) saíram da configuração: não estão mais disponíveis no "
                         u"arquivo.".format(len(dropped)))
        self.tb_file_warn.Text = u"\n".join(warns)
        self._refresh_cfg()
        self._fill_proj()
        self._fill_xgroups()

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

    def _fill_params(self):
        self._commit(self.grid_params)
        self._bulk = True
        try:
            self.t_params.Rows.Clear()
            groups, types = [], set()
            for p in file_order(self.spf):
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
                row["Sel"] = False
                row["Usavel"] = usable
                row["Guid"] = p.guid_key or u"#linha{}".format(p.line_no)
                row["Nome"] = p.name
                row["GrupoArq"] = p.group_name
                row["Tipo"] = datatype
                row["GuidTxt"] = p.guid
                row["Situacao"] = sit
                row["SitKey"] = key
                row["Nivel"] = level
                row["Lote"] = u"✓" if p.guid_key in self.cfg else u""
                row["Obs"] = u" · ".join(notes)
                row["Busca"] = fold(u" ".join((p.name, p.guid, p.group_name, datatype, p.datatype,
                                               p.description)))
                self.t_params.Rows.Add(row)
                if p.group_name not in groups:
                    groups.append(p.group_name)
                types.add(datatype)
        finally:
            self._bulk = False
        self.cb_fgroup.ItemsSource = [ALL_GROUPS] + groups
        self.cb_fgroup.SelectedIndex = 0
        self.cb_ftype.ItemsSource = [ALL_TYPES] + sorted(t for t in types if t)
        self.cb_ftype.SelectedIndex = 0
        self._filter_params()

    # ---------------------------------------------------------- tabela de parâmetros
    def _commit(self, grid):
        """Fecha edição pendente da tabela (evita erro ao refiltrar)."""
        try:
            grid.CommitEdit(DataGridEditingUnit.Row, True)
        except Exception:
            try:
                grid.CancelEdit()
            except Exception:
                pass

    def _filter_params(self):
        self._commit(self.grid_params)
        parts = []
        text = search_filter(self.f_search.Text or u"")
        if text:
            parts.append(text)
        group = self.cb_fgroup.SelectedItem
        if group and group != ALL_GROUPS:
            parts.append(eq_filter(u"GrupoArq", group))
        dtype = self.cb_ftype.SelectedItem
        if dtype and dtype != ALL_TYPES:
            parts.append(eq_filter(u"Tipo", dtype))
        status = STATUS_FILTERS.get(self.cb_fstatus.SelectedItem or u"", u"")
        if status:
            parts.append(status)
        if self.chk_only_marked.IsChecked:
            parts.append(u"Sel = true")
        self.t_params.DefaultView.RowFilter = u" AND ".join(u"({})".format(p) for p in parts)
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

    def _set_marks(self, rows, value):
        self._commit(self.grid_params)
        self._bulk = True
        try:
            for row in rows:
                if row["Usavel"] or not value:
                    row["Sel"] = value
        finally:
            self._bulk = False
        self._update_count()

    def _on_select_all(self, sender, args):
        self._set_marks([drv.Row for drv in self.t_params.DefaultView], True)

    def _on_select_none(self, sender, args):
        self._set_marks(list(self.t_params.Rows), False)

    def _on_params_key(self, sender, args):
        if args.Key != Key.Space:
            return
        rows = [drv.Row for drv in list(self.grid_params.SelectedItems)]
        if rows:
            value = not all(bool(r["Sel"]) for r in rows if r["Usavel"])
            self._set_marks(rows, value)
            args.Handled = True

    def _on_configure(self, sender, args):
        self._commit(self.grid_params)
        marked = [r for r in self.t_params.Select(u"Sel = true") if r["Usavel"]]
        if not marked:
            self.tb_params_msg.Text = u"Marque ao menos um parâmetro disponível."
            return
        self.tb_params_msg.Text = u""
        group_key, kind, cats = self._controls()
        by_guid = self.spf_by_guid
        added, already = [], 0
        for row in marked:
            guid = to_unicode(row["Guid"])
            if guid in self.cfg:
                already += 1
                continue
            p = by_guid.get(guid)
            if p is None:
                continue
            self.cfg[guid] = {"guid": guid, "name": p.name, "group_key": group_key,
                              "kind": kind, "cats": set(cats)}
            added.append(guid)
        self._set_marks(marked, False)
        self._refresh_lote()
        self._refresh_cfg(select=added)
        msg = u"{} parâmetro(s) adicionados com as opções atuais dos controles.".format(len(added))
        if already:
            msg += u" {} já estavam na configuração.".format(already)
        if not cats:
            msg += u" Escolha as categorias e clique em 'Aplicar às selecionadas'."
        self.tb_cfg_msg.Text = msg
        self._go(TAB_CFG)

    def _refresh_lote(self):
        self._bulk = True
        try:
            for row in self.t_params.Rows:
                row["Lote"] = u"✓" if to_unicode(row["Guid"]) in self.cfg else u""
        finally:
            self._bulk = False

    # ---------------------------------------------------------- configuração
    def _controls(self):
        """(grupo, vínculo, categorias) escolhidos nos controles à direita."""
        idx = self.cb_group.SelectedIndex
        group_key = self.groups[idx][0] if 0 <= idx < len(self.groups) else None
        kind = P.TYPE if self.rb_type.IsChecked else P.INSTANCE
        self._commit(self.grid_cats)
        cats = set(int(r["Id"]) for r in self.t_cats.Select(u"Sel = true"))
        return group_key, kind, cats

    def _selected_cfg_guids(self):
        return [to_unicode(drv["Guid"]) for drv in list(self.grid_cfg.SelectedItems)]

    def _apply_controls(self, selected_only):
        if not self.cfg:
            self.tb_cfg_msg.Text = u"A configuração está vazia. Use '+ Adicionar parâmetros'."
            return
        targets = self._selected_cfg_guids() if selected_only else list(self.cfg)
        if not targets:
            self.tb_cfg_msg.Text = u"Selecione ao menos uma linha da configuração (ou use 'Aplicar a todas')."
            return
        set_group, set_kind, set_cats = (bool(self.chk_set_group.IsChecked),
                                         bool(self.chk_set_kind.IsChecked),
                                         bool(self.chk_set_cats.IsChecked))
        if not (set_group or set_kind or set_cats):
            self.tb_cfg_msg.Text = u"Marque o que aplicar: Grupo, Vínculo e/ou Categorias."
            return
        group_key, kind, cats = self._controls()
        if set_group and group_key is None:
            self.tb_cfg_msg.Text = u"Escolha o grupo de propriedades."
            return
        if set_cats and not cats:
            self.tb_cfg_msg.Text = u"Marque ao menos uma categoria (ou desmarque 'Categorias' em Aplicar)."
            return
        for guid in targets:
            item = self.cfg[guid]
            if set_group:
                item["group_key"] = group_key
            if set_kind:
                item["kind"] = kind
            if set_cats:
                item["cats"] = set(cats)
        self._refresh_cfg(select=targets if selected_only else None)
        what = [w for w, on in ((u"grupo", set_group), (u"vínculo", set_kind),
                                (u"categorias", set_cats)) if on]
        self.tb_cfg_msg.Text = u"{} aplicado(s) em {} parâmetro(s).".format(
            u", ".join(what).capitalize(), len(targets))

    def _on_copy_from(self, sender, args):
        guids = self._selected_cfg_guids()
        if not guids:
            self.tb_cfg_msg.Text = u"Selecione uma linha da configuração para copiar as opções."
            return
        item = self.cfg[guids[0]]
        keys = [k for k, _ in self.groups]
        if item["group_key"] in keys:
            self.cb_group.SelectedIndex = keys.index(item["group_key"])
        self.rb_type.IsChecked = item["kind"] == P.TYPE
        self.rb_inst.IsChecked = item["kind"] != P.TYPE
        self._commit(self.grid_cats)
        self._bulk = True
        try:
            for row in self.t_cats.Rows:
                row["Sel"] = int(row["Id"]) in item["cats"]
        finally:
            self._bulk = False
        self._update_cat_count()
        self.tb_cfg_msg.Text = u"Opções de '{}' copiadas para os controles.".format(item["name"])

    def _on_remove(self, sender, args):
        guids = self._selected_cfg_guids()
        if not guids:
            self.tb_cfg_msg.Text = u"Selecione as linhas que quer tirar da configuração."
            return
        for guid in guids:
            self.cfg.pop(guid, None)
        self._refresh_lote()
        self._refresh_cfg()
        self.tb_cfg_msg.Text = u"{} parâmetro(s) removido(s) da configuração.".format(len(guids))

    def _build_plan(self):
        items = list(self.cfg.values())
        defs = dict((g, self._def_info_by_guid(g)) for g in self.cfg)
        self.plan_rows = P.build_plan(items, defs, self.project, self.cat_names, self.group_labels,
                                      self.no_type, allow_extend=bool(self.chk_extend.IsChecked))
        return self.plan_rows

    def _refresh_cfg(self, select=None):
        """Remonta a tabela da configuração e a revisão a partir de self.cfg."""
        keep = set(select) if select is not None else set(self._selected_cfg_guids())
        rows = self._build_plan()
        self.t_cfg.Rows.Clear()
        for r in rows:
            item = self.cfg[r.guid]
            row = self.t_cfg.NewRow()
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
        self._update_footer()

    # ---------------------------------------------------------- categorias
    def _filter_cats(self):
        self._commit(self.grid_cats)
        parts = []
        text = search_filter(self.f_cat_search.Text or u"")
        if text:
            parts.append(text)
        ctype = self.cb_cat_type.SelectedItem
        if ctype and ctype != ALL_CAT_TYPES:
            parts.append(eq_filter(u"Tipo", ctype))
        if self.chk_cat_only.IsChecked:
            parts.append(u"Sel = true")
        self.t_cats.DefaultView.RowFilter = u" AND ".join(u"({})".format(p) for p in parts)
        self._update_cat_count()

    def _on_cat_changed(self, sender, args):
        if not self._bulk and args.Column.ColumnName == "Sel":
            self._update_cat_count()

    def _update_cat_count(self):
        marked = len(self.t_cats.Select(u"Sel = true"))
        self.tb_cat_count.Text = u"{} marcada(s) · {} de {} na lista".format(
            marked, self.t_cats.DefaultView.Count, self.t_cats.Rows.Count)

    def _set_cat_marks(self, rows, value):
        self._commit(self.grid_cats)
        self._bulk = True
        try:
            for row in rows:
                row["Sel"] = value
        finally:
            self._bulk = False
        self._update_cat_count()

    def _on_cat_all(self, sender, args):
        self._set_cat_marks([drv.Row for drv in self.t_cats.DefaultView], True)

    def _on_cat_none(self, sender, args):
        self._set_cat_marks(list(self.t_cats.Rows), False)

    def _on_cats_key(self, sender, args):
        if args.Key != Key.Space:
            return
        rows = [drv.Row for drv in list(self.grid_cats.SelectedItems)]
        if rows:
            self._set_cat_marks(rows, not all(bool(r["Sel"]) for r in rows))
            args.Handled = True

    # ---------------------------------------------------------- revisão e execução
    def _fill_review(self, rows):
        self.t_review.Rows.Clear()
        for r in rows:
            row = self.t_review.NewRow()
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
            text, style = u"A configuração está vazia: adicione parâmetros na aba 1.", "oca.Msg.Info"
        elif count[P.INCOMPLETE]:
            text = (u"{} parâmetro(s) incompleto(s) (sem categoria ou sem grupo): corrija na aba "
                    u"Configuração antes de aplicar.".format(count[P.INCOMPLETE]))
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

    def _update_footer(self):
        if self.proj_mode:
            self.tb_footer.Text = u"{} parâmetro(s) compartilhado(s) no projeto".format(
                len(self.proj_items))
            return
        n = len(self.cfg)
        self.tb_footer.Text = u"Configuração: {} parâmetro(s)".format(n) if n else u""
        self.b_apply.IsEnabled = self.tabs.SelectedIndex == TAB_REVIEW and self._ready()

    def _set_mode(self, proj):
        """Troca a aba principal: o botão principal do rodapé acompanha a aba."""
        from System.Windows import Visibility
        self.proj_mode = bool(proj)
        shown, hidden = Visibility.Visible, Visibility.Collapsed
        self.pn_proj.Visibility = shown if proj else hidden
        self.pn_link.Visibility = hidden if proj else shown
        self.b_export.Visibility = shown if proj else hidden
        self.b_apply.Visibility = hidden if proj else shown
        self._update_footer()

    def _go(self, index):
        self.tabs.SelectedIndex = index

    def _on_tab(self, sender, args):
        # SelectionChanged também sobe das tabelas e listas internas
        if not System.Object.ReferenceEquals(args.OriginalSource, self.tabs):
            return
        if self.tabs.SelectedIndex == TAB_REVIEW:
            self._refresh_cfg()
        self._update_footer()

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
        """Remonta a tabela dos parâmetros compartilhados do projeto."""
        self._commit(self.grid_proj)
        marked = set(to_unicode(r["Guid"]) for r in self.t_proj.Select(u"Sel = true"))
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
            for item in items:
                text, key, level, can = self._txt_status(item)
                kind = P.KIND_LABELS.get(item["kind"], u"—") if item["bound"] else u"—"
                cats = (P.cats_summary(item["cats"], self.cat_names) if item["bound"] else u"—")
                origin = u"Parâmetro de projeto" if item["bound"] else u"Só em famílias"
                row = self.t_proj.NewRow()
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
        self._show_detail(None)
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
        self._update_proj_count()

    def _on_proj_changed(self, sender, args):
        if not self._bulk and args.Column.ColumnName == "Sel":
            self._update_proj_count()

    def _update_proj_count(self):
        total = self.t_proj.Rows.Count
        if not total:
            self.tb_pcount.Text = u"Nenhum parâmetro compartilhado no projeto."
            return
        self.tb_pcount.Text = u"{} marcado(s) · {} de {} exibido(s) · {} fora do TXT".format(
            len(self.t_proj.Select(u"Sel = true")), self.t_proj.DefaultView.Count, total,
            len(self.t_proj.Select(u"Pode = true")))

    def _set_proj_marks(self, rows, value):
        self._commit(self.grid_proj)
        self._bulk = True
        try:
            for row in rows:
                if row["Pode"] or not value:
                    row["Sel"] = value
        finally:
            self._bulk = False
        self._update_proj_count()

    def _on_proj_key(self, sender, args):
        if args.Key != Key.Space:
            return
        rows = [drv.Row for drv in list(self.grid_proj.SelectedItems)]
        if rows:
            self._set_proj_marks(rows, not all(bool(r["Sel"]) for r in rows if r["Pode"]))
            args.Handled = True

    def _on_proj_select(self, sender, args):
        if not System.Object.ReferenceEquals(args.OriginalSource, self.grid_proj):
            return
        selected = list(self.grid_proj.SelectedItems)
        self._show_detail(to_unicode(selected[0]["Guid"]) if selected else None)

    def _show_detail(self, guid):
        item = self.proj_items.get(guid) if guid else None
        if item is None:
            self.tb_d_name.Text = u"Selecione uma linha da tabela."
            self.tb_d_guid.Text = u""
            self.tb_d_info.Text = u""
            self.tb_d_cats.Text = u""
            return
        text, _, _, _ = self._txt_status(item)
        self.tb_d_name.Text = item["name"]
        self.tb_d_guid.Text = item["guid"]
        if item["bound"]:
            lines = [u"Tipo de dado: {}".format(item["datatype"] or u"?"),
                     u"Vínculo: {}".format(P.KIND_LABELS.get(item["kind"], u"?")),
                     u"Grupo de propriedades: {}".format(item["group_label"] or u"?"),
                     u"No TXT: {}".format(text)]
            cats = u"\n".join(item["cats_names"]) or u"(nenhuma)"
        else:
            lines = [u"Tipo de dado: {}".format(item["datatype"] or u"?"),
                     u"Origem: só em famílias carregadas (não é parâmetro de projeto)",
                     u"No TXT: {}".format(text)]
            cats = (u"Sem vínculo de projeto: as categorias dependem das famílias que usam o "
                    u"parâmetro.")
        self.tb_d_info.Text = u"\n".join(lines)
        self.tb_d_cats.Text = cats

    def _fill_xgroups(self):
        """Grupos de definição do TXT carregado + 'Novo grupo…'."""
        names = [self.spf.groups[g] for g in self.spf.group_order] if self.spf else []
        current = self.cb_xgroup.SelectedItem
        self.cb_xgroup.ItemsSource = names + [NEW_GROUP]
        if current in names:
            self.cb_xgroup.SelectedItem = current
        elif u"Exported Parameters" in names:          # grupo que o próprio Revit usa ao exportar
            self.cb_xgroup.SelectedItem = u"Exported Parameters"
        else:
            self.cb_xgroup.SelectedItem = NEW_GROUP
        if not (self.f_xgroup.Text or u"").strip():
            self.f_xgroup.Text = DEFAULT_NEW_GROUP
        self._on_xgroup(None, None)

    def _on_xgroup(self, sender, args):
        self.f_xgroup.IsEnabled = self.cb_xgroup.SelectedItem == NEW_GROUP

    def _export_group(self):
        """(nome do grupo, é novo?) escolhido para a exportação."""
        if self.cb_xgroup.SelectedItem == NEW_GROUP:
            name = (self.f_xgroup.Text or u"").strip()
            existing = self.spf.groups.values() if self.spf else []
            return name, name not in existing
        return to_unicode(self.cb_xgroup.SelectedItem or u""), False

    def _on_export(self, sender, args):
        self._commit(self.grid_proj)
        self.tb_proj_msg.Text = u""
        if self.spf is None or not self.path:
            self.tb_proj_msg.Text = u"Carregue um arquivo TXT antes de exportar ('Carregar arquivo…')."
            return
        if self.revit_error:
            self.tb_proj_msg.Text = u"O Revit não abriu o TXT carregado; não é possível exportar para ele."
            return
        guids = [to_unicode(r["Guid"]) for r in self.t_proj.Select(u"Sel = true") if r["Pode"]]
        if not guids:
            self.tb_proj_msg.Text = u"Marque ao menos um parâmetro que esteja fora do TXT."
            return
        group, is_new = self._export_group()
        error = R.valid_group_name(group)
        if error:
            self.tb_proj_msg.Text = error
            return
        details = [u"Arquivo: {}".format(self.path),
                   u"Grupo do TXT: {}{}".format(group, u" (novo)" if is_new else u""),
                   u"GUID, nome e tipo de dado são os do projeto. Nada muda no projeto.",
                   u"Antes de gravar, uma cópia do TXT é salva na mesma pasta.",
                   u"Descrição fica vazia e 'editável pelo usuário' fica sim: a API não informa "
                   u"esses dados no projeto."]
        if not forms.alert(u"Exportar {} parâmetro(s) para o TXT?".format(len(guids)),
                           sub_msg=u"\n".join(details), title=ALERT_TITLE, yes=True, no=True):
            return
        try:
            written, backup = R.export_definitions(self.doc, self.path, guids, group)
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
                    sub_msg=u"Cópia do arquivo anterior: {}".format(backup), title=ALERT_TITLE)


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
