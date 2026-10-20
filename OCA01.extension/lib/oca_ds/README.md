# oca_ds: Design System OCA

Pacote compartilhado pelos comandos da aba OCA. O pyRevit adiciona a pasta
`lib` da extensão ao caminho de importação de todos os scripts, então qualquer
comando pode usar `from oca_ds import ...`.

```
lib/oca_ds/
├─ __init__.py          atalhos: build_xaml, alert_title, alert, brand_report, output_header, is_dark
├─ tokens.py            cores claro/escuro, raios, fontes, medidas, paleta do Excel (único lugar com cores)
├─ theme.py             tema automático: segue o tema do Revit (UIThemeManager)
├─ wpf.py               estilos WPF de todos os controles + build_xaml (cabeçalho, corpo, rodapé)
├─ messages.py          título e texto dos diálogos nativos
├─ report.py            relatórios HTML (cores, tema e logo) e cabeçalho da janela de saída
├─ excel.py             planilhas .xlsx no padrão (estilos, tabela, filtros, impressão, aba Resumo)
├─ assets/logo/logo.png logo original da empresa (não editar; substituir pelo novo arquivo, mesmo nome)
└─ assets/logo/icon.png logo quadrado 64 px, ícone da barra de título (gerado do logo.png)
```

## Tema

- **Automático:** cada janela e cada relatório usa o tema do Revit no momento em que é aberto
  (`UIThemeManager.CurrentTheme`, Revit 2024+). Revit 2023 ou anterior: tema claro.
- **Claro:** fundo `#F3F5F6`, painéis brancos. **Escuro (opção A):** fundo `#2A2F35`, painéis `#33393F`.
- **Marca:** `#2FBAA7` nos dois temas (logo e botão principal).
- **Limites:** a janela aberta não troca de tema sozinha (reabrir aplica). A barra de título e o
  contorno externo são do Windows. Os diálogos (TaskDialog) e a barra de progresso do pyRevit são nativos.
- **A testar no Revit:** com "usar a configuração do sistema" ligada, confirmar que o tema
  seguido é o efetivamente aplicado.

## Janelas

```python
from pyrevit import forms
from oca_ds import build_xaml, alert_title

ALERT_TITLE = alert_title(__title__)          # "OCA · Nome do Comando"

BODY = u"""
  <StackPanel>
    <HeaderedContentControl Header="SEÇÃO" Style="{StaticResource oca.Section}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="{StaticResource oca.LabelWidth}"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="Rótulo" Style="{StaticResource oca.Label}"/>
        <ComboBox x:Name="campo" Grid.Column="1"/>
      </Grid>
    </HeaderedContentControl>
    <Border Style="{StaticResource oca.Msg.Err}" Visibility="{Binding Visibility, ElementName=error}">
      <TextBlock x:Name="error" Style="{StaticResource oca.MsgText}"/>
    </Border>
  </StackPanel>"""

FOOTER = u"""
  <Button x:Name="ok" Content="Criar vistas" Style="{StaticResource oca.Primary}"/>
  <Button x:Name="cancel" Content="Cancelar" Margin="8,0,0,0"/>"""

XAML = build_xaml(title=__title__, subtitle=__doc__, body=BODY,
                  footer_right=FOOTER, size="S")   # S 460, M 580, L 1160 (+ height); width=, min_width=, min_height= se precisar
```

- A classe da janela continua igual (`forms.WPFWindow.__init__(self, XAML, literal_string=True)`).
- **Ícone da barra de título:** o pyRevit troca o ícone de toda janela pelo dele depois de carregar
  o XAML (`WPFWindow.setup_icon`). Ao importar `oca_ds`, esse passo passa a usar `icon.png`, só nas
  janelas montadas por `build_xaml` (marcadas com `Tag="oca-light|oca-dark"`); as outras janelas do
  pyRevit não mudam.
- **Botão principal:** verbo explícito, só a primeira letra maiúscula ("Atualizar descrições"),
  `Style="{StaticResource oca.Primary}"`, primeiro do rodapé à direita. Um por janela.
- **Nunca** escreva cores no XAML do comando: use os pincéis `oca.*`.
- **Raios:** 4 px em campos, listas e botões; 6 px em seções, mensagens, tabelas, cartões e dicas;
  8 px em menus suspensos. Nenhum canto vivo dentro da janela.

| Chave | Uso |
|---|---|
| `oca.Section` | `HeaderedContentControl`: caixa com título em caixa-alta; o conteúdo ocupa o espaço restante |
| `oca.Label` / `oca.LabelWidth` | rótulo da coluna esquerda / largura padrão dessa coluna (120) |
| `oca.Hint`, `oca.Caption`, `oca.Value`, `oca.Mono` | ajuda (11), legenda, valor em destaque, código (Consolas) |
| `oca.Card` | `Border` de pré-visualização |
| `oca.Msg.Info` / `.Ok` / `.Warn` / `.Err` + `oca.MsgText` | faixa de mensagem; some sozinha quando o texto está vazio |
| `oca.Primary` / `oca.Ghost` / `oca.Compact` | botão principal / sem borda ("← Voltar") / compacto (24 px) |
| `oca.RadiusControl`, `oca.RadiusGroup`, `oca.RadiusPanel` | `CornerRadius` 4 · 6 · 8 para bordas próprias do comando |
| `oca.Bg`, `oca.Surface`, `oca.Surface2`, `oca.Field`, `oca.Line`, `oca.FieldLine`, `oca.Ink`, `oca.Ink2`, `oca.Accent`, `oca.AccentSoft`, `oca.AccentInk`, `oca.Brand`, `oca.BrandInk`, `oca.Ok(Soft)`, `oca.Warn(Soft)`, `oca.Err(Soft)`, `oca.Violet` | pincéis, quando um elemento precisar de cor explícita |

Código que pinta controles em tempo de execução usa `self.FindResource("oca.Warn")`, nunca um hexadecimal.

## Redimensionamento das janelas

Toda janela montada por `build_xaml` é redimensionável pelas bordas e pelos cantos (cursor de
redimensionar do Windows), com maximizar, restaurar e minimizar. Não é preciso fazer nada no
comando: o comportamento vem do `build_xaml` e do ajuste que a lib faz na abertura da janela.

| Tipo | `size` | Abertura | Ao redimensionar |
|---|---|---|---|
| Formulário | `"S"` (460) ou `"M"` (580) | altura pelo conteúdo | a largura do conteúdo acompanha a janela; quando a altura não basta, o corpo rola na vertical |
| Tabela | `"L"` (1160) + `height` | largura e altura informadas | as linhas e colunas `*` (tabelas) ocupam o espaço; cada tabela rola sozinha |

- **Tamanho mínimo:** padrão de 75% da largura inicial (nunca menos que 420 em formulários e 760 em
  tabelas) e altura mínima de 300 (formulários) ou 460 (tabelas). Use `min_width=` e `min_height=`
  quando a janela precisar de outro mínimo (ex.: Acabamentos 900 × 600).
- **Cabe na tela:** na abertura, a janela é limitada à área útil do monitor onde abriu (sem passar
  da barra de tarefas) e reposicionada se estiver fora dela; formulários mais altos que a tela
  passam a rolar. Isso roda no mesmo passo do ícone (`fit_to_screen`, ver "Ícone da barra de título").
- **Rolagem:** formulários ganham um `ScrollViewer` vertical da lib. Se o corpo do comando já começa
  com `<ScrollViewer`, a lib não envolve de novo (evita rolagem dentro de rolagem). Para forçar,
  `scroll=True/False`.
- **Textos que acompanham a largura:** os textos de opções (`RadioButton`, `CheckBox`) e os títulos
  de seção quebram linha em vez de serem cortados.
- **Exceção:** `resizable=False` deixa a janela com tamanho fixo, como antes. Use só se a janela
  realmente não puder mudar de tamanho.

Ao montar o corpo de um comando novo, para que ele aproveite o espaço:
- use colunas e linhas `*` (e não larguras fixas) nas áreas que devem crescer; campos e listas
  esticam sozinhos dentro das seções;
- em janelas de tabela, deixe a tabela numa linha `Height="*"`; em formulários, empilhe as
  seções num `StackPanel` (a lib cuida da rolagem);
- reserve larguras fixas para campos numéricos curtos (ex.: margem em metros, `Width="72"`).

## Mensagens

Diálogos nativos (`forms.alert`), com `title=ALERT_TITLE`:

| Tipo | Frase principal (`msg`) | Detalhe (`sub_msg`) |
|---|---|---|
| Confirmação | a pergunta: "Gravar 221 campos em 48 ambientes?" | o que muda e como desfazer |
| Sucesso | o resultado: "Arquivo exportado. Abrir agora?" | caminho do arquivo |
| Alerta | o que impede: "Nenhuma linha está marcada." | como resolver |
| Erro | o que falhou: "A atualização falhou e foi desfeita." | mensagem técnica do Revit |

## Relatórios HTML e janela de saída

- No modelo `.html`: `/*__OCA_TOKENS__*/` dentro do `:root`, cores com `var(--oca-Bg)`,
  `var(--oca-Accent)`...; `__OCA_LOGO__` como `src` do logo; `__OCA_THEME__` (light/dark) se precisar;
  `/*__OCA_XLSX__*/null` para receber os estilos do Excel gerado no navegador.
- Na geração: `html = brand_report(html)` (tema do Revit nesse momento).
- Janela de saída do pyRevit: `output_header(output, __title__, u"Log de execução")`.

## Excel

```python
from oca_ds import excel as oca_excel

parts = oca_excel.workbook_parts(
    [{"name": u"Ambientes", "rows": [cabecalho] + linhas, "widths": larguras,
      "skip_empty": True, "mono_cols": (3, 4)}],
    title=u"Memorial de acabamentos", project=u"Projeto.rvt",
    summary=[(u"Projeto", u"Projeto.rvt"), (u"Ambientes", 48)])   # None = sem aba Resumo
# parts = {caminho: texto}; o logo do Resumo vem como bytearray (grave em binário)
```

Integridade: o módulo grava os valores exatamente como recebe. A regra de cada aba diz só
**como** gravar (`numbers`, `num_format`, `none_text`, `skip_empty`); textos continuam texto
(zeros à esquerda e IDs preservados), nada é arredondado, nenhuma fórmula é criada, os dados
começam na linha 1 e a aba Resumo é sempre a última.

Visual: cabeçalho `#1D7266` com texto branco em negrito, linha 1 congelada, tabela estruturada
com filtro e linhas alternadas (estilo de tabela "OCA"; quando os títulos se repetem, só filtro; sem dados, nem tabela nem filtro, como antes), Segoe UI 10, códigos em Consolas, A4 paisagem ajustado à
largura com a linha 1 repetida, cabeçalho "OCA · título | projeto" e rodapé "aba | Página x de y".
