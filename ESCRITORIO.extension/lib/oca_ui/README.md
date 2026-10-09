# oca_ui: padrão visual OCA

Pacote compartilhado pelos comandos da aba OCA. O pyRevit adiciona a pasta
`ESCRITORIO.extension/lib` ao caminho de importação de todos os scripts da
extensão, então qualquer comando pode usar `from oca_ui import ...`.

```
lib/oca_ui/
├─ __init__.py          build_xaml(), alert_title(), brand_report(), output_header()
├─ theme.py             paleta escura (e clara, não usada), larguras e estilos WPF (único lugar com cores)
└─ assets/logo/logo.png logo original da empresa (não editar; substituir pelo novo arquivo, mesmo nome)
```

## O que o padrão faz

- **Tema:** sempre escuro, independente do tema do Revit (decisão de padrão).
  A paleta clara (`theme.LIGHT`) só é usada no cabeçalho da janela de saída do pyRevit, que tem fundo claro.
- **Cabeçalho:** logo, nome do comando (`__title__`) e descrição (`__doc__`).
  O logo também vira o ícone da janela.
- **Rodapé:** botão principal primeiro, depois Cancelar; ações auxiliares à esquerda.
- **Controles:** TextBox, ComboBox, RadioButton, CheckBox e Button já saem
  estilizados. Não defina cores no XAML do comando.
- **Logo:** é carregado com `CacheOption=OnLoad`, então o arquivo não fica
  travado e pode ser substituído com o Revit aberto. Se o arquivo faltar ou
  estiver corrompido, a janela abre sem logo.

## Como criar ou converter um comando

1. Escreva só o conteúdo da janela (corpo) e os botões (rodapé), mantendo os
   `x:Name` que o código usa:

```python
from pyrevit import forms
from oca_ui import build_xaml, alert_title

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
  <Button x:Name="ok" Content="Executar" Style="{StaticResource oca.Primary}"/>
  <Button x:Name="cancel" Content="Cancelar" Margin="8,0,0,0"/>"""

XAML = build_xaml(title=__title__, subtitle=__doc__, body=BODY,
                  footer_right=FOOTER, size="S")   # S 460 px, M 580 px, L 1160 px (+ height)
```

2. A classe da janela continua igual (`forms.WPFWindow.__init__(self, XAML, literal_string=True)`).
3. Em todo `forms.alert`, passe `title=ALERT_TITLE`. Escreva uma frase
   principal curta em `msg` e coloque o detalhe ou o caminho de correção em `sub_msg`.

## Estilos disponíveis

| Chave | Uso |
|---|---|
| `oca.Section` | `HeaderedContentControl`: título de seção em caixa-alta com filete |
| `oca.Label` / `oca.LabelWidth` | rótulo da coluna esquerda / largura padrão dessa coluna |
| `oca.Hint` | texto de ajuda, 11 px, cor secundária |
| `oca.Caption` / `oca.Value` | legenda e valor em cartões de pré-visualização |
| `oca.Card` | `Border` de pré-visualização |
| `oca.Msg.Info` / `.Ok` / `.Warn` / `.Err` + `oca.MsgText` | faixa de mensagem; some sozinha quando o texto está vazio |
| `oca.Primary` / `oca.Ghost` | botão principal / botão sem borda ("Voltar") |
| `oca.Bg`, `oca.Surface`, `oca.Ink`, `oca.Ink2`, `oca.Line`, `oca.Accent`, `oca.Brand`, `oca.Warn`, `oca.Err`, `oca.Ok`... | pincéis, quando um elemento precisar de cor explícita |

## Relatórios HTML e janela de saída

- **Relatório:** no modelo `.html`, use `/*__OCA_TOKENS__*/` dentro do `:root` e
  `var(--oca-Bg)`, `var(--oca-Accent)`... nas cores; use `__OCA_LOGO__` como `src` do logo.
  Na geração: `html = brand_report(html)`. O logo vai embutido (base64), porque o
  relatório é gravado fora da extensão.
- **Janela de saída:** `output_header(output, __title__, u"Log de execução")` no início do script.
