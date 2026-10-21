# Gerenciador de Parâmetros Compartilhados

Vincula ao projeto, em lote, parâmetros de um arquivo TXT de parâmetros compartilhados,
mantendo os GUIDs originais, e mostra os parâmetros compartilhados que já estão no projeto, com
exportação para o TXT. Ambiente-alvo: pyRevit 5.x com IronPython 2.7, Revit 2024/2025/2026.
Os grupos de parâmetro também têm tratamento para Revit 2021–2023.

## Arquivos

```
OCA01.extension/
├─ OCA.tab/
│  ├─ bundle.yaml                       ← precisa de "- Parâmetros" no layout (ver Instalação)
│  └─ Parâmetros.panel/                 NOVO
│     ├─ bundle.yaml                    cor do rodapé do painel (claro/escuro), lida pelo startup.py
│     └─ 01_Gerenciador de Parâmetros Compartilhados.pushbutton/
│        ├─ script.py                   janela WPF (2 abas principais) e fluxo do comando
│        ├─ icon.png                    96 px, traço na cor do painel
│        └─ LEIA-ME.md                  este arquivo
└─ lib/
   └─ oca_params/                       NOVO
      ├─ __init__.py
      ├─ spfile.py                      leitura do TXT (Python puro, só leitura)
      ├─ plan.py                        regras: o que criar, ampliar ou bloquear (Python puro)
      └─ revit_ops.py                   API do Revit: versões, categorias, vínculos, transação
```

Nenhum arquivo existente foi alterado. A identidade visual vem de `lib/oca_ds` (`build_xaml`,
pincéis `oca.*`, tema claro/escuro seguindo o Revit, janela redimensionável).

## Janela

Duas abas principais no topo, no estilo das abas do Explorador de Arquivos, com a barra do arquivo
TXT logo abaixo (vale para as duas):

- **Vincular do TXT ao projeto**: etapas 1 · Parâmetros do arquivo, 2 · Configuração, 3 · Revisão.
- **Parâmetros no projeto**: todos os parâmetros compartilhados do modelo (`SharedParameterElement`),
  com tipo de dado, Instância/Tipo, grupo de propriedades, categorias, origem (parâmetro de projeto
  ou só em famílias), GUID e situação no TXT carregado (no TXT, no TXT com outro nome, conflito de
  nome, fora do TXT). Busca, filtros, detalhe com todas as categorias da linha selecionada e
  **Exportar para o TXT** (grupo existente do TXT ou grupo novo).

O botão principal do rodapé acompanha a aba: "Aplicar no projeto" ou "Exportar para o TXT".

## Arquitetura

1. **Leitura do TXT (`spfile`)**: detecta a codificação (UTF-16 LE do Revit, UTF-8, ANSI), lê as
   colunas pelo cabeçalho `*PARAM` (aceita arquivos antigos com menos colunas) e aponta problemas
   sem corrigir nada: GUID vazio/inválido/repetido, grupo inexistente, nome vazio (bloqueiam);
   oculto, não editável, nome repetido no grupo, espaços no nome (avisos).
2. **Leitura pelo Revit (`revit_ops.read_definitions`)**: as definições usadas na vinculação vêm
   do próprio Revit (`Application.OpenSharedParameterFile`), associadas pelo GUID. Uma definição
   só é "disponível" se o TXT estiver correto **e** o Revit a tiver lido.
3. **Estado do projeto (`scan_project`)**: percorre `Document.ParameterBindings` (tipo de vínculo,
   categorias, grupo, GUID) e os `SharedParameterElement` (parâmetros que só existem por famílias).
4. **Plano (`plan.build_plan`)**: para cada item do lote decide `Criar vínculo`, `Adicionar
   categorias`, `Sem alteração`, `Não será aplicado` (com o motivo) ou `Incompleto`. A aba Revisão
   mostra exatamente esse plano; o plano é refeito com o projeto relido antes da confirmação.
5. **Execução (`apply_plan`)**: uma única `Transaction`. Qualquer recusa, exceção ou commit não
   confirmado → `RollBack` e o projeto fica como estava.
6. **Parâmetros do projeto (`list_project_shared`)**: lê os `SharedParameterElement` e cruza com
   os vínculos lidos em `scan_project`. Só leitura, sem transação.
7. **Exportação (`export_definitions`)**: valida tudo antes da primeira gravação, faz a cópia do TXT
   e grava pelo Revit (`Definitions.Create` com `ExternalDefinitionCreationOptions.GUID` = GUID do
   projeto). Depois, o TXT é relido e as duas abas são atualizadas.

## Regras de segurança

- Vincular não grava no TXT. O TXT só é gravado por **Exportar para o TXT**, sempre com
  confirmação e com cópia antes (`<nome>_backup_<AAAAMMDD_HHMMSS>.txt`, na mesma pasta). Sem cópia,
  nada é gravado.
- Exportar mantém o GUID do projeto (nenhum GUID novo é gerado) e só grava parâmetros que **não**
  estão no TXT. Bloqueia: mesmo GUID já no TXT, ou mesmo nome no TXT com outro GUID. Nada muda
  no projeto.
- Nada é excluído: `BindingMap.Remove` e `Document.Delete` não são usados.
- Parâmetro já vinculado com o **mesmo** tipo de vínculo: recebe só as categorias que faltam
  (`ReInsert` com categorias atuais + novas, mantendo o grupo atual). Antes do commit, até 300
  valores já preenchidos são relidos; se algum mudar ou sumir, tudo é desfeito. Pode ser
  desligado na aba Revisão.
- Parâmetro já vinculado com tipo **diferente** (Instância ↔ Tipo): não é aplicado. Trocar exigiria
  excluir o parâmetro e os valores.
- Nome igual a outro parâmetro de projeto (outro GUID ou não compartilhado), ou dois itens do lote
  com o mesmo nome: não é aplicado.

## Riscos e limitações da API (conhecidos antes da implementação)

Exportação para o TXT:

| Risco | Tratamento |
|---|---|
| O Revit regrava o arquivo inteiro no formato dele ao criar uma definição (pode reordenar linhas). | Cópia do TXT antes de gravar; o arquivo continua sendo um TXT do Revit (UTF-16). |
| Descrição e "editável pelo usuário" de um parâmetro não são expostos pela API no projeto. | Exportados como descrição vazia e editável = sim. Visível e "ocultar sem valor" vêm do projeto. |
| O Revit grava cada definição na hora; não há transação no arquivo. | Tudo é validado antes da primeira gravação. Se algo falhar no meio, a mensagem lista o que já foi gravado e o caminho da cópia. |
| Arquivo somente leitura, em rede sem permissão ou aberto por outro programa. | Conferido antes; nada é gravado e a mensagem explica. |
| Parâmetros só em famílias não têm categorias de projeto. | Mostrados com origem "Só em famílias", sem vínculo nem categorias. |

Vinculação:

| Risco | Tratamento |
|---|---|
| A API só abre o arquivo configurado em `Application.SharedParametersFilename` (configuração do usuário). | O caminho é trocado só durante a leitura e a gravação e restaurado em `finally`. Se o Revit for encerrado à força nesse intervalo, o caminho do usuário pode ficar apontando para o arquivo escolhido. |
| `BuiltInParameterGroup` foi removido no Revit 2024; grupos por `ForgeTypeId` só existem do 2022 em diante. | Detecção por recurso (`ParameterUtils.GetAllBuiltInGroups`). 2022/2023 usam `ForgeTypeId`. |
| `GetAllBuiltInGroups` devolve todos os grupos internos, não exatamente o subconjunto da janela nativa. | Lista os grupos com rótulo no idioma do Revit, em ordem alfabética. Grupos raros podem aparecer. |
| `BindingMap.ReInsert` é documentado como "remove e recria o vínculo". | Usado só para acrescentar categorias, com o mesmo tipo e grupo, valores conferidos e rollback se algo mudar. Teste numa cópia do modelo antes do uso em produção. |
| Categorias sem tipo (Ambientes, Áreas, Espaços, Informações do projeto, Folhas, Materiais) aceitam vínculo de Tipo pela API, mas o parâmetro não aparece. | Aviso na revisão; não bloqueia. |
| Arquivo editado fora do Revit (UTF-8/ANSI) pode ser lido com outro nome pelo Revit. | A revisão avisa quando o nome lido pelo Revit difere do TXT; vale o nome do Revit. |
| Tipo de dado de versão mais nova (ex.: especificação inexistente na versão em uso). | O Revit não lê a definição → linha indisponível com o motivo. |
| Modelo central: o Revit precisa conseguir emprestar os elementos dos parâmetros. | Se outro usuário estiver com eles, a transação falha e é desfeita, com a mensagem do Revit. |
| A janela não troca de tema sozinha se o tema do Revit mudar com ela aberta. | Limitação do `oca_ds`: reabrir aplica o tema. |

## Instalação

1. Copie `OCA.tab/Parâmetros.panel/` e `lib/oca_params/` para a extensão (mesmas posições).
2. Inclua o painel no layout da aba, em `OCA.tab/bundle.yaml` (o pyRevit **não cria** painéis
   fora do `layout`):
   ```yaml
   layout:
     - Sobre
     - Keynote
     - Vistas
     - Coordenadas
     - Visualizar
     - Parâmetros
   ```
3. pyRevit → **Reload**. O botão aparece no painel Parâmetros da aba OCA.

## Roteiro de teste (numa cópia do modelo)

1. Abrir o comando: o TXT configurado no Revit (ou o último usado) é carregado sozinho. Conferir
   contagem, codificação e linhas cinza (motivo em Observações).
2. Buscar com e sem acento, filtrar por grupo, tipo de dado e situação; "Só marcados".
3. Marcar vários (checkbox, ou Shift+clique e Espaço), "Configurar marcados".
4. Na Configuração: escolher grupo, Instância/Tipo e categorias; "Aplicar a todas"; depois
   selecionar uma linha, mudar para Tipo e "Aplicar às selecionadas"; "+ Adicionar parâmetros"
   e incluir mais um.
5. Revisão: conferir ações e observações; "Aplicar no projeto" → confirmar.
6. No Revit: Gerenciar → Parâmetros de projeto. Conferir nome,
   Instância/Tipo, categorias e grupo na paleta Propriedades. Ctrl+Z deve desfazer tudo de uma vez.
7. Rodar de novo com um parâmetro já vinculado: acrescentar uma categoria (Adicionar categorias)
   e confirmar que valores já preenchidos continuam iguais.
8. Pedir Tipo para um parâmetro que já está como Instância: deve aparecer "Não será aplicado".
9. Em Opções → Arquivo, conferir que o caminho do arquivo de parâmetros compartilhados do usuário
   voltou ao original.
10. Aba **Parâmetros no projeto**: conferir grupo, vínculo e categorias contra Gerenciar →
    Parâmetros de projeto; filtrar "Só em famílias" e "Fora do TXT".
11. Numa **cópia** do TXT: marcar dois parâmetros "Fora do TXT", exportar para um grupo novo.
    Conferir a cópia `_backup_` criada, o grupo novo no TXT, os GUIDs iguais aos do projeto e que
    a situação passou para "No TXT". Exportar de novo os mesmos não deve ser possível.
12. Repetir no tema escuro do Revit e em pelo menos duas versões (ex.: 2024 e 2025).
