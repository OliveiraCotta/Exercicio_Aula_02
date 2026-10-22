# Atualização das Descrições por Keynote — auditoria técnica (versão 2026.10.21)

> **Nada disto foi executado dentro do Revit.** Este ambiente é Linux, sem Revit, sem
> pyRevit e sem IronPython. A validação foi feita com:
> - um simulador da API do Revit, do WPF e do .NET escrito em Python (CPython 3.11);
> - testes automatizados do TXT e do backup;
> - análise estática, incluindo a gramática do Python 2.
>
> Tudo que depende do Revit real está marcado como **"testar no Revit"**. Nenhum teste
> de integração foi aprovado no Revit.

---

## 0e. Coluna "DESCRIÇÃO EM IN" do Excel → parâmetro "Descrição em IN" (revisão 2026.10.22c)

O valor da coluna **DESCRIÇÃO EM IN** da aba MATERIAIS vai para o parâmetro
**Descrição em IN**. A coluna é reconhecida sem diferenciar maiúsculas, acentos e
espaços repetidos, e também com texto depois do nome, como "DESCRIÇÃO EM IN (inglês)" ou
com quebra de linha no cabeçalho.

- Se a aba tiver **as duas** colunas ("DESCRIÇÃO EM IN" e a antiga "DESCRIÇÃO IN"), vale
  a **DESCRIÇÃO EM IN**. Antes valia a que estivesse mais à esquerda.
- Planilhas antigas, só com "DESCRIÇÃO IN", continuam funcionando.
- Sem nenhuma das duas, a leitura para com a mensagem "Colunas não encontradas na aba
  'MATERIAIS': DESCRIÇÃO EM IN".
- O log registra a coluna usada, por exemplo:
  `Descrição em IN = coluna G 'DESCRICAO EM IN'`.
- A mesma coluna alimenta o texto do TXT de keynotes ("DESCRIÇÃO | DESCRIÇÃO EM IN").
- Alterado: `HEADER_ALIASES["en"]`, `EN_HEADER_PREFIX` e a escolha da coluna em
  `read_excel`. Nada mais.

## 0d. Parâmetro renomeado para "Descrição em IN" (revisão 2026.10.22b)

Na janela, nas mensagens, no log e no dashboard, o parâmetro agora se chama
**"Descrição em IN"** (`PARAM_IN_NAME`). A lógica não mudou.

- **O comando encontra o parâmetro pelo GUID** `70307f6e-3b31-4049-bed4-62dfc3570fec`,
  não pelo nome. No `OCA_Parametros_Template.txt` do escritório (01/10), esse GUID está
  como `Descrição IN`, tipo `TEXT`.
  - Se o parâmetro foi **renomeado no arquivo de parâmetros compartilhados mantendo o
    GUID**, tudo funciona, mesmo em modelos que ainda mostram o nome antigo (testado no
    simulador).
  - Se foi **criado um parâmetro novo** "Descrição em IN", ele tem **outro GUID**. Nesse
    caso o comando não grava nele e mostra no topo: "Atenção: 'Descrição em IN' com GUID
    diferente do oficial … não é usado" (testado no simulador). Para usar o GUID novo,
    troque `PARAM_IN_GUID_STR` no `script.py`.
- Não mudaram: os cabeçalhos aceitos na planilha. A coluna pode se chamar "DESCRIÇÃO IN"
  ou "DESCRIÇÃO EM IN".

## 0c. Pasta de backup `<nome>_obsoletos` e botão Atualizar (revisão 2026.10.22)

**Pasta de backup.** O TXT anterior não vai mais para `Obsoletos`. Agora vai para uma
pasta com o nome do próprio TXT e o final `_obsoletos`, ao lado dele:

```
<pasta>\OCA_693_Keynote.txt
<pasta>\OCA_693_Keynote_obsoletos\OCA_693_Keynote_20261021_173527.txt
```

- O resto não muda: o backup é conferido antes da troca e nunca é apagado nem
  sobrescrito. Os arquivos `_REJEITADO` e `_FALHOU` vão para a mesma pasta.
- Uma pasta `Obsoletos` criada por versões anteriores **não é mexida**. Mova ou apague
  à mão se quiser.
- Funções alteradas: `obsolete_backup_path` (constante `OBSOLETE_SUFFIX`) e as mensagens
  (`_backup_label`).

**Botão "Atualizar"** (no canto direito da linha da busca):
- relê o **modelo** e refaz as 3 tabelas pelo mesmo caminho usado depois de gravar
  (`collect_model` → `analyze_model` → `reload`), via ExternalEvent e sem transação;
- não relê o Excel, não gera nem recarrega o TXT e não grava nada;
- mantém a busca e o "Agrupar por";
- as marcações voltam ao padrão, como depois de gravar;
- se houver Keynote digitada na tabela 2 e ainda não gravada, pede confirmação antes de
  descartar.

Funções novas: `VerifyWindow._refresh` e `refresh_run`. Nada mais foi alterado.

## 0b. TXT de Keynote em UTF-16 LE com BOM (revisão 2026.10.21)

**A instrução de ANSI / Windows-1252 foi substituída: o TXT agora é gravado em UTF-16 LE
com BOM (bytes iniciais `FF FE`).**

### O que a investigação encontrou

| Pergunta | Evidência | Resposta |
|----------|-----------|----------|
| O que o plugin gravava antes? | Histórico do git, `write_keynote_txt`, de `f6fb24f` (01/10) a `54c3a08` (20/10): `text.encode("cp1252")`; só se falhasse, `codecs.BOM_UTF16_LE + text.encode("utf-16-le")`. | **ANSI (cp1252) por padrão.** UTF-16 **LE com BOM** só quando havia um caractere fora do ANSI (o caso do print com "UTF-16"). Nunca gravou UTF-16 BE nem UTF-16 sem BOM. |
| E o `codecs.open(..., "utf-16")` do código antigo? | Esse trecho gravava o **arquivo de parâmetros compartilhados** temporário, não o TXT de keynotes (foi removido em `75cf66d`). | Não tem relação com o TXT de keynotes. |
| Em que codificação está o TXT de keynotes do escritório? | `OCA_XXX_Lista_de_Materiais_00.txt` (enviado em 01/10): 1.787 bytes, **sem BOM**, nenhum byte `00`, acentos em bytes únicos (`C1 C3 C7 C9 CD D3 D5`), inválido como UTF-8, 59 linhas CRLF. | **ANSI / Windows-1252, sem BOM.** |
| Qual arquivo do escritório é UTF-16? | `OCA_Parametros_Template.txt` começa com `FF FE`. | É o arquivo de **parâmetros compartilhados** do Revit (UTF-16 LE com BOM), não o de keynotes. |
| O que o Revit aceita para keynotes? | Ajuda do Revit 2024/2025/2027, "About Editing the Keynote File": ANSI, ou "Unicode" pelo Bloco de Notas (= UTF-16 LE com BOM). Artigo de suporte da Autodesk: para caracteres fora do ANSI, salvar como **UTF-16 LE**. | ANSI e UTF-16 LE com BOM. **UTF-16 BE e UTF-16 sem BOM não são citados pela Autodesk e não foram adotados.** |

**Conclusão:** a premissa de que o escritório e o plugin usavam UTF-16 não se confirma
nos arquivos disponíveis. Mesmo assim, UTF-16 LE com BOM é um formato **documentado pela
Autodesk** para keynotes e aceita qualquer caractere (≥, Ω, →, CJK etc.). Por isso adotei
esse formato, como pedido. A compatibilidade **não foi testada no Revit real** (ver H.2).
Se algum Revit da equipe não ler o arquivo, o backup ANSI anterior está em `Obsoletos`.

### O que mudou no código (`script.py`)

| Trecho | Antes | Agora |
|--------|-------|-------|
| `cp1252_encode`, `cp1252_decode`, `_CP1252_*`, `non_cp1252_report` | Codificador ANSI estrito | **Removidos.** No lugar: `txt_encode` / `txt_decode` (UTF-16 LE com BOM, montados byte a byte, sem depender de `codecs`, `StreamWriter` ou `System.Text.Encoding`) e `txt_char_report`. |
| `detect_txt_encoding` (novo) | — | Identifica, pelos bytes, um TXT existente: UTF-16 LE/BE com ou sem BOM, UTF-8 com ou sem BOM, ANSI ou ASCII. Usado no log, na mensagem e antes da recarga. |
| `validate_txt_bytes` | ANSI sem BOM | Exige BOM `FF FE`, tamanho par, nenhum par substituto quebrado, nenhum NUL, CRLF, 2 TAB por linha, nenhuma linha vazia, a contagem de linhas e **texto lido de volta = texto gerado**. |
| `write_keynote_txt` | Gravava cp1252 | Grava os bytes UTF-16 LE com BOM com `File.WriteAllBytes` (nenhuma API escolhe a codificação por nós). Confere o `.tmp`, faz o backup **byte a byte** (o anterior ANSI continua ANSI), troca e confere de novo os bytes **e** a codificação detectada no disco. Devolve `(backup, codificação gravada, codificação do anterior)`. |
| `load_keynote_txt` | — | Antes de recarregar, lê os bytes do disco: se não for UTF-16 LE com BOM, **não recarrega**. |
| `make_keynote_txt` | Bloqueava caracteres fora do ANSI | Bloqueia só o que não pode ir para o arquivo: controles (exceto TAB/CR/LF), metades soltas de par substituto, U+FEFF/U+FFFE/U+FFFF. A mensagem só é montada depois da validação. |
| `restore_txt_after_failed_reload`, backups | — | Sem mudança: cópias byte a byte, nunca convertidas. |
| Logs, HTML, Excel | UTF-8 / XML do Excel | **Sem mudança.** Só o TXT de keynotes é UTF-16. |

**Mensagem na janela** (só aparece se os bytes gravados conferirem):

`TXT de Keynotes gravado em UTF-16 LE, com BOM, e validado com sucesso · anterior (ANSI (Windows-1252)) em Obsoletos\<nome>_<data>.txt · recarregado no projeto e conferido (N de N entradas, iguais ao arquivo).`

**Linhas vazias e quebras de linha.** O TXT é gerado do Excel, uma linha por keynote,
como o do escritório (que não tem linhas vazias). Linhas vazias do Excel não viram linha
no TXT. Uma quebra de linha **dentro** de uma célula (CRLF, LF ou CR) e TAB viram espaço,
porque quebrariam o formato do Revit. Todas as linhas terminam em CRLF.

## 0a. pyRevit 6.x: a janela abre, mas os eventos não funcionam (revisão 2026.10.20b)

**Sintomas no pyRevit 6.5.5 + Revit 2025.2** (print de 20/10):
- trocar "Agrupar por" não muda a tabela;
- clicar numa linha não pré-seleciona nada (Propriedades continua em "3D View");
- digitar a Keynote na tabela 2 e clicar em "Verificar no Excel" não faz nada;
- nenhuma mensagem de erro aparece.

**Causa provável (hipótese forte, não reproduzida no Revit):** o motor Python não fica
vivo depois que o script termina.
- O botão declarava o motor persistente só pelo `__persistentengine__ = True` dentro do
  `script.py`.
- No pyRevit 6.x/7.x, o loader lê essa configuração do `bundle.yaml` do botão
  (`engine: persistent: true`). É assim que está no exemplo oficial de janela não modal
  do pyRevit. As notas do 7.0 dizem que o `bundle.yaml` é lido primeiro e que a leitura
  das variáveis do script ficou desligada por padrão.
- Sem a configuração, a janela abre (isso acontece enquanto o script ainda roda), mas o
  motor é descartado em seguida. A partir daí, **nenhum evento da janela roda** — nem o
  tratamento de erro, por isso nada aparece.

Os três sintomas têm essa mesma causa. O agrupamento, em si, não mudou.

**Correções:**
- `bundle.yaml` novo na pasta do botão, com `engine: persistent: true`, título e dica.
- **Autoteste** 2 s depois de a janela abrir:
  - confere se o motor Python continua vivo;
  - executa uma ação externa sem transação.
  - Se algo falhar, a linha verde mostra "ATENÇÃO: o pyRevit encerrou o motor…" ou "o
    Revit não executou o autoteste…". **O comando não falha mais em silêncio.**
- A seleção tem um vigia de 5 s: avisa se o Revit não executar.
- O log registra `bundle_persistent`, `engine_id`, `engine_ver` e `cached_engine`, e
  avisa se o `bundle.yaml` faltar.
- **Filtros por coluna removidos**, a pedido. Ficaram a busca geral (Buscar / Enter /
  Limpar busca) e o "Agrupar por".

**Depois de copiar os arquivos, recarregue o pyRevit (Reload)**, porque o `bundle.yaml` só
é lido no carregamento. Se o aviso continuar, feche e abra o Revit.

## 0. Leia primeiro: a cópia instalada estava desatualizada

O zip enviado (`…950c`) e o `main` do repositório estão no commit `75cf66d`. As correções
de 15/10 e 16/10 foram feitas em `ESCRITORIO.extension/…`, e o `main` apagou essa pasta
quando moveu o comando para `OCA01.extension/…`. **Por isso nenhuma correção posterior a
`75cf66d` chegou ao seu Revit.** Os prints confirmam: o subtítulo da janela não mostra
"versão …".

Esta versão está em `OCA01.extension/OCA.tab/Keynote.panel/02_Atualização Descrição por
Keynotes.pushbutton/`. A janela de verificação mostra **"versão 2026.10.22c"** no subtítulo,
e o log também. Se o texto não aparecer, o Revit está rodando outra cópia.

---

## A. Diagnóstico

**Status:**
- **Confirmado** = reproduzido no simulador ou demonstrável pela leitura do código e pelo
  comportamento documentado da API.
- **Hipótese** = plausível, mas só o Revit real, o log ou o journal confirmam.

| # | Problema | Local | Causa / evidência | Impacto | Versões | Correção | Validação | Status |
|---|----------|-------|-------------------|---------|---------|----------|-----------|--------|
| 1 | O Revit roda o script antigo | Instalação / repositório | O zip e o `main` estão em `75cf66d`. As correções foram para `ESCRITORIO.extension`, apagada no `main`. O subtítulo não mostra versão. | Nenhuma correção nova chega ao Revit, e o comando "continua sem funcionar". | Todas | Código movido para `OCA01.extension` (merge `54c3a08`). A versão aparece na janela e no log. | Ver "versão 2026.10.20" no subtítulo. | **Confirmado** |
| 2 | Excel aberto impede a leitura | `read_excel` (antes `ZipFile.OpenRead`) | `OpenRead` só permite que outros processos **leiam** o arquivo. Como o Excel o mantém aberto para gravação, a abertura falha ("being used by another process"). | O comando para na etapa 1. | Todas (.NET 4.8 e .NET 8) | `open_xlsx`: `FileStream` com `FileShare.ReadWrite \| Delete` e `ZipArchive` só leitura. Lê o que está **salvo**. Mensagens claras para arquivo inexistente, não-.xlsx, bloqueado ou corrompido. | Simulador ok. **Testar no Revit com a planilha aberta no Excel.** | Confirmado pelo código e pela documentação do .NET; não reproduzido no Windows |
| 3 | `unicode` / `unichr` não existem em IronPython 3.4 / CPython | Topo do script, `to_unicode`, cp1252 | Nomes que só existem no Python 2. | Com o motor padrão do pyRevit trocado para IronPython 3.4, o comando falha **antes** de abrir qualquer janela (`NameError`). | pyRevit 5.2 a 7.0 com IronPython 3.4 selecionado | Equivalentes definidos no início; `to_unicode` trata `bytes`. | O script inteiro roda em CPython 3.11 no simulador. A gramática do Python 2 aceita o arquivo. | **Confirmado** |
| 4 | *(substituído na 2026.10.21 — ver 0b)* TXT gravado em UTF-16 quando há caractere fora do Windows-1252 | `write_keynote_txt` (versão anterior) | Troca automática para UTF-16 com "≥" e símbolos parecidos. Já tinha acontecido: o print mostrava "UTF-16". | TXT fora do padrão ANSI do escritório. | Todas | O TXT **não é gravado**. O comando lista caractere, código, keynote, linha do Excel e sugestão (≥ → ">="). | 40 testes automáticos + simulador | **Confirmado** |
| 5 | Backups misturados na pasta do TXT | `backup_path` (removida) | Os backups ficavam ao lado do TXT, com nome parecido. | Pasta poluída e risco de carregar um backup por engano. | Todas | Pasta `Obsoletos` (seção C.2). O backup é conferido byte a byte **antes** de trocar o TXT. | Testes 24–34 | **Confirmado** |
| 6 | Executor pelo evento `Idling` (introduzido em `20a44c1`) | Fila de ações | Um handler de `Idling` inscrito pelo script continua no Revit se o pyRevit recarregar com a janela aberta. Ele aponta para um motor descartado. | Possível encerramento nativo do Revit. | Todas; o risco é maior no pyRevit 7, que compartilha o motor | **Removido.** Só `ExternalEvent`. O resultado do `Raise` é registrado, e um pedido negado sai da fila. Um vigia de 8 s avisa na janela. | Regressão (cenário `deny`) | Hipótese de risco; removido por precaução |
| 7 | Seleção percorria a vista inteira a cada clique | `show_in_view` | `FilteredElementCollector(doc, view.Id)` em toda seleção de linha. | Lentidão em vistas grandes; parece travamento. | Todas | `VisibleInViewFilter` aplicado **só** aos IDs da linha. | Simulador (visível e não visível) | **Confirmado** |
| 8 | O cabeçalho das colunas sumia ao rolar | XAML das três tabelas | Quem rolava era a página, não a tabela, então o cabeçalho saía da tela. | Difícil ler tabelas longas. | Todas | Cada tabela rola por dentro, com a altura ajustada à janela, e o cabeçalho fica fixo. A roda do mouse passa para a página no fim da tabela. | Simulador (eventos). **Testar no Revit (visual).** | Confirmado pelo código |
| 9 | Recarga com erro de sintaxe era confirmada | `load_keynote_txt` | Os erros de sintaxe eram lidos **depois** do `Commit`. | O projeto podia ficar com uma tabela parcial e o disco com o TXT novo. | Todas | Erro de sintaxe → exceção **dentro** da transação → `RollBack`. O TXT anterior é restaurado do backup e o recusado vai para `Obsoletos\…_REJEITADO.txt`. | Simulador (`--syntaxerr` e `--txfail=Carregar`) | **Confirmado** |
| 10 | O resumo final não separava os resultados | `save_run` | Bloqueios e erros apareciam numa lista só; ignorados não eram contados. | O usuário não sabia o que foi recusado e por quê. | Todas | Concluído / Ignorados / Bloqueados / Erros, com até 10 itens de cada e o restante no log. | Simulador (`refuse`, `txfail`) | **Confirmado** |
| 11 | Uma cópia de backup que falhou na conferência parecia válida | `write_keynote_txt` | Uma cópia corrompida ficava em `Obsoletos` com o nome normal. | Restaurar dessa cópia poderia gravar um TXT ruim. | Todas | A cópia é renomeada para `_FALHOU.txt`. Nada é apagado. | Teste 32 | **Confirmado** |
| 12 | `OCA01.extension/startup.py` (fora deste comando) | `startup.py` | `ThemeChanged` nunca é removido. Variáveis globais (`DEBUG_TXT`, `cores_paineis`) ficam no motor; o **pyRevit 7 compartilha um motor entre os scripts de inicialização de todas as extensões**. | Depois de recarregar o pyRevit, um handler antigo pode rodar ao trocar o tema: risco de erro ou encerramento. | pyRevit 7 (motor compartilhado); todas (handler) | **Não alterado** (não faz parte do pedido). Recomendação na seção G. | — | Hipótese |
| 13 | "Unrecoverable error" com a janela aberta | Interface WPF | Ver H1 a H3 da revisão anterior. Desde `75cf66d`, os eventos da janela passam por `_safe` e há uma rede de segurança no Dispatcher. | Encerramento do Revit | 2025 e 2026 (.NET 8) | Mantido. Diagnóstico **H** (títulos de grupo simples) isola o template da linha de grupo. | **Só no Revit** | Hipótese |
| 14 | "Não grava no 2026.4" | `save_run` | O caminho `Set` → `Commit` → releitura não tem erro de lógica no simulador. A explicação mais provável é o item 1 (código antigo). | — | 2026.4 | Diagnóstico **I** (Shift+Clique) grava, confere e **desfaz**. Mostra se o Revit aceita a gravação. | **Só no Revit** | Hipótese (causa provável: item 1) |

**Um `try/except` não impede todo encerramento do Revit.** Erros nativos não passam
pelo Python: driver de vídeo, falta de memória, falha interna do Revit ou do WPF fora dos
eventos tratados. Para esses casos existe o log, gravado linha a linha (a última linha
mostra onde o comando estava), e o journal do Revit (seção H.3).

---

## B. Correções por prioridade

1. **Instalar a cópia certa** (item 1). Sem isso, nada do resto chega ao Revit.
2. **Leitura do Excel aberto** (item 2): falhava na etapa 1.
3. **Compatibilidade de motor** (item 3): falha total se o motor for trocado.
4. **Segurança do Revit**: retirada do executor `Idling` (item 6). Desfazer a recarga do
   TXT com erro (item 9).
5. **TXT ANSI estrito e backup conferido** (itens 4, 5 e 11).
6. **Relatório final e regras de negócio** (item 10 e seção D).
7. **Interface**: cabeçalho fixo e seleção (itens 7 e 8, seção C.3). Os filtros por coluna
   entraram e foram retirados na 2026.10.20b, a pedido.
8. **Log**: ambiente completo (pyRevit, motor, .NET, documento, workshared). Arquivo, backup,
   recarga, transação e `ExternalEvent` registrados.

---

## C. O que mudou no script (`script.py`, versão 2026.10.20)

### C.1 TXT em ANSI / Windows-1252 — *substituído pela seção 0b (UTF-16 LE com BOM)*

- O codificador próprio (`cp1252_encode` / `cp1252_decode`) não depende do codec do
  motor Python, então se comporta igual no IronPython 2.7, no IronPython 3.4 e no CPython.
  - Caracteres aceitos: 0x20–0x7E, 0xA0–0xFF e os 27 caracteres próprios do Windows-1252
    em 0x80–0x9F, como € “ ” ‘ ’ – — … • ™ ‰ Œ œ Š š Ž ž Ÿ ƒ.
  - TAB, CR e LF só aparecem como separadores.
  - **Nada vira "?".**
- **Antes de tocar em qualquer arquivo**, `non_cp1252_report` lista cada caractere fora do
  Windows-1252 (caractere, `U+XXXX`, keynote, linha do Excel e sugestão de troca).
  Caracteres de controle vindos do Excel também entram na lista.
  - Se a lista não estiver vazia, o TXT **não é gravado** e o anterior continua no disco e
    no projeto.
- Depois de gravar, `validate_txt_bytes` confere o arquivo:
  - sem BOM;
  - sem os bytes indefinidos 81, 8D, 8F, 90 e 9D;
  - quebras de linha CRLF, sem CR solto;
  - nenhuma linha vazia;
  - 2 tabulações por linha (chave, texto, pai);
  - a quantidade de linhas.
- Depois da recarga, o comando confere a tabela de keynotes do projeto contra o arquivo:
  mesmas chaves, textos e pais. As diferenças vão para o log, e a mensagem mostra
  "N de M entradas, iguais ao arquivo".
- O conteúdo não muda: o TXT do escritório volta idêntico, byte a byte, depois de lido e
  regravado.

### C.2 Backup em `Obsoletos`

**A estrutura adotada:**

```
<pasta do TXT>\OCA_XXX_Lista de Materiais_XX.txt          ← sempre o vigente
<pasta do TXT>\Obsoletos\OCA_XXX_Lista de Materiais_XX_20261020_120000.txt
<pasta do TXT>\Obsoletos\OCA_XXX_Lista de Materiais_XX_20261020_120000_2.txt   (mesmo segundo)
```

O exemplo do pedido (`Backups/OCA_XXX_…_XX_Obsoletos_20261020_120000.txt`) misturava
"Backups" (pasta) e "Obsoletos" (no nome). Usei **uma** pasta, `Obsoletos`, como no texto
do pedido, e mantive no arquivo o nome original mais a data e a hora. Assim o backup
ordena por data e o nome da lista continua igual. Se o escritório preferir o outro
formato, a troca fica em `OBSOLETE_DIR` e `obsolete_backup_path`.

**A ordem de `write_keynote_txt`:**
1. Codifica. Se houver caractere inválido, para sem tocar no disco.
2. Grava `<TXT>.tmp` e confere os bytes e a estrutura.
3. Se o TXT existe:
   - cria `Obsoletos`;
   - copia o TXT para lá **sem sobrescrever** (`File.Copy(…, False)`);
   - confere a cópia byte a byte.
   - **Se qualquer parte falhar, o TXT original não é trocado.** Uma cópia ruim vira
     `_FALHOU.txt`.
4. Troca o TXT e confere os bytes finais. Se não conferir, restaura o backup.
5. Apaga só o `.tmp`. **Nenhum backup é apagado ou sobrescrito.**

Se o Revit **recusar** o TXT novo (erro de sintaxe ou transação desfeita):
- a recarga é desfeita e o projeto continua com a tabela anterior;
- o TXT anterior volta ao disco a partir do backup;
- o recusado fica em `Obsoletos\…_REJEITADO.txt`.

O disco e o projeto voltam a ficar iguais.

### C.3 Interface

- **Cabeçalho fixo:** cada tabela rola por dentro, com a altura ajustada à janela
  (`_fit_tables`). A roda do mouse rola a tabela e, no fim dela, passa para a página
  (`_wheel`). Não há duas barras disputando a mesma roda.
- Mantidos: seções recolhíveis, colunas ajustáveis, visual OCA DS e virtualização das
  linhas.
- **Busca:** a busca geral continua (Buscar / Enter / Limpar busca). Os filtros por
  coluna foram retirados na 2026.10.20b.
- **Seleção:**
  - só seleciona se o documento ativo for o do comando;
  - ignora elementos apagados;
  - `SetElementIds` sem transação, rodando no contexto da API (`ExternalEvent`), e a janela
    continua aberta;
  - enquadra só os elementos visíveis na vista ativa;
  - **se nenhum estiver visível, a pré-seleção continua** e a mensagem orienta a abrir
    outra vista.
- **Preservados:**
  - Voltar, Finalizar, Dashboard HTML e agrupamento;
  - Shift+Clique (modo de diagnóstico), oca_ds e log;
  - leitura do Excel sem Excel instalado.

### C.4 Log e ambiente

O log grava as seguintes informações:
- Revit: ano, atualização, build e família de build;
- pyRevit: `versionmgr`, com o `sessioninfo` como alternativa no 6/7;
- motor Python e versão (`platform.python_implementation`), Python 2 ou 3, e motor
  persistente;
- .NET (`Environment.Version` e `RuntimeInformation`), sistema operacional, documento,
  caminho e workshared.

Cada etapa registra início, fim e duração. Também entram no log:
- operações de arquivo: `.tmp`, backup conferido, troca, `_FALHOU` e `_REJEITADO`;
- recarga e conferência da tabela;
- `Start`, `Commit` e `RollBack` das transações;
- `ExternalEvent.Raise` e o seu resultado;
- tracebacks.

**Uma falha ao gravar o próprio log nunca interrompe o comando.**

---

## D. Regras de negócio auditadas

| Regra | Onde | Resultado |
|-------|------|-----------|
| Célula vazia no Excel nunca apaga o modelo | `plan_to_model`, `save_run` (tabela 2: `if not new … continue`) | Ok. A tabela mostra "(mantém)". |
| A Description é atualizada mesmo sem Descrição em IN | `plan_to_model`, `save_run` | Ok. No cenário `noin`: Description 4, IN 0, sem erro. |
| A IN é identificada **pelo GUID** `70307f6e-3b31-4049-bed4-62dfc3570fec` | `in_param` (`get_Parameter(Guid)`) | Ok. |
| Parâmetro com o mesmo nome e outro GUID nunca é gravado | `in_param_clash` | Ok. Aviso no topo e no log ("Generic Models … ignorado"); nenhum valor gravado nele. |
| Somente leitura e emprestado (workshared) | Tabela 1: o plano marca o bloqueio, e a linha fica desabilitada com o motivo. Tabela 2: `IsReadOnly` / `owned_by_other` na gravação. | Ok. Vão para **Bloqueados** no resumo. |
| As marcações correspondem ao que pode ser gravado | Tabela 1: linhas bloqueadas não podem ser marcadas. Tabela 2: só grava o que mudou. | Ok. Limite: na tabela 2, somente leitura e workshared só são detectados **na gravação**, e a linha continua marcável. Fica no resumo como bloqueio; não é erro silencioso. |
| A confirmação lista os campos | `_apply` | Ok: "Description N · Descrição em IN N · Keynote N" e quantas linhas substituem valor. |
| O relatório final separa os resultados | `save_run` | Concluído / Ignorados / Bloqueados / Erros. |
| As tabelas atualizam sem duplicar | `reload` + reanálise | Ok. O simulador confere "duplicadas: 0" em todos os modos. |
| Nada é gravado sem confirmação | `_apply` (MessageBox Sim/Não) | Ok. "Não" → `sets: 0`. |

**Atenção, decisão do usuário:** o TXT é gravado e recarregado quando você marca a opção
de TXT na primeira aba e clica em Executar, **sem** uma segunda pergunta. Foi um pedido
anterior seu ("sempre substituir … não precisa aparecer esse aviso"). A recarga é uma
transação no projeto e pode ser desfeita com Ctrl+Z. Se quiser uma confirmação também
para o TXT, é uma mudança pequena.

### D.1 Parâmetros adicionais do projeto (avaliado, **não implementado**)

Gravar outros parâmetros vindos do Excel é viável com o mesmo mecanismo:
- identificação por GUID;
- a mesma confirmação;
- conferência depois de gravar.

Os riscos:
- parâmetros de projeto sem GUID: dá para identificar pelo nome, mas é ambíguo;
- tipos de dado diferentes de texto: número, unidade, sim/não;
- parâmetros de instância: muito mais elementos, e mais emprestados em workshared.

Recomendação: só com uma lista fechada de GUIDs aprovada pelo escritório.

---

## E. Transações e recuperação

O fluxo é este: validação → prévia (tabelas) → confirmação → transação → conferência →
relatório.

| Passo | Atômico? | Se falhar |
|-------|----------|-----------|
| Grava o TXT | Sim, por troca conferida | O TXT original fica (C.2). |
| Recarga do TXT (transação "Carregar TXT de Keynotes") | Sim, uma transação | `RollBack` + TXT anterior restaurado (C.2). |
| Tabela 1 (transação "Atualizar descrições por Keynote") | Sim, uma transação | Tudo desfeito; vai para **Erros**. |
| Tabela 2 (transação "Completar Keynote / Descrição dos tipos") | Sim, uma transação | Tudo desfeito; vai para **Erros**. |

**O TXT e o modelo não formam uma operação atômica.** O arquivo fica fora do Revit, e
não há como prendê-lo à mesma transação.

Cada passo é conferido antes do próximo:
- a recarga só acontece com o TXT conferido;
- as gravações no modelo só acontecem depois da sua confirmação, e são relidas depois do
  `Commit`.

Se um passo falhar, o comando faz o seguinte:
- não continua com base nele;
- desfaz o que pode (transação e TXT restaurado);
- diz no resumo e no log o que ficou como estava.

Cada transação aparece separada no Ctrl+Z do Revit.

---

## F. Compatibilidade

### F.1 O que o script usa

| Recurso | Revit 2024 (.NET 4.8) | 2025 (.NET 8) | 2026 (.NET 8) | Observação |
|---------|------|------|------|-----------|
| `FileStream` + `ZipArchive` (`System.IO.Compression`) | ok | ok | ok | O mesmo assembly nos dois .NET. |
| `File.WriteAllBytes/ReadAllBytes/Copy/Move`, `Directory.CreateDirectory` | ok | ok | ok | — |
| `System.Array[System.Byte](lista)` | ok | ok | ok | Conversão do IronPython. |
| `KeynoteTable.LoadFrom`, `KeyBasedTreeEntriesLoadResults` | ok | ok | ok | — |
| `Transaction` + `IFailuresPreprocessor` | ok | ok | ok | — |
| `ExternalEvent` / `IExternalEventHandler` | ok | ok | ok | Executor único (sem `Idling`). |
| `VisibleInViewFilter`, `SetElementIds`, `ShowElements` | ok | ok | ok | — |
| `ElementId.Value` | ok | ok | ok | `IntegerValue` só como alternativa. |
| WPF DataGrid (rolagem interna, virtualização) | ok | **testar** | **testar** | Diferenças do .NET 8 só aparecem no Revit. |

### F.2 pyRevit × motor

| pyRevit | Motor padrão | Situação do script | Observações das notas de versão |
|---------|--------------|--------------------|---------------------------------|
| 5.2.0.25181 | IronPython 2.7.12 | Alvo principal | — |
| 6.4.0.26100 | IronPython 2.7.12 (3.4.2 opcional) | Compatível pelo código; não executado | Há relatos de recursos do pyRevit falhando com IronPython 3.4.2 como motor principal (issue #3087). |
| 6.5.5.26237 | IronPython 2.7.12 | Compatível pelo código; não executado | O 6.5 cria um motor novo para cada script de inicialização. |
| 7.0.0.26278 | IronPython 2.7.12 ou 3.4.2 | Compatível pelo código; não executado | Mudanças que afetam este comando: um único motor para preload, startup e postload; saída em WebView2; `to_net_list` removido (não é usado aqui); Revit < 2021 não é mais suportado. |
| Qualquer versão, com **CPython** | — | **Não suportado** | O comando herda classes .NET (`IExternalEventHandler`, `IFailuresPreprocessor`, `WPFWindow`), e isso exige IronPython. O script não tem `#! python3`, então roda no IronPython. |

**A compatibilidade com um ano do Revit não garante a das atualizações dele.** O log
grava o build exato. Teste em cada combinação que a equipe usa (seção H).

---

## G. `startup.py` da extensão (não alterado)

Riscos:
- `__revit__.ThemeChanged += ao_mudar_tema` **nunca é removido**. Depois de recarregar o
  pyRevit, o handler antigo continua inscrito. No pyRevit 7, com o motor compartilhado,
  também ficam as variáveis globais `DEBUG_TXT` e `cores_paineis`.
- `UIThemeManager` e `ThemeChanged` só existem no Revit 2024 em diante. Para 2024–2026,
  tudo bem.

Recomendações:
- proteger com `try/except` dentro de `ao_mudar_tema`;
- guardar o handler num atributo único, por exemplo em `AppDomain.CurrentDomain.SetData`,
  para remover o anterior antes de inscrever o novo;
- usar nomes com prefixo (`_oca_…`).

Se quiser, faço isso num commit separado.

---

## H. Testes

### H.1 Executados aqui (simulador + testes automatizados), todos aprovados

**Regressão** (`regress.sh`): 21 cenários, sem traceback inesperado.

| Cenário | O que confere |
|---------|---------------|
| `full` | Fluxo completo. Agrupamentos, busca e Limpar. Seleção visível e **não visível** (pré-seleção mantida). Dashboard. Gravação recusada (`sets 0`) e confirmada: 11 gravados e conferidos. Finalizar. |
| `noin` | Modelo sem Descrição em IN: Description grava, IN = "Não aplicável". |
| `typeonly` | Keynote digitada na tabela 2 e gravada sem "Verificar". |
| `back` | Voltar e reabrir. |
| `txfail1` / `txfail2` | O Revit desfaz a transação da tabela 1 ou da 2. Aparece em **Erros**, e a outra tabela grava. |
| `refuse` | `Set()` recusado na Description: 4 erros listados. |
| `deny` | `ExternalEvent` negado: nada gravado, aviso na janela. |
| `uiexc` | Exceção da interface tratada. O único traceback é proposital, registrado no log. |
| `diag_A` … `diag_I`, `diag_L` | Todos os modos de diagnóstico. |
| `realds_repo` / `realds_old` | Com a `oca_ds` real do repositório e com a cópia antiga (sem `min_width`). |

**Autoteste do motor (2026.10.20b):**
- `--probe`: o motor está vivo e o ExternalEvent é executado (log "Autoteste").
- `--dead`: as variáveis do script são apagadas depois que a janela abre, simulando o
  pyRevit descartando o motor. A linha verde mostra o "ATENÇÃO…".

**Extras do TXT:**
- `--syntaxerr` / `--txfail=Carregar`: recarga desfeita, TXT anterior restaurado,
  recusado em `_REJEITADO`.
- Excel com "≥": gravado em UTF-16 LE com BOM, com "≥" no arquivo, e recarregado (2026.10.21).
- TXT anterior em ANSI: backup idêntico ao original; o novo é UTF-16 LE com BOM e tem o mesmo texto.

**`tests/keynotes/test_txt16.py` (2026.10.21): 84 de 84.** Rodar na raiz do repositório:
`python3 tests/keynotes/test_txt16.py`. Não precisa de Revit; usa `fake_dotnet.py` e o TXT
real do escritório em `fixtures/`. Bytes comparados com o codec `utf-16-le` do
CPython, que é independente do código testado:
- acentos e cedilha (á ã é í ó ú ç, maiúsculas): bytes, ida e volta; `ç` = `FF FE E7 00`;
- símbolos (€ “ ” – — … ™ ® © § ° ± × ÷ ² ³ µ ½ ‰ ≥ ≤ ≈ ≠ Ω → ∅ ✓), CJK e emoji (par
  substituto `D83D DE00`);
- 9 não representáveis bloqueados (NUL, BEL, C0, DEL, C1, U+FEFF, U+FFFE e metades soltas),
  e o relatório com código, keynote, origem e dica;
- quebras dentro da célula viram espaço; LF, CR solto, linha vazia, TAB a mais, contagem e
  texto diferente detectados;
- detecção de 8 codificações (UTF-16 LE/BE com e sem BOM, UTF-8 com e sem BOM, ANSI,
  ASCII), e a validação recusa todas menos UTF-16 LE com BOM. Arquivo truncado (ímpar), NUL
  e par quebrado no arquivo detectados;
- TXT anterior em cada uma das 7 outras codificações: backup **idêntico byte a byte** e o
  novo em UTF-16 LE com BOM;
- falhas (cópia do backup, backup corrompido, criar pasta, `.tmp`, **arquivo bloqueado**,
  pasta inexistente, **nome inválido no Windows**): erro claro e original intacto;
- 1 linha e 20.000 linhas (com emoji): bytes iguais à referência;
- TXT real do escritório: confirmado ANSI; convertido, ida e volta idêntica e estrutura ok.

**`tests/keynotes/test_py2path.py`: 3 de 3.** Simula o IronPython (texto UTF-16 por dentro): par
substituto vira os bytes certos, a leitura de volta é igual e a metade solta é bloqueada.

**`test_txt.py` (ANSI, 2026.10.20): 40 de 40 — histórico, não vale mais.** Cobria:
- acentos, ç, € e aspas tipográficas, iguais ao codec cp1252;
- ida e volta;
- 8 caracteres não representáveis e 5 controles bloqueados, sem "?";
- relatório com código, keynote, origem e sugestão;
- BOM, LF, linha vazia, tabulação a mais, contagem de linhas e byte indefinido detectados;
- 1ª gravação sem backup;
- backup conferido;
- nome `_2` no mesmo segundo, com o backup anterior intacto;
- falhas: cópia, cópia corrompida (`_FALHOU`), criar pasta, `.tmp` e pasta inexistente,
  com o original sempre mantido;
- arquivo grande: 20.000 linhas, 1,2 MB;
- TXT do escritório idêntico byte a byte.

**Análise estática:**
- compila no Python 3;
- a gramática do Python 2 aceita o arquivo;
- nenhum recurso exclusivo do Python 2 ou do 3 (`iteritems`, `has_key`, f-strings,
  `nonlocal` etc.).

Limites desses testes:
- O simulador **imita** a API; não é o Revit.
- O WPF não é renderizado, então nada visual foi testado.
- O IronPython não foi executado. O desempenho com 20.000 linhas foi medido no CPython
  (0,4 s); no IronPython deve ser algumas vezes mais lento.

### H.2 Dependem do Revit real (pendentes)

Use **cópias** dos modelos. A ordem sugerida em cada combinação de pyRevit e Revit
(2024 / 2025 / 2026 e as atualizações usadas):

1. **Shift+Clique → I** (1 minuto): grava, confere e desfaz a Description. Diz se o Revit
   aceita a gravação.
2. **Execução normal com a planilha ABERTA no Excel**: o comando deve ler sem erro.
3. **TXT**:
   - abra o TXT no Bloco de Notas → Salvar como: a codificação deve aparecer como
     **"UTF-16 LE"** (ou "Unicode"); confira a pasta `Obsoletos`;
   - coloque "≥", "Ω" e "ç" numa descrição: devem aparecer certos nas tags de keynote e
     em Anotar > Keynote > Configurações de Keynote > Exibir;
   - teste em **cada** versão do Revit usada (2024, 2025, 2026): a leitura de UTF-16 LE
     com BOM é documentada pela Autodesk, mas não foi testada aqui;
   - deixe o TXT somente leitura e confira se aparece um erro claro e o original continua.
4. **Interface**:
   - role tabelas longas e confira se o cabeçalho fica fixo e a roda do mouse não briga
     entre a tabela e a página;
   - teste a busca e Limpar busca;
   - redimensione colunas e recolha seções.
5. **Seleção**:
   - linha com elemento visível: seleciona e enquadra;
   - linha com elemento oculto: a pré-seleção continua (confira no Revit);
   - com outro projeto ativo: aviso, sem selecionar.
6. **Gravação**:
   - marque linhas → "Não" (nada muda) → "Sim" (resumo com os 4 grupos);
   - confira nas Propriedades do tipo;
   - Ctrl+Z.
7. **Workshared**: um tipo emprestado por outro usuário deve aparecer em **Bloqueados**.
8. **Modelo sem Descrição em IN e com IN de outro GUID**.
9. **Shift+Clique → H**: se a versão normal fecha o Revit e a H não, a causa é a linha de
   título do DataGrid.

Anote em cada teste: a versão (do log), o teste, o resultado e, se o Revit fechar, a
última linha do log.

### H.3 Se o Revit fechar

1. **Não abra o Revit de novo antes de copiar os arquivos.**
2. Copie o log do dia (seção I).
3. Copie o `journal.XXXX.txt` mais recente de
   `%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 20XX\Journals`, mais algum `.dmp`, se
   houver.
4. No journal, procure perto do fim por `ExternalEvent`, `API_ERROR`, `fatal error` e
   `unrecoverable`.

---

## I. Instalação, teste seguro e logs

**Como substituir o script:**
1. Feche a janela do comando no Revit.
2. Substitua `script.py`, `bundle.yaml` (novo), `keynotes.html` e este LEIA-ME em
   `…\OCA01.extension\OCA.tab\Keynote.panel\02_Atualização Descrição por Keynotes.pushbutton\`.
3. Apague a pasta antiga `ESCRITORIO.extension\…\02_Atualização Descrição por
   Keynotes.pushbutton` se ela ainda existir **e** estiver registrada no pyRevit. Duas
   cópias do botão causam exatamente o problema do item 1.
4. pyRevit → **Reload**.
5. Abra o comando e confira **"versão 2026.10.22c"** no subtítulo. Espere 3 s: não pode
   aparecer nenhum "ATENÇÃO" na linha verde.

**Como testar com segurança:**
- use uma cópia do .rvt;
- use uma cópia do TXT, ou confira a pasta `Obsoletos`;
- comece pelo diagnóstico **I** (não altera o modelo);
- só depois grave poucas linhas.

**Logs:**
- arquivo: `%LOCALAPPDATA%\OCA\pyRevit\logs\Keynotes\keynotes_AAAAMMDD.log`;
- atalho: Shift+Clique → "Abrir a pasta dos logs".

Cada linha tem hora com milissegundos, o identificador da execução e o nível (INFO, STEP,
WARN ou ERROR). As linhas `>>` e `<<` marcam o início e o fim de cada etapa.
