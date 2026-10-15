# Atualização das Descrições por Keynote — diagnóstico, alterações e testes

Revisão de 15/10/2026. **Nada disto foi executado dentro do Revit.** A validação foi feita
com compilação do script, análise estática e um simulador da API do Revit/WPF em Python.
O que depende do Revit real está marcado como "testar no Revit".

---

## 1. Diagnóstico

### 1.1 Evidências (fatos observados nos prints e no código)

| # | Evidência | O que indica |
|---|-----------|--------------|
| E1 | A janela de verificação abriu com as três tabelas cheias e a linha verde dizendo "TXT … carregado no projeto". | As etapas 1 a 10 terminaram: leitura do Excel, TXT, recarga, coleta, análise e janela. O encerramento aconteceu **depois**, com a janela aberta (interação ou ociosidade), e não durante a recarga. Não dá para descartar um efeito tardio da recarga, como o redesenho das vistas com as novas keynotes. |
| E2 | Mensagem do Revit 2026: "An unrecoverable error has occurred". Em outro modelo: "Revit has experienced a fatal error". | No Revit, esse texto costuma aparecer quando uma exceção .NET não tratada chega à fila de eventos da interface (o *Dispatcher*), e janelas não modais de add-ins são a origem típica. |
| E3 | Código anterior: **nenhum** evento da janela não modal tinha `try/except` (cliques, seleção de linha, edição de célula, combos, rolagem). | Qualquer erro Python num desses eventos subia direto para o Revit. Essa é a causa mais provável do encerramento (H1). |
| E4 | Código anterior: `Parameter.Set()` devolve `False` quando o Revit recusa o valor, e o retorno era ignorado. `Transaction.Commit()` devolve `RolledBack` quando há erro, e o status também era ignorado. | Explica o "às vezes funciona mas não grava nada": a janela informava sucesso sem que o valor tivesse ficado no modelo. |
| E5 | Código anterior: o resultado de `ExternalEvent.Raise()` era ignorado. | Se o Revit estiver ocupado (edição de esboço ou outro comando), o pedido é negado. O botão ficava travado em "Gravando no modelo..." e nada era gravado. |
| E6 | Código anterior: `show_in_view` usava `revit.uidoc` (o documento **ativo**) com IDs do documento do comando. | Se o usuário trocasse de projeto com a janela aberta, a seleção ia para o documento errado ou dava erro. |
| E7 | Print 1: muitos tipos com "(sem parâmetro no tipo)". | O modelo não tem "Descrição IN" em várias categorias. Na versão anterior isso gerava avisos, bloqueios e itens "não gravados". |
| E8 | TXT gravado em UTF-16, com avisos "PF 24 / PF 25 repetido no Excel". | Há caracteres fora do cp1252 (por exemplo "≥") e chaves repetidas na planilha. O TXT anterior não validava a árvore: chave repetida, pai inexistente e ciclo. |

### 1.2 Hipóteses (precisam do log ou do journal para confirmar)

| # | Hipótese | Probabilidade | O que foi feito |
|---|----------|---------------|-----------------|
| H1 | Exceção Python num evento da janela, que chega ao Revit e encerra o programa. | Alta | Todos os eventos passam por `_safe` (registra no log e mostra na linha verde). Também há uma rede de segurança no *Dispatcher*. |
| H2 | Exceção interna do DataGrid do WPF (.NET 8 no 2025/2026) nas linhas de título de grupo, que trocam o template da linha. | Média | Essas linhas agora são "invisíveis" ao mouse e ao teclado (`IsHitTestVisible=False`). A rede de segurança marca a exceção como tratada quando a janela do comando está em uso. O modo de diagnóstico **H** usa títulos de grupo sem troca de template, para isolar. |
| H3 | Árvore de keynotes inválida (ciclo, pai inexistente ou chave repetida), que faz o Revit falhar ao redesenhar as tags depois da recarga. | Média-baixa | O TXT é validado e corrigido antes de gravar, com avisos no log. A recarga só é dada como feita depois do commit e da conferência das entradas. Use os testes **B** e **C** para isolar. |
| H4 | `ShowElements` com a vista ativa sendo folha, tabela ou outro documento. | Baixa | O comando confere o documento ativo e o tipo de vista antes de selecionar e enquadrar. |
| H5 | Varredura de cores antigas pesada em modelos grandes. | Não é a causa do encerramento: ela roda **antes** da janela (E1). | Foi otimizada mesmo assim (ver 2.3). |

**Conclusão:** a correção cobre H1 a H4 e o log passa a mostrar a última etapa e a
última ação antes de qualquer encerramento. Se o Revit ainda fechar, o log e o journal
apontam se foi um evento da janela, uma ação do ExternalEvent ou algo do próprio Revit
(seção 6).

---

## 2. Funções corrigidas e alterações

### 2.1 `Descrição IN` opcional (requisito funcional)

- `plan_to_model`: o campo IN é ignorado quando o parâmetro não existe no tipo ou tem
  outro GUID. Não gera bloqueio, e a Description segue normalmente.
- `type_status`, `verify_incomplete`, `classify_type`: a ausência do parâmetro não conta
  como falta, aviso, divergência nem "não verificado".
- `row_view`, `_vals2`, `_vals3`, `_cmp2`: as tabelas mostram **"Não aplicável"**,
  diferente de "(vazio)", que é campo existente sem valor.
- `save_run`: nunca tenta gravar IN onde não existe.
- Parâmetro com o mesmo nome e GUID diferente: aparece no log e no resumo do topo
  ("Atenção: … não é usado"). Nunca é gravado e não bloqueia nada.
- **Removido do fluxo e do código:** `offer_in_binding`, `bind_in_param`,
  `_insert_binding`, `in_param_missing_categories` e `SHARED_PARAM_TXT`. O comando não
  cria, não vincula e não pergunta mais. Saiu também a troca temporária do arquivo de
  parâmetros compartilhados do Revit.
- Dashboard: os tipos sem o parâmetro mostram "Não aplicável" e ficam fora de todos os
  indicadores; os critérios explicam isso.

### 2.1b Description não era gravada pela tabela 2 (correção de 15/10)

- **Causa:** na tabela 2, digitar a Keynote só gravava a Keynote. A Description e a
  Descrição IN do Excel só entravam depois de clicar em "Verificar no Excel", e só quando
  o campo do modelo estava **vazio**. Quem digitava a Keynote e clicava direto em
  "Atualizar descrições" gravava só a Keynote; a Description ficava para uma segunda
  rodada pela tabela 1.
- **Correção (`verify_incomplete`, `_cmp2`, `_on_changed2`, `analyze_model`):** com a
  Keynote no Excel, a Description e a Descrição IN da planilha ficam **propostas** na
  linha, tanto ao digitar quanto ao abrir. A proposta preenche campo vazio (célula verde)
  ou substitui um valor diferente (célula âmbar), igual à tabela 1. Só é gravada com a
  linha marcada e a confirmação. Excel vazio continua sem apagar nada.
- A confirmação passou a mostrar quantos campos serão gravados (Description, Descrição
  IN, Keynote). O resumo depois de gravar mostra quantos de cada um foram **gravados e
  conferidos**, e o log registra cada valor (antes → depois).
- A conferência ignora diferenças apenas de espaço ou quebra de linha, que o Revit pode
  normalizar, para não acusar falha num valor que foi gravado.

### 2.2 Gravação no modelo (`save_run`, `run_transaction`, `_set_param`)

- `run_transaction`:
  - confere se o documento está aberto, se não é somente leitura e se não há outra
    transação aberta (não aninha);
  - usa um `IFailuresPreprocessor` que descarta os avisos e desfaz a transação em caso de
    erro, **sem abrir diálogo modal** dentro do ExternalEvent;
  - só devolve sucesso se o `Commit` for `Committed`;
  - registra Start, Commit e RollBack no log.
- `_set_param`: trata `Set()` com retorno `False` como falha.
- Depois do commit, cada parâmetro gravado é **relido**. Só entra como "gravado" o que o
  modelo mostra de fato.
- Ao final aparece um **resumo**: N gravados e conferidos, mais a lista do que não foi
  gravado e o motivo de cada item (somente leitura, workshared, Revit recusou, transação
  desfeita). Se não houver nada para gravar, o resumo explica por quê.
- O botão "Atualizar descrições" sempre volta a ficar ativo (`finally`).
- Se o Revit negar o pedido (`Raise` diferente de Accepted ou Pending), a janela avisa e
  nada fica pendurado.

### 2.3 Fluxo, TXT, interface e limpeza

| Função | Alteração |
|--------|-----------|
| `RunLog` / `LOG` (novo) | Log diário em arquivo, gravado linha a linha. Cada execução tem um identificador próprio. Se o log falhar, o comando continua. |
| `revit_env_info` (novo) | Registra a versão do Revit (ano, atualização tipo `2025.4` e família de build `24.x`/`25.x`/`26.x`), pyRevit, motor Python, .NET, sistema operacional e documento. |
| `start_session` / `_step` (novo) | Cada etapa (1. Excel, 3-6. TXT, 7. coleta, 8-9. análise, 10. janela) tem início, fim, duração e contagens no log. Uma falha mostra a etapa e o caminho do log. |
| `build_keynote_txt` + `validate_keynote_entries` (novo) | Valida a árvore: chave vazia, chave repetida, pai inexistente e ciclo. O conteúdo de arquivos válidos não muda: os testes deram byte a byte igual à versão anterior e ao TXT do escritório. |
| `write_keynote_txt` | Grava num arquivo temporário e confere. Só então faz o backup com data e hora e substitui. Uma falha mantém o TXT anterior intacto. |
| `load_keynote_txt` | Recarga feita por `run_transaction` (commit conferido). Registra erros de sintaxe e de entrada informados pelo Revit e confere quantas entradas ficaram no projeto contra quantas tem o arquivo. |
| `make_keynote_txt` | Não grava um TXT vazio. Só diz "carregado" depois da confirmação do Revit. |
| `_RevitAction.Execute` / `run_in_revit` | Fila com nome da ação. Confere se o documento ainda existe. Registra exceções com traceback e trata o resultado de `Raise()`. |
| `show_in_view` / `_active_uidoc` (novo) | Usa o documento ativo **somente** se for o do comando. Ignora IDs apagados. Não enquadra em folha nem em tabela. Sem transação. |
| `VerifyWindow._safe`, `_watch_unhandled`, `_on_ui_exception` (novos) | Todos os eventos da janela ficam protegidos. Exceções do WPF fora do Python vão para o log e são tratadas quando a janela do comando está ativa. Os handlers são removidos ao fechar. |
| `scan_types`, `collect_model` | Elemento ou tipo inacessível é pulado, com contagem no log. Coletor iterado sem carregar a lista inteira. |
| `analyze_model` | **Uma** passada pelo modelo, reaproveitada pelas tabelas 2 e 3, pela seleção e pelo dashboard. Antes eram três. |
| `cleanup_old_colors` | Só confere elementos que as versões antigas pintavam e só em vistas de modelo. Varredura completa em partes de até 90 s, com o progresso salvo; depois disso, só a vista ativa. Transação conferida. Tudo registrado no log. |
| Modo de diagnóstico (novo, **Shift+Clique** no botão) | Etapas separadas A a H, descritas na seção 4. |

A interface, o Design System OCA, os temas, a busca, o agrupamento, as colunas
ajustáveis, a seleção, o dashboard, o TXT e o backup foram preservados. As únicas
mudanças visíveis são o texto "Não aplicável", o resumo depois da gravação e os avisos
de diagnóstico.

---

## 3. Matriz de compatibilidade (auditoria das APIs usadas)

| Recurso | 2024 (.NET 4.8) | 2025 (.NET 8) | 2026 (.NET 8) | Observação |
|---------|-----------------|---------------|---------------|------------|
| `ElementId.Value` / `IntegerValue` | ok | ok | ok | `eid_int` usa `.Value` quando existe. `IntegerValue` está obsoleto desde 2024. |
| `KeynoteTable.GetKeynoteTable`, `LoadFrom`, `KeyBasedTreeEntriesLoadResults`, `ExternalResourceReference.CreateLocalResource` | ok | ok | ok | Sem mudança de assinatura desde 2018. |
| `Transaction`, `FailureHandlingOptions`, `IFailuresPreprocessor` | ok | ok | ok | — |
| `ExternalEvent`, `IExternalEventHandler`, `ExternalEventRequest` | ok | ok | ok | — |
| `UIDocument.ShowElements`, `Selection.SetElementIds` | ok | ok | ok | — |
| `OverrideGraphicSettings` (cores e padrões) | ok | ok | ok | Propriedades de 2019 em diante; as antigas são lidas com `getattr`. |
| `WorksharingUtils.GetCheckoutStatus` | ok | ok | ok | — |
| `BuiltInParameterGroup` / `ParameterType` | — | — | — | **Não são mais usados** (só apareciam no vínculo removido). |
| `System.IO.Compression` (leitura do .xlsx) | `…FileSystem` | `…ZipFile` | `…ZipFile` | O script tenta os dois nomes de assembly. |
| WPF, `System.Data` (DataTable) | ok | ok | ok | Comportamento do DataGrid no .NET 8: **testar no Revit** (H2). |
| pyRevit `forms.alert(sub_msg)`, `CommandSwitchWindow`, `ProgressBar`, `EXEC_PARAMS`, `get_document_data_file` | 4.8+ | 4.8+ / 5.x | 5.x | Você usa a 5.2. |
| `lib/oca_ds` (Design System OCA) | dependência local | idem | idem | Se faltar, o comando avisa e para. |

A compatibilidade com um ano não garante a das atualizações dele. Por isso cada execução
grava no log o build exato (`VersionBuild`) e a atualização (`SubVersionNumber`).

---

## 4. Plano de testes (use CÓPIAS dos modelos)

O **modo de diagnóstico** abre com Shift+Clique no botão.

| Teste | Como | Esperado | Log |
|-------|------|----------|-----|
| A. Limpeza de cores antigas | Diagnóstico A | Pergunta se há cores; não fecha o Revit. | Vistas conferidas, tempo e elementos. |
| B. TXT sem recarga | Diagnóstico B | TXT e backup na pasta; o projeto não muda. | Entradas e avisos da validação. |
| C. Recarga sem interface | Diagnóstico C (TXT gerado no B) | "Carregado (N de N entradas)". | Commit e erros do Revit. **Depois do C, navegue por vistas com tags de keynote: se fechar aqui, é H3.** |
| D. Coleta e análise | Diagnóstico D | Resumo com contagens. | Tipos, instâncias e Descrição IN (com / sem / GUID diferente). |
| E. Interface com dados de teste | Diagnóstico E | Janela com linhas fictícias; clique, agrupe, busque, recolha e role. "Atualizar" não grava. | Eventos da interface. **Se fechar aqui, é interface (H1/H2), não o modelo.** |
| F. Completa com dados reais | Diagnóstico F (ou clique normal) | Fluxo normal. | Todas as etapas. |
| G. Gravação com confirmação | No F: marque linhas → "Atualizar descrições" → Não, depois Sim | Não = nada muda. Sim = resumo "N gravados e conferidos"; os valores aparecem nas Propriedades do tipo. | Start/Commit/RollBack e cada falha. |
| H. Sem `Descrição IN` | Modelo sem o parâmetro | "Não aplicável" nas colunas IN; a Description grava normalmente; o dashboard sem erros de IN. | "Descrição IN: não existe neste modelo". |
| I. IN só em algumas categorias | Modelo misto | Grava IN só onde existe; nos demais tipos aparece "Não aplicável". | "existe em X de Y tipo(s)". |
| J. Mesmo nome, outro GUID | Parâmetro "Descrição IN" de projeto com outro GUID | Aviso no topo; nada gravado nele. | WARN "GUID diferente". |
| Extra | Diagnóstico G (sem limpeza e sem recarga) e H (títulos de grupo simples) | Se G não fecha e F fecha, a causa está na limpeza ou na recarga. Se H não fecha e F fecha, a causa é a linha de título do DataGrid. | — |

### Matriz por versão

Rode **F, G, H, I e J** em cada uma. A versão exata está em Ajuda > Sobre e no log
(`revit_build`, `revit_update`).

| Revit | Builds a cobrir |
|-------|-----------------|
| 2024 | build inicial (24.0) e a última atualização instalada (24.x) |
| 2025 | build inicial (25.0), 2025.2 e 2025.4 (as suas) |
| 2026 | build inicial (26.0) e 2026.4 (a sua) |

Para cada combinação, anote: versão (do log), teste, resultado, e se fechou, a última
linha do log.

---

## 5. Onde fica o log

`%LOCALAPPDATA%\OCA\pyRevit\logs\Keynotes\keynotes_AAAAMMDD.log`

Ou use o atalho: Shift+Clique > "Abrir a pasta dos logs".

Cada linha tem hora (com milissegundos), o identificador da execução e o nível
(INFO/STEP/WARN/ERROR). Linhas `>>` e `<<` marcam o início e o fim de cada etapa.
**Se o Revit fechar, a última linha diz onde estava.**

## 6. Como coletar o journal do Revit se o encerramento continuar

1. **Não abra o Revit de novo antes de copiar os arquivos**, porque cada abertura cria um
   journal novo.
2. Abra `%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 2026\Journals` (troque o ano
   conforme a versão).
3. Ordene por data e copie o `journal.XXXX.txt` **mais recente**, mais algum `.dmp` da
   mesma pasta ou da pasta do relatório de erros, se existir.
4. Copie também o log do comando do dia (seção 5).
5. Envie os arquivos junto com a versão (Ajuda > Sobre) e o que você estava fazendo na
   hora: clicou numa linha, rolou, trocou o agrupamento, estava parado, etc.
6. No journal, procure perto do fim por `Jrn.RibbonEvent` (botão), `API_SUCCESS` /
   `API_ERROR`, `ExternalEvent` e `fatal error` / `unrecoverable`. As linhas antes do
   erro mostram a última operação do Revit.
