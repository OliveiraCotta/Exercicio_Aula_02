# Instalador do Plugin OCA (pyRevit)

Gera `dist/Instalador_Plugin_OCA.exe`, que instala a extensão `OCA.extension` na
pasta de extensões do pyRevit de outro computador, sem depender dos arquivos
originais nem do caminho onde o código foi desenvolvido.

## Arquivos

| Arquivo | Função |
|---|---|
| `installer/Instalador_Plugin_OCA.iss` | Código-fonte e configuração do instalador (Inno Setup). É o único arquivo necessário para compilar. |
| `installer/build.ps1` | Compila no Windows: localiza o Inno Setup, compila e mostra o SHA-256 do `.exe`. |
| `installer/build-wine.sh` | Compila no Linux via Wine (foi assim que o `.exe` deste repositório foi gerado). |
| `installer/LEIA_ANTES_DE_INSTALAR.txt` | Texto exibido no assistente antes da instalação. |
| `installer/assets/` | Ícone e imagens do assistente (gerados por `installer/tools/gerar_assets.py` a partir do logo OCA). |
| `dist/Instalador_Plugin_OCA.exe` | Instalador pronto. |

Nenhum arquivo de `OCA.extension` é alterado pelo processo de build.

## Estrutura da extensão analisada

Versão atual: **1.01** (recebida como `OCA1.01.extension.zip`).

```
OCA.extension/
  startup.py                    pinta a barra de título dos painéis (lê os bundle.yaml)
  Instaladores/                 scripts da impressora "OCA PDF24" (.cmd/.ps1) - só copiados
  lib/oca_ds/                   biblioteca compartilhada (WPF, Excel, relatórios, logo)
  lib/oca_params/               parâmetros compartilhados (usada pelo painel Gerência)
  OCA.tab/                      aba "OCA"; ordem dos painéis em bundle.yaml
    Sobre.panel/Sobre.pushbutton                         (contém VERSAO = versão do plugin)
    Keynote.panel/01_Acabamentos de Ambiente.pushbutton  (+ .xlsx e .html usados pelo script)
    Keynote.panel/02_Atualização Descrição por Keynotes.pushbutton
    Vistas.panel/01_Abertura de Vista, 02_Abertura de Multiplas Vistas por Ambientes
    Coordenadas.panel/01_Exportar Coordenadas.pushbutton
    Gerência.panel/01_Gerenciador de Parâmetros Compartilhados.pushbutton
    Exportação.panel/ExportarPranchas.pushbutton         (título "Exportação DWG/PDF")
    Visualizar.panel/01_Visualizador de IFC.pushbutton   (+ ifc-hud-viewer.html)
```

51 arquivos, 7 painéis, 9 botões.

**Nome da pasta:** o zip veio como `OCA1.01.extension`, mas o instalador grava
sempre como `OCA.extension`. Com o número da versão no nome, o pyRevit
carregaria a versão nova *ao lado* da anterior (aba OCA com botões duplicados), e
cada versão futura somaria mais uma pasta. Os scripts encontram seus arquivos por
caminho relativo, então o nome da pasta não muda o funcionamento. Se já existir
uma `OCA1.01.extension` copiada à mão, o instalador avisa (tela de verificação e
relatório) para apagá-la; ele não a remove sozinho.

**Pasta Instaladores / impressora PDF24:** o botão Exportação DWG/PDF (opção
PDF24) precisa da impressora "OCA PDF24". Os scripts que a criam pedem
administrador, alteram o registro do PDF24 (HKLM) e criam porta e impressora. O
instalador **não** os executa; apenas os copia. O próprio botão oferece
instalar a impressora, e o relatório final indica o `.cmd`. Os scripts encontram seus recursos por caminho
relativo (`os.path.dirname(__file__)`), então funcionam em qualquer pasta de
extensões. O `startup.py` tem `%APPDATA%\pyRevit\Extensions\OCA.extension` apenas
como reserva e usa a pasta real de `__file__`.

**Onde ficam as configurações do usuário:** os comandos usam
`script.get_config()` / `script.save_config()` (gravam no `pyRevit_config.ini`) e
`script.get_document_data_file()` / `get_universal_data_file()` (gravam em
`%APPDATA%\pyRevit\...`). Nada disso fica dentro de `OCA.extension`, então
atualizar ou reinstalar o plugin não apaga as opções salvas. O instalador não
toca no `pyRevit_config.ini`.

## Como gerar o .exe

### Windows (recomendado)

1. Instale o **Inno Setup 6.4 ou superior**: <https://jrsoftware.org/isdl.php>.
2. Na pasta do repositório, rode:
   ```powershell
   powershell -ExecutionPolicy Bypass -File installer\build.ps1
   ```
   ou abra `installer\Instalador_Plugin_OCA.iss` no Inno Setup e clique em
   *Build > Compile*.
3. O resultado sai em `dist\Instalador_Plugin_OCA.exe`.

Se o Inno Setup estiver em outro local: `build.ps1 -IsccPath "C:\...\ISCC.exe"`.

### Linux (Wine)

Veja o cabeçalho de `installer/build-wine.sh`. O locale precisa ser UTF-8, senão o
Wine corrompe o nome da pasta `02_Atualização Descrição por Keynotes.pushbutton`.

### O que a compilação faz

- **Valida a extensão** e aborta com mensagem clara se faltar `startup.py`,
  `OCA.tab`, `OCA.tab\bundle.yaml`, `lib\oca_ds\__init__.py`, o logo, ou se algum
  `.pushbutton` não tiver `script.py`.
- **Lê a versão** de `VERSAO` no botão Sobre: `u"1.01"` é usada como está; o
  formato antigo `u"13/10/2026"` vira `2026.10.13`. Para forçar outra:
  `build.ps1 -AppVersion 1.02`. A comparação é numérica (1.01 < 1.2 < 1.10) e a
  mudança da numeração por data para 1.01 é reconhecida (não gera falso aviso de
  "versão instalada mais nova").
- **Gera a lista de arquivos e o SHA-256** de cada um a partir da mesma
  varredura que alimenta a seção `[Files]`. Ignora `__pycache__`, `*.pyc`,
  `Thumbs.db`, `desktop.ini`, `.DS_Store` e arquivos temporários do Office (`~$*`).

## O que o instalador faz no computador de destino

1. **Boas-vindas e informações** (requisitos, o que será feito).
2. **Verificação do pyRevit**, sem alterar nada. Mostra:
   - `pyRevit_config.ini` do usuário (`%APPDATA%\pyRevit`) e de administrador
     (`%PROGRAMDATA%\pyRevit`), e as pastas da chave `[core] userextensions`;
   - se o pyRevit está conectado ao Revit 2025 e 2026 (arquivos `.addin` com o
     `PyRevitLoader` em `%APPDATA%\Autodesk\Revit\Addins\<ano>` e
     `%PROGRAMDATA%\Autodesk\Revit\Addins\<ano>`) e onde o pyRevit está instalado;
   - o pyRevit em *Programas instalados*;
   - as pastas de extensões, e se alguma já contém `OCA.extension` (com a versão);
   - outras extensões que também tenham `OCA.tab` (por exemplo, uma antiga
     `ESCRITORIO.extension`), que duplicariam a aba;
   - se o Revit está aberto.

   Se o pyRevit não for encontrado, explica o problema e oferece cancelar
   (o botão padrão é *Não continuar*).
3. **Escolha da pasta de extensões**: a pasta padrão do pyRevit, as pastas
   registradas no `pyRevit_config.ini` ou *Outra pasta* (configuração manual). A
   pasta pré-selecionada é a da instalação anterior, se houver.
4. **Validações da pasta**: recusa a raiz de um disco ou uma pasta `*.extension`
   (erro comum: escolher a própria `OCA.extension`); avisa se a pasta é interna
   do pyRevit (apagada quando o pyRevit é atualizado), se ela não é carregada
   pelo pyRevit, ou se já existe outra `OCA.extension` em outra pasta.
5. **Instalação existente**: mostra versão instalada × nova e oferece
   *Atualizar*, *Reinstalação limpa* ou *Cancelar*. Avisa quando a versão
   instalada é mais nova que a do instalador.
6. **Resumo** (tela *Pronto para instalar*) e **progresso** da cópia.
7. **Backup**: antes de qualquer alteração, copia a pasta atual para
   `%LOCALAPPDATA%\OCA\Plugin OCA\Backups\OCA.extension_<versão>_<data>` e confere
   cada arquivo copiado. Se o backup falhar, a instalação para sem mexer em nada.
8. **Conferência**: depois de copiar, confere existência e SHA-256 de todos os
   arquivos e mostra o **relatório** (pasta, tipo de instalação, backup, arquivos
   instalados, atualizados, removidos e mantidos, próximos passos). O relatório é
   salvo em `%LOCALAPPDATA%\OCA\Plugin OCA\relatorio_instalacao_<data>.txt`.

### Atualizar × reinstalação limpa

- **Atualizar** substitui os arquivos do pacote. Arquivos que não pertencem ao
  pacote são **mantidos**. Um arquivo só é **removido** quando a instalação
  anterior foi feita por este instalador (lista em
  `manifesto_instalado.txt`), ele não existe mais na versão nova **e** não foi
  modificado desde então. Arquivos antigos modificados são mantidos e listados.
  Em instalações feitas à mão (sem manifesto), nada é removido: arquivos que
  sobrarem são listados no relatório para revisão.
- **Reinstalação limpa** renomeia a pasta atual (operação atômica: se algum
  arquivo estiver em uso, nada muda) e só então a apaga. Tudo continua no backup.

### Permissões

- Por padrão roda **sem administrador** (`PrivilegesRequired=lowest`): a pasta
  padrão do pyRevit fica no perfil do usuário.
- Se a pasta escolhida não aceitar gravação, o instalador explica e oferece
  **reabrir como administrador** (o Windows pede confirmação). Só nesse caso há
  elevação.
- Nunca abre o Revit, nunca executa comandos externos nem acessa a internet. O
  único programa que ele pode iniciar é ele mesmo, na reabertura como
  administrador, e só depois de perguntar.
- Atenção: se a elevação usar **outra conta** (usuário comum + senha de um
  administrador), `%APPDATA%` passa a ser o da conta de administrador. Nesse caso,
  confira a pasta de destino na tela de escolha.

### Instalação silenciosa (TI)

```
Instalador_Plugin_OCA.exe /VERYSILENT /SUPPRESSMSGBOXES /LOG="C:\temp\oca.log"
```

| Parâmetro | Efeito |
|---|---|
| `/DIR="C:\...\Extensions"` | Pasta de extensões (padrão: a detectada). |
| `/MODO=limpa` | Reinstalação limpa quando já houver uma versão (padrão: atualizar). |
| `/ALLUSERS` | Modo administrador (para pastas protegidas). |

Em modo silencioso, as perguntas assumem a resposta padrão segura: continua se o
pyRevit não for encontrado (útil para preparar a máquina antes do pyRevit), não
instala em pasta interna do pyRevit e falha se não houver permissão de gravação.

## Procedimento para atualizar o plugin no futuro

1. Altere os arquivos em `OCA.extension` normalmente.
2. Atualize `VERSAO` em `OCA.tab/Sobre.panel/Sobre.pushbutton/script.py` (ex.:
   `u"1.02"`). Esse valor vira a versão do instalador e é o mesmo mostrado no
   botão Sobre. Mantenha a pasta com o nome `OCA.extension` no repositório (não
   `OCA1.02.extension`).
3. Gere o `.exe` (`installer\build.ps1`).
4. **Não altere o `AppGuid`** no `.iss`. Ele identifica o produto; é o que faz a
   versão nova reconhecer e atualizar a anterior.
5. Nos computadores, rode o novo instalador e escolha **Atualizar**.

O que é preservado: as opções dos comandos (no `pyRevit_config.ini` e nas pastas
de dados do pyRevit, fora de `OCA.extension`), arquivos que você tenha
adicionado à pasta do plugin e um backup completo da versão anterior.

Para renomear ou remover um botão: basta fazê-lo em `OCA.extension`. Na
atualização, o instalador remove os arquivos antigos não modificados (se a
instalação anterior foi feita por ele) e lista no relatório o que sobrou.

## Desinstalação

*Configurações > Aplicativos > Plugin OCA - Automatizações em BIM > Desinstalar.*
Remove os arquivos instalados. Arquivos adicionados por você, os backups e os
relatórios são mantidos. Depois, use pyRevit > Reload ou reinicie o Revit.

## Como confirmar no Revit

1. Revit fechado: abra o Revit 2025/2026 normalmente (não precisa ser como
   administrador). Revit aberto: aba **pyRevit > Reload**.
2. A aba **OCA** deve aparecer com os painéis Sobre, Keynote, Vistas,
   Coordenadas, Gerência, Exportação e Visualizar.
3. **OCA > Sobre** deve mostrar a versão instalada.
4. Se a aba não aparecer: confirme que o pyRevit está conectado a essa versão do
   Revit (`pyrevit attached` no Prompt de Comando) e que a pasta escolhida é a
   padrão ou está em pyRevit > Settings > *Custom Extension Directories*.

## Testes feitos

O `.exe` foi compilado com Inno Setup 6.4.3 e executado sob Wine (Windows 10
emulado) nos cenários abaixo. A árvore instalada foi comparada byte a byte com
`OCA.extension` (`diff -r`), além da conferência SHA-256 do próprio instalador.

- Instalação nova sem pyRevit (aviso e opção de cancelar) e com pyRevit simulado
  (`pyRevit_config.ini` com `userextensions`, `.addin` para 2025/2026).
- Atualização de instalação manual antiga em pasta personalizada registrada:
  versão lida do botão Sobre, backup conferido, arquivos extras mantidos.
- Atualização sobre instalação feita pelo instalador: arquivo obsoleto não
  modificado removido (com a pasta vazia), obsoleto modificado mantido,
  arquivo do pacote alterado localmente substituído e listado.
- Reinstalação limpa (`/MODO=limpa`), modo administrador (`/ALLUSERS /DIR=`),
  desinstalação (arquivo do usuário preservado).
- Interface: todas as telas capturadas, incluindo *Outra pasta* com aviso de
  pasta não registrada e cancelamento na tela final sem criar nada.
- Versão 1.01: instalação nova (51 arquivos idênticos ao zip) e atualização de
  uma máquina com a versão 2026.10.13 instalada pelo instalador anterior: os
  botões removidos/renomeados na 1.01 (`02_Abertura de Multiplas Vistas`,
  `03_Abertura de Multiplas Vistas por Ambientes`) foram apagados com as pastas,
  sem deixar botão duplicado; `OCA1.01.extension` copiada à mão é detectada.
- Erros de compilação: botão sem `script.py`, `startup.py` ausente, pasta da
  extensão inexistente; versão forçada por `/DAppVersion`.

**Não testado (exige Windows real):** o pedido de UAC na reabertura como
administrador, pasta sem permissão de gravação, aviso do SmartScreen e o
carregamento no Revit/pyRevit reais. Recomenda-se uma instalação de teste em uma
máquina com Revit antes de distribuir.

## Assinatura digital

O `.exe` não é assinado. O Windows SmartScreen pode mostrar "O Windows protegeu
o computador" (*Mais informações > Executar assim mesmo*). Para evitar isso,
assine com um certificado de assinatura de código (Inno Setup: diretiva
`SignTool`).

## Solução de problemas

- **Log detalhado**: `%TEMP%\Setup Log <data> #<n>.txt` (ou `/LOG="arquivo"`).
- **"Não foi possível liberar a pasta atual"**: feche o Revit e o Excel (as
  planilhas modelo do botão Acabamentos podem estar abertas) e tente de novo.
- **Botões duplicados na aba OCA**: há duas cópias do plugin (veja os avisos da
  tela de verificação), por exemplo uma `ESCRITORIO.extension` antiga.
