; =============================================================================
;  Instalador do Plugin OCA - extensão pyRevit (OCA<versão>.extension)
;  Compilador: Inno Setup 6.4 ou superior (https://jrsoftware.org/isinfo.php)
;
;  Compilar:   installer\build.ps1            (recomendado - ver installer\LEIA-ME.md)
;         ou   ISCC.exe installer\Instalador_Plugin_OCA.iss
;         ou   abrir este arquivo no Inno Setup e clicar em Build > Compile
;
;  Saída:      dist\Instalador_Plugin_<nome da extensão>.exe
;              (ex.: OCA1.01.extension -> dist\Instalador_Plugin_OCA1.01.exe)
;
;  Este script NÃO altera nenhum arquivo do plugin. Na compilação ele:
;    1. valida a estrutura da extensão (startup.py, OCA.tab, lib\oca_ds,
;       script.py em cada .pushbutton...);
;    2. lê a versão do plugin em OCA.tab\Sobre.panel\Sobre.pushbutton\script.py
;       (VERSAO = u"1.01", ou no formato antigo u"dd/mm/aaaa" -> aaaa.mm.dd),
;       a mesma exibida no botão Sobre;
;    3. gera a lista de arquivos do pacote com o SHA-256 de cada um, usada para
;       conferir a cópia no computador de destino.
; =============================================================================

#if Ver < EncodeVer(6, 4, 0)
  #error Este script requer o Inno Setup 6.4 ou superior.
#endif

#define AppName       "Plugin OCA - Automatizações em BIM"
#define AppShortName  "Plugin OCA"
#define AppPublisher  "Oliveira Cotta"
#define AppURL        "https://www.oliveiracotta.com.br"
; GUID fixo: identifica o produto no Windows. NUNCA altere - é ele que permite
; que versões futuras reconheçam e atualizem esta instalação.
#define AppGuid       "0EDB5C30-7A63-400F-9133-7C591AFBFA0B"

; Pasta com a extensão: a ÚNICA pasta *.extension na raiz do repositório (ao
; lado da pasta installer). O nome dela é o nome instalado no pyRevit e define
; o nome do .exe: OCA1.01.extension -> Instalador_Plugin_OCA1.01.exe.
; Pode ser trocada na linha de comando: ISCC /DSourceRoot="C:\caminho\OCA1.01.extension"
#ifndef SourceRoot
  ; lista "Nome1|Nome2|" das pastas *.extension ao lado da pasta installer
  #define ListarExt(int Achou, int H, str Acc) \
    Achou \
      ? ( Local[0] = FindGetFileName(H), \
          ListarExt(FindNext(H), H, DirExists(SourcePath + "..\" + Local[0]) ? Acc + Local[0] + "|" : Acc) ) \
      : Acc
  #define ExtH FindFirst(SourcePath + "..\*.extension", faAnyFile)
  #define ExtLista (ExtH ? ListarExt(1, ExtH, "") : "")
  #if ExtLista == ""
    #error Nenhuma pasta *.extension encontrada ao lado da pasta installer (ou use /DSourceRoot=...).
  #endif
  #define ExtFound Copy(ExtLista, 1, Pos("|", ExtLista) - 1)
  #if Len(ExtLista) > Len(ExtFound) + 1
    #error Há mais de uma pasta *.extension ao lado da pasta installer. Deixe só a versão a empacotar (ou use /DSourceRoot=...).
  #endif
  #define SourceRoot  SourcePath + "..\" + ExtFound
#endif
#define ExtName       ExtractFileName(RemoveBackslash(SourceRoot))
#if LowerCase(ExtractFileExt(ExtName)) != "extension"
  #error O nome da pasta da extensão precisa terminar em .extension (encontrado: {#ExtName}).
#endif
#define ExtBase       Copy(ExtName, 1, Len(ExtName) - Len(".extension"))

; -----------------------------------------------------------------------------
; 1. Validação da estrutura da extensão (na compilação)
; -----------------------------------------------------------------------------
#if !DirExists(SourceRoot)
  #error Pasta da extensão não encontrada (verifique /DSourceRoot=...).
#endif

#define RequireFile(str Rel) \
  FileExists(SourceRoot + "\" + Rel) ? "" : Error("Arquivo obrigatório ausente em " + ExtName + ": " + Rel)
#define RequireDir(str Rel) \
  DirExists(SourceRoot + "\" + Rel) ? "" : Error("Pasta obrigatória ausente em " + ExtName + ": " + Rel)

#expr RequireFile("startup.py")
#expr RequireDir("OCA.tab")
#expr RequireFile("OCA.tab\bundle.yaml")
#expr RequireDir("lib\oca_ds")
#expr RequireFile("lib\oca_ds\__init__.py")
#expr RequireFile("lib\oca_ds\assets\logo\logo.png")
#expr RequireFile("OCA.tab\Sobre.panel\Sobre.pushbutton\script.py")

; -----------------------------------------------------------------------------
; 2. Versão: lida do botão Sobre. Aceita VERSAO = u"1.01" (usada como está)
;    ou o formato antigo VERSAO = u"13/10/2026" (-> 2026.10.13).
;    Pode ser forçada com: ISCC /DAppVersion=1.02
; -----------------------------------------------------------------------------
#ifndef AppVersion
  #define SobreScript SourceRoot + "\OCA.tab\Sobre.panel\Sobre.pushbutton\script.py"
  #define VersaoLinha ""
  #define FH FileOpen(SobreScript)
  #if !FH
    #error Não foi possível abrir o script do botão Sobre para ler a versão.
  #endif
  #sub LerLinhaVersao
    #define public LinhaAtual Trim(FileRead(FH))
    #if VersaoLinha == "" && Copy(LinhaAtual, 1, 6) == "VERSAO"
      #define public VersaoLinha LinhaAtual
    #endif
  #endsub
  #define LinhaAtual ""
  #for {0; !FileEof(FH); 0} LerLinhaVersao
  #expr FileClose(FH)
  #if VersaoLinha == ""
    #error Linha VERSAO não encontrada no script do botão Sobre. Defina a versão com /DAppVersion=1.01
  #endif
  #define VersaoAspas1 Pos("""", VersaoLinha)
  #define VersaoResto  Copy(VersaoLinha, VersaoAspas1 + 1)
  #define VersaoData   Copy(VersaoResto, 1, Pos("""", VersaoResto) - 1)
  #define SoDigitosEPontos(str S) \
    Len(S) == 0 ? 1 : (Pos(Copy(S, 1, 1), "0123456789.") > 0 ? SoDigitosEPontos(Copy(S, 2)) : 0)
  #if Len(VersaoData) == 10 && Copy(VersaoData, 3, 1) == "/" && Copy(VersaoData, 6, 1) == "/"
    #define AppVersion Copy(VersaoData, 7, 4) + "." + Copy(VersaoData, 4, 2) + "." + Copy(VersaoData, 1, 2)
  #elif Len(VersaoData) > 0 && SoDigitosEPontos(VersaoData) && \
        Pos(Copy(VersaoData, 1, 1), "0123456789") > 0 && \
        Pos(Copy(VersaoData, Len(VersaoData), 1), "0123456789") > 0
    #define AppVersion VersaoData
  #else
    #error VERSAO do botão Sobre deve ser numérica (ex.: 1.01) ou dd/mm/aaaa. Ou defina a versão com /DAppVersion=1.01
  #endif
#endif

#if Pos(AppVersion, ExtName) == 0
  #expr Warning("A versão do botão Sobre (" + AppVersion + ") não aparece no nome da pasta " + ExtName + ". Confira se ambos foram atualizados.")
#endif

; -----------------------------------------------------------------------------
; 3. Varredura da extensão: gera as entradas de [Files] e o manifesto SHA-256
;    a partir da MESMA lista, então o que é instalado é exatamente o que é
;    conferido. Ignora lixo de sistema/compilação (__pycache__, *.pyc...).
; -----------------------------------------------------------------------------
#define Excluido(str Nome) \
  (LowerCase(Nome) == "__pycache__") || (LowerCase(Nome) == ".git") || \
  (LowerCase(Nome) == ".ds_store") || (LowerCase(Nome) == "thumbs.db") || \
  (LowerCase(Nome) == "desktop.ini") || (LowerCase(ExtractFileExt(Nome)) == "pyc") || \
  (Copy(Nome, 1, 2) == "~$")

#define PascalStr(str S) "'" + StringChange(S, "'", "''") + "'"

#define ValidarBotao(str Pasta) \
  (LowerCase(ExtractFileExt(Pasta)) == "pushbutton") && !FileExists(Pasta + "\script.py") \
    ? Error("Botão sem script.py: " + Pasta) : ""

; Modo 1 = linha de [Files]; Modo 2 = linha do manifesto em [Code]
#define EmitirArquivo(int Modo, str Completo, str Rel, str RelDir) \
  Modo == 1 \
    ? "Source: """ + Completo + """; DestDir: ""{app}\" + ExtName + RelDir + """; " + \
      "Flags: ignoreversion overwritereadonly uninsremovereadonly" + NewLine \
    : "  M(" + PascalStr(Rel) + ", " + Str(FileSize(Completo)) + ", '" + \
      LowerCase(GetSHA256OfFile(Completo)) + "');" + NewLine

#define ProcessarEntrada(int Modo, str Pasta, str RelDir, int Achou, int H) \
  Achou \
    ? ( Local[0] = FindGetFileName(H), \
        Local[1] = Pasta + "\" + Local[0], \
        Local[2] = (Local[0] == "." || Local[0] == ".." || Excluido(Local[0])) \
          ? "" \
          : ( DirExists(Local[1]) \
              ? ValidarBotao(Local[1]) + ProcessarPasta(Modo, Local[1], RelDir + "\" + Local[0]) \
              : EmitirArquivo(Modo, Local[1], Copy(RelDir + "\" + Local[0], 2), RelDir) ), \
        Local[2] + ProcessarEntrada(Modo, Pasta, RelDir, FindNext(H), H) ) \
    : ""

#define ProcessarPasta(int Modo, str Pasta, str RelDir) \
  Local[0] = FindFirst(Pasta + "\*", faAnyFile), \
  Local[0] ? ProcessarEntrada(Modo, Pasta, RelDir, 1, Local[0]) : ""

#define ListaArquivos   ProcessarPasta(1, SourceRoot, "")
#define ListaManifesto  ProcessarPasta(2, SourceRoot, "")

#if Pos("M(", ListaManifesto) == 0
  #error Nenhum arquivo encontrado na pasta da extensão.
#endif

[Setup]
AppId={{{#AppGuid}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
VersionInfoVersion={#AppVersion}
VersionInfoDescription=Instalador do {#AppName}
VersionInfoProductName={#AppName}
VersionInfoCompany={#AppPublisher}
VersionInfoCopyright=© {#Copy(AppVersion, 1, 4)} {#AppPublisher}

; {app} = pasta de EXTENSÕES do pyRevit (a pasta que contém as *.extension).
; O plugin vai para {app}\{#ExtName}. O valor real é decidido pela detecção
; em [Code] (GetDefaultDir) e pode ser trocado pelo usuário.
DefaultDirName={code:GetDefaultDir}
AppendDefaultDirName=no
UsePreviousAppDir=no
DirExistsWarning=no
DisableDirPage=no
DisableProgramGroupPage=yes
DisableWelcomePage=no
DisableReadyPage=no
DisableFinishedPage=no

; Sem administrador por padrão: a pasta padrão do pyRevit fica no perfil do
; usuário (%APPDATA%\pyRevit\Extensions). Se o usuário escolher uma pasta
; protegida, o próprio instalador oferece reabrir como administrador
; (parâmetro /ALLUSERS), e só nesse caso o Windows pede a elevação.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline

; Não fechar nem reiniciar programas (nunca mexe no Revit).
CloseApplications=no
RestartApplications=no
RestartIfNeededByRun=no

; Desinstalador fica FORA da pasta de extensões do pyRevit.
UninstallFilesDir={localappdata}\OCA\Plugin OCA\Desinstalador
UninstallDisplayName={#AppName}
UninstallDisplayIcon={uninstallexe}

OutputDir={#SourcePath}..\dist
OutputBaseFilename=Instalador_Plugin_{#ExtBase}
SetupIconFile={#SourcePath}assets\oca_setup.ico
WizardStyle=modern
WizardSizePercent=120
WizardImageFile={#SourcePath}assets\wizard_large_100.bmp,{#SourcePath}assets\wizard_large_150.bmp,{#SourcePath}assets\wizard_large_200.bmp
WizardSmallImageFile={#SourcePath}assets\wizard_small_100.bmp,{#SourcePath}assets\wizard_small_150.bmp,{#SourcePath}assets\wizard_small_200.bmp
InfoBeforeFile={#SourcePath}LEIA_ANTES_DE_INSTALAR.txt
Compression=lzma2/max
SolidCompression=yes
; Revit 2025/2026 exigem Windows 10/11 64 bits.
MinVersion=10.0
ArchitecturesAllowed=x64compatible
SetupLogging=yes
ShowLanguageDialog=no

[Languages]
Name: "ptbr"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Messages]
WelcomeLabel2=Este assistente vai instalar o [name/ver] (extensão {#ExtName} do pyRevit) neste computador.%n%nO instalador procura a pasta de extensões do pyRevit, mostra o que encontrou e pede confirmação antes de copiar qualquer arquivo. Se já houver uma versão instalada (inclusive uma versão antiga com outro nome de pasta, como OCA.extension), você poderá escolher o que fazer, e um backup será feito antes.%n%nRecomenda-se fechar o Revit antes de continuar (não é obrigatório).
SelectDirLabel3=O plugin será instalado na subpasta {#ExtName}, dentro da pasta de EXTENSÕES do pyRevit indicada abaixo.
SelectDirBrowseLabel=Informe a pasta que CONTÉM as extensões do pyRevit (não a pasta {#ExtName}). Para trocar, clique em Procurar.
SelectDirDesc=Em qual pasta de extensões do pyRevit o plugin deve ser instalado?
FinishedHeadingLabel=Instalação do [name] concluída
UninstalledAll=O %1 foi removido.%n%nSe o Revit estiver aberto, clique em pyRevit > Reload (ou reinicie o Revit) para a aba OCA sumir.%n%nArquivos que não foram instalados por este instalador (por exemplo, arquivos que você adicionou à pasta {#ExtName}) e os backups (pasta AppData\Local\OCA\Plugin OCA\Backups) foram mantidos.

[Files]
#emit ListaArquivos

[UninstallDelete]
Type: files; Name: "{localappdata}\OCA\Plugin OCA\estado_instalacao.txt"
Type: files; Name: "{localappdata}\OCA\Plugin OCA\manifesto_instalado.txt"

[Code]
const
  EXT_NAME = '{#ExtName}';
  NEW_VERSION = '{#AppVersion}';
  PYREVIT_ADDIN_ID = 'B39107C3-A1D7-47F4-A5A1-532DDF6EDB5D';  { PyRevitConsts.AddinId }
  CRLF = #13#10;
  MODE_NEW = 0;
  MODE_UPDATE = 1;
  MODE_CLEAN = 2;

type
  TIntArray = array of Integer;

var
  { Manifesto do pacote (gerado na compilação) }
  PkgRel, PkgHash: TArrayOfString;
  PkgSize: TIntArray;
  PkgCount: Integer;

  { Manifesto da instalação anterior feita por este instalador (se houver) }
  PrevRel, PrevHash: TArrayOfString;
  PrevCount: Integer;
  PrevStateDir, PrevStateVersion: String;

  { Ambiente detectado }
  PyRevitRoot, DefaultExtDir: String;
  CandDirs, CandLabels: TStringList;     { pastas onde o pyRevit procura extensões }
  RegisteredDirs: TStringList;           { pastas registradas em pyRevit_config.ini }
  EnvLines: TStringList;                 { texto da página de diagnóstico }
  ExistingInstalls: TStringList;         { ...\<EXT_NAME> encontrados }
  ConflictExts: TStringList;             { outras *.extension com OCA.tab (versões antigas) }
  OldPage: TInputOptionWizardPage;
  MigratedOld: TStringList;
  PrevStateExt: String;
  PyRevitFound, RevitRunning: Boolean;
  AttachedYears: String;

  { Páginas }
  EnvPage: TOutputMsgMemoWizardPage;
  DirChoicePage: TInputOptionWizardPage;
  ExistingPage: TInputOptionWizardPage;
  ReportPage: TOutputMsgMemoWizardPage;

  { Instalação }
  InstallMode: Integer;
  TargetExt, BackupDir, InstalledVersion: String;
  RemovedObsolete, KeptModified, KeptExtras, OverwrittenModified, UpdatedFiles, MovedToBackup: TStringList;
  ReportText, ReportFile: String;
  VerifyOk: Boolean;
  VerifyFailures: TStringList;
  CloseWithoutConfirm: Boolean;

{ ---------------------------------------------------------------------------
  Manifesto do pacote
  --------------------------------------------------------------------------- }
procedure M(const Rel: String; const Size: Integer; const Hash: String);
begin
  SetArrayLength(PkgRel, PkgCount + 1);
  SetArrayLength(PkgHash, PkgCount + 1);
  SetArrayLength(PkgSize, PkgCount + 1);
  PkgRel[PkgCount] := Rel;
  PkgHash[PkgCount] := Hash;
  PkgSize[PkgCount] := Size;
  PkgCount := PkgCount + 1;
end;

procedure LoadPackageManifest;
begin
  PkgCount := 0;
#emit ListaManifesto
end;

function IndexInPkg(const Rel: String): Integer;
var I: Integer;
begin
  Result := -1;
  for I := 0 to PkgCount - 1 do
    if CompareText(PkgRel[I], Rel) = 0 then begin Result := I; Exit; end;
end;

function IndexInPrev(const Rel: String): Integer;
var I: Integer;
begin
  Result := -1;
  for I := 0 to PrevCount - 1 do
    if CompareText(PrevRel[I], Rel) = 0 then begin Result := I; Exit; end;
end;

{ ---------------------------------------------------------------------------
  Utilitários
  --------------------------------------------------------------------------- }
function StateDir: String;
begin
  Result := ExpandConstant('{localappdata}\OCA\Plugin OCA');
end;

function JoinRel(const A, B: String): String;
begin
  if A = '' then Result := B else Result := A + '\' + B;
end;

function SafeSHA256(const FileName: String): String;
begin
  try
    Result := Lowercase(GetSHA256OfFile(FileName));
  except
    Result := '';
  end;
end;

{ Expande %VARIAVEL% como o pyRevit faz (Environment.ExpandEnvironmentVariables). }
function ExpandEnvVars(const S: String): String;
var Rest, Name, Value: String; P: Integer;
begin
  Result := '';
  Rest := S;
  while True do
  begin
    P := Pos('%', Rest);
    if P = 0 then break;
    Result := Result + Copy(Rest, 1, P - 1);
    Rest := Copy(Rest, P + 1, Length(Rest));
    P := Pos('%', Rest);
    if P = 0 then begin Result := Result + '%'; break; end;
    Name := Copy(Rest, 1, P - 1);
    Value := GetEnv(Name);
    if (Name <> '') and (Value <> '') then
      Result := Result + Value
    else
      Result := Result + '%' + Name + '%';
    Rest := Copy(Rest, P + 1, Length(Rest));
  end;
  Result := Result + Rest;
end;

function NormDir(const S: String): String;
var T: String;
begin
  T := Trim(S);
  StringChangeEx(T, '/', '\', True);
  T := ExpandEnvVars(T);
  if T = '' then begin Result := ''; Exit; end;
  Result := RemoveBackslashUnlessRoot(ExpandFileName(T));
end;

function SameDir(const A, B: String): Boolean;
begin
  Result := (A <> '') and (B <> '') and (CompareText(NormDir(A), NormDir(B)) = 0);
end;

function ListIndexOfDir(L: TStringList; const Dir: String): Integer;
var I: Integer;
begin
  Result := -1;
  for I := 0 to L.Count - 1 do
    if SameDir(L[I], Dir) then begin Result := I; Exit; end;
end;

{ Lê arquivo de texto UTF-8 (com ou sem BOM); se não for UTF-8 válido, usa ANSI. }
function ReadTextFile(const FileName: String; var Text: String): Boolean;
var Raw: AnsiString; Decoded: String; I: Integer;
begin
  Result := False;
  Text := '';
  if not FileExists(FileName) then Exit;
  try
    if not LoadStringFromFile(FileName, Raw) then Exit;
  except
    Exit;
  end;
  if (Length(Raw) >= 3) and (Ord(Raw[1]) = $EF) and (Ord(Raw[2]) = $BB) and (Ord(Raw[3]) = $BF) then
    Delete(Raw, 1, 3);
  Decoded := UTF8Decode(Raw);
  if (Decoded = '') and (Raw <> '') then
    Decoded := String(Raw)
  else
    { U+FFFD = byte inválido em UTF-8: o arquivo está em ANSI }
    for I := 1 to Length(Decoded) do
      if Ord(Decoded[I]) = $FFFD then
      begin
        Decoded := String(Raw);
        Break;
      end;
  Text := Decoded;
  Result := True;
end;

function WriteTextFileUTF8(const FileName, Text: String): Boolean;
begin
  ForceDirectories(ExtractFileDir(FileName));
  Result := SaveStringToFile(FileName, UTF8Encode(Text), False);
end;

{ Valor de uma chave de um .ini (aceita continuação em linhas indentadas). }
function IniValueFromText(const Text, Section, Key: String): String;
var L: TStringList; I, P: Integer; Line, CurSec, K: String; InKey: Boolean;
begin
  Result := '';
  L := TStringList.Create;
  try
    L.Text := Text;
    CurSec := '';
    InKey := False;
    for I := 0 to L.Count - 1 do
    begin
      Line := L[I];
      if InKey then
      begin
        if (Length(Line) > 0) and ((Line[1] = ' ') or (Line[1] = #9)) then
        begin
          Result := Result + Trim(Line);
          Continue;
        end;
        Exit;
      end;
      Line := Trim(Line);
      if (Line = '') or (Line[1] = ';') or (Line[1] = '#') then Continue;
      if Line[1] = '[' then
      begin
        CurSec := Trim(Copy(Line, 2, Pos(']', Line) - 2));
        Continue;
      end;
      if CompareText(CurSec, Section) <> 0 then Continue;
      P := Pos('=', Line);
      if P = 0 then Continue;
      K := Trim(Copy(Line, 1, P - 1));
      if CompareText(K, Key) = 0 then
      begin
        Result := Trim(Copy(Line, P + 1, Length(Line)));
        InKey := True;
      end;
    end;
  finally
    L.Free;
  end;
end;

{ Converte o valor de "userextensions" em caminhos. O valor é uma lista
  JSON/Python, por exemplo: userextensions = ["C:\\pasta1", "D:\\pasta2"] }
procedure ParsePathList(const Value: String; Dest: TStringList);
var I: Integer; C, Q: Char; Cur, V: String; InStr, Any: Boolean;
begin
  V := Trim(Value);
  InStr := False;
  Any := False;
  Cur := '';
  Q := '"';
  I := 1;
  while I <= Length(V) do
  begin
    C := V[I];
    if InStr then
    begin
      if (C = '\') and (I < Length(V)) and ((V[I + 1] = '\') or (V[I + 1] = Q)) then
      begin
        Cur := Cur + V[I + 1];
        I := I + 1;
      end
      else if C = Q then
      begin
        InStr := False;
        if Trim(Cur) <> '' then Dest.Add(Trim(Cur));
        Cur := '';
      end
      else
        Cur := Cur + C;
    end
    else if (C = '"') or (C = '''') then
    begin
      InStr := True;
      Any := True;
      Q := C;
      Cur := '';
    end;
    I := I + 1;
  end;
  { Valor sem aspas (um único caminho) }
  if (not Any) and (V <> '') and (V <> '[]') then
  begin
    if (V[1] = '[') and (V[Length(V)] = ']') then V := Trim(Copy(V, 2, Length(V) - 2));
    if V <> '' then Dest.Add(V);
  end;
end;

procedure ListFilesRec(const Root, Rel: String; Files: TStringList);
var FR: TFindRec; P: String;
begin
  if Rel = '' then P := Root else P := Root + '\' + Rel;
  if FindFirst(P + '\*', FR) then
  try
    repeat
      if (FR.Name <> '.') and (FR.Name <> '..') then
      begin
        if (FR.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0 then
        begin
          if (FR.Attributes and FILE_ATTRIBUTE_REPARSE_POINT) = 0 then
            ListFilesRec(Root, JoinRel(Rel, FR.Name), Files);
        end
        else
          Files.Add(JoinRel(Rel, FR.Name));
      end;
    until not FindNext(FR);
  finally
    FindClose(FR);
  end;
end;

{ Versão gravada no botão Sobre de uma instalação existente (dd/mm/aaaa -> aaaa.mm.dd) }
function VersionFromSobre(const ExtDir: String): String;
var Text, Line, D: String; L: TStringList; I, P: Integer;
begin
  Result := '';
  if not ReadTextFile(ExtDir + '\OCA.tab\Sobre.panel\Sobre.pushbutton\script.py', Text) then Exit;
  L := TStringList.Create;
  try
    L.Text := Text;
    for I := 0 to L.Count - 1 do
    begin
      Line := Trim(L[I]);
      if Copy(Line, 1, 6) = 'VERSAO' then
      begin
        P := Pos('"', Line);
        if P = 0 then Exit;
        D := Copy(Line, P + 1, Length(Line));
        P := Pos('"', D);
        if P = 0 then Exit;
        D := Copy(D, 1, P - 1);
        if (Length(D) = 10) and (D[3] = '/') and (D[6] = '/') then
          Result := Copy(D, 7, 4) + '.' + Copy(D, 4, 2) + '.' + Copy(D, 1, 2)
        else
          Result := D;
        Exit;
      end;
    end;
  finally
    L.Free;
  end;
end;

{ Compara versões numéricas componente a componente ("1.01" < "1.2" < "1.10").
  Retorna -1, 0 ou 1. }
function NextVersionPart(var S: String): Integer;
var P: Integer;
begin
  P := Pos('.', S);
  if P = 0 then
  begin
    Result := StrToIntDef(S, 0);
    S := '';
  end
  else
  begin
    Result := StrToIntDef(Copy(S, 1, P - 1), 0);
    S := Copy(S, P + 1, Length(S));
  end;
end;

function CompareVersions(A, B: String): Integer;
var X, Y: Integer;
begin
  Result := 0;
  while (A <> '') or (B <> '') do
  begin
    X := NextVersionPart(A);
    Y := NextVersionPart(B);
    if X < Y then begin Result := -1; Exit; end;
    if X > Y then begin Result := 1; Exit; end;
  end;
end;

{ Versões antigas usavam a data (aaaa.mm.dd); as novas usam 1.01, 1.02... }
function IsDateVersion(V: String): Boolean;
begin
  Result := NextVersionPart(V) >= 1900;
end;

function DescribeVersion(const V: String): String;
begin
  if V = '' then Result := 'desconhecida' else Result := V;
end;

{ Testa se é possível gravar na pasta (ou, se ela ainda não existe, na
  primeira pasta-mãe existente). Não cria a pasta de destino. }
function CanWriteTo(const Dir: String): Boolean;
var D, Parent, TestFile: String;
begin
  Result := False;
  D := NormDir(Dir);
  while (D <> '') and (not DirExists(D)) do
  begin
    Parent := ExtractFileDir(D);
    if (Parent = D) or (Parent = '') then Exit;
    D := Parent;
  end;
  if D = '' then Exit;
  TestFile := AddBackslash(D) + 'oca_teste_gravacao_' + GetDateTimeString('hhnnsszzz', #0, #0) + '.tmp';
  try
    Result := SaveStringToFile(TestFile, 'teste', False);
  except
    Result := False;
  end;
  if Result then DeleteFile(TestFile);
end;

function IsRevitRunningNow: Boolean;
var Locator, Service, Procs: Variant;
begin
  Result := False;
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Service := Locator.ConnectServer('.', 'root\CIMV2');
    Procs := Service.ExecQuery('SELECT ProcessId FROM Win32_Process WHERE Name = ''Revit.exe''');
    Result := Procs.Count > 0;
  except
    Result := False;
  end;
end;

{ ---------------------------------------------------------------------------
  Detecção do ambiente pyRevit
  Baseada no código do pyRevit (pyrevitlib/pyrevit/__init__.py e
  pyRevitLabs): extensões de terceiros = %APPDATA%\pyRevit\Extensions (sempre)
  + pastas da chave [core] userextensions do pyRevit_config.ini.
  --------------------------------------------------------------------------- }
procedure AddCandidate(const Dir, Origin: String);
var N: String; Idx: Integer;
begin
  N := NormDir(Dir);
  if N = '' then Exit;
  Idx := ListIndexOfDir(CandDirs, N);
  if Idx >= 0 then
  begin
    if Pos(Origin, CandLabels[Idx]) = 0 then
      CandLabels[Idx] := CandLabels[Idx] + '; ' + Origin;
    Exit;
  end;
  CandDirs.Add(N);
  CandLabels.Add(Origin);
end;

procedure ReadPyRevitConfig(const CfgFile, Origin: String);
var Text, Value: String; L: TStringList; I: Integer;
begin
  if not FileExists(CfgFile) then
  begin
    EnvLines.Add('  - ' + Origin + ': não encontrado (' + CfgFile + ')');
    Exit;
  end;
  PyRevitFound := True;
  if not ReadTextFile(CfgFile, Text) then
  begin
    EnvLines.Add('  - ' + Origin + ': encontrado, mas não foi possível lê-lo (' + CfgFile + ')');
    Exit;
  end;
  Value := IniValueFromText(Text, 'core', 'userextensions');
  L := TStringList.Create;
  try
    ParsePathList(Value, L);
    if L.Count = 0 then
      EnvLines.Add('  - ' + Origin + ': ' + CfgFile + ' (nenhuma pasta personalizada de extensões)')
    else
    begin
      EnvLines.Add('  - ' + Origin + ': ' + CfgFile);
      for I := 0 to L.Count - 1 do
      begin
        if ListIndexOfDir(RegisteredDirs, L[I]) < 0 then RegisteredDirs.Add(NormDir(L[I]));
        if DirExists(NormDir(L[I])) then
        begin
          EnvLines.Add('      pasta registrada: ' + NormDir(L[I]));
          AddCandidate(L[I], 'registrada no pyRevit_config.ini');
        end
        else
          EnvLines.Add('      pasta registrada que NÃO EXISTE (o pyRevit a ignora): ' + NormDir(L[I]));
      end;
    end;
  finally
    L.Free;
  end;
end;

function ExtractTag(const Text, Tag: String): String;
var P1, P2: Integer; Low: String;
begin
  Result := '';
  Low := Lowercase(Text);
  P1 := Pos('<' + Lowercase(Tag) + '>', Low);
  if P1 = 0 then Exit;
  P1 := P1 + Length(Tag) + 2;
  P2 := Pos('</' + Lowercase(Tag) + '>', Low);
  if P2 <= P1 then Exit;
  Result := Trim(Copy(Text, P1, P2 - P1));
end;

procedure ScanAddinFolder(const Folder: String; const Year: String; var Found: Boolean);
var FR: TFindRec; Text, Asm, Root: String; P: Integer;
begin
  if not DirExists(Folder) then Exit;
  if FindFirst(Folder + '\*.addin', FR) then
  try
    repeat
      if ReadTextFile(Folder + '\' + FR.Name, Text) then
        if (Pos(Lowercase(PYREVIT_ADDIN_ID), Lowercase(Text)) > 0) or
           (Pos('pyrevitloader', Lowercase(Text)) > 0) then
        begin
          Found := True;
          PyRevitFound := True;
          Asm := ExtractTag(Text, 'Assembly');
          EnvLines.Add('  - Revit ' + Year + ': pyRevit conectado (' + Folder + '\' + FR.Name + ')');
          if Asm <> '' then
          begin
            if FileExists(Asm) then
            begin
              P := Pos('\bin\', Lowercase(Asm));
              if P > 0 then Root := Copy(Asm, 1, P - 1) else Root := ExtractFileDir(Asm);
              EnvLines.Add('      instalação do pyRevit: ' + Root);
            end
            else
              EnvLines.Add('      ATENÇÃO: o carregador indicado no .addin não existe: ' + Asm);
          end;
        end;
    until not FindNext(FR);
  finally
    FindClose(FR);
  end;
end;

procedure ScanUninstallKey(const RootKey: Integer; const RootName: String);
var Names: TArrayOfString; I: Integer; Key, DName, DVer: String;
begin
  Key := 'Software\Microsoft\Windows\CurrentVersion\Uninstall';
  if not RegGetSubkeyNames(RootKey, Key, Names) then Exit;
  for I := 0 to GetArrayLength(Names) - 1 do
    if RegQueryStringValue(RootKey, Key + '\' + Names[I], 'DisplayName', DName) then
      if Pos('pyrevit', Lowercase(DName)) > 0 then
      begin
        DVer := '';
        RegQueryStringValue(RootKey, Key + '\' + Names[I], 'DisplayVersion', DVer);
        PyRevitFound := True;
        EnvLines.Add('  - Programa instalado: ' + DName + ' ' + DVer + ' (' + RootName + ')');
      end;
end;

procedure FindExistingInstalls;
var I: Integer; FR: TFindRec; D: String;
begin
  ExistingInstalls.Clear;
  ConflictExts.Clear;
  for I := 0 to CandDirs.Count - 1 do
  begin
    D := CandDirs[I];
    if DirExists(D + '\' + EXT_NAME) then ExistingInstalls.Add(D + '\' + EXT_NAME);
    if FindFirst(D + '\*.extension', FR) then
    try
      repeat
        if ((FR.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0) and
           (CompareText(FR.Name, EXT_NAME) <> 0) and
           DirExists(D + '\' + FR.Name + '\OCA.tab') then
          ConflictExts.Add(D + '\' + FR.Name);
      until not FindNext(FR);
    finally
      FindClose(FR);
    end;
  end;
end;

procedure LoadPreviousState;
var Text, Line: String; L: TStringList; I, P1, P2: Integer;
begin
  PrevStateDir := '';
  PrevStateVersion := '';
  PrevStateExt := 'OCA.extension';   { instalações antigas não gravavam o nome }
  PrevCount := 0;
  if ReadTextFile(StateDir + '\estado_instalacao.txt', Text) then
  begin
    L := TStringList.Create;
    try
      L.Text := Text;
      for I := 0 to L.Count - 1 do
      begin
        Line := L[I];
        if Copy(Line, 1, 6) = 'pasta=' then PrevStateDir := Copy(Line, 7, Length(Line));
        if Copy(Line, 1, 7) = 'versao=' then PrevStateVersion := Copy(Line, 8, Length(Line));
        if Copy(Line, 1, 9) = 'extensao=' then PrevStateExt := Copy(Line, 10, Length(Line));
      end;
    finally
      L.Free;
    end;
  end;
  if ReadTextFile(StateDir + '\manifesto_instalado.txt', Text) then
  begin
    L := TStringList.Create;
    try
      L.Text := Text;
      for I := 0 to L.Count - 1 do
      begin
        { formato: sha256|tamanho|caminho relativo }
        Line := L[I];
        P1 := Pos('|', Line);
        if P1 = 0 then Continue;
        P2 := Pos('|', Copy(Line, P1 + 1, Length(Line)));
        if P2 = 0 then Continue;
        SetArrayLength(PrevRel, PrevCount + 1);
        SetArrayLength(PrevHash, PrevCount + 1);
        PrevHash[PrevCount] := Lowercase(Copy(Line, 1, P1 - 1));
        PrevRel[PrevCount] := Copy(Line, P1 + P2 + 1, Length(Line));
        PrevCount := PrevCount + 1;
      end;
    finally
      L.Free;
    end;
  end;
end;

procedure DetectEnvironment;
var Years: TArrayOfString; I: Integer; Found: Boolean; AdminRoot, Ovr: String;
begin
  EnvLines.Clear;
  CandDirs.Clear;
  CandLabels.Clear;
  RegisteredDirs.Clear;
  PyRevitFound := False;
  AttachedYears := '';

  Ovr := GetEnv('PYREVIT_PATH_OVERRIDE');
  if Ovr <> '' then PyRevitRoot := NormDir(Ovr)
  else PyRevitRoot := ExpandConstant('{userappdata}\pyRevit');
  Ovr := GetEnv('PYREVIT_PROGRAMDATA_PATH_OVERRIDE');
  if Ovr <> '' then AdminRoot := NormDir(Ovr)
  else AdminRoot := ExpandConstant('{commonappdata}\pyRevit');
  DefaultExtDir := PyRevitRoot + '\Extensions';
  { A pasta padrão é sempre carregada pelo pyRevit e vem primeiro na lista }
  AddCandidate(DefaultExtDir, 'pasta padrão do pyRevit');

  EnvLines.Add('VERIFICAÇÃO DO AMBIENTE');
  EnvLines.Add('Usuário do Windows: ' + GetUserNameString);
  if IsAdmin then
    EnvLines.Add('Instalador rodando COM privilégios de administrador. As pastas abaixo pertencem ao usuário acima.')
  else
    EnvLines.Add('Instalador rodando sem privilégios de administrador.');
  EnvLines.Add('');

  { 1. Configuração do pyRevit }
  EnvLines.Add('1) Configuração do pyRevit (pyRevit_config.ini)');
  ReadPyRevitConfig(PyRevitRoot + '\pyRevit_config.ini', 'configuração do usuário');
  ReadPyRevitConfig(AdminRoot + '\pyRevit_config.ini', 'configuração de administrador');
  if DirExists(PyRevitRoot) then PyRevitFound := True;
  EnvLines.Add('');

  { 2. pyRevit conectado ao Revit 2025/2026 (manifestos .addin) }
  EnvLines.Add('2) pyRevit conectado ao Revit (arquivos .addin)');
  SetArrayLength(Years, 2);
  Years[0] := '2025';
  Years[1] := '2026';
  for I := 0 to GetArrayLength(Years) - 1 do
  begin
    Found := False;
    ScanAddinFolder(ExpandConstant('{userappdata}\Autodesk\Revit\Addins\') + Years[I], Years[I], Found);
    ScanAddinFolder(ExpandConstant('{commonappdata}\Autodesk\Revit\Addins\') + Years[I], Years[I], Found);
    ScanAddinFolder(ExpandConstant('{commonpf64}\Autodesk\Revit\Addins\') + Years[I], Years[I], Found);
    if Found then
    begin
      if AttachedYears <> '' then AttachedYears := AttachedYears + ', ';
      AttachedYears := AttachedYears + Years[I];
    end
    else
      EnvLines.Add('  - Revit ' + Years[I] + ': pyRevit NÃO está conectado (nenhum .addin do pyRevit encontrado)');
    if FileExists(ExpandConstant('{commonpf64}\Autodesk\Revit ') + Years[I] + '\Revit.exe') then
      EnvLines.Add('      Revit ' + Years[I] + ' instalado em ' + ExpandConstant('{commonpf64}\Autodesk\Revit ') + Years[I])
    else
      EnvLines.Add('      Revit ' + Years[I] + ' não encontrado no local padrão (pode estar em outra pasta)');
  end;
  EnvLines.Add('');

  { 3. Programas instalados (Painel de Controle) }
  EnvLines.Add('3) pyRevit em Programas instalados');
  I := EnvLines.Count;
  ScanUninstallKey(HKCU, 'usuário');
  ScanUninstallKey(HKLM64, 'computador');
  ScanUninstallKey(HKLM32, 'computador 32 bits');
  if EnvLines.Count = I then EnvLines.Add('  - nenhum registro encontrado');
  EnvLines.Add('');

  { 4. Pastas de extensões }
  { (a pasta padrão e as registradas já foram adicionadas acima) }
  LoadPreviousState;
  if (PrevStateDir <> '') and (DirExists(PrevStateDir + '\' + EXT_NAME) or DirExists(PrevStateDir + '\' + PrevStateExt)) then
    AddCandidate(PrevStateDir, 'instalação anterior deste plugin');
  FindExistingInstalls;

  EnvLines.Add('4) Pastas onde o pyRevit procura extensões');
  for I := 0 to CandDirs.Count - 1 do
  begin
    EnvLines.Add('  - ' + CandDirs[I] + '  [' + CandLabels[I] + ']');
    if DirExists(CandDirs[I] + '\' + EXT_NAME) then
      EnvLines.Add('      já contém ' + EXT_NAME + ' (versão ' +
        DescribeVersion(VersionFromSobre(CandDirs[I] + '\' + EXT_NAME)) + ')')
    else if not DirExists(CandDirs[I]) then
      EnvLines.Add('      (a pasta ainda não existe; será criada na instalação)');
  end;
  if ConflictExts.Count > 0 then
  begin
    EnvLines.Add('');
    EnvLines.Add('ATENÇÃO: outra(s) extensão(ões) também contêm a aba OCA (OCA.tab):');
    for I := 0 to ConflictExts.Count - 1 do
      EnvLines.Add('  - ' + ConflictExts[I] + '  (versão ' + DescribeVersion(VersionFromSobre(ConflictExts[I])) + ')');
    EnvLines.Add('  Provavelmente são versões antigas do plugin com outro nome de pasta. Se ficarem,');
    EnvLines.Add('  a aba OCA terá botões DUPLICADOS. Mais adiante você poderá movê-las para o backup.');
  end;
  if ExistingInstalls.Count > 1 then
  begin
    EnvLines.Add('');
    EnvLines.Add('ATENÇÃO: há mais de uma ' + EXT_NAME + ' em pastas carregadas pelo pyRevit.');
    EnvLines.Add('  Mantenha apenas uma, para evitar botões duplicados ou conflito de versões.');
  end;

  RevitRunning := IsRevitRunningNow;
  EnvLines.Add('');
  if RevitRunning then
    EnvLines.Add('O Revit está ABERTO. Você pode instalar mesmo assim; ao final será preciso usar pyRevit > Reload ou reiniciar o Revit.')
  else
    EnvLines.Add('O Revit não está aberto.');
  EnvLines.Add('');
  if PyRevitFound then
  begin
    if AttachedYears <> '' then
      EnvLines.Add('RESULTADO: pyRevit encontrado e conectado ao Revit ' + AttachedYears + '.')
    else
      EnvLines.Add('RESULTADO: pyRevit encontrado, mas não está conectado ao Revit 2025 nem ao 2026. Os comandos só aparecem nas versões do Revit em que o pyRevit estiver conectado.');
  end
  else
    EnvLines.Add('RESULTADO: pyRevit NÃO encontrado neste computador.');
end;

{ ---------------------------------------------------------------------------
  Pasta de destino
  --------------------------------------------------------------------------- }
function IsRegisteredDir(const Dir: String): Boolean;
begin
  Result := SameDir(Dir, DefaultExtDir) or (ListIndexOfDir(RegisteredDirs, Dir) >= 0);
end;

function GetDefaultDir(Param: String): String;
var I: Integer;
begin
  Result := DefaultExtDir;
  { 1. pasta da instalação anterior feita por este instalador }
  if (PrevStateDir <> '') and (DirExists(PrevStateDir + '\' + EXT_NAME) or DirExists(PrevStateDir + '\' + PrevStateExt)) then
  begin
    Result := NormDir(PrevStateDir);
    Exit;
  end;
  { 2. primeira cópia do plugin (mesmo nome ou versão antiga) numa pasta do pyRevit }
  for I := 0 to ExistingInstalls.Count - 1 do
  begin
    Result := ExtractFileDir(ExistingInstalls[I]);
    Exit;
  end;
  for I := 0 to ConflictExts.Count - 1 do
  begin
    Result := ExtractFileDir(ConflictExts[I]);
    Exit;
  end;
end;

procedure RelaunchElevated(const Dir: String);
var ErrorCode: Integer;
begin
  if ShellExec('', ExpandConstant('{srcexe}'), '/ALLUSERS /DIR="' + Dir + '"', '', SW_SHOWNORMAL, ewNoWait, ErrorCode) then
  begin
    CloseWithoutConfirm := True;
    WizardForm.Close;
  end
  else
    MsgBox('Não foi possível reabrir o instalador como administrador (código ' + IntToStr(ErrorCode) + ').' + CRLF +
      'Feche o instalador, clique nele com o botão direito e escolha "Executar como administrador", ou escolha outra pasta.',
      mbError, MB_OK);
end;

{ Valida a pasta de extensões escolhida. Retorna False para ficar na página. }
function ValidateTargetDir(const Dir: String): Boolean;
var D, Msg: String; I: Integer;
begin
  Result := False;
  D := NormDir(Dir);
  if D = '' then
  begin
    SuppressibleMsgBox('Informe a pasta de extensões do pyRevit.', mbError, MB_OK, IDOK);
    Exit;
  end;
  if (Length(D) <= 3) then
  begin
    SuppressibleMsgBox('Escolha uma pasta, não a raiz de um disco (' + D + ').', mbError, MB_OK, IDOK);
    Exit;
  end;
  if CompareText(ExtractFileExt(D), '.extension') = 0 then
  begin
    SuppressibleMsgBox('A pasta escolhida termina em ".extension":' + CRLF + D + CRLF + CRLF +
      'Escolha a pasta que CONTÉM as extensões (por exemplo, ' + DefaultExtDir + '), e não a pasta de uma extensão. ' +
      'O instalador cria a subpasta ' + EXT_NAME + ' sozinho.', mbError, MB_OK, IDOK);
    Exit;
  end;
  if (Pos('\pyrevit-master\', Lowercase(D) + '\') > 0) or DirExists(ExtractFileDir(D) + '\pyrevitlib') then
  begin
    if SuppressibleMsgBox('A pasta escolhida parece ser uma pasta INTERNA do pyRevit:' + CRLF + D + CRLF + CRLF +
      'Ela é substituída quando o pyRevit é atualizado, e o plugin seria apagado junto. ' +
      'Recomenda-se usar ' + DefaultExtDir + '.' + CRLF + CRLF + 'Deseja usar esta pasta mesmo assim?',
      mbConfirmation, MB_YESNO or MB_DEFBUTTON2, IDNO) <> IDYES then Exit;
  end;
  if not IsRegisteredDir(D) then
  begin
    if SuppressibleMsgBox('O pyRevit NÃO procura extensões nesta pasta:' + CRLF + D + CRLF + CRLF +
      'Ela não é a pasta padrão (' + DefaultExtDir + ') nem está registrada no pyRevit_config.ini. ' +
      'Depois de instalar, será preciso adicioná-la no Revit em pyRevit > Settings > ' +
      '"Custom Extension Directories" (o relatório final explica como).' + CRLF + CRLF +
      'Deseja continuar com esta pasta?', mbConfirmation, MB_YESNO or MB_DEFBUTTON2, IDYES) <> IDYES then Exit;
  end;
  for I := 0 to ExistingInstalls.Count - 1 do
    if not SameDir(ExtractFileDir(ExistingInstalls[I]), D) then
    begin
      Msg := 'Já existe uma ' + EXT_NAME + ' em OUTRA pasta carregada pelo pyRevit:' + CRLF + ExistingInstalls[I] + CRLF + CRLF +
        'Se você instalar em ' + D + ', o pyRevit encontrará duas cópias do plugin. O instalador não remove a outra cópia. ' +
        'Recomenda-se instalar na mesma pasta (para atualizá-la) ou remover a cópia antiga manualmente.' + CRLF + CRLF +
        'Deseja continuar com ' + D + '?';
      if SuppressibleMsgBox(Msg, mbConfirmation, MB_YESNO or MB_DEFBUTTON2, IDYES) <> IDYES then Exit;
    end;
  if not CanWriteTo(D) then
  begin
    if IsAdmin then
    begin
      SuppressibleMsgBox('Não é possível gravar na pasta:' + CRLF + D + CRLF + CRLF +
        'Mesmo com privilégios de administrador o acesso foi negado. Verifique as permissões da pasta ' +
        'ou escolha outra.', mbError, MB_OK, IDOK);
      Exit;
    end;
    if WizardSilent then
    begin
      Log('Sem permissão de gravação em ' + D + '. Rode com /ALLUSERS para instalar como administrador.');
      Exit;
    end;
    if MsgBox('Sua conta não tem permissão para gravar na pasta:' + CRLF + D + CRLF + CRLF +
      'Para instalar nela é preciso permissão de administrador. Deseja reabrir o instalador como administrador? ' +
      '(O Windows vai pedir confirmação.)' + CRLF + CRLF +
      'Clique em "Não" para escolher outra pasta.', mbConfirmation, MB_YESNO) = IDYES then
      RelaunchElevated(D);
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------
  Páginas do assistente
  --------------------------------------------------------------------------- }
function OtherDirIndex: Integer;
begin
  Result := CandDirs.Count;   { última opção: "Outra pasta..." }
end;

procedure FillDirChoicePage;
var I, Sel: Integer; Lbl, CmdDir, Def: String;
begin
  for I := 0 to CandDirs.Count - 1 do
  begin
    Lbl := CandDirs[I] + '   (' + CandLabels[I] + ')';
    if DirExists(CandDirs[I] + '\' + EXT_NAME) then
      Lbl := Lbl + ' - contém ' + EXT_NAME + ' ' + DescribeVersion(VersionFromSobre(CandDirs[I] + '\' + EXT_NAME));
    DirChoicePage.Add(Lbl);
  end;
  DirChoicePage.Add('Outra pasta (escolher na próxima tela)');

  CmdDir := ExpandConstant('{param:DIR|}');
  if CmdDir <> '' then Def := CmdDir else Def := GetDefaultDir('');
  Sel := ListIndexOfDir(CandDirs, Def);
  if Sel < 0 then Sel := OtherDirIndex;
  DirChoicePage.SelectedValueIndex := Sel;
  if Sel = OtherDirIndex then WizardForm.DirEdit.Text := NormDir(Def);
end;

procedure InitializeWizard;
var Sub: String; I: Integer;
begin
  EnvPage := CreateOutputMsgMemoPage(wpInfoBefore,
    'Verificação do pyRevit',
    'Resultado da busca pelo pyRevit e pelas pastas de extensões neste computador.',
    'Nada foi alterado até aqui. Confira o diagnóstico e clique em Avançar.', '');
  EnvPage.RichEditViewer.Lines.Text := EnvLines.Text;

  if CandDirs.Count > 1 then
    Sub := 'O pyRevit carrega extensões das pastas abaixo. Escolha onde instalar o plugin. ' +
      'Se já houver uma versão instalada, escolha a pasta dela para atualizá-la.'
  else
    Sub := 'O pyRevit carrega extensões da pasta abaixo (pasta padrão, recomendada). ' +
      'Use "Outra pasta" apenas se o seu pyRevit estiver configurado com uma pasta de extensões personalizada.';
  DirChoicePage := CreateInputOptionPage(EnvPage.ID,
    'Pasta de extensões do pyRevit',
    'Onde o plugin deve ser instalado?', Sub, True, False);
  FillDirChoicePage;

  ExistingPage := CreateInputOptionPage(wpSelectDir,
    'Instalação existente encontrada',
    'Já existe uma versão do plugin nesta pasta. O que deseja fazer?', '', True, False);
  ExistingPage.Add('Atualizar: substitui os arquivos do plugin e mantém arquivos extras que você tenha adicionado (recomendado)');
  ExistingPage.Add('Reinstalação limpa: move todo o conteúdo atual para o backup e instala do zero');
  ExistingPage.Add('Cancelar a instalação (nada será alterado)');
  { Em instalações silenciosas: /MODO=limpa faz reinstalação limpa; o padrão é atualizar. }
  if CompareText(ExpandConstant('{param:MODO|atualizar}'), 'limpa') = 0 then
    ExistingPage.SelectedValueIndex := 1
  else
    ExistingPage.SelectedValueIndex := 0;

  OldPage := CreateInputOptionPage(ExistingPage.ID,
    'Versões antigas do plugin',
    'Outras pastas com a aba OCA foram encontradas.',
    'Estas pastas também criam a aba OCA no Revit. Se ficarem junto com ' + EXT_NAME + ', os botões ' +
    'aparecerão DUPLICADOS. As marcadas serão copiadas para o backup (' + StateDir + '\Backups), conferidas ' +
    'e só então retiradas da pasta de extensões. Desmarque as que não forem versões antigas deste plugin.',
    False, False);
  for I := 0 to ConflictExts.Count - 1 do
  begin
    OldPage.Add(ConflictExts[I] + '   (versão ' + DescribeVersion(VersionFromSobre(ConflictExts[I])) + ')');
    { Em instalações silenciosas: /ANTIGAS=manter deixa as versões antigas onde estão }
    OldPage.Values[I] := CompareText(ExpandConstant('{param:ANTIGAS|mover}'), 'manter') <> 0;
  end;

  ReportPage := CreateOutputMsgMemoPage(wpInstalling,
    'Relatório da instalação',
    'Resultado da conferência dos arquivos instalados.',
    'Este relatório também foi salvo em arquivo (caminho no final do texto).', '');
end;

function SelectedTargetDir: String;
begin
  Result := NormDir(WizardDirValue);
end;

procedure UpdateExistingPageText;
var Sub, V: String;
begin
  TargetExt := SelectedTargetDir + '\' + EXT_NAME;
  V := VersionFromSobre(TargetExt);
  InstalledVersion := V;
  Sub := 'Pasta: ' + TargetExt + CRLF +
    'Versão instalada: ' + DescribeVersion(V) + '      Nova versão: ' + NEW_VERSION + CRLF;
  if (V <> '') and (IsDateVersion(V) <> IsDateVersion(NEW_VERSION)) then
    Sub := Sub + 'A versão instalada usa a numeração antiga (por data); a atualização é segura.' + CRLF
  else if (V <> '') and (CompareVersions(V, NEW_VERSION) > 0) then
    Sub := Sub + 'ATENÇÃO: a versão instalada é MAIS NOVA do que a deste instalador.' + CRLF
  else if (V <> '') and (CompareVersions(V, NEW_VERSION) = 0) then
    Sub := Sub + 'Esta mesma versão já está instalada (a atualização reinstala os arquivos).' + CRLF;
  Sub := Sub + 'Antes de qualquer alteração, um backup completo da pasta atual será salvo em ' +
    StateDir + '\Backups. As configurações dos comandos (guardadas pelo pyRevit fora desta pasta) não são alteradas.';
  { O rótulo tem altura fixa: ajusta ao texto e desce a lista de opções }
  ExistingPage.SubCaptionLabel.Caption := Sub;
  WizardForm.AdjustLabelHeight(ExistingPage.SubCaptionLabel);
  ExistingPage.CheckListBox.Top := ExistingPage.SubCaptionLabel.Top + ExistingPage.SubCaptionLabel.Height + ScaleY(12);
  ExistingPage.CheckListBox.Height := ExistingPage.SurfaceHeight - ExistingPage.CheckListBox.Top;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if PageID = wpSelectDir then
    Result := DirChoicePage.SelectedValueIndex <> OtherDirIndex
  else if PageID = ExistingPage.ID then
    Result := not DirExists(SelectedTargetDir + '\' + EXT_NAME)
  else if PageID = OldPage.ID then
    Result := ConflictExts.Count = 0;
end;

procedure CurPageChanged(CurPageID: Integer);
var Extra: String;
begin
  Extra := '';
  if IndexInPkg('Instaladores\Instalar impressora OCA PDF24.ps1') >= 0 then
    Extra := '- Exportação em PDF pelo PDF24: veja o item 5 do relatório.' + CRLF;
  if CurPageID = ExistingPage.ID then
    UpdateExistingPageText
  else if CurPageID = wpFinished then
  begin
    if VerifyOk then
      WizardForm.FinishedLabel.Caption :=
        'O plugin foi instalado e conferido com sucesso em:' + CRLF + TargetExt + CRLF + CRLF +
        'Próximos passos:' + CRLF +
        '- Revit fechado: basta abrir o Revit 2025/2026. A aba OCA aparece na faixa de opções.' + CRLF +
        '- Revit aberto: clique em pyRevit > Reload (ou feche e abra o Revit).' + CRLF +
        '- Para conferir a versão, clique em OCA > Sobre.' + CRLF + Extra + CRLF +
        'Relatório completo: ' + ReportFile
    else
      WizardForm.FinishedLabel.Caption :=
        'A instalação terminou COM PROBLEMAS: alguns arquivos não conferem com o pacote.' + CRLF + CRLF +
        'Veja o relatório em: ' + ReportFile + CRLF + CRLF +
        'Feche o Revit e programas que possam estar usando arquivos da pasta (ex.: Excel) e execute o instalador novamente.';
    WizardForm.AdjustLabelHeight(WizardForm.FinishedLabel);
  end;
end;

procedure CancelButtonClick(CurPageID: Integer; var Cancel, Confirm: Boolean);
begin
  if CloseWithoutConfirm then
  begin
    Confirm := False;
    Cancel := True;
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = EnvPage.ID then
  begin
    if not PyRevitFound then
    begin
      if SuppressibleMsgBox('O pyRevit não foi encontrado neste computador.' + CRLF + CRLF +
        'O Plugin OCA é uma extensão do pyRevit e só funciona com ele instalado e conectado ao Revit ' +
        '(2025/2026). Sem o pyRevit, os arquivos serão copiados, mas nenhum comando aparecerá no Revit.' + CRLF + CRLF +
        'Recomenda-se cancelar, instalar o pyRevit (https://github.com/pyrevitlabs/pyRevit/releases) e ' +
        'executar este instalador novamente.' + CRLF + CRLF +
        'Deseja continuar mesmo assim?', mbConfirmation, MB_YESNO or MB_DEFBUTTON2, IDYES) <> IDYES then
      begin
        CloseWithoutConfirm := True;
        WizardForm.Close;
        Result := False;
      end;
    end;
  end
  else if CurPageID = DirChoicePage.ID then
  begin
    if DirChoicePage.SelectedValueIndex <> OtherDirIndex then
    begin
      WizardForm.DirEdit.Text := CandDirs[DirChoicePage.SelectedValueIndex];
      Result := ValidateTargetDir(WizardForm.DirEdit.Text);
    end;
  end
  else if CurPageID = wpSelectDir then
    Result := ValidateTargetDir(WizardDirValue)
  else if CurPageID = ExistingPage.ID then
  begin
    if ExistingPage.SelectedValueIndex = 2 then
    begin
      if SuppressibleMsgBox('Cancelar a instalação? Nenhum arquivo foi alterado.', mbConfirmation, MB_YESNO, IDYES) = IDYES then
      begin
        CloseWithoutConfirm := True;
        WizardForm.Close;
      end;
      Result := False;
    end;
  end;
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo,
  MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var S, T: String; I: Integer;
begin
  T := SelectedTargetDir + '\' + EXT_NAME;
  S := 'Plugin:' + NewLine + Space + '{#AppName} ' + NEW_VERSION + NewLine + NewLine +
    'Pasta de extensões do pyRevit:' + NewLine + Space + SelectedTargetDir + NewLine + NewLine +
    'O plugin será instalado em:' + NewLine + Space + T + NewLine + NewLine +
    'Arquivos no pacote:' + NewLine + Space + IntToStr(PkgCount) + NewLine + NewLine;
  if DirExists(T) then
  begin
    S := S + 'Instalação existente:' + NewLine + Space + 'versão ' + DescribeVersion(VersionFromSobre(T)) + NewLine;
    if ExistingPage.SelectedValueIndex = 1 then
      S := S + Space + 'Ação: reinstalação limpa (conteúdo atual vai para o backup)' + NewLine
    else
      S := S + Space + 'Ação: atualizar' + NewLine;
    S := S + Space + 'Backup em: ' + StateDir + '\Backups' + NewLine + NewLine;
  end
  else
    S := S + 'Ação:' + NewLine + Space + 'nova instalação' + NewLine + NewLine;
  for I := 0 to ConflictExts.Count - 1 do
    if OldPage.Values[I] then
      S := S + 'Versão antiga a retirar (vai para o backup):' + NewLine + Space + ConflictExts[I] + NewLine + NewLine
    else
      S := S + 'ATENÇÃO: versão antiga MANTIDA (botões duplicados):' + NewLine + Space + ConflictExts[I] + NewLine + NewLine;
  if not IsRegisteredDir(SelectedTargetDir) then
    S := S + 'ATENÇÃO: esta pasta não está registrada no pyRevit (registre-a depois em pyRevit > Settings).' + NewLine + NewLine;
  if IsRevitRunningNow then
    S := S + 'O Revit está aberto: depois de instalar, use pyRevit > Reload ou reinicie o Revit.' + NewLine;
  S := S + 'O instalador não executa o Revit nem outros programas.';
  Result := S;
end;

{ ---------------------------------------------------------------------------
  Instalação: backup, atualização e conferência
  --------------------------------------------------------------------------- }
function CopyTreeVerified(const Src, Dest: String; var Err: String): Boolean;
var Files: TStringList; I: Integer; S, D: String;
begin
  Result := False;
  Files := TStringList.Create;
  try
    ListFilesRec(Src, '', Files);
    if not ForceDirectories(Dest) then
    begin
      Err := 'não foi possível criar a pasta de backup ' + Dest;
      Exit;
    end;
    for I := 0 to Files.Count - 1 do
    begin
      S := Src + '\' + Files[I];
      D := Dest + '\' + Files[I];
      ForceDirectories(ExtractFileDir(D));
      if not CopyFile(S, D, False) then
      begin
        Err := 'falha ao copiar ' + S + ' para o backup';
        Exit;
      end;
      if SafeSHA256(S) <> SafeSHA256(D) then
      begin
        Err := 'a cópia de backup de ' + Files[I] + ' não confere';
        Exit;
      end;
    end;
    Log(Format('Backup: %d arquivo(s) copiados e conferidos em %s', [Files.Count, Dest]));
    Result := True;
  finally
    Files.Free;
  end;
end;

procedure RemoveEmptyParents(const Root, Rel: String);
var D: String;
begin
  D := ExtractFileDir(Root + '\' + Rel);
  while (Length(D) > Length(Root)) and (CompareText(Copy(D, 1, Length(Root)), Root) = 0) do
  begin
    if not RemoveDir(D) then Exit;   { só remove se estiver vazia }
    D := ExtractFileDir(D);
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var Existing: TStringList; I, K: Integer; Err, Rel, H, OldDir: String; HavePrev: Boolean;
begin
  Result := '';
  RemovedObsolete.Clear;
  KeptModified.Clear;
  KeptExtras.Clear;
  OverwrittenModified.Clear;
  UpdatedFiles.Clear;
  MovedToBackup.Clear;
  MigratedOld.Clear;
  BackupDir := '';
  TargetExt := SelectedTargetDir + '\' + EXT_NAME;

  if not CanWriteTo(SelectedTargetDir) then
  begin
    Result := 'Sem permissão para gravar em ' + SelectedTargetDir + '.' + CRLF +
      'Escolha outra pasta ou execute o instalador como administrador.';
    Exit;
  end;

  { 0. Versões antigas com outro nome de pasta: backup conferido e retirada }
  for I := 0 to ConflictExts.Count - 1 do
  begin
    if not OldPage.Values[I] then Continue;
    OldDir := ConflictExts[I];
    if not DirExists(OldDir) then Continue;
    Err := StateDir + '\Backups\' + ExtractFileName(OldDir) + '_' + DescribeVersion(VersionFromSobre(OldDir)) + '_' +
      GetDateTimeString('yyyy-mm-dd_hh-nn-ss', '-', '-');
    Rel := Err;
    if not CopyTreeVerified(OldDir, Rel, Err) then
    begin
      Result := 'Não foi possível fazer o backup da versão antiga ' + OldDir + ' (' + Err + ').' + CRLF + CRLF +
        'Nada foi alterado nessa pasta. Verifique o espaço em disco e tente novamente, ou desmarque essa pasta.';
      Exit;
    end;
    H := OldDir + '.antigo-' + GetDateTimeString('yyyymmddhhnnss', #0, #0);
    if not RenameFile(OldDir, H) then
    begin
      Result := 'Não foi possível retirar a versão antiga:' + CRLF + OldDir + CRLF + CRLF +
        'Feche o Revit e programas que possam estar usando arquivos dessa pasta e tente de novo. ' +
        'Há uma cópia completa em:' + CRLF + Rel;
      Exit;
    end;
    if not DelTree(H, True, True, True) then
      KeptModified.Add('A pasta antiga não pôde ser apagada totalmente e pode ser removida à mão: ' + H);
    MigratedOld.Add(OldDir + '  ->  backup: ' + Rel);
  end;

  if not DirExists(TargetExt) then
  begin
    InstallMode := MODE_NEW;
    Exit;
  end;

  InstalledVersion := VersionFromSobre(TargetExt);
  if ExistingPage.SelectedValueIndex = 1 then InstallMode := MODE_CLEAN else InstallMode := MODE_UPDATE;

  { 1. Backup completo e conferido da pasta atual }
  BackupDir := StateDir + '\Backups\' + EXT_NAME + '_' + DescribeVersion(InstalledVersion) + '_' +
    GetDateTimeString('yyyy-mm-dd_hh-nn-ss', '-', '-');
  if not CopyTreeVerified(TargetExt, BackupDir, Err) then
  begin
    Result := 'Não foi possível fazer o backup da instalação existente (' + Err + ').' + CRLF + CRLF +
      'Nada foi alterado na pasta do plugin. Verifique o espaço em disco e tente novamente.';
    Exit;
  end;

  { O manifesto anterior só vale se a instalação anterior foi nesta mesma pasta }
  HavePrev := (PrevCount > 0) and SameDir(PrevStateDir, SelectedTargetDir) and (CompareText(PrevStateExt, EXT_NAME) = 0);

  Existing := TStringList.Create;
  try
    ListFilesRec(TargetExt, '', Existing);

    if InstallMode = MODE_CLEAN then
    begin
      { Renomeia primeiro (operação atômica: ou tudo, ou nada) e só depois apaga.
        O nome temporário não termina em ".extension", então o pyRevit o ignora. }
      OldDir := TargetExt + '.antigo-' + GetDateTimeString('yyyymmddhhnnss', #0, #0);
      if not RenameFile(TargetExt, OldDir) then
      begin
        Result := 'Não foi possível liberar a pasta atual do plugin para a reinstalação limpa:' + CRLF + TargetExt + CRLF + CRLF +
          'Feche o Revit e programas que possam estar usando arquivos dessa pasta (por exemplo, o Excel com a planilha ' +
          'modelo aberta) e tente de novo. Nada foi alterado. Há também uma cópia completa em:' + CRLF + BackupDir;
        Exit;
      end;
      if not DelTree(OldDir, True, True, True) then
        KeptModified.Add('A pasta antiga não pôde ser apagada totalmente e pode ser removida à mão: ' + OldDir);
      for I := 0 to Existing.Count - 1 do
        if IndexInPkg(Existing[I]) < 0 then MovedToBackup.Add(Existing[I]);
      Exit;
    end;

    { MODE_UPDATE }
    for I := 0 to Existing.Count - 1 do
    begin
      Rel := Existing[I];
      if IndexInPkg(Rel) >= 0 then
      begin
        { arquivo do pacote: será substituído; avisa se foi alterado localmente }
        H := SafeSHA256(TargetExt + '\' + Rel);
        if H <> PkgHash[IndexInPkg(Rel)] then UpdatedFiles.Add(Rel);
        if HavePrev then
        begin
          K := IndexInPrev(Rel);
          if K >= 0 then
          begin
            if (H <> PrevHash[K]) and (H <> PkgHash[IndexInPkg(Rel)]) then
              OverwrittenModified.Add(Rel);
          end;
        end;
        Continue;
      end;
      { arquivo que não está no novo pacote }
      K := -1;
      if HavePrev then K := IndexInPrev(Rel);
      if K < 0 then
        KeptExtras.Add(Rel)                          { não foi este instalador que criou: mantém }
      else if SafeSHA256(TargetExt + '\' + Rel) = PrevHash[K] then
      begin
        if DeleteFile(TargetExt + '\' + Rel) then
        begin
          RemovedObsolete.Add(Rel);                  { era da versão anterior e não mudou: remove }
          RemoveEmptyParents(TargetExt, Rel);
        end
        else
          KeptModified.Add(Rel + '  (não foi possível remover)');
      end
      else
        KeptModified.Add(Rel);                       { era da versão anterior mas foi alterado: mantém }
    end;
  finally
    Existing.Free;
  end;
end;

procedure AddList(L: TStringList; const Title: String; Items: TStringList);
var I: Integer;
begin
  if Items.Count = 0 then Exit;
  L.Add('');
  L.Add(Title + ' (' + IntToStr(Items.Count) + '):');
  for I := 0 to Items.Count - 1 do L.Add('  ' + Items[I]);
end;

function ButtonTitle(const ScriptFile, Fallback: String): String;
var Text, Line, Q: String; L: TStringList; I, P: Integer;
begin
  Result := Fallback;
  if not ReadTextFile(ScriptFile, Text) then Exit;
  L := TStringList.Create;
  try
    L.Text := Text;
    for I := 0 to L.Count - 1 do
    begin
      Line := Trim(L[I]);
      if Copy(Line, 1, 9) <> '__title__' then Continue;
      P := Pos('=', Line);
      if P = 0 then Exit;
      Line := Trim(Copy(Line, P + 1, Length(Line)));
      if Copy(Line, 1, 1) = 'u' then Line := Copy(Line, 2, Length(Line));
      if Line = '' then Exit;
      Q := Line[1];
      if (Q <> '"') and (Q <> '''') then Exit;
      Line := Copy(Line, 2, Length(Line));
      P := Pos(Q, Line);
      if P = 0 then Exit;
      Line := Copy(Line, 1, P - 1);
      StringChangeEx(Line, '\n', ' ', True);
      Result := Line;
      Exit;
    end;
  finally
    L.Free;
  end;
end;

{ title: "..." de um bundle.yaml (só a chave de primeiro nível) }
function BundleTitle(const BundleFile: String): String;
var Text, Line, Q: String; L: TStringList; I, P: Integer;
begin
  Result := '';
  if not ReadTextFile(BundleFile, Text) then Exit;
  L := TStringList.Create;
  try
    L.Text := Text;
    for I := 0 to L.Count - 1 do
    begin
      Line := L[I];
      if Copy(Line, 1, 6) <> 'title:' then Continue;
      Line := Trim(Copy(Line, 7, Length(Line)));
      if Line = '' then Exit;
      Q := Line[1];
      if (Q = '"') or (Q = '''') then
      begin
        Line := Copy(Line, 2, Length(Line));
        P := Pos(Q, Line);
        if P > 0 then Line := Copy(Line, 1, P - 1);
      end;
      StringChangeEx(Line, '\n', ' ', True);
      Result := Trim(Line);
      Exit;
    end;
  finally
    L.Free;
  end;
end;

{ "OCA.tab\Vistas.panel\01_Abertura de Vista.pushbutton\script.py" -> linha do guia }
procedure AddCommandGuide(L: TStringList);
var I, P: Integer; Rel, Panel, Btn, Rest, Title: String;
begin
  for I := 0 to PkgCount - 1 do
  begin
    Rel := PkgRel[I];
    if CompareText(ExtractFileName(Rel), 'script.py') <> 0 then Continue;
    if CompareText(Copy(Rel, 1, 8), 'OCA.tab\') <> 0 then Continue;
    Rest := Copy(Rel, 9, Length(Rel));
    P := Pos('.panel\', Rest);
    if P = 0 then Continue;
    Panel := Copy(Rest, 1, P - 1);
    Btn := ExtractFileName(ExtractFileDir(Rel));
    Btn := Copy(Btn, 1, Length(Btn) - Length(ExtractFileExt(Btn)));
    Title := BundleTitle(TargetExt + '\' + ExtractFileDir(Rel) + '\bundle.yaml');
    if Title = '' then Title := ButtonTitle(TargetExt + '\' + Rel, Btn);
    L.Add('  - Painel ' + Panel + ': botão "' + Title + '"');
  end;
end;

procedure VerifyAndReport;
var L, Manifest, Others: TStringList; I, OkCount: Integer; P, H, ModeText: String;
begin
  VerifyFailures.Clear;
  OkCount := 0;
  L := TStringList.Create;
  Manifest := TStringList.Create;
  try
    { Conferência arquivo a arquivo (existência + SHA-256) }
    for I := 0 to PkgCount - 1 do
    begin
      P := TargetExt + '\' + PkgRel[I];
      if not FileExists(P) then
        VerifyFailures.Add(PkgRel[I] + '  -> AUSENTE')
      else
      begin
        H := SafeSHA256(P);
        if H <> PkgHash[I] then
          VerifyFailures.Add(PkgRel[I] + '  -> CONTEÚDO DIFERENTE do pacote')
        else
          OkCount := OkCount + 1;
      end;
      Manifest.Add(PkgHash[I] + '|' + IntToStr(PkgSize[I]) + '|' + PkgRel[I]);
    end;
    VerifyOk := (VerifyFailures.Count = 0) and (OkCount = PkgCount);

    case InstallMode of
      MODE_UPDATE: ModeText := 'Atualização (versão anterior: ' + DescribeVersion(InstalledVersion) + ')';
      MODE_CLEAN:  ModeText := 'Reinstalação limpa (versão anterior: ' + DescribeVersion(InstalledVersion) + ')';
    else
      if MigratedOld.Count > 0 then
        ModeText := 'Nova instalação em ' + EXT_NAME + ', substituindo versão antiga com outro nome de pasta'
      else
        ModeText := 'Nova instalação';
    end;

    L.Add('RELATÓRIO DE INSTALAÇÃO - {#AppName}');
    L.Add('Versão instalada: ' + NEW_VERSION);
    L.Add('Data: ' + GetDateTimeString('dd/mm/yyyy hh:nn:ss', '/', ':'));
    L.Add('Usuário do Windows: ' + GetUserNameString);
    L.Add('');
    if VerifyOk then
      L.Add('RESULTADO: SUCESSO - ' + IntToStr(OkCount) + ' de ' + IntToStr(PkgCount) + ' arquivos conferidos (SHA-256).')
    else
      L.Add('RESULTADO: FALHA - ' + IntToStr(VerifyFailures.Count) + ' arquivo(s) com problema; ' +
        IntToStr(OkCount) + ' de ' + IntToStr(PkgCount) + ' conferidos.');
    L.Add('');
    L.Add('Pasta de extensões do pyRevit: ' + SelectedTargetDir);
    L.Add('Pasta do plugin: ' + TargetExt);
    L.Add('Tipo: ' + ModeText);
    if BackupDir <> '' then L.Add('Backup da versão anterior: ' + BackupDir);
    if IsRegisteredDir(SelectedTargetDir) then
      L.Add('Pasta carregada pelo pyRevit: sim')
    else
    begin
      L.Add('Pasta carregada pelo pyRevit: NÃO. Para registrar: no Revit, pyRevit > Settings >');
      L.Add('  "Custom Extension Directories" > adicione ' + SelectedTargetDir + ' > Save Settings and Reload.');
    end;

    AddList(L, 'PROBLEMAS ENCONTRADOS NA CONFERÊNCIA', VerifyFailures);
    AddList(L, 'Versões antigas do plugin retiradas da pasta de extensões (cópia completa no backup)', MigratedOld);
    Others := TStringList.Create;
    try
      for I := 0 to ConflictExts.Count - 1 do
        if DirExists(ConflictExts[I]) then Others.Add(ConflictExts[I]);
      AddList(L, 'ATENÇÃO: estas extensões também têm a aba OCA e foram MANTIDAS (botões podem aparecer ' +
        'DUPLICADOS). Se forem cópias antigas do plugin, rode o instalador de novo e marque-as, ou apague-as', Others);
    finally
      Others.Free;
    end;
    AddList(L, 'Arquivos com conteúdo diferente da versão anterior, ATUALIZADOS', UpdatedFiles);
    AddList(L, 'Arquivos obsoletos da versão anterior removidos', RemovedObsolete);
    AddList(L, 'Arquivos da versão anterior alterados localmente e MANTIDOS (revise)', KeptModified);
    AddList(L, 'Arquivos alterados localmente que foram SUBSTITUÍDOS (o original está no backup)', OverwrittenModified);
    AddList(L, 'Arquivos que não fazem parte do plugin, MANTIDOS sem alteração', KeptExtras);
    AddList(L, 'Arquivos que não fazem parte do pacote: NÃO foram reinstalados (cópia no backup acima)', MovedToBackup);

    L.Add('');
    L.Add('ARQUIVOS INSTALADOS (' + IntToStr(PkgCount) + '):');
    for I := 0 to PkgCount - 1 do
      L.Add('  ' + PkgRel[I] + '  (' + IntToStr(PkgSize[I]) + ' bytes)');

    L.Add('');
    L.Add('PRÓXIMOS PASSOS');
    L.Add('1. Revit fechado: abra o Revit 2025 ou 2026 normalmente (não é preciso abrir como administrador).');
    L.Add('   Revit aberto: clique na aba pyRevit > Reload, ou feche e abra o Revit.');
    L.Add('2. Confirme que a aba "OCA" aparece na faixa de opções do Revit, com os comandos:');
    AddCommandGuide(L);
    L.Add('3. Clique em OCA > Sobre: a janela deve mostrar a versão ' + NEW_VERSION + '.');
    L.Add('4. Se a aba não aparecer: confira se o pyRevit está conectado a essa versão do Revit');
    L.Add('   (pyRevit > Settings, ou o comando "pyrevit attached" no Prompt de Comando) e se a pasta acima');
    L.Add('   está entre as pastas de extensões do pyRevit.');
    if IndexInPkg('Instaladores\Instalar impressora OCA PDF24.ps1') >= 0 then
    begin
      L.Add('5. Exportação DWG/PDF com a opção PDF24 precisa da impressora "OCA PDF24" (e do PDF24 Creator).');
      L.Add('   Este instalador NÃO a instala nem executa scripts. O próprio botão oferece instalá-la, ou execute');
      L.Add('   ' + TargetExt + '\Instaladores\Instalar impressora OCA PDF24.cmd');
      L.Add('   (pede administrador; altera configurações do PDF24 e cria uma impressora).');
    end;
    L.Add('');
    L.Add('As configurações dos comandos (pastas, opções salvas) ficam no pyRevit_config.ini e nas pastas de');
    L.Add('dados do pyRevit, fora da pasta do plugin, e não foram alteradas.');

    ReportFile := StateDir + '\relatorio_instalacao_' + GetDateTimeString('yyyy-mm-dd_hh-nn-ss', '-', '-') + '.txt';
    L.Add('');
    L.Add('Relatório salvo em: ' + ReportFile);
    ReportText := L.Text;
    if not WriteTextFileUTF8(ReportFile, ReportText) then
      Log('Não foi possível salvar o relatório em ' + ReportFile);
    Log(ReportText);

    { Estado para a próxima atualização }
    WriteTextFileUTF8(StateDir + '\manifesto_instalado.txt', Manifest.Text);
    WriteTextFileUTF8(StateDir + '\estado_instalacao.txt',
      'versao=' + NEW_VERSION + CRLF + 'pasta=' + SelectedTargetDir + CRLF + 'extensao=' + EXT_NAME + CRLF +
      'data=' + GetDateTimeString('yyyy-mm-dd hh:nn:ss', '-', ':') + CRLF);
  finally
    L.Free;
    Manifest.Free;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    VerifyAndReport;
    ReportPage.RichEditViewer.Lines.Text := ReportText;
    if not VerifyOk then
      SuppressibleMsgBox('A instalação terminou, mas a conferência encontrou ' + IntToStr(VerifyFailures.Count) +
        ' arquivo(s) com problema. Veja o relatório na próxima tela.' + CRLF + CRLF +
        'Feche o Revit e programas que possam estar usando arquivos da pasta do plugin e execute o instalador novamente.',
        mbError, MB_OK, IDOK);
  end;
end;

function InitializeSetup: Boolean;
begin
  CandDirs := TStringList.Create;
  CandLabels := TStringList.Create;
  RegisteredDirs := TStringList.Create;
  EnvLines := TStringList.Create;
  ExistingInstalls := TStringList.Create;
  ConflictExts := TStringList.Create;
  MigratedOld := TStringList.Create;
  RemovedObsolete := TStringList.Create;
  KeptModified := TStringList.Create;
  KeptExtras := TStringList.Create;
  OverwrittenModified := TStringList.Create;
  UpdatedFiles := TStringList.Create;
  MovedToBackup := TStringList.Create;
  VerifyFailures := TStringList.Create;
  LoadPackageManifest;
  DetectEnvironment;
  Log(EnvLines.Text);
  Result := True;
end;
