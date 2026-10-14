<#
.SYNOPSIS
  Gera dist\Instalador_Plugin_OCA.exe a partir de OCA.extension.

.DESCRIPTION
  1. Localiza o compilador do Inno Setup (ISCC.exe, versão 6.4 ou superior).
  2. Compila installer\Instalador_Plugin_OCA.iss. A validação da estrutura da
     extensão, a leitura da versão (botão Sobre) e o manifesto SHA-256 são
     feitos pelo próprio .iss durante a compilação.
  3. Mostra o caminho e o SHA-256 do instalador gerado.

.PARAMETER IsccPath
  Caminho do ISCC.exe, se não estiver em um local padrão.

.PARAMETER AppVersion
  Força a versão (aaaa.mm.dd). Por padrão é lida de VERSAO no botão Sobre.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File installer\build.ps1
#>
param(
    [string]$IsccPath = "",
    [string]$AppVersion = ""
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$iss  = Join-Path $here "Instalador_Plugin_OCA.iss"
$out  = Join-Path $root "dist\Instalador_Plugin_OCA.exe"

function Find-Iscc {
    if ($IsccPath) {
        if (Test-Path $IsccPath) { return (Resolve-Path $IsccPath).Path }
        throw "ISCC.exe não encontrado em: $IsccPath"
    }
    $cmd = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $candidates = @()
    foreach ($v in @("7", "6")) {
        $candidates += Join-Path ${env:ProgramFiles(x86)} "Inno Setup $v\ISCC.exe"
        $candidates += Join-Path $env:ProgramFiles "Inno Setup $v\ISCC.exe"
        $candidates += Join-Path $env:LOCALAPPDATA "Programs\Inno Setup $v\ISCC.exe"
    }
    # Local registrado pelo instalador do Inno Setup
    foreach ($key in @("HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                       "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                       "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")) {
        foreach ($v in @("7", "6")) {
            $p = Get-ItemProperty -Path "$key\Inno Setup $($v)_is1" -ErrorAction SilentlyContinue
            if ($p -and $p.InstallLocation) { $candidates += Join-Path $p.InstallLocation "ISCC.exe" }
        }
    }
    foreach ($c in $candidates) { if ($c -and (Test-Path $c)) { return $c } }
    throw ("Inno Setup não encontrado. Instale o Inno Setup 6.4 ou superior " +
           "(https://jrsoftware.org/isdl.php) ou informe -IsccPath ""C:\caminho\ISCC.exe"".")
}

$iscc = Find-Iscc
Write-Host "Compilador: $iscc"
$isccArgs = @()
if ($AppVersion) { $isccArgs += "/DAppVersion=$AppVersion" }
$isccArgs += $iss

& $iscc @isccArgs
if ($LASTEXITCODE -ne 0) {
    throw "A compilação falhou (código $LASTEXITCODE). Veja a mensagem do Inno Setup acima."
}
if (-not (Test-Path $out)) { throw "Instalador não encontrado após compilar: $out" }

$hash = (Get-FileHash -Algorithm SHA256 $out).Hash.ToLower()
$size = [math]::Round((Get-Item $out).Length / 1MB, 2)
Write-Host ""
Write-Host "Instalador gerado: $out ($size MB)"
Write-Host "SHA-256: $hash"
