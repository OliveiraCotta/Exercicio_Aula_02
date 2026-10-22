<#
.SYNOPSIS
  Gera dist\Instalador_Plugin_<extensão>.exe a partir da pasta *.extension do
  repositório (ex.: OCA1.01.extension -> dist\Instalador_Plugin_OCA1.01.exe).

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

$inicio = Get-Date
& $iscc @isccArgs
if ($LASTEXITCODE -ne 0) {
    throw "A compilação falhou (código $LASTEXITCODE). Veja a mensagem do Inno Setup acima."
}
$gerado = Get-ChildItem -Path (Join-Path $root "dist") -Filter "Instalador_Plugin_*.exe" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -ge $inicio.AddSeconds(-2) } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $gerado) { throw "Instalador não encontrado em dist após compilar." }
$out = $gerado.FullName

$hash = (Get-FileHash -Algorithm SHA256 $out).Hash.ToLower()
$size = [math]::Round((Get-Item $out).Length / 1MB, 2)
Write-Host ""
Write-Host "Instalador gerado: $out ($size MB)"
Write-Host "SHA-256: $hash"
