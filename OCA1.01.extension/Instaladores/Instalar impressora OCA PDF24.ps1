# Instala a impressora "OCA PDF24" usada pelo botão Exportar Pranchas (opção Printer > PDF24).
#
# O Revit não deixa imprimir em arquivo com impressoras de PDF. Por isso o botão usa uma
# impressora PDF24 configurada para SALVAR SOZINHA numa pasta, sem abrir janela.
# Este instalador cria:
#   - a porta  "\\.\pipe\PDFPrint - oca.exportarpranchas"
#   - a impressora "OCA PDF24" (driver PDF24)
#   - a configuração do PDF24 para essa impressora salvar em %LOCALAPPDATA%\OCA_ExportarPranchas\PDF24
# Precisa do PDF24 Creator instalado e de permissão de administrador (uma vez por computador).
# Para remover tudo: "Remover impressora OCA PDF24.cmd" (ou este arquivo com -Remover).
param([switch]$Remover)

$service = 'oca.exportarpranchas'
$printer = 'OCA PDF24'
$port    = "\\.\pipe\PDFPrint - $service"
$svcKey  = "HKLM:\SOFTWARE\PDF24\Services\$service"

function Fechar($codigo) {
    Write-Host ''
    Read-Host 'Pressione Enter para fechar' | Out-Null
    exit $codigo
}

# pede administrador: abre de novo este arquivo com "Executar como administrador"
$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) {
    $argumentos = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    if ($Remover) { $argumentos += ' -Remover' }
    try {
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argumentos
    } catch {
        Write-Host 'É preciso permissão de administrador para instalar a impressora.' -ForegroundColor Red
        Fechar 1
    }
    exit 0
}

try {
    if ($Remover) {
        Write-Host 'Removendo a impressora OCA PDF24...'
        if (Get-Printer -Name $printer -ErrorAction SilentlyContinue) { Remove-Printer -Name $printer }
        if (Get-PrinterPort -Name $port -ErrorAction SilentlyContinue) { Remove-PrinterPort -Name $port }
        if (Test-Path $svcKey) { Remove-Item $svcKey -Recurse -Force }
        Restart-Service -Name PDF24 -ErrorAction SilentlyContinue
        Write-Host 'Pronto: a impressora OCA PDF24 foi removida.' -ForegroundColor Green
        Fechar 0
    }

    if (-not (Get-PrinterDriver -Name 'PDF24' -ErrorAction SilentlyContinue)) {
        throw 'O PDF24 Creator não está instalado neste computador. Instale o PDF24 Creator (pdf24.org) e rode este instalador de novo.'
    }

    Write-Host 'Configurando o PDF24 para salvar sozinho...'
    New-Item -Path $svcKey -Force | Out-Null
    $texto = @{
        Port               = $port
        Handler            = 'autoSave'
        AutoSaveDir        = '%localappdata%\OCA_ExportarPranchas\PDF24'
        AutoSaveFilename   = '$fileName'
        AutoSaveFileCmd    = ''
        FilenameErasements = ''
        ShellCmd           = ''
    }
    foreach ($nome in $texto.Keys) {
        New-ItemProperty -Path $svcKey -Name $nome -Value $texto[$nome] -PropertyType String -Force | Out-Null
    }
    $numeros = @{
        AutoSaveOpenDir        = 0
        AutoSaveOverwriteFile  = 1
        AutoSaveShowProgress   = 0
        AutoSaveUseFileChooser = 0
        AutoSaveUseFileCmd     = 0
        LoadInCreatorIfOpen    = 0
    }
    foreach ($nome in $numeros.Keys) {
        New-ItemProperty -Path $svcKey -Name $nome -Value $numeros[$nome] -PropertyType DWord -Force | Out-Null
    }

    Write-Host 'Criando a porta e a impressora...'
    if (-not (Get-PrinterPort -Name $port -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $port }
    if (-not (Get-Printer -Name $printer -ErrorAction SilentlyContinue)) {
        Add-Printer -Name $printer -DriverName 'PDF24' -PortName $port
    }

    Write-Host 'Reiniciando o serviço do PDF24 para reconhecer a impressora nova...'
    Restart-Service -Name PDF24

    Write-Host ''
    Write-Host 'Pronto: a impressora "OCA PDF24" foi instalada.' -ForegroundColor Green
    Write-Host 'No Revit, use Exportar Pranchas > Opções de PDF > Printer > PDF24.'
    Fechar 0
} catch {
    Write-Host ''
    Write-Host "ERRO: $($_.Exception.Message)" -ForegroundColor Red
    Fechar 1
}
