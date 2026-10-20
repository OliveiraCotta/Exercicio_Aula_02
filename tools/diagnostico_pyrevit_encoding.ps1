# Diagnostico de codificacao dos scripts pyRevit (OCA).
#
# Para cada script.py da extensao instalada, verifica:
#   - BOM UTF-8 e decodificacao UTF-8 estrita (bytes invalidos);
#   - docstring de modulo com caractere nao-ASCII (causa o UnicodeDecodeError
#     do pyRevit 5.x no IronPython 2.7: ScriptFileParser.get_docstring faz
#     doc_str.decode('utf-8'));
#   - se informado -Repositorio, compara o SHA256 com o arquivo do repositorio.
#
# Uso (PowerShell):
#   powershell -ExecutionPolicy Bypass -File diagnostico_pyrevit_encoding.ps1
#   powershell -ExecutionPolicy Bypass -File diagnostico_pyrevit_encoding.ps1 `
#       -Instalado "$env:APPDATA\pyRevit\Extensions\OCA.extension" `
#       -Repositorio "C:\caminho\Exercicio_Aula_02\OCA01.extension"
#
# Este arquivo e so ASCII de proposito: o PowerShell 5.1 le .ps1 sem BOM como ANSI.

param(
    [string]$Instalado = (Join-Path $env:APPDATA "pyRevit\Extensions\OCA.extension"),
    [string]$Repositorio = ""
)

$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)

function Get-ModuleDocstring([string]$text) {
    # Pula linhas em branco e comentarios; se a primeira instrucao for uma
    # string entre aspas triplas (com ou sem prefixo u/r), devolve o conteudo.
    $m = [regex]::Match($text, '^(?:[ \t]*(?:#[^\n]*)?\r?\n)*[ \t]*[uUrR]?("""|'''''')(.*?)\1', 'Singleline')
    if ($m.Success) { return $m.Groups[2].Value }
    return $null
}

if (-not (Test-Path $Instalado)) {
    Write-Host "Pasta instalada nao encontrada: $Instalado" -ForegroundColor Red
    exit 1
}

$problemas = 0
Get-ChildItem -Path $Instalado -Recurse -Filter "script.py" | ForEach-Object {
    $arq = $_.FullName
    $rel = $arq.Substring($Instalado.Length).TrimStart('\')
    $bytes = [System.IO.File]::ReadAllBytes($arq)
    $bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    $status = @()
    $text = $null
    try { $text = $strictUtf8.GetString($bytes) } catch { $status += "NAO e UTF-8 valido" }
    if ($bom) { $status += "tem BOM UTF-8" }
    if ($text) {
        $doc = Get-ModuleDocstring $text
        if ($doc -and ($doc.ToCharArray() | Where-Object { [int]$_ -gt 127 })) {
            $status += "docstring de modulo com acento (causa o UnicodeDecodeError)"
        }
    }
    if ($Repositorio) {
        $par = Join-Path $Repositorio $rel
        if (Test-Path $par) {
            $h1 = (Get-FileHash $arq -Algorithm SHA256).Hash
            $h2 = (Get-FileHash $par -Algorithm SHA256).Hash
            if ($h1 -ne $h2) { $status += "DIFERENTE do repositorio" }
        } else {
            $status += "nao existe no repositorio"
        }
    }
    if ($status.Count -eq 0) {
        Write-Host ("OK     " + $rel) -ForegroundColor Green
    } else {
        $problemas++
        Write-Host ("ATENCAO " + $rel + " -> " + ($status -join "; ")) -ForegroundColor Yellow
    }
}
Write-Host ""
Write-Host ("Arquivos com problema: " + $problemas)
