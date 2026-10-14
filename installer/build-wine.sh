#!/usr/bin/env bash
# Compila o instalador no Linux usando o Inno Setup via Wine (alternativa ao
# build.ps1 do Windows; foi assim que o .exe deste repositório foi gerado).
#
# Requisitos: wine (64 bits) e o Inno Setup 6.4+ extraído em uma pasta, por
# exemplo do pacote NuGet "Tools.InnoSetup":
#   curl -L -o is.nupkg https://api.nuget.org/v3-flatcontainer/tools.innosetup/6.4.3/tools.innosetup.6.4.3.nupkg
#   unzip is.nupkg -d innosetup
#
# Uso: installer/build-wine.sh /caminho/para/innosetup/tools/ISCC.exe [/DAppVersion=aaaa.mm.dd]
#
# IMPORTANTE: o locale precisa ser UTF-8; caso contrário o Wine corrompe os
# nomes de pasta com acento (ex.: "02_Atualização Descrição por Keynotes").
set -euo pipefail
ISCC="${1:?informe o caminho do ISCC.exe}"; shift
export LANG=C.UTF-8 LC_ALL=C.UTF-8 WINEDEBUG=-all
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
RUN=(wine)
if [ -z "${DISPLAY:-}" ] && command -v xvfb-run >/dev/null; then RUN=(xvfb-run -a wine); fi
"${RUN[@]}" "$ISCC" "$@" "Z:$ROOT/installer/Instalador_Plugin_OCA.iss"
sha256sum "$ROOT/dist/Instalador_Plugin_OCA.exe"
