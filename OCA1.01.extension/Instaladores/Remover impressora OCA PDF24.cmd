@echo off
rem Remove a impressora OCA PDF24 (pede permissao de administrador).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Instalar impressora OCA PDF24.ps1" -Remover
