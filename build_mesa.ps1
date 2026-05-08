# Gera dist\mesa_itau.exe (onefile) com Playwright + Chromium embutido.
# Objetivo: rodar em PC do usuário sem precisar instalar pip/dependências.
# Se houver Chrome/Edge, o app usa channel; se não houver, usa fallback Chromium embutido.

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

# Embute browsers do Playwright no ambiente/pacote.
$env:PLAYWRIGHT_BROWSERS_PATH = "0"

python -m pip install -r requirements-build.txt -q
python -m playwright install chromium
python -m PyInstaller mesa_itau.spec --noconfirm --clean

Write-Host ""
Write-Host "Pronto. Execute: dist\mesa_itau.exe"
Write-Host "Pode enviar apenas esse arquivo."
