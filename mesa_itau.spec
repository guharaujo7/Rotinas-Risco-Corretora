# -*- mode: python ; coding: utf-8 -*-
# Um único .exe (onefile): PyPDF2 + pywin32 + Playwright.
# Com build_mesa.ps1 (PLAYWRIGHT_BROWSERS_PATH=0 + playwright install chromium),
# o Chromium embutido é incluído no pacote e o app pode usar fallback sem Chrome/Edge.
# Uso: pyinstaller mesa_itau.spec

from pathlib import Path
from PyInstaller.utils.hooks import collect_all, collect_dynamic_libs

block_cipher = None

datas = []
binaries = []
hiddenimports = []

try:
    d2, b2, h2 = collect_all("PyPDF2")
    datas += d2
    binaries += b2
    hiddenimports += h2
except Exception:
    hiddenimports.append("PyPDF2")

try:
    dp, bp, hp = collect_all("playwright")
    datas += dp
    binaries += bp
    hiddenimports += hp
except Exception:
    hiddenimports.append("playwright")

try:
    import playwright
    pw_pkg = Path(playwright.__file__).resolve().parent
    local_browsers = pw_pkg / "driver" / "package" / ".local-browsers"
    if local_browsers.exists():
        datas.append((str(local_browsers), "playwright/driver/package/.local-browsers"))
except Exception:
    pass

try:
    binaries += collect_dynamic_libs("pywin32")
except Exception:
    pass

hiddenimports += [
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    "win32timezone",
    "PyPDF2",
]

a = Analysis(
    ["itau.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="mesa_itau",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
