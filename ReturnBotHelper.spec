# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_root = Path(SPECPATH)
templates = [
    "mail-in template.xlsx",
    "mail-in swollen template.xlsx",
    "kbb template.xlsx",
    "battery kbb template.xlsx",
]

a = Analysis(
    [str(project_root / "returnbot_cli.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=[(str(project_root / name), ".") for name in templates],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ReturnBotHelper",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch="arm64",
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="ReturnBotHelper",
)
