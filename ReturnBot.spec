# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['ReturnBot.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.png', '.'), ('mail-in template.xlsx', '.'), ('mail-in swollen template.xlsx', '.'), ('kbb template.xlsx', '.'), ('battery kbb template.xlsx', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='退料機器人',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['MyIcon.icns'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='退料機器人',
)
app = BUNDLE(
    coll,
    name='退料機器人.app',
    icon='MyIcon.icns',
    bundle_identifier='com.returnbot.app',  # 這裡可以隨便取一個 ID
    info_plist={
        'NSAppleEventsUsageDescription': '請點選「好」以允許機器人控制 Excel 進行退料單填寫。',
        'NSPrincipalClass': 'NSApplication',
        'NSHighResolutionCapable': 'True',
        'LSBackgroundOnly': 'False'
    },
)