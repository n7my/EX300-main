# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('ControlCAN.dll', '.'), ('config.txt', '.'), ('取消全屏按钮.png', '.'), ('取消全屏按钮2.png', '.'), ('全屏按钮.png', '.'), ('全屏按钮2.png', '.'), ('logo.png', '.'), ('pass.png', '.'), ('fail.png', '.'), ('ControlCAN.lib', '.'), ('pushbutton.qrc', '.'), ('beast5.png', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
