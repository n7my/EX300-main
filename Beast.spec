# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('ControlCAN.dll', '.'), ('config.txt', '.'), ('取消全屏按钮.png', '.'), ('取消全屏按钮2.png', '.'), ('全屏按钮.png', '.'), ('全屏按钮2.png', '.'), ('logo.png', '.'), ('pass.png', '.'), ('fail.png', '.'), ('ControlCAN.lib', '.'), ('pushbutton.qrc', '.'), ('beast5.png', '.'), ('0000.png', '.'), ('485拨码.png', '.'), ('5555.png', '.'), ('AAAA.png', '.'), ('AE_00.png', '.'), ('AE_55.png', '.'), ('AE_aa.png', '.'), ('BAT_LOW灯.png', '.'), ('ERROR灯.png', '.'), ('HOST灯.png', '.'), ('PRG灯.png', '.'), ('RS-232C灯.png', '.'), ('RS-485灯.png', '.'), ('RUN灯.png', '.'), ('所有灯.png', '.'), ('beast5.ico', '.')],
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
    name='Beast',
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
    icon=['beast5.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Beast',
)
