# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('icon', 'icon'), ('ControlCAN.dll', '.'), ('config.txt', '.'), ('取消全屏按钮.png', '.'), ('取消全屏按钮2.png', '.'), ('全屏按钮.png', '.'), ('全屏按钮2.png', '.'), ('logo.png', '.'), ('pass.png', '.'), ('fail.png', '.'), ('ControlCAN.lib', '.'), ('pushbutton.qrc', '.'), ('beast5.png', '.'), ('0000.png', '.'), ('485拨码.png', '.'), ('5555.png', '.'), ('AAAA.png', '.'), ('AE_00.png', '.'), ('AE_55.png', '.'), ('AE_aa.png', '.'), ('BAT_LOW灯.png', '.'), ('ERROR灯.png', '.'), ('HOST灯.png', '.'), ('PRG灯.png', '.'), ('RS-232C灯.png', '.'), ('RS-485灯.png', '.'), ('RUN灯.png', '.'), ('所有灯.png', '.'), ('AE_ERROR.png', '.'), ('AE_RUN.png', '.'), ('AQ_ERROR.png', '.'), ('AQ_RUN.png', '.'), ('DO_0.png', '.'), ('DO_1.png', '.'), ('DO_10.png', '.'), ('DO_11.png', '.'), ('DO_12.png', '.'), ('DO_13.png', '.'), ('DO_14.png', '.'), ('DO_15.png', '.'), ('DO_2.png', '.'), ('DO_3.png', '.'), ('DO_4.png', '.'), ('DO_5.png', '.'), ('DO_6.png', '.'), ('DO_7.png', '.'), ('DO_8.png', '.'), ('DO_9.png', '.'), ('DO_ERROR.png', '.'), ('DO_FF.png', '.'), ('DO_RUN.png', '.'), ('ET1600_0.png', '.'), ('ET1600_1.png', '.'), ('ET1600_10.png', '.'), ('ET1600_11.png', '.'), ('ET1600_12.png', '.'), ('ET1600_13.png', '.'), ('ET1600_14.png', '.'), ('ET1600_15.png', '.'), ('ET1600_2.png', '.'), ('ET1600_3.png', '.'), ('ET1600_4.png', '.'), ('ET1600_5.png', '.'), ('ET1600_6.png', '.'), ('ET1600_7.png', '.'), ('ET1600_8.png', '.'), ('ET1600_9.png', '.'), ('ET1600_ERROR.png', '.'), ('ET1600_FF.png', '.'), ('ET1600_RUN.png', '.')],
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
    icon=['./icon/beast5.ico'],
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
