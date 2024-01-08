# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:/Users/LDCC_HWANG/OneDrive/문서/futurechem_OrderAuto/autoDocx/autoOrder.py'],
    pathex=['C:/Users/LDCC_HWANG/OneDrive/문서/futurechem_OrderAuto/autoDocx'],
    binaries=[],
    datas=[],
    hiddenimports=['autoDocx'],
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
    name='구매요구서 문서 자동화 ver 0.1',
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
    icon_resources=[(2, 'C:/Users/LDCC_HWANG/OneDrive/문서/futurechem_OrderAuto/autoDocx/futureMain.ico')],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='autoOrder',
    icon='C:/Users/LDCC_HWANG/OneDrive/문서/futurechem_OrderAuto/autoDocx/future.ico'
)
