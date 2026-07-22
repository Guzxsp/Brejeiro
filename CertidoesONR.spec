# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['consulta_certidoes_onr.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl', 'openpyxl.cell._writer', 'openpyxl.cell', 'openpyxl.workbook'],
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
    a.binaries,
    a.datas,
    [],
    name='CertidoesONR',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='NONE',
)
