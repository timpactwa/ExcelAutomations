# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['NormalizeAdditions.py'],
    pathex=[],
    binaries=[],
    datas=[],
    # pandas only imports openpyxl when read_excel is actually called, so
    # PyInstaller can't see it — without this the frozen app raises
    # "Missing optional dependency 'openpyxl'" on every Excel open.
    # tzdata is Windows' only source of zone rules for zoneinfo, needed to
    # stamp batch filenames in Eastern time.
    hiddenimports=['openpyxl', 'tzdata'],
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
    name='Clean_Additions',
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
)
