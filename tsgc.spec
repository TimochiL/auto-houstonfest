# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:\\My Stuff\\School\\2024-2025 HS\\Houstonfest\\auto-houstonfest\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('D:\\My Stuff\\School\\2024-2025 HS\\Houstonfest\\auto-houstonfest\\boomer_utils.py', '.'), ('D:\\My Stuff\\School\\2024-2025 HS\\Houstonfest\\auto-houstonfest\\generate_reports.py', '.'), ('D:\\My Stuff\\School\\2024-2025 HS\\Houstonfest\\auto-houstonfest\\models.py', '.')],
    hiddenimports=['pony.orm.dbproviders.sqlite'],
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
    name='tsgc',
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
    contents_directory='bin',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='tsgc',
)
