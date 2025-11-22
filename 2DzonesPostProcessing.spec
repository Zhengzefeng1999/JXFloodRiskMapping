# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['2DzonesPostProcessing.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
    'matplotlib', 'scipy', 'tensorflow', 'torch', 
    'sklearn', 'seaborn', 'plotly', 'dash', 'flask', 
    'django', 'sqlalchemy', 'psycopg2', 'pymongo'
    ],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='2DzonesPostProcessing',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='2DzonesPostProcessing',
)
