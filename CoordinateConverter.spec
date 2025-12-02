# -*- mode: python ; coding: utf-8 -*-

# 排除不必要的大型库以减小exe文件大小
excluded_libs = [
    # 数据科学和机器学习库
    'matplotlib', 'mpl_toolkits', 'scipy', 'numpy.random._pickle',
    'sklearn', 'skimage', 'PIL', 'pillow',
    'tensorflow', 'torch', 'keras',
    
    # GIS相关但不需要的库
    'geopandas', 'rasterio', 'fiona', 'shapely',
    'cartopy', 'pygmt', 'osgeo', 'gdal', 'ogr', 'osr',
    
    # 其他大型库
    'IPython', 'jupyter', 'notebook', 'ipykernel',
    'distributed', 'dask', 'bokeh', 'holoviews', 'panel',
    'plotly', 'dash', 'flask', 'django',
    'requests', 'urllib3', 'httpx',
    'sqlalchemy', 'pymysql', 'psycopg2',
    'pytest', 'unittest',
    'xmlrpc', 'html', 'cgi'
]

a = Analysis(
    ['CoordinateConverter.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded_libs,
    noarchive=False,
    optimize=1,  # 启用基本优化
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CoordinateConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 设置为False以避免显示控制台窗口
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
    name='CoordinateConverter',
)