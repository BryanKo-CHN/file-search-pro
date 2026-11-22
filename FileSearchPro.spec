# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = []
for pkg in ('pypdf', 'pymupdf', 'mammoth', 'pptx', 'openpyxl', 'lxml', 'PIL',
            'cobble', 'et_xmlfile', 'win32com', 'comtypes'):
    try:
        tmp_ret = collect_all(pkg)
        datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
    except Exception:
        pass
# COM 兼容 Microsoft Office / WPS，pywin32 与 comtypes 两条路都打进去
hiddenimports += ['pythoncom', 'pywintypes', 'win32com', 'win32com.client',
                  'comtypes', 'comtypes.client']
# 确保多进程 spawn 相关模块被打进去
hiddenimports += ['multiprocessing', 'multiprocessing.spawn', 'multiprocessing.resource_tracker']

a = Analysis(
    ['../file_search_pro.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='FileSearchPro',
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
