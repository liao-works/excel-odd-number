# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates'), ('assets', 'assets')],
    hiddenimports=['pandas', 'pandas._libs.tslibs.timedeltas', 'pandas._libs.tslibs.np_datetime', 'pandas._libs.tslibs.nattype', 'pandas._libs.skiplist', 'openpyxl', 'openpyxl.workbook', 'openpyxl.worksheet.worksheet', 'ttkbootstrap', 'PIL', 'tkinter', 'numpy', 'numpy.core._methods', 'numpy.lib.format', 'numpy.core._dtype_ctypes'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PyQt6', 'matplotlib', 'scipy', 'streamlit'],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Excel数据处理工具',
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
