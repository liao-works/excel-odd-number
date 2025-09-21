# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from pathlib import Path

# 项目根目录 - 在spec文件中使用当前工作目录
project_root = Path.cwd()

a = Analysis(
    ['main.py'],  # 主程序入口
    pathex=[str(project_root)],  # 搜索路径
    binaries=[],
    datas=[
        # 包含模板目录
        (str(project_root / 'templates'), 'templates'),
        # 包含资源目录  
        (str(project_root / 'assets'), 'assets'),
    ],
    hiddenimports=[
        'pandas',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime', 
        'pandas._libs.tslibs.nattype',
        'pandas._libs.skiplist',
        'openpyxl', 
        'openpyxl.workbook',
        'openpyxl.worksheet.worksheet',
        'ttkbootstrap',
        'PIL',
        'PIL._tkinter_finder',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'numpy',
        'numpy.core._methods',
        'numpy.lib.format',
        'numpy.core._dtype_ctypes',
        'pkg_resources.py2_warn',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'PyQt5',
        'PyQt6', 
        'PySide2',
        'PySide6',
        'matplotlib',
        'scipy',
        'streamlit',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel数据处理工具',  # 生成的exe文件名
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以在这里指定图标文件路径
)