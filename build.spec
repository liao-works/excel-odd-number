# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

# 项目根目录
project_root = Path(__file__).parent

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
        'openpyxl', 
        'ttkbootstrap',
        'PIL',
        'PIL._tkinter_finder',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
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
        'numpy.distutils',
        'distutils',
        'pathlib',  # 排除冲突的pathlib包
        'streamlit',  # 排除streamlit相关冲突
    ],
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
    version_file=None,
)