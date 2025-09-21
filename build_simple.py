# -*- coding: utf-8 -*-
"""
单文件打包脚本 - 解决NumPy兼容性问题
"""
import subprocess
import sys
from pathlib import Path

def onefile_build():
    """使用单文件模式打包"""
    print("使用单文件模式打包...")
    
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',  # 单文件模式
        '--windowed',  # 无控制台
        '--name=Excel数据处理工具',
        '--add-data=templates;templates',
        '--add-data=assets;assets',
        
        # 关键的隐藏导入，解决NumPy问题
        '--hidden-import=pandas',
        '--hidden-import=pandas._libs.tslibs.timedeltas',
        '--hidden-import=pandas._libs.tslibs.np_datetime',
        '--hidden-import=pandas._libs.tslibs.nattype',
        '--hidden-import=pandas._libs.skiplist',
        '--hidden-import=openpyxl',
        '--hidden-import=openpyxl.workbook',
        '--hidden-import=openpyxl.worksheet.worksheet', 
        '--hidden-import=ttkbootstrap',
        '--hidden-import=PIL',
        '--hidden-import=tkinter',
        '--hidden-import=numpy',
        '--hidden-import=numpy.core._methods',
        '--hidden-import=numpy.lib.format',
        '--hidden-import=numpy.core._dtype_ctypes',
        
        # 排除不需要的包
        '--exclude-module=PyQt5',
        '--exclude-module=PyQt6',
        '--exclude-module=matplotlib',
        '--exclude-module=scipy',
        '--exclude-module=streamlit',
        
        # 其他选项
        '--clean',
        '--noconfirm',
        
        'main.py'
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("[OK] 单文件打包成功！")
        print("exe文件位置: dist/Excel数据处理工具.exe")
        
        # 检查文件大小
        exe_file = Path('dist/Excel数据处理工具.exe')
        if exe_file.exists():
            size_mb = exe_file.stat().st_size / (1024 * 1024)
            print(f"文件大小: {size_mb:.1f} MB")
        
        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] 打包失败: {e}")
        print("STDOUT:", e.stdout)
        print("STDERR:", e.stderr)
        return False

if __name__ == "__main__":
    onefile_build()