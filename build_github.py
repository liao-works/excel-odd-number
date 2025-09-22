# -*- coding: utf-8 -*-
"""
GitHub Actions Compatible Build Script
"""
import subprocess
import sys
import shutil
from pathlib import Path
import datetime
import os

def onefile_build():
    """Build using single file mode - GitHub Actions compatible"""
    print("Starting single file packaging...")

    # Delete dist directory
    if Path('dist').exists():
        shutil.rmtree('dist')
    # Delete build directory
    if Path('build').exists():
        shutil.rmtree('build')

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--windowed',
        '--name=Excel数据处理工具',
        '--add-data=templates;templates',
        '--add-data=assets;assets',
        
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
        
        '--exclude-module=PyQt5',
        '--exclude-module=PyQt6',
        '--exclude-module=matplotlib',
        '--exclude-module=scipy',
        '--exclude-module=streamlit',
        
        '--clean',
        '--noconfirm',
        
        'main.py'
    ]
    
    print("Executing PyInstaller command...")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='utf-8', errors='ignore')
        print("[OK] Single file packaging successful!")
        print("EXE location: dist/Excel数据处理工具.exe")
        
        # Check file size
        exe_file = Path('dist/Excel数据处理工具.exe')
        if exe_file.exists():
            size_mb = exe_file.stat().st_size / (1024 * 1024)
            print(f"File size: {size_mb:.1f} MB")
            
            # Create English user guide for GitHub release
            create_english_guide()

        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Packaging failed: {e}")
        print("STDOUT:", e.stdout if e.stdout else "No stdout")
        print("STDERR:", e.stderr if e.stderr else "No stderr")
        return False

def create_english_guide():
    """Create English user guide for international users"""
    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    guide_content = f"""# Excel Data Processor - User Guide

## Build Information
- Build Time: {current_time}
- Application: Excel Data Processor

## Quick Start Guide

### 1. First Time Setup
1. Launch Excel Data Processor.exe
2. Go to Settings > Template Settings
3. Configure UPS and DPD templates
4. Save settings

### 2. Daily Usage
1. Select Excel file to process
2. Choose processing type (UPS or DPD)
3. Click Start Processing
4. Results saved to desktop automatically

### 3. System Requirements
- Windows 10 or higher
- 500MB disk space
- Supports .xlsx and .xls files

### 4. Troubleshooting
- Run as administrator if startup fails
- Check template files are .xlsx format
- Ensure desktop write permissions

---
Build Time: {current_time}
"""
    
    try:
        guide_file = Path('dist/UserGuide.txt')
        with open(guide_file, 'w', encoding='utf-8') as f:
            f.write(guide_content)
        print(f"[OK] Created user guide: {guide_file}")
    except Exception as e:
        print(f"[WARNING] Failed to create user guide: {e}")

if __name__ == "__main__":
    # Set UTF-8 encoding
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    
    try:
        onefile_build()
    except UnicodeError as e:
        print(f"Encoding error: {e}")
        print("Falling back to ASCII output mode")
        # Fallback logic here if needed