# -*- coding: utf-8 -*-
"""
Single File Packaging Script - Resolves NumPy compatibility issues
English version for GitHub Actions compatibility
"""
import subprocess
import sys
import shutil
from pathlib import Path
import datetime

def create_usage_instructions():
    """Create detailed usage instructions file"""
    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    usage_content = f"""
# Excel Data Processor - User Guide

## Package Information
- Build Time: {current_time}
- File Location: ExcelDataProcessor.exe

## 🚀 Quick Start

### 1. First Use - Template Setup
Before using this tool, you need to set up Excel template files:

#### Step 1: Launch Program
Double-click "ExcelDataProcessor.exe" to start the program

#### Step 2: Open Settings
- Click "Settings" in the menu bar
- Select "Template Settings"

#### Step 3: Configure Templates
**UPS Summary Template Setup:**
- Click "Select UPS Summary Template" button
- Choose your UPS Excel template file (.xlsx format)
- Ensure template contains necessary column headers and formatting

**DPD Data Report Template Setup:**
- Click "Select DPD Data Report Template" button
- Choose your DPD Excel template file (.xlsx format)
- Ensure template contains necessary column headers and formatting

#### Step 4: Save Settings
- Click "Save Settings" button
- System will confirm settings saved successfully

### 2. Daily Usage Instructions

#### Processing Excel Files Steps:

**Step 1: Upload File**
- Click "Select File" button on main interface
- Choose Excel file to process (supports .xlsx and .xls formats)
- File path will be displayed on interface

**Step 2: Select Processing Type**
- Choose processing type on interface:
  - 🚚 UPS Processing: Use UPS summary template
  - 📦 DPD Processing: Use DPD data report template

**Step 3: Start Processing**
- Click "Start Processing" button
- Program will automatically process data and fill corresponding template
- Progress information will be displayed during processing

**Step 4: View Results**
- After processing completion, file will be automatically saved to desktop
- File name format: ProcessingType_Timestamp.xlsx
- System will display save location

## 📋 Supported File Formats

### Input Files
- Excel files: .xlsx, .xls
- Recommend using .xlsx format for best compatibility

### Output Files
- Unified output as .xlsx format
- Automatically saved to Windows desktop
- Filename includes timestamp to avoid overwriting

## ⚙️ System Requirements

### Minimum Requirements
- Windows 10 or higher
- 500MB available disk space
- 4GB memory (8GB recommended)

### Permission Requirements
- Desktop write permission (for saving output files)
- File read permission (for reading Excel files and templates)

## 🔧 Troubleshooting

### Common Issues and Solutions

**Issue 1: Program Cannot Start**
- Solution: Run program as administrator
- Check if blocked by antivirus software, add to trusted list

**Issue 2: Template File Not Found**
- Solution: Reset template path
- Ensure template file exists and is readable
- Check template file format is .xlsx

**Issue 3: Cannot Save Output File**
- Solution: Check desktop write permissions
- Ensure sufficient disk space
- Close other programs that might be using the file

**Issue 4: Processing Failed or Data Error**
- Solution: Check input Excel file format
- Confirm data column headers match template
- Check for empty rows or format anomalies

**Issue 5: Program Running Slowly**
- Solution: Please wait patiently when processing large files
- Close other memory-intensive programs
- Consider batch processing for large amounts of data

## 📞 Technical Support

### Getting Help
For technical support, please provide:
1. Error screenshots or error messages
2. Detailed operation steps
3. Example Excel files used (if shareable)
4. System version information

### Notes
- Please backup important Excel files regularly
- Recommend testing template configuration in test environment first
- Ensure system stability when processing large files
- First startup may take longer time

### Version Information
- Tool Version: 1.0
- Build Time: {current_time}
- Supported Formats: Excel .xlsx/.xls

---
© Excel Data Processor - Professional Excel Data Processing Solution
"""

    # Save instructions to dist directory
    dist_dir = Path('dist')
    if dist_dir.exists():
        usage_file = dist_dir / 'UserGuide.txt'
        try:
            with open(usage_file, 'w', encoding='utf-8') as f:
                f.write(usage_content)
            print(f"[OK] Generated user guide: {usage_file}")
        except Exception as e:
            print(f"[WARNING] Failed to generate user guide: {e}")
    
    # Also output to console
    print("\n" + "="*60)
    print("📖 Key Usage Points:")
    print("="*60)
    print("1. First use requires template setup:")
    print("   - Launch program → Settings → Template Settings")
    print("   - Select UPS and DPD template files respectively")
    print("   - Click Save Settings")
    print()
    print("2. Daily usage workflow:")
    print("   - Select Excel file → Choose processing type → Start processing")
    print("   - Results automatically saved to desktop")
    print()
    print("3. Important notes:")
    print("   - Supports .xlsx and .xls formats")
    print("   - Output files saved to desktop")
    print("   - First startup may be slow")
    print("="*60)

def onefile_build():
    """Build using single file mode"""
    print("Starting single file packaging...")

    # Delete dist directory
    if Path('dist').exists():
        shutil.rmtree('dist')
    # Delete build directory
    if Path('build').exists():
        shutil.rmtree('build')

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',  # Single file mode
        '--windowed',  # No console window
        '--name=Excel数据处理工具',
        '--add-data=templates;templates',
        '--add-data=assets;assets',
        
        # Critical hidden imports to solve NumPy issues
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
        
        # Exclude unnecessary packages
        '--exclude-module=PyQt5',
        '--exclude-module=PyQt6',
        '--exclude-module=matplotlib',
        '--exclude-module=scipy',
        '--exclude-module=streamlit',
        
        # Other options
        '--clean',
        '--noconfirm',
        
        'main.py'
    ]
    
    print(f"Executing command: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("[OK] Single file packaging successful!")
        print("EXE location: dist/Excel数据处理工具.exe")
        
        # Check file size
        exe_file = Path('dist/Excel数据处理工具.exe')
        if exe_file.exists():
            size_mb = exe_file.stat().st_size / (1024 * 1024)
            print(f"File size: {size_mb:.1f} MB")
            
            # Generate usage instructions
            create_usage_instructions()

        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Packaging failed: {e}")
        print("STDOUT:", e.stdout)
        print("STDERR:", e.stderr)
        return False

if __name__ == "__main__":
    onefile_build()