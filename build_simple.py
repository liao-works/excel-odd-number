# -*- coding: utf-8 -*-
"""
单文件打包脚本 - 解决NumPy兼容性问题
"""
import subprocess
import sys
import shutil
from pathlib import Path
import datetime

def create_usage_instructions():
    """创建详细的使用说明文件"""
    current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    usage_content = f"""
# Excel数据处理工具 - 使用说明

## 打包信息
- 打包时间: {current_time}
- 文件位置: Excel数据处理工具.exe

## 🚀 快速开始

### 1. 首次使用 - 设置模板
在使用本工具之前，您需要先设置Excel模板文件：

#### 步骤1：启动程序
双击"Excel数据处理工具.exe"启动程序

#### 步骤2：打开设置
- 点击菜单栏的"设置"
- 选择"模板设置"

#### 步骤3：配置模板
**UPS总结单模板设置:**
- 点击"选择UPS总结单模板"按钮
- 选择您的UPS Excel模板文件（.xlsx格式）
- 确保模板包含必要的列标题和格式

**DPD数据预报模板设置:**
- 点击"选择DPD数据预报模板"按钮  
- 选择您的DPD Excel模板文件（.xlsx格式）
- 确保模板包含必要的列标题和格式

#### 步骤4：保存设置
- 点击"保存设置"按钮
- 系统会提示设置保存成功

### 2. 日常使用说明

#### 处理Excel文件的步骤：

**第1步：上传文件**
- 点击主界面的"选择文件"按钮
- 选择需要处理的Excel文件（支持.xlsx和.xls格式）
- 文件路径会显示在界面上

**第2步：选择处理类型**
- 在界面上选择处理类型：
  - 🚚 UPS处理：使用UPS总结单模板
  - 📦 DPD处理：使用DPD数据预报模板

**第3步：开始处理**
- 点击"开始处理"按钮
- 程序会自动处理数据并填充到对应模板
- 处理过程中会显示进度信息

**第4步：查看结果**
- 处理完成后，文件会自动保存到桌面
- 文件名格式：处理类型_时间戳.xlsx
- 系统会显示保存位置

## 📋 支持的文件格式

### 输入文件
- Excel文件：.xlsx, .xls
- 建议使用.xlsx格式以获得最佳兼容性

### 输出文件
- 统一输出为.xlsx格式
- 自动保存到Windows桌面
- 文件名包含时间戳，避免覆盖

## ⚙️ 系统要求

### 最低要求
- Windows 10 或更高版本
- 500MB 可用磁盘空间
- 4GB 内存（推荐8GB）

### 权限要求
- 桌面写入权限（保存输出文件）
- 文件读取权限（读取Excel文件和模板）

## 🔧 故障排除

### 常见问题及解决方案

**问题1：程序无法启动**
- 解决方案：以管理员身份运行程序
- 检查是否被杀毒软件拦截，添加信任

**问题2：找不到模板文件**
- 解决方案：重新设置模板路径
- 确保模板文件存在且可读
- 检查模板文件格式是否为.xlsx

**问题3：无法保存输出文件**
- 解决方案：检查桌面写入权限
- 确保磁盘空间充足
- 关闭可能占用文件的其他程序

**问题4：处理失败或数据错误**
- 解决方案：检查输入Excel文件格式
- 确认数据列标题与模板匹配
- 检查是否有空行或格式异常

**问题5：程序运行缓慢**
- 解决方案：处理大文件时请耐心等待
- 关闭其他占用内存的程序
- 考虑分批处理大量数据

## 📞 技术支持

### 获取帮助
如需技术支持，请提供以下信息：
1. 错误截图或错误信息
2. 详细的操作步骤
3. 使用的Excel文件示例（如可分享）
4. 系统版本信息

### 注意事项
- 请定期备份重要的Excel文件
- 建议在测试环境中先验证模板配置
- 处理大文件时请确保系统稳定
- 首次运行可能需要较长启动时间

### 版本信息
- 工具版本：1.0
- 打包时间：{current_time}
- 支持格式：Excel .xlsx/.xls

---
© Excel数据处理工具 - 专业的Excel数据处理解决方案
"""

    # 保存说明文件到dist目录
    dist_dir = Path('dist')
    if dist_dir.exists():
        usage_file = dist_dir / '使用说明.txt'
        try:
            with open(usage_file, 'w', encoding='utf-8') as f:
                f.write(usage_content)
            print(f"[OK] 已生成使用说明: {usage_file}")
        except Exception as e:
            print(f"[WARNING] 生成使用说明失败: {e}")
    
    # 同时输出到控制台
    print("\n" + "="*60)
    print("📖 使用说明要点：")
    print("="*60)
    print("1. 首次使用需要设置模板：")
    print("   - 启动程序 → 设置 → 模板设置")
    print("   - 分别选择UPS和DPD模板文件")
    print("   - 点击保存设置")
    print()
    print("2. 日常使用流程：")
    print("   - 选择Excel文件 → 选择处理类型 → 开始处理")
    print("   - 结果自动保存到桌面")
    print() 
    print("3. 注意事项：")
    print("   - 支持.xlsx和.xls格式")
    print("   - 输出文件保存到桌面")
    print("   - 首次启动可能较慢")
    print("="*60)

def onefile_build():
    """使用单文件模式打包"""
    print("使用单文件模式打包...")

    # 删除dist目录
    if Path('dist').exists():
        shutil.rmtree('dist')
    # 删除build目录
    if Path('build').exists():
        shutil.rmtree('build')

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
            
            # 生成使用说明文件
            create_usage_instructions()

        return True
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] 打包失败: {e}")
        print("STDOUT:", e.stdout)
        print("STDERR:", e.stderr)
        return False

if __name__ == "__main__":
    onefile_build()