# -*- coding: utf-8 -*-
"""
PyInstaller打包脚本
用于将Python应用程序打包为Windows exe文件
"""
import os
import sys
import shutil
import subprocess
from pathlib import Path

def clean_build_dirs():
    """清理构建目录"""
    print("清理旧的构建文件...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        if Path(dir_name).exists():
            shutil.rmtree(dir_name)
            print(f"已删除目录: {dir_name}")
    
    # 清理.pyc文件
    for pyc_file in Path('.').rglob('*.pyc'):
        pyc_file.unlink()
        
    print("清理完成！\n")

def build_exe():
    """使用PyInstaller构建exe文件"""
    print("开始使用PyInstaller构建exe文件...")
    
    try:
        # 使用spec文件构建
        cmd = [
            sys.executable, '-m', 'PyInstaller',
            '--clean',  # 清理缓存
            'build.spec'
        ]
        
        print(f"执行命令: {' '.join(cmd)}")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore'
        )
        
        if result.returncode == 0:
            print("[OK] PyInstaller构建成功！")
            return True
        else:
            print("[ERROR] PyInstaller构建失败：")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            return False
            
    except Exception as e:
        print(f"[ERROR] 构建过程出错: {str(e)}")
        return False

def check_output():
    """检查输出文件"""
    dist_dir = Path('dist')
    
    if not dist_dir.exists():
        print("[ERROR] dist目录不存在")
        return False
        
    exe_files = list(dist_dir.glob('*.exe'))
    
    if not exe_files:
        print("[ERROR] 未找到exe文件")
        return False
        
    for exe_file in exe_files:
        size_mb = exe_file.stat().st_size / (1024 * 1024)
        print(f"[OK] 生成exe文件: {exe_file.name}")
        print(f"   文件大小: {size_mb:.1f} MB")
        print(f"   文件路径: {exe_file.absolute()}")
        
    return True

def create_installer_info():
    """创建安装说明"""
    info_content = """
# Excel数据处理工具 - 使用说明

## 安装
1. 将生成的exe文件复制到任意目录
2. 确保templates目录包含所需的模板文件
3. 双击exe文件即可运行

## 使用方法
1. 启动程序
2. 点击"选择文件"上传Excel文件
3. 选择处理类型（UPS或DPD）
4. 点击"开始处理"
5. 处理完成后，文件将自动保存到桌面

## 设置模板
1. 点击菜单栏的"设置" -> "模板设置"
2. 分别选择UPS总结单模板和DPD数据预报模板
3. 点击"保存设置"

## 注意事项
- 支持.xlsx和.xls格式的Excel文件
- 模板文件必须是.xlsx格式
- 输出文件将自动命名并保存到桌面
- 程序运行需要Windows 10或更高版本

## 技术支持
如遇问题，请检查：
1. Excel文件格式是否正确
2. 模板文件是否存在且可读
3. 桌面是否有写入权限
"""
    
    info_file = Path('dist') / '使用说明.txt'
    
    try:
        with open(info_file, 'w', encoding='utf-8') as f:
            f.write(info_content)
        print(f"[OK] 已创建使用说明: {info_file}")
    except Exception as e:
        print(f"[ERROR] 创建使用说明失败: {str(e)}")

def main():
    """主函数"""
    print("=" * 60)
    print("Excel数据处理工具 - PyInstaller打包脚本")
    print("=" * 60)
    
    # 检查依赖
    try:
        import PyInstaller
        print(f"[OK] PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("[ERROR] PyInstaller未安装，请运行: pip install pyinstaller")
        return
    
    # 执行构建步骤
    success = True
    
    # 步骤1：清理
    clean_build_dirs()
    
    # 步骤2：构建
    success &= build_exe()
    
    if success:
        # 步骤3：检查输出
        success &= check_output()
        
        if success:
            # 步骤4：创建说明
            create_installer_info()
            
            print("\n" + "=" * 60)
            print("[SUCCESS] 打包完成！")
            print("exe文件位置: dist/Excel数据处理工具.exe")
            print("请将整个dist目录分发给用户")
            print("=" * 60)
        else:
            print("\n[ERROR] 输出检查失败")
    else:
        print("\n[ERROR] 构建失败，请检查错误信息")

if __name__ == "__main__":
    main()