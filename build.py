# -*- coding: utf-8 -*-
"""
Excel数据处理工具 - PyInstaller打包脚本
用于将Python应用程序打包为Windows exe文件

解决常见打包问题：
1. numpy导入路径冲突
2. 依赖版本检查
3. 资源文件包含
4. 环境兼容性检查
"""
import os
import sys
import shutil
import subprocess
import importlib
import pkg_resources
from pathlib import Path

def check_environment():
    """检查和修复环境问题"""
    print("检查Python环境...")
    
    # 检查Python版本
    python_version = sys.version_info
    print(f"Python版本: {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    if python_version < (3, 8):
        print("[ERROR] 需要Python 3.8或更高版本")
        return False
    
    # 检查numpy问题
    try:
        import numpy
        print(f"[OK] numpy版本: {numpy.__version__}")
        
        # 检查numpy路径是否正确
        numpy_path = Path(numpy.__file__).parent
        if "site-packages" not in str(numpy_path):
            print(f"[WARNING] numpy路径可能有问题: {numpy_path}")
            print("建议在虚拟环境中重新安装numpy")
        
    except ImportError as e:
        print(f"[ERROR] numpy导入失败: {e}")
        print("请尝试重新安装: pip uninstall numpy && pip install numpy")
        return False
    except Exception as e:
        print(f"[ERROR] numpy检查失败: {e}")
        return False
    
    # 检查当前目录是否包含numpy源码
    current_dir = Path.cwd()
    if (current_dir / 'numpy').exists():
        print("[WARNING] 当前目录包含numpy源码，可能导致导入冲突")
        print("建议切换到其他目录运行打包脚本")
    
    return True

def check_dependencies():
    """检查所需依赖包"""
    print("检查依赖包...")
    
    required_packages = {
        'pandas': '2.1.4',
        'openpyxl': '3.1.2', 
        'ttkbootstrap': '1.12.2',
        'PyInstaller': '6.3.0'  # 升级到更稳定的版本
    }
    
    missing_packages = []
    version_mismatches = []
    
    for package, required_version in required_packages.items():
        try:
            if package == 'PyInstaller':
                import PyInstaller
                installed_version = PyInstaller.__version__
            else:
                installed_version = pkg_resources.get_distribution(package).version
                
            print(f"[OK] {package}: {installed_version}")
            
            # 检查版本是否匹配（允许补丁版本差异）
            if not installed_version.startswith(required_version.rsplit('.', 1)[0]):
                version_mismatches.append((package, installed_version, required_version))
                
        except (ImportError, pkg_resources.DistributionNotFound):
            missing_packages.append(package)
            print(f"[ERROR] 缺少包: {package}")
    
    if missing_packages:
        print("\n缺少以下依赖包，请安装：")
        for package in missing_packages:
            print(f"  pip install {package}=={required_packages[package]}")
        return False
    
    if version_mismatches:
        print("\n以下包版本可能不兼容：")
        for package, installed, required in version_mismatches:
            print(f"  {package}: 已安装{installed}, 推荐{required}")
        print("如遇问题，建议安装推荐版本")
    
    return True

def check_project_files():
    """检查项目文件完整性"""
    print("检查项目文件...")
    
    required_files = [
        'main.py',
        'build.spec', 
        'config.py',
        'requirements.txt'
    ]
    
    required_dirs = [
        'src',
        'src/ui',
        'src/core',
        'src/utils',
        'templates',
        'assets'
    ]
    
    missing_files = []
    missing_dirs = []
    
    # 检查文件
    for file_path in required_files:
        if not Path(file_path).exists():
            missing_files.append(file_path)
            print(f"[ERROR] 缺少文件: {file_path}")
        else:
            print(f"[OK] 文件存在: {file_path}")
    
    # 检查目录
    for dir_path in required_dirs:
        if not Path(dir_path).exists():
            missing_dirs.append(dir_path)
            print(f"[WARNING] 缺少目录: {dir_path}")
        else:
            print(f"[OK] 目录存在: {dir_path}")
    
    if missing_files:
        print(f"\n[ERROR] 缺少关键文件，无法继续打包")
        return False
    
    if missing_dirs:
        print(f"\n[WARNING] 缺少部分目录，可能影响功能")
    
    return True

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
            '--noconfirm',  # 不询问覆盖
            'build.spec'
        ]
        
        print(f"执行命令: {' '.join(cmd)}")
        
        # 设置环境变量以避免numpy冲突
        env = os.environ.copy()
        env['PYTHONPATH'] = ''  # 清空PYTHONPATH
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            env=env,
            cwd=str(Path.cwd())
        )
        
        if result.returncode == 0:
            print("[OK] PyInstaller构建成功！")
            
            # 检查常见警告
            if "WARNING" in result.stdout or "WARNING" in result.stderr:
                print("\n构建过程中有警告信息：")
                warnings = []
                for line in (result.stdout + result.stderr).split('\n'):
                    if "WARNING" in line:
                        warnings.append(line.strip())
                
                for warning in warnings[:5]:  # 只显示前5个警告
                    print(f"  {warning}")
                
                if len(warnings) > 5:
                    print(f"  ... 还有{len(warnings)-5}个警告")
            
            return True
        else:
            print("[ERROR] PyInstaller构建失败：")
            print("STDOUT:", result.stdout)
            print("STDERR:", result.stderr)
            
            # 分析常见错误
            error_output = result.stdout + result.stderr
            if "numpy" in error_output.lower():
                print("\n[提示] 检测到numpy相关错误，请尝试：")
                print("1. 在虚拟环境中重新安装numpy")
                print("2. 确保不在numpy源码目录中运行")
                print("3. 清理Python缓存：python -Bc \"import pathlib; [p.unlink() for p in pathlib.Path('.').rglob('*.py[co]')]\"")
            
            if "module not found" in error_output.lower():
                print("\n[提示] 检测到模块缺失错误，请检查：")
                print("1. 所有依赖是否已安装")
                print("2. 虚拟环境是否正确激活")
                print("3. PYTHONPATH是否设置正确")
            
            return False
            
    except Exception as e:
        print(f"[ERROR] 构建过程出错: {str(e)}")
        return False

def check_output():
    """检查输出文件"""
    print("检查输出文件...")
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
        
        # 检查文件大小是否合理
        if size_mb < 20:
            print(f"   [WARNING] 文件大小偏小，可能缺少依赖")
        elif size_mb > 200:
            print(f"   [WARNING] 文件大小较大，考虑优化")
        else:
            print(f"   [OK] 文件大小正常")
        
    # 检查资源文件是否包含
    resource_dirs = ['templates', 'assets']
    for resource_dir in resource_dirs:
        if Path(resource_dir).exists():
            dist_resource = dist_dir / resource_dir
            if dist_resource.exists():
                print(f"[OK] 资源目录已包含: {resource_dir}")
            else:
                print(f"[WARNING] 资源目录未包含: {resource_dir}")
        
    return True

def create_installer_info():
    """创建安装和使用说明"""
    info_content = f"""
# Excel数据处理工具 - 使用说明

## 版本信息
- 打包时间: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- Python版本: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}

## 安装步骤
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

## 文件说明
- Excel数据处理工具.exe: 主程序文件
- templates/: 模板文件目录
- assets/: 资源文件目录（如有）
- 使用说明.txt: 本说明文件

## 系统要求
- Windows 10或更高版本
- 至少500MB可用磁盘空间
- 支持Excel文件格式：.xlsx, .xls

## 注意事项
- 支持.xlsx和.xls格式的Excel文件
- 模板文件必须是.xlsx格式
- 输出文件将自动命名并保存到桌面
- 首次运行可能需要几秒钟启动时间
- 如果杀毒软件误报，请添加信任

## 故障排除
如遇问题，请检查：
1. Excel文件格式是否正确
2. 模板文件是否存在且可读
3. 桌面是否有写入权限
4. 是否有足够的磁盘空间
5. Windows防火墙是否阻止程序运行

## 常见错误解决
- "找不到模板文件": 检查templates目录是否完整
- "无法保存文件": 检查桌面写入权限
- "程序无法启动": 以管理员身份运行
- "处理失败": 检查Excel文件是否损坏

## 联系方式
如需技术支持，请提供：
- 错误截图
- 操作步骤
- 使用的Excel文件示例（如可分享）
"""
    
    info_file = Path('dist') / '使用说明.txt'
    
    try:
        with open(info_file, 'w', encoding='utf-8') as f:
            f.write(info_content)
        print(f"[OK] 已创建使用说明: {info_file}")
    except Exception as e:
        print(f"[ERROR] 创建使用说明失败: {str(e)}")

def copy_resources():
    """复制必要的资源文件到dist目录"""
    print("复制资源文件...")
    
    dist_dir = Path('dist')
    if not dist_dir.exists():
        print("[ERROR] dist目录不存在")
        return False
    
    resources = [
        ('templates', '模板文件'),
        ('assets', '资源文件'),
        ('settings.json', '配置文件')
    ]
    
    success = True
    
    for resource, description in resources:
        src = Path(resource)
        if src.exists():
            dst = dist_dir / resource
            try:
                if src.is_dir():
                    if dst.exists():
                        shutil.rmtree(dst)
                    shutil.copytree(src, dst)
                else:
                    shutil.copy2(src, dst)
                print(f"[OK] 已复制{description}: {resource}")
            except Exception as e:
                print(f"[ERROR] 复制{description}失败: {e}")
                success = False
        else:
            print(f"[WARNING] {description}不存在: {resource}")
    
    return success

def main():
    """主函数"""
    print("=" * 60)
    print("Excel数据处理工具 - PyInstaller打包脚本")
    print("=" * 60)
    
    # 检查系统环境
    print("\n[步骤1] 环境检查")
    if not check_environment():
        print("[FAILED] 环境检查失败，请修复后重试")
        return
    
    # 检查依赖包
    print("\n[步骤2] 依赖检查")
    if not check_dependencies():
        print("[FAILED] 依赖检查失败，请安装缺失的包")
        return
    
    # 检查项目文件
    print("\n[步骤3] 项目文件检查")  
    if not check_project_files():
        print("[FAILED] 项目文件检查失败，请补充缺失文件")
        return
    
    print("\n[INFO] 所有预检查通过，开始构建...")
    
    # 执行构建步骤
    success = True
    
    # 步骤4：清理
    print("\n[步骤4] 清理构建目录")
    clean_build_dirs()
    
    # 步骤5：构建
    print("\n[步骤5] 执行PyInstaller构建")
    success &= build_exe()
    
    if success:
        # 步骤6：检查输出
        print("\n[步骤6] 检查构建输出")
        success &= check_output()
        
        if success:
            # 步骤7：复制资源
            print("\n[步骤7] 复制资源文件")
            copy_resources()
            
            # 步骤8：创建说明
            print("\n[步骤8] 创建使用说明")
            create_installer_info()
            
            print("\n" + "=" * 60)
            print("[SUCCESS] 打包完成！")
            print("exe文件位置: dist/Excel数据处理工具.exe")
            print("请将整个dist目录分发给用户")
            print("\n打包内容：")
            print("- Excel数据处理工具.exe (主程序)")
            print("- templates/ (模板文件)")
            print("- assets/ (资源文件，如有)")
            print("- 使用说明.txt (用户手册)")
            print("=" * 60)
        else:
            print("\n[ERROR] 输出检查失败")
    else:
        print("\n[ERROR] 构建失败，请检查错误信息")
        print("\n常见解决方案：")
        print("1. 创建干净的虚拟环境重新安装依赖")
        print("2. 确保不在numpy源码目录中运行")
        print("3. 清理Python缓存文件")
        print("4. 检查PyInstaller和依赖包版本兼容性")

if __name__ == "__main__":
    main()