# GitHub Actions 构建和发布说明

## 🚀 自动化工作流功能

本项目配置了GitHub Actions自动化工作流，用于：
- 自动构建Windows可执行文件
- 创建GitHub Release
- 上传构建产物到Release

## 📋 工作流触发条件

### 1. 标签推送触发（推荐）
当推送以`v`开头的标签时自动触发：

```bash
# 创建并推送版本标签
git tag v1.0.0
git push origin v1.0.0
```

### 2. 手动触发
在GitHub仓库的Actions页面可以手动运行工作流

## 🔧 工作流步骤

1. **环境准备**
   - 使用Windows latest环境
   - 设置Python 3.11
   - 安装项目依赖

2. **构建过程**
   - 执行`build_simple.py`进行单文件打包
   - 验证生成的EXE文件
   - 创建发布文件结构

3. **发布准备**
   - 复制主程序和说明文件
   - 包含模板和资源文件
   - 生成版本信息文件
   - 创建完整的ZIP压缩包

4. **GitHub Release**
   - 自动创建Release页面
   - 上传ZIP压缩包（完整版）
   - 上传单独的EXE文件（简化版）
   - 生成详细的Release说明

## 📦 发布产物

每次成功构建后，Release将包含：

1. **完整压缩包**: `Excel数据处理工具-v1.0.0-Windows.zip`
   - Excel数据处理工具.exe
   - 使用说明.txt
   - 版本信息.txt
   - templates/ (如有)
   - assets/ (如有)

2. **单独EXE**: `Excel数据处理工具-v1.0.0.exe`
   - 仅包含主程序文件

## 🏷️ 版本标签规范

建议使用语义化版本标签：

```bash
v1.0.0    # 主版本.次版本.修订版本
v1.1.0    # 新功能版本
v1.0.1    # 修复版本
v2.0.0    # 重大更新版本
```

## 📝 发布流程示例

```bash
# 1. 完成开发和测试
git add .
git commit -m "feat: 添加新功能"
git push origin main

# 2. 创建版本标签
git tag v1.0.0 -m "发布版本 1.0.0"
git push origin v1.0.0

# 3. GitHub Actions自动构建
# 查看进度: https://github.com/用户名/仓库名/actions

# 4. 自动创建Release
# 查看发布: https://github.com/用户名/仓库名/releases
```

## ⚙️ 配置要求

确保GitHub仓库已配置：

1. **Actions权限**: 仓库设置中启用GitHub Actions
2. **Token权限**: GITHUB_TOKEN有创建Release的权限
3. **分支保护**: 如有需要，配置分支保护规则

## 🔍 故障排除

### 常见问题

1. **构建失败**: 检查dependencies是否完整
2. **权限错误**: 确认GITHUB_TOKEN权限
3. **文件缺失**: 确保templates/assets目录存在

### 调试方法

1. 查看Actions运行日志
2. 检查工作流YAML语法
3. 验证本地构建是否成功

## 📈 使用统计

Release创建后，可以在GitHub Analytics中查看：
- 下载次数统计
- 用户反馈和Issues
- 版本使用分布

---

使用此工作流，每次推送版本标签都会自动构建并发布新版本，大大简化了发布流程。