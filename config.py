# -*- coding: utf-8 -*-
"""
配置文件
包含应用程序的全局配置项
"""
import os
from pathlib import Path

# 应用程序信息
APP_NAME = "Excel数据处理工具"
APP_VERSION = "1.0.0"
WINDOW_TITLE = f"{APP_NAME} v{APP_VERSION}"

# 窗口尺寸
MAIN_WINDOW_WIDTH = 650
MAIN_WINDOW_HEIGHT = 600
SETTINGS_WINDOW_WIDTH = 500
SETTINGS_WINDOW_HEIGHT = 400

# 文件路径配置
PROJECT_ROOT = Path(__file__).parent
TEMPLATES_DIR = PROJECT_ROOT / "templates"
ASSETS_DIR = PROJECT_ROOT / "assets"

# 支持的文件格式
SUPPORTED_EXCEL_FORMATS = [
    ("Excel文件", "*.xlsx"),
    ("Excel文件", "*.xls"),
    ("所有文件", "*.*")
]

# 输出路径（桌面）
DESKTOP_PATH = Path.home() / "Desktop"

# 模板类型
TEMPLATE_TYPES = {
    "UPS": "UPS总结单模板",
    "DPD": "DPD数据预报模板"
}

# UI主题
UI_THEME = "cosmo"  # ttkbootstrap主题

# 日志配置
LOG_LEVEL = "INFO"
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"