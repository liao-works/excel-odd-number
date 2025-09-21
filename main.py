# -*- coding: utf-8 -*-
"""
Excel数据处理工具 - 主程序入口
功能：读取Excel文件，根据选择的模板类型进行数据填充，输出到桌面
"""
import sys
import os
import logging
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from config import *
from src.ui.main_window import MainWindow

def setup_logging():
    """设置日志配置"""
    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL),
        format=LOG_FORMAT,
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('app.log', encoding='utf-8')
        ]
    )

def main():
    """主函数"""
    try:
        # 设置日志
        setup_logging()
        logger = logging.getLogger(__name__)
        logger.info(f"启动 {APP_NAME} v{APP_VERSION}")
        
        # 创建主窗口并运行
        app = MainWindow()
        app.run()
        
    except Exception as e:
        logging.error(f"应用程序启动失败: {e}")
        input("按回车键退出...")
        sys.exit(1)

if __name__ == "__main__":
    main()