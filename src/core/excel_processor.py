# -*- coding: utf-8 -*-
"""
Excel文件处理核心逻辑
负责读取Excel文件、数据验证、错误处理等
"""
import pandas as pd
import logging
import json
import sys
from pathlib import Path
from datetime import datetime
from src.core.ups.ups_processor import UPSDataProcessor
from src.core.dpd.dpd_processor import DPDProcessor
# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from config import *

class ExcelProcessor:
    """Excel文件处理器"""

    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def process_file(self, input_file, template_type, detail_file=None):
        """
        处理Excel文件

        Args:
            input_file (str): 输入文件路径
            template_type (str): 模板类型 ("UPS" 或 "DPD")
            detail_file (str, optional): 单件明细表文件路径

        Returns:
            str: 输出文件路径，失败返回None
        """
        try:
          output_path = self.get_output_path(template_type)
          template_path = self.get_template_path(template_type)

          original_file_data = self.get_original_file_data(input_file, 0)
          original_detail_file_data = self.get_original_file_data(detail_file, 0)

          if template_type == "UPS":
              ups_processor = UPSDataProcessor()
              ups_processor.process_ups_data(original_file_data, original_detail_file_data, template_path, output_path)
          elif template_type == "DPD":
              dpd_processor = DPDProcessor()
              dpd_processor.process_dpd_data(original_file_data, original_detail_file_data, template_path, output_path)

          return output_path

        except Exception as e:
            self.logger.error(f"处理Excel文件时出错: {str(e)}")
            return None

    def get_original_file_data(self, input_file, sheet_index):
        """
        使用pandas获取指定sheet index的数据

        Args:
            input_file (str): Excel文件路径
            sheet_index (int|str): sheet索引或名称
                - int: sheet的索引位置 (0为第一个sheet)
                - str: sheet的名称

        Returns:
            pd.DataFrame: 指定sheet的数据，失败返回None
        """
        try:
            self.logger.info(f"开始读取Excel文件: {input_file}")

            # 验证文件是否存在
            if not Path(input_file).exists():
                self.logger.error(f"文件不存在: {input_file}")
                return None

            # 先获取所有sheet信息
            excel_file = pd.ExcelFile(input_file)
            sheet_names = excel_file.sheet_names
            self.logger.info(f"文件包含的sheet: {sheet_names}")

            # 处理不同类型的sheet_index参数
            target_sheet = None

            if isinstance(sheet_index, int):
                # 使用索引访问sheet
                if 0 <= sheet_index < len(sheet_names):
                    target_sheet = sheet_names[sheet_index]
                    self.logger.info(f"使用索引 {sheet_index} 访问sheet: '{target_sheet}'")
                else:
                    self.logger.error(f"sheet索引 {sheet_index} 超出范围 (0-{len(sheet_names)-1})")
                    return None

            elif isinstance(sheet_index, str):
                # 使用名称访问sheet
                if sheet_index in sheet_names:
                    target_sheet = sheet_index
                    self.logger.info(f"使用名称访问sheet: '{target_sheet}'")
                else:
                    self.logger.error(f"未找到名为 '{sheet_index}' 的sheet")
                    self.logger.info(f"可用的sheet名称: {sheet_names}")
                    return None
            else:
                self.logger.error(f"不支持的sheet_index类型: {type(sheet_index)}")
                return None

            # 读取指定sheet的数据
            df = pd.read_excel(input_file, sheet_name=target_sheet)

            # 数据清理：删除完全空白的行和列
            original_shape = df.shape
            df = df.dropna(how='all').dropna(axis=1, how='all')
            cleaned_shape = df.shape

            self.logger.info(f"sheet '{target_sheet}' 数据读取成功:")
            self.logger.info(f"  - 原始大小: {original_shape[0]}行 x {original_shape[1]}列")
            self.logger.info(f"  - 清理后大小: {cleaned_shape[0]}行 x {cleaned_shape[1]}列")
            self.logger.info(f"  - 列名: {df.columns.tolist()}")

            # 验证数据是否为空
            if df.empty:
                self.logger.warning(f"sheet '{target_sheet}' 没有有效数据")
                return df  # 返回空DataFrame而不是None，便于后续处理

            # 显示前几行数据样本（用于调试）
            if len(df) > 0:
                self.logger.debug(f"数据样本（前3行）:")
                for i, row in df.head(3).iterrows():
                    self.logger.debug(f"  第{i+1}行: {row.to_dict()}")

            return df

        except FileNotFoundError:
            self.logger.error(f"文件未找到: {input_file}")
            return None

        except PermissionError:
            self.logger.error(f"文件访问权限不足: {input_file}")
            return None

        except pd.errors.EmptyDataError:
            self.logger.error(f"文件为空或无有效数据: {input_file}")
            return None

        except pd.errors.ParserError as e:
            self.logger.error(f"文件解析错误: {str(e)}")
            return None

        except Exception as e:
            self.logger.error(f"读取Excel文件时发生未知错误: {str(e)}")
            import traceback
            self.logger.error(traceback.format_exc())
            return None

        finally:
            # 确保ExcelFile对象被正确关闭
            try:
                if 'excel_file' in locals():
                    excel_file.close()
            except:
                pass

    def get_output_path(self, template_type):
        """
        获取输出文件路径
        """
        return DESKTOP_PATH / f"{template_type}总结单-{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    def get_template_path(self, template_type):
        """
        获取模板文件路径

        Args:
            template_type (str): 模板类型

        Returns:
            str: 模板文件路径，失败返回None
        """
        try:
            # 尝试从设置文件读取
            settings_file = PROJECT_ROOT / "settings.json"

            if settings_file.exists():
                with open(settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)

                if template_type == "UPS":
                    template_path = settings.get("ups_template")
                elif template_type == "DPD":
                    template_path = settings.get("dpd_template")
                else:
                    self.logger.error(f"未知的模板类型: {template_type}")
                    return None

                if template_path and template_path != "未设置" and Path(template_path).exists():
                    return template_path

            # 使用默认模板路径
            default_templates = {
                "UPS": TEMPLATES_DIR / "UPS总结单模板.xlsx",
                "DPD": TEMPLATES_DIR / "DPD数据预报模板.xlsx"
            }

            default_path = default_templates.get(template_type)
            if default_path and default_path.exists():
                return str(default_path)

            self.logger.error(f"找不到{template_type}模板文件")
            return None

        except Exception as e:
            self.logger.error(f"获取模板路径时出错: {str(e)}")
            return None
