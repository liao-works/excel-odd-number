# -*- coding: utf-8 -*-
"""
文件处理工具类
提供文件操作相关的辅助功能
"""
import os
import shutil
import logging
from pathlib import Path
from datetime import datetime

class FileHandler:
    """文件处理工具类"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
    def ensure_directory_exists(self, directory_path):
        """
        确保目录存在，不存在则创建
        
        Args:
            directory_path (str|Path): 目录路径
            
        Returns:
            bool: 操作结果
        """
        try:
            Path(directory_path).mkdir(parents=True, exist_ok=True)
            return True
        except Exception as e:
            self.logger.error(f"创建目录失败: {str(e)}")
            return False
            
    def safe_copy_file(self, source, destination):
        """
        安全复制文件
        
        Args:
            source (str|Path): 源文件路径
            destination (str|Path): 目标文件路径
            
        Returns:
            bool: 操作结果
        """
        try:
            # 确保目标目录存在
            dest_dir = Path(destination).parent
            self.ensure_directory_exists(dest_dir)
            
            # 复制文件
            shutil.copy2(source, destination)
            return True
        except Exception as e:
            self.logger.error(f"复制文件失败: {str(e)}")
            return False
            
    def get_unique_filename(self, file_path):
        """
        获取唯一的文件名（如果文件已存在，则添加序号）
        
        Args:
            file_path (str|Path): 文件路径
            
        Returns:
            Path: 唯一的文件路径
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            return file_path
            
        # 获取文件名和扩展名
        stem = file_path.stem
        suffix = file_path.suffix
        parent = file_path.parent
        
        # 添加序号直到找到不存在的文件名
        counter = 1
        while True:
            new_name = f"{stem}_{counter}{suffix}"
            new_path = parent / new_name
            if not new_path.exists():
                return new_path
            counter += 1
            
    def get_file_info(self, file_path):
        """
        获取文件信息
        
        Args:
            file_path (str|Path): 文件路径
            
        Returns:
            dict: 文件信息字典
        """
        try:
            file_path = Path(file_path)
            
            if not file_path.exists():
                return None
                
            stat = file_path.stat()
            
            return {
                'name': file_path.name,
                'stem': file_path.stem,
                'suffix': file_path.suffix,
                'size': stat.st_size,
                'size_mb': round(stat.st_size / (1024 * 1024), 2),
                'created': datetime.fromtimestamp(stat.st_ctime),
                'modified': datetime.fromtimestamp(stat.st_mtime),
                'is_file': file_path.is_file(),
                'is_dir': file_path.is_dir(),
                'absolute_path': str(file_path.absolute())
            }
            
        except Exception as e:
            self.logger.error(f"获取文件信息失败: {str(e)}")
            return None
            
    def cleanup_temp_files(self, temp_dir, max_age_hours=24):
        """
        清理临时文件
        
        Args:
            temp_dir (str|Path): 临时文件目录
            max_age_hours (int): 最大保留时间（小时）
        """
        try:
            temp_dir = Path(temp_dir)
            
            if not temp_dir.exists():
                return
                
            current_time = datetime.now()
            
            for file_path in temp_dir.iterdir():
                if file_path.is_file():
                    file_modified = datetime.fromtimestamp(file_path.stat().st_mtime)
                    age_hours = (current_time - file_modified).total_seconds() / 3600
                    
                    if age_hours > max_age_hours:
                        try:
                            file_path.unlink()
                            self.logger.info(f"已删除过期临时文件: {file_path.name}")
                        except Exception as e:
                            self.logger.warning(f"删除临时文件失败: {str(e)}")
                            
        except Exception as e:
            self.logger.error(f"清理临时文件失败: {str(e)}")