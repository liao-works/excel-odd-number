# -*- coding: utf-8 -*-
"""
设置窗口UI
包含UPS和DPD模板选择区域
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttk_boot
from ttkbootstrap.constants import *
import json
import os
import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from config import *

class SettingsWindow:
    """设置窗口类"""
    
    def __init__(self, parent):
        self.parent = parent
        self.window = ttk_boot.Toplevel(parent)
        self.window.title("模板设置")
        self.window.geometry(f"{SETTINGS_WINDOW_WIDTH}x{SETTINGS_WINDOW_HEIGHT}")
        self.window.resizable(False, False)
        
        # 设置为模态窗口
        self.window.transient(parent)
        self.window.grab_set()
        
        # 居中显示
        self.center_window()
        
        # 初始化变量
        self.ups_template_path = tk.StringVar()
        self.dpd_template_path = tk.StringVar()
        
        # 加载现有配置
        self.load_settings()
        
        # 创建UI组件
        self.create_widgets()
        self.create_layout()
        
    def center_window(self):
        """将窗口居中显示"""
        self.window.update_idletasks()
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        x = parent_x + (parent_width // 2) - (SETTINGS_WINDOW_WIDTH // 2)
        y = parent_y + (parent_height // 2) - (SETTINGS_WINDOW_HEIGHT // 2)
        
        self.window.geometry(f"{SETTINGS_WINDOW_WIDTH}x{SETTINGS_WINDOW_HEIGHT}+{x}+{y}")
        
    def create_widgets(self):
        """创建UI组件"""
        # 主标题
        self.title_label = ttk_boot.Label(
            self.window,
            text="模板文件设置",
            font=("微软雅黑", 14, "bold"),
            bootstyle="primary"
        )
        
        # UPS模板设置框架
        self.ups_frame = ttk_boot.LabelFrame(
            self.window,
            text="UPS总结单模板",
            padding=15,
            bootstyle="info"
        )
        
        # UPS模板路径显示
        self.ups_path_label = ttk_boot.Label(
            self.ups_frame,
            textvariable=self.ups_template_path,
            wraplength=350,
            justify="left",
            font=("微软雅黑", 9),
            bootstyle="secondary"
        )
        
        # UPS模板选择按钮
        self.ups_select_btn = ttk_boot.Button(
            self.ups_frame,
            text="📁 选择UPS模板",
            command=lambda: self.select_template("UPS"),
            bootstyle="outline-info",
            width=20
        )
        
        # DPD模板设置框架
        self.dpd_frame = ttk_boot.LabelFrame(
            self.window,
            text="DPD数据预报模板",
            padding=15,
            bootstyle="success"
        )
        
        # DPD模板路径显示
        self.dpd_path_label = ttk_boot.Label(
            self.dpd_frame,
            textvariable=self.dpd_template_path,
            wraplength=350,
            justify="left",
            font=("微软雅黑", 9),
            bootstyle="secondary"
        )
        
        # DPD模板选择按钮
        self.dpd_select_btn = ttk_boot.Button(
            self.dpd_frame,
            text="📁 选择DPD模板",
            command=lambda: self.select_template("DPD"),
            bootstyle="outline-success",
            width=20
        )
        
        # 按钮框架
        self.button_frame = ttk_boot.Frame(self.window)
        
        # 保存按钮
        self.save_btn = ttk_boot.Button(
            self.button_frame,
            text="💾 保存设置",
            command=self.save_settings,
            bootstyle="success",
            width=12
        )
        
        # 取消按钮
        self.cancel_btn = ttk_boot.Button(
            self.button_frame,
            text="❌ 取消",
            command=self.close_window,
            bootstyle="outline-secondary",
            width=12
        )
        
        # 重置按钮
        self.reset_btn = ttk_boot.Button(
            self.button_frame,
            text="🔄 重置",
            command=self.reset_settings,
            bootstyle="outline-warning",
            width=12
        )
        
    def create_layout(self):
        """布局UI组件"""
        # 主标题
        self.title_label.pack(pady=15)
        
        # UPS模板设置
        self.ups_frame.pack(fill="x", padx=20, pady=10)
        self.ups_path_label.pack(anchor="w", pady=5)
        self.ups_select_btn.pack(anchor="w", pady=5)
        
        # DPD模板设置
        self.dpd_frame.pack(fill="x", padx=20, pady=10)
        self.dpd_path_label.pack(anchor="w", pady=5)
        self.dpd_select_btn.pack(anchor="w", pady=5)
        
        # 按钮区域
        self.button_frame.pack(fill="x", padx=20, pady=20)
        self.save_btn.pack(side="left", padx=5)
        self.cancel_btn.pack(side="left", padx=5)
        self.reset_btn.pack(side="right", padx=5)
        
    def select_template(self, template_type):
        """选择模板文件"""
        file_path = filedialog.askopenfilename(
            title=f"选择{TEMPLATE_TYPES[template_type]}",
            filetypes=SUPPORTED_EXCEL_FORMATS,
            initialdir=str(TEMPLATES_DIR)
        )
        
        if file_path:
            if template_type == "UPS":
                self.ups_template_path.set(file_path)
            elif template_type == "DPD":
                self.dpd_template_path.set(file_path)
                
    def load_settings(self):
        """加载现有设置"""
        settings_file = PROJECT_ROOT / "settings.json"
        
        if settings_file.exists():
            try:
                with open(settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    
                self.ups_template_path.set(settings.get("ups_template", "未设置"))
                self.dpd_template_path.set(settings.get("dpd_template", "未设置"))
                
            except Exception as e:
                print(f"加载设置失败: {e}")
                self.set_default_paths()
        else:
            self.set_default_paths()
            
    def set_default_paths(self):
        """设置默认路径"""
        default_ups = TEMPLATES_DIR / "UPS总结单模板.xlsx"
        default_dpd = TEMPLATES_DIR / "DPD数据预报模板.xlsx"
        
        self.ups_template_path.set(str(default_ups) if default_ups.exists() else "未设置")
        self.dpd_template_path.set(str(default_dpd) if default_dpd.exists() else "未设置")
        
    def save_settings(self):
        """保存设置"""
        settings = {
            "ups_template": self.ups_template_path.get(),
            "dpd_template": self.dpd_template_path.get()
        }
        
        settings_file = PROJECT_ROOT / "settings.json"
        
        try:
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
                
            messagebox.showinfo("成功", "设置保存成功！")
            self.close_window()
            
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败:\\n{str(e)}")
            
    def reset_settings(self):
        """重置设置"""
        result = messagebox.askyesno("确认", "确定要重置所有设置吗？")
        
        if result:
            self.set_default_paths()
            messagebox.showinfo("完成", "设置已重置为默认值")
            
    def close_window(self):
        """关闭窗口"""
        self.window.grab_release()
        self.window.destroy()