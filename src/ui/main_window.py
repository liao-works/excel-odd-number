# -*- coding: utf-8 -*-
"""
主窗口UI
包含文件上传区域、UPS/DPD选择、处理按钮等功能
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttk_boot
from ttkbootstrap.constants import *
import os
import sys
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root))

from config import *
from src.ui.settings_window import SettingsWindow
from src.core.excel_processor import ExcelProcessor

class MainWindow:
    """主窗口类"""

    def __init__(self):
        self.root = ttk_boot.Window(themename=UI_THEME)
        self.root.title(WINDOW_TITLE)
        self.root.geometry(f"{MAIN_WINDOW_WIDTH}x{MAIN_WINDOW_HEIGHT}")
        self.root.resizable(False, False)

        # 居中显示窗口
        self.center_window()

        # 初始化变量
        self.selected_file = tk.StringVar()
        self.detail_file = tk.StringVar()  # 新增：单件明细表文件
        self.template_type = tk.StringVar(value="UPS")
        self.status_message = tk.StringVar(value="请选择Excel文件")

        # 初始化处理器
        self.processor = ExcelProcessor()

        # 创建UI组件
        self.create_menu()
        self.create_widgets()
        self.create_layout()

    def center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (MAIN_WINDOW_WIDTH // 2)
        y = (self.root.winfo_screenheight() // 2) - (MAIN_WINDOW_HEIGHT // 2)
        self.root.geometry(f"{MAIN_WINDOW_WIDTH}x{MAIN_WINDOW_HEIGHT}+{x}+{y}")

    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="设置", menu=settings_menu)
        settings_menu.add_command(label="模板设置", command=self.open_settings)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)

    def create_widgets(self):
        """创建UI组件"""
        # 主标题
        self.title_label = ttk_boot.Label(
            self.root,
            text="Excel数据处理工具",
            font=("微软雅黑", 14, "bold"),
            bootstyle="primary"
        )

        # 文件上传区域框架
        self.upload_frame = ttk_boot.LabelFrame(
            self.root,
            text="1. 选择主数据文件",
            padding=12,
            bootstyle="info"
        )

        # 主数据按钮框架
        self.main_btn_frame = ttk_boot.Frame(self.upload_frame)

        # 文件选择按钮
        self.select_file_btn = ttk_boot.Button(
            self.main_btn_frame,
            text="📁 选择文件",
            command=self.select_file,
            bootstyle="outline-primary",
            width=15
        )

        # 主数据清除按钮
        self.clear_file_btn = ttk_boot.Button(
            self.main_btn_frame,
            text="🗑️ 清除文件",
            command=self.clear_main_file,
            bootstyle="outline-secondary",
            width=15
        )

        # 文件路径显示
        self.file_path_label = ttk_boot.Label(
            self.upload_frame,
            textvariable=self.selected_file,
            wraplength=400,
            justify="center",
            font=("微软雅黑", 9),
            bootstyle="secondary"
        )

        # 拖拽提示
        self.drag_label = ttk_boot.Label(
            self.upload_frame,
            text="或将Excel文件拖拽到此处",
            font=("微软雅黑", 9),
            bootstyle="light"
        )

        # 单件明细表上传区域框架
        self.detail_frame = ttk_boot.LabelFrame(
            self.root,
            text="2. 选择单件明细表文件（可选）",
            padding=12,
            bootstyle="warning"
        )

        # 明细表按钮框架
        self.detail_btn_frame = ttk_boot.Frame(self.detail_frame)

        # 明细表文件选择按钮
        self.select_detail_btn = ttk_boot.Button(
            self.detail_btn_frame,
            text="📋 选择明细表",
            command=self.select_detail_file,
            bootstyle="outline-warning",
            width=15
        )

        # 明细表清除按钮
        self.clear_detail_btn = ttk_boot.Button(
            self.detail_btn_frame,
            text="🗑️ 清除明细表",
            command=self.clear_detail_file,
            bootstyle="outline-secondary",
            width=15
        )

        # 明细表文件路径显示
        self.detail_path_label = ttk_boot.Label(
            self.detail_frame,
            textvariable=self.detail_file,
            wraplength=400,
            justify="center",
            font=("微软雅黑", 9),
            bootstyle="secondary"
        )

        # 明细表提示
        self.detail_tip_label = ttk_boot.Label(
            self.detail_frame,
            text="单件明细表用于补充详细信息（可选上传）",
            font=("微软雅黑", 9),
            bootstyle="light"
        )

        # 处理类型选择框架
        self.type_frame = ttk_boot.LabelFrame(
            self.root,
            text="3. 选择处理类型",
            padding=12,
            bootstyle="success"
        )

        # UPS选项
        self.ups_radio = ttk_boot.Radiobutton(
            self.type_frame,
            text="UPS总结单",
            variable=self.template_type,
            value="UPS",
            bootstyle="success"
        )

        # DPD选项
        self.dpd_radio = ttk_boot.Radiobutton(
            self.type_frame,
            text="DPD数据预报",
            variable=self.template_type,
            value="DPD",
            bootstyle="success"
        )

        # 处理按钮
        self.process_btn = ttk_boot.Button(
            self.root,
            text="🚀 开始处理",
            command=self.process_file,
            bootstyle="success",
            width=20,
            style="success.TButton"
        )

        # 状态栏
        self.status_frame = ttk_boot.Frame(self.root)

        self.status_label = ttk_boot.Label(
            self.status_frame,
            textvariable=self.status_message,
            font=("微软雅黑", 9),
            bootstyle="info"
        )

        # 进度条
        self.progress = ttk_boot.Progressbar(
            self.status_frame,
            mode="indeterminate",
            bootstyle="success"
        )

    def create_layout(self):
        """布局UI组件"""
        # 主标题
        self.title_label.pack(pady=15)

        # 文件上传区域
        self.upload_frame.pack(fill="x", padx=20, pady=8)
        self.main_btn_frame.pack(pady=8)
        self.select_file_btn.pack(side="left", padx=5)
        self.clear_file_btn.pack(side="left", padx=5)
        self.file_path_label.pack(pady=3)
        self.drag_label.pack(pady=3)

        # 单件明细表上传区域
        self.detail_frame.pack(fill="x", padx=20, pady=8)
        self.detail_btn_frame.pack(pady=8)
        self.select_detail_btn.pack(side="left", padx=5)
        self.clear_detail_btn.pack(side="left", padx=5)
        self.detail_path_label.pack(pady=3)
        self.detail_tip_label.pack(pady=3)

        # 处理类型选择
        self.type_frame.pack(fill="x", padx=20, pady=8)
        self.ups_radio.pack(side="left", padx=20)
        self.dpd_radio.pack(side="left", padx=20)

        # 处理按钮
        self.process_btn.pack(pady=20)

        # 状态栏
        self.status_frame.pack(fill="x", side="bottom", padx=20, pady=10)
        self.status_label.pack(side="left")

        # 绑定拖拽事件
        self.bind_drag_drop()

    def bind_drag_drop(self):
        """绑定拖拽事件"""
        def on_drag_enter(event):
            self.upload_frame.config(bootstyle="warning")

        def on_drag_leave(event):
            self.upload_frame.config(bootstyle="info")

        def on_drop(event):
            self.upload_frame.config(bootstyle="info")
            # 注意：完整的拖拽功能需要额外的库支持
            # 这里先保留接口，后续可以扩展

        # 绑定事件（简化版本）
        self.upload_frame.bind("<Button-1>", lambda e: self.select_file())

    def select_file(self):
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=SUPPORTED_EXCEL_FORMATS,
            initialdir=str(Path.home() / "Desktop")
        )

        if file_path:
            self.selected_file.set(file_path)
            filename = os.path.basename(file_path)
            self.status_message.set(f"已选择文件: {filename}")
            self.process_btn.config(state="normal")

    def select_detail_file(self):
        """选择单件明细表文件"""
        file_path = filedialog.askopenfilename(
            title="选择单件明细表文件",
            filetypes=SUPPORTED_EXCEL_FORMATS,
            initialdir=str(Path.home() / "Desktop")
        )

        if file_path:
            self.detail_file.set(file_path)
            filename = os.path.basename(file_path)
            self.status_message.set(f"已选择明细表: {filename}")
        else:
            self.detail_file.set("")

    def clear_main_file(self):
        """清除主数据文件选择"""
        self.selected_file.set("")
        self.status_message.set("请选择Excel文件")
        self.process_btn.config(state="disabled")

    def clear_detail_file(self):
        """清除明细表文件选择"""
        self.detail_file.set("")
        if self.selected_file.get():
            filename = os.path.basename(self.selected_file.get())
            self.status_message.set(f"已选择文件: {filename}")
        else:
            self.status_message.set("请选择Excel文件")

    def process_file(self):
        """处理文件"""
        if not self.selected_file.get():
            messagebox.showwarning("警告", "请先选择Excel文件")
            return

        try:
            # 显示进度条
            self.progress.pack(side="right", padx=10)
            self.progress.start()
            self.status_message.set("正在处理文件...")
            self.process_btn.config(state="disabled")

            # 更新界面
            self.root.update()

            # 执行处理
            detail_file_path = self.detail_file.get() if self.detail_file.get() else None
            result = self.processor.process_file(
                # 主数据文件
                input_file=self.selected_file.get(),
                # 明细表文件
                detail_file=detail_file_path,
                # 模板类型
                template_type=self.template_type.get(),
            )

            # 停止进度条
            self.progress.stop()
            self.progress.pack_forget()

            if result:
                self.status_message.set("处理完成！文件已保存到桌面")
                messagebox.showinfo("成功", f"文件处理完成！\\n输出文件: {result}")
            else:
                self.status_message.set("处理失败")
                messagebox.showerror("错误", "文件处理失败，请检查文件格式")

        except Exception as e:
            # 停止进度条
            self.progress.stop()
            self.progress.pack_forget()

            self.status_message.set("处理出错")
            messagebox.showerror("错误", f"处理文件时出错:\\n{str(e)}")

        finally:
            self.process_btn.config(state="normal")

    def open_settings(self):
        """打开设置窗口"""
        settings_window = SettingsWindow(self.root)

    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "关于",
            f"{APP_NAME} v{APP_VERSION}\\n\\n"
            "功能：Excel数据模板填充工具\\n"
            "支持UPS总结单和DPD数据预报模板\\n\\n"
            "开发：Python + ttkbootstrap"
        )

    def run(self):
        """运行应用程序"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.root.quit()