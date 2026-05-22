#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
政务协助系统 - 主入口
Government Affairs Assistance System - Main Entry

使用方法:
    python main.py

或者直接运行各个模块:
    python -m src.modules.export_issue_detail
    python -m src.modules.update_progress
    python -m src.modules.add_return_visit_department
"""

import sys
import tkinter as tk
from tkinter import messagebox, ttk

# 添加src目录到路径
sys.path.insert(0, '.')

from src.core.config import get_config
from src.utils.ui_utils import center_window


class MainApplication:
    """主应用程序类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("政务协助系统 v1.0")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        self.config = get_config()
        self._setup_ui()
        center_window(self.root)
    
    def _setup_ui(self):
        """设置界面"""
        # 标题
        title_label = tk.Label(
            self.root,
            text="政务协助系统",
            font=("微软雅黑", 24, "bold"),
            fg="#333"
        )
        title_label.pack(pady=20)
        
        subtitle_label = tk.Label(
            self.root,
            text="Government Affairs Assistance System",
            font=("Arial", 12),
            fg="#666"
        )
        subtitle_label.pack(pady=(0, 30))
        
        # 功能按钮框架
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=20, padx=40, fill=tk.X)
        
        # 功能列表
        functions = [
            ("📊 综合数据处理器", self._run_integrated_processor),
            ("📤 导出纠纷详情", self._run_export_detail),
            ("📥 上传研判记录", self._run_update_progress),
            ("🏢 回访部门管理", self._run_return_visit),
            ("🌐 打开数据台账", self._run_web_dashboard),
        ]
        
        for text, command in functions:
            btn = tk.Button(
                button_frame,
                text=text,
                font=("微软雅黑", 12),
                width=25,
                height=2,
                command=command,
                bg="#f0f0f0",
                relief=tk.FLAT,
                cursor="hand2"
            )
            btn.pack(pady=8, fill=tk.X)
            
            # 悬停效果
            btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#e0e0e0"))
            btn.bind("<Leave>", lambda e, b=btn: b.config(bg="#f0f0f0"))
        
        # 配置检查提示
        if not self.config.is_configured():
            warning_label = tk.Label(
                self.root,
                text="⚠️ 请先配置 config/config.json 文件",
                font=("微软雅黑", 10),
                fg="#ff6b6b"
            )
            warning_label.pack(pady=10)
        
        # 底部信息
        info_frame = tk.Frame(self.root)
        info_frame.pack(side=tk.BOTTOM, pady=15)
        
        tk.Label(
            info_frame,
            text="版本: 1.0.0 | 作者: XiaoDus",
            font=("微软雅黑", 9),
            fg="#999"
        ).pack()
    
    def _run_integrated_processor(self):
        """运行综合数据处理器"""
        try:
            from src.modules.integrated_processor import main
            main()
        except Exception as e:
            messagebox.showerror("错误", f"启动失败: {e}")
    
    def _run_export_detail(self):
        """运行导出纠纷详情"""
        try:
            from src.modules.export_issue_detail import main
            main()
        except Exception as e:
            messagebox.showerror("错误", f"启动失败: {e}")
    
    def _run_update_progress(self):
        """运行上传研判记录"""
        try:
            from src.modules.update_progress import main
            main()
        except Exception as e:
            messagebox.showerror("错误", f"启动失败: {e}")
    
    def _run_return_visit(self):
        """运行回访部门管理"""
        try:
            from src.modules.add_return_visit_department import main
            main()
        except Exception as e:
            messagebox.showerror("错误", f"启动失败: {e}")
    
    def _run_web_dashboard(self):
        """打开Web数据台账"""
        import webbrowser
        import os
        
        html_path = os.path.abspath("web/矛盾纠纷数据台账.html")
        if os.path.exists(html_path):
            webbrowser.open(f"file://{html_path}")
        else:
            messagebox.showerror("错误", "找不到数据台账文件")
    
    def run(self):
        """运行应用程序"""
        self.root.mainloop()


def main():
    """主函数"""
    app = MainApplication()
    app.run()


if __name__ == "__main__":
    main()
