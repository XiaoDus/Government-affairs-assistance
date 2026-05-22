# -*- coding: utf-8 -*-
"""
UI工具模块
处理界面相关的通用功能
"""

import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional, Dict, Any


def center_window(window: tk.Tk or tk.Toplevel, width: int = None, height: int = None) -> None:
    """
    将窗口居中显示
    
    Args:
        window: 窗口对象
        width: 窗口宽度，如为None则使用当前宽度
        height: 窗口高度，如为None则使用当前高度
    """
    window.update_idletasks()
    
    w = width or window.winfo_width()
    h = height or window.winfo_height()
    
    x = (window.winfo_screenwidth() // 2) - (w // 2)
    y = (window.winfo_screenheight() // 2) - (h // 2)
    
    window.geometry(f'{w}x{h}+{x}+{y}')


def create_token_dialog(parent: tk.Tk = None, current_token: str = "", 
                       title: str = "权限不足") -> Dict[str, Any]:
    """
    创建Token输入对话框
    
    Args:
        parent: 父窗口
        current_token: 当前token值
        title: 窗口标题
        
    Returns:
        包含token和cancelled的字典
    """
    if parent is None:
        parent = tk._default_root
        if parent is None:
            parent = tk.Tk()
            parent.withdraw()
    
    new_token = tk.StringVar(value=current_token)
    token_window = tk.Toplevel(parent)
    token_window.title(title)
    token_window.geometry("500x200")
    token_window.resizable(False, False)
    token_window.grab_set()
    
    result = {"token": "", "cancelled": True}
    
    def confirm():
        result["token"] = new_token.get()
        result["cancelled"] = False
        token_window.destroy()
    
    def cancel():
        result["cancelled"] = True
        token_window.destroy()
    
    # 提示标签
    tk.Label(
        token_window, 
        text="权限不足，请重新输入用户令牌：", 
        font=("微软雅黑", 10)
    ).pack(pady=20)
    
    # 输入框
    token_frame = tk.Frame(token_window)
    token_frame.pack(pady=10)
    tk.Entry(token_frame, textvariable=new_token, width=50).pack(side=tk.LEFT, padx=5)
    
    # 按钮
    button_frame = tk.Frame(token_window)
    button_frame.pack(pady=20)
    tk.Button(
        button_frame, 
        text="确定", 
        command=confirm, 
        width=12, 
        bg="#4CAF50", 
        fg="white"
    ).pack(side=tk.LEFT, padx=10)
    tk.Button(
        button_frame, 
        text="取消", 
        command=cancel, 
        width=12
    ).pack(side=tk.LEFT, padx=10)
    
    # 居中显示
    center_window(token_window)
    
    token_window.wait_window()
    return result


def create_progress_dialog(parent: tk.Tk, title: str = "处理进度") -> tuple:
    """
    创建进度对话框
    
    Args:
        parent: 父窗口
        title: 窗口标题
        
    Returns:
        (窗口对象, 更新进度函数, 关闭函数)
    """
    progress_window = tk.Toplevel(parent)
    progress_window.title(title)
    progress_window.geometry("400x120")
    progress_window.resizable(False, False)
    progress_window.transient(parent)
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(
        progress_window,
        variable=progress_var,
        maximum=100,
        length=350,
        mode='determinate'
    )
    progress_bar.pack(pady=20, padx=20)
    
    status_label = tk.Label(
        progress_window, 
        text="准备开始...", 
        font=("微软雅黑", 10)
    )
    status_label.pack(pady=10)
    
    center_window(progress_window)
    
    def update_progress(current: int, total: int, status: str = ""):
        """更新进度"""
        percent = (current / total) * 100 if total > 0 else 0
        progress_var.set(percent)
        if status:
            status_label.config(text=status)
        progress_window.update()
    
    def close():
        """关闭窗口"""
        progress_window.destroy()
    
    return progress_window, update_progress, close


def create_file_dialog(title: str = "选择文件", file_types: list = None) -> Optional[str]:
    """
    创建文件选择对话框
    
    Args:
        title: 对话框标题
        file_types: 文件类型列表
        
    Returns:
        选择的文件路径，取消则返回None
    """
    from tkinter import filedialog
    
    if file_types is None:
        file_types = [("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
    
    return filedialog.askopenfilename(title=title, filetypes=file_types)


def create_folder_dialog(title: str = "选择文件夹") -> Optional[str]:
    """
    创建文件夹选择对话框
    
    Args:
        title: 对话框标题
        
    Returns:
        选择的文件夹路径，取消则返回None
    """
    from tkinter import filedialog
    
    return filedialog.askdirectory(title=title)
