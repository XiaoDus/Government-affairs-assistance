# -*- coding: utf-8 -*-
"""
文件工具模块
处理文件相关的通用功能
"""

import os
from typing import Dict


def clean_filename(filename: str) -> str:
    """
    清理文件名，移除非法字符
    
    Args:
        filename: 原始文件名
        
    Returns:
        清理后的文件名
    """
    illegal_chars = '\\/:*?"<>|'
    for char in illegal_chars:
        filename = filename.replace(char, '')
    filename = filename.replace('\n', '').replace('\r', '').replace('\t', '')
    return filename.strip()


def get_file_type(file_path: str) -> str:
    """
    根据文件扩展名获取MIME类型
    
    Args:
        file_path: 文件路径
        
    Returns:
        MIME类型
    """
    ext = os.path.splitext(file_path)[1].lower()
    
    mime_types: Dict[str, str] = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
        '.doc': 'application/msword',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.pdf': 'application/pdf',
        '.txt': 'text/plain',
        '.xls': 'application/vnd.ms-excel',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    
    return mime_types.get(ext, 'application/octet-stream')


def ensure_dir(path: str) -> bool:
    """
    确保目录存在，如不存在则创建
    
    Args:
        path: 目录路径
        
    Returns:
        是否成功
    """
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except OSError:
        return False


def get_safe_path(base_path: str, filename: str) -> str:
    """
    获取安全的文件路径，避免文件名冲突
    
    Args:
        base_path: 基础目录
        filename: 文件名
        
    Returns:
        安全的完整路径
    """
    clean_name = clean_filename(filename)
    full_path = os.path.join(base_path, clean_name)
    
    # 如果文件已存在，添加序号
    if os.path.exists(full_path):
        name, ext = os.path.splitext(clean_name)
        counter = 1
        while os.path.exists(full_path):
            new_name = f"{name}_{counter}{ext}"
            full_path = os.path.join(base_path, new_name)
            counter += 1
    
    return full_path
