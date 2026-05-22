# -*- coding: utf-8 -*-
"""
配置管理模块
统一管理应用程序配置
"""

import os
import json
from pathlib import Path
from typing import Dict, Any, Optional


class Config:
    """配置管理类"""
    
    # 默认配置文件路径
    DEFAULT_CONFIG_PATH = 'config/config.json'
    
    # 默认配置值
    DEFAULTS = {
        'server_host': 'your_server_ip_here',
        'server_port': 'your_server_port_here',
        'token': 'your_token_here',
        'default_save_path': 'your_save_path_here'
    }
    
    def __init__(self, config_path: str = None):
        """
        初始化配置管理器
        
        Args:
            config_path: 配置文件路径，默认为 config/config.json
        """
        self.config_path = config_path or self.DEFAULT_CONFIG_PATH
        self._config: Dict[str, Any] = {}
        self._load()
    
    def _load(self) -> None:
        """从文件加载配置"""
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self._config = json.load(f)
            except (json.JSONDecodeError, IOError) as e:
                print(f"加载配置文件失败: {e}")
                self._config = {}
        else:
            self._config = {}
    
    def save(self) -> bool:
        """
        保存配置到文件
        
        Returns:
            bool: 保存是否成功
        """
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
            return True
        except IOError as e:
            print(f"保存配置文件失败: {e}")
            return False
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        获取配置项
        
        Args:
            key: 配置键名
            default: 默认值
            
        Returns:
            配置值
        """
        return self._config.get(key, default or self.DEFAULTS.get(key))
    
    def set(self, key: str, value: Any) -> None:
        """
        设置配置项
        
        Args:
            key: 配置键名
            value: 配置值
        """
        self._config[key] = value
    
    def get_base_url(self) -> str:
        """
        获取服务器基础URL
        
        Returns:
            基础URL，如 https://ip:port
        """
        host = self.get('server_host')
        port = self.get('server_port')
        return f'https://{host}:{port}'
    
    def get_token(self) -> str:
        """
        获取访问令牌
        
        Returns:
            Token字符串
        """
        return self.get('token', '')
    
    def get_save_path(self) -> str:
        """
        获取默认保存路径
        
        Returns:
            保存路径
        """
        return self.get('default_save_path', '')
    
    def is_configured(self) -> bool:
        """
        检查是否已配置
        
        Returns:
            是否所有必要配置都已设置
        """
        required = ['server_host', 'server_port', 'token']
        return all(
            self.get(key) and 
            self.get(key) != self.DEFAULTS.get(key) 
            for key in required
        )


# 全局配置实例
_config_instance: Optional[Config] = None


def get_config() -> Config:
    """
    获取全局配置实例（单例模式）
    
    Returns:
        Config实例
    """
    global _config_instance
    if _config_instance is None:
        _config_instance = Config()
    return _config_instance
