# 上传图片到服务器
import tkinter as tk
from tkinter import filedialog
import requests
import urllib3
import json
import os

# 抑制urllib3的安全警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 配置文件路径
CONFIG_FILE = "config.json"

def load_config():
    """
    加载配置
    :return: 配置字典
    """
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载配置文件失败: {e}")
    return {}

def save_config(config):
    """
    保存配置
    :param config: 配置字典
    """
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置文件失败: {e}")

def get_base_url():
    """从配置文件获取服务器基础URL"""
    config = load_config()
    host = config.get('server_host', 'your_server_ip_here')
    port = config.get('server_port', '39999')
    return f'https://{host}:{port}'

def upload_file(file_path, token):
    """
    上传图片到服务器
    :param file_path: 要上传的文件路径
    :param token: 用户令牌
    :return: 上传结果
    """
    # 上传URL
    url = f'{get_base_url()}/sys/common/upload'
    
    # 准备请求头
    headers = {
        'x-access-token': token
    }
    
    try:
        # 打开文件并准备上传
        with open(file_path, 'rb') as f:
            files = {'file': f}
            # 发送POST请求，设置30秒超时
            response = requests.post(url, headers=headers, files=files, verify=False, timeout=30)
        
        # 检查响应状态
        response.raise_for_status()
        
        # 解析响应数据
        result = response.json()
        return result
    except Exception as error:
        print(f'上传错误: {error}')
        raise

def select_and_upload(token=None):
    """
    打开文件选择对话框并上传选中的图片
    :param token: 用户令牌，如果为None则从配置文件加载
    """
    # 如果没有提供token，从配置文件加载
    if token is None:
        config = load_config()
        token = config.get('token', '')
    
    # 创建Tkinter根窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏根窗口
    
    # 打开文件选择对话框
    file_path = filedialog.askopenfilename(
        title="选择要上传的图片",
        filetypes=[("图片文件", "*.jpg *.jpeg *.png *.gif *.bmp"), ("所有文件", "*.*")]
    )
    
    if file_path:
        print(f"选择的文件: {file_path}")
        try:
            result = upload_file(file_path, token)
            print(f"上传成功: {result}")
            return result
        except Exception as e:
            print(f"上传失败: {e}")
    else:
        print("未选择文件")

if __name__ == "__main__":
    # 这里可以设置一个默认的token用于测试
    default_token = ""
    select_and_upload(default_token)
