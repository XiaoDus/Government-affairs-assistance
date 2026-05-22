# -*- coding: utf-8 -*-
"""
API客户端模块
统一处理HTTP请求和API调用
"""

import requests
import urllib3
from typing import Dict, Any, Optional, BinaryIO

from .config import get_config

# 禁用SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class APIClient:
    """API客户端类"""
    
    def __init__(self, token: str = None):
        """
        初始化API客户端
        
        Args:
            token: 访问令牌，如不传入则从配置获取
        """
        self.config = get_config()
        self.token = token or self.config.get_token()
        self.base_url = self.config.get_base_url()
    
    def _get_headers(self) -> Dict[str, str]:
        """
        获取请求头
        
        Returns:
            请求头字典
        """
        return {
            'x-access-token': self.token,
            'Content-Type': 'application/json'
        }
    
    def get(self, endpoint: str, params: Dict = None, **kwargs) -> requests.Response:
        """
        发送GET请求
        
        Args:
            endpoint: API端点（不含基础URL）
            params: 查询参数
            **kwargs: 其他requests参数
            
        Returns:
            Response对象
        """
        url = f"{self.base_url}{endpoint}"
        headers = self._get_headers()
        
        return requests.get(
            url, 
            headers=headers, 
            params=params, 
            verify=False,
            **kwargs
        )
    
    def post(self, endpoint: str, data: Dict = None, json_data: Dict = None, **kwargs) -> requests.Response:
        """
        发送POST请求
        
        Args:
            endpoint: API端点（不含基础URL）
            data: 表单数据
            json_data: JSON数据
            **kwargs: 其他requests参数
            
        Returns:
            Response对象
        """
        url = f"{self.base_url}{endpoint}"
        headers = self._get_headers()
        
        return requests.post(
            url,
            headers=headers,
            data=data,
            json=json_data,
            verify=False,
            **kwargs
        )
    
    def upload_file(self, file_path: str, endpoint: str = '/sys/common/upload') -> Dict[str, Any]:
        """
        上传文件
        
        Args:
            file_path: 文件路径
            endpoint: 上传端点
            
        Returns:
            上传结果
        """
        url = f"{self.base_url}{endpoint}"
        headers = {'x-access-token': self.token}
        
        try:
            with open(file_path, 'rb') as f:
                files = {'file': f}
                response = requests.post(
                    url,
                    headers=headers,
                    files=files,
                    verify=False,
                    timeout=30
                )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            raise Exception(f"文件上传失败: {e}")
    
    def update_progress(self, issue_id: str, circulation_id: str, opinion: str, 
                       attachment_list: list = None) -> Dict[str, Any]:
        """
        更新进度
        
        Args:
            issue_id: 问题ID
            circulation_id: 流转ID
            opinion: 处理意见
            attachment_list: 附件列表
            
        Returns:
            更新结果
        """
        endpoint = '/event-center/api/circulation/updateProgress'
        
        data = {
            "issueId": issue_id,
            "circulationId": circulation_id,
            "operationType": 100009,
            "mediationType": 2,
            "opinion": opinion
        }
        
        if attachment_list:
            data["attachmentList"] = attachment_list
        
        response = self.post(endpoint, json_data=data)
        response.raise_for_status()
        return response.json()
    
    def get_todo_list(self, page_type: int = None, **kwargs) -> Dict[str, Any]:
        """
        获取待办列表
        
        Args:
            page_type: 页面类型（1=矛盾纠纷登记, 2=贵和码）
            **kwargs: 其他查询参数
            
        Returns:
            待办列表数据
        """
        endpoint = '/event-center/homePage/getMyTodoList'
        
        import time
        params = {
            '_t': int(time.time()),
            'pageNo': 1,
            'pageSize': 1000,
            'sort': 'createTime',
            'order': 'descend',
            'serialNumber': '',
            'subject': ''
        }
        
        if page_type:
            params['pageType'] = page_type
        
        params.update(kwargs)
        
        response = self.get(endpoint, params=params)
        response.raise_for_status()
        return response.json()
