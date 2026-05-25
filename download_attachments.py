# -*- coding: utf-8 -*-
"""
附件下载模块
从API下载纠纷详情中的附件文件
"""

import os
import json
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 配置
from src.core.config import get_config

def clean_filename(filename):
    """清理文件名中的非法字符"""
    illegal_chars = '\\/:*?"<>|'
    for char in illegal_chars:
        filename = filename.replace(char, '')
    filename = filename.replace('\n', '').replace('\r', '').replace('\t', '')
    filename = filename.strip()
    return filename

def download_file(url, token, save_path):
    """下载单个文件"""
    try:
        headers = {'x-access-token': token} if token else {}
        response = requests.get(url, headers=headers, verify=False)
        response.raise_for_status()
        with open(save_path, 'wb') as f:
            f.write(response.content)
        return True
    except Exception as e:
        print(f"下载文件失败 {url}: {e}")
        return False

def main():
    """主函数"""
    config = get_config()
    
    # 从配置获取API地址和Token
    base_url = config.get_base_url()
    token = config.get_token()
    
    # 示例：获取指定ID的纠纷详情
    # 请修改为您需要查询的纠纷ID
    issue_id = "your_issue_id_here"
    api_url = f"{base_url}/event-center/issueDetail/getIssueDetail?id={issue_id}&_t={int(time.time())}"
    
    try:
        print("正在请求API接口...")
        headers = {'x-access-token': token} if token else {}
        response = requests.get(api_url, headers=headers, verify=False)
        response.raise_for_status()
        data = response.json()
        
        print("API响应数据结构:", list(data.keys()))
        
        if 'result' in data:
            detail_data = data['result']
            
            issue_id = detail_data.get('id', '')
            issue_name = detail_data.get('subject', '')
            
            folder_name = clean_filename(f"{issue_id}({issue_name})")
            # 从配置获取保存路径
            save_base_path = config.get_save_path()
            if not save_base_path or save_base_path == 'your_save_path_here':
                print("请先在config.json中配置default_save_path")
                return
                
            folder_path = os.path.join(save_base_path, folder_name)
            
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                print(f"创建文件夹: {folder_path}")
            
            if 'operationLogs' in detail_data and detail_data['operationLogs']:
                operation_logs = detail_data['operationLogs']
                print(f"找到 {len(operation_logs)} 条操作日志")
                
                total_attachments = 0
                
                for log in operation_logs:
                    process_dept = log.get('processDept', '')
                    
                    if '街道' in process_dept:  # 通用匹配
                        print(f"\n找到包含'街道'的记录:")
                        print(f"  处理部门: {process_dept}")
                        print(f"  处理人: {log.get('processor', '')}")
                        print(f"  操作名称: {log.get('operationName', '')}")
                        
                        if 'attachmentList' in log and log['attachmentList']:
                            attachments = log['attachmentList']
                            print(f"  附件数量: {len(attachments)}")
                            
                            for attachment in attachments:
                                full_url = attachment.get('fullUrl', '')
                                file_name = attachment.get('fileName', '')
                                
                                if full_url and file_name:
                                    file_name = clean_filename(file_name)
                                    save_path = os.path.join(folder_path, file_name)
                                    
                                    print(f"  正在下载: {file_name}")
                                    if download_file(full_url, token, save_path):
                                        print(f"  成功保存到: {save_path}")
                                        total_attachments += 1
                                    else:
                                        print(f"  下载失败: {file_name}")
                
                print(f"\n下载完成！共下载 {total_attachments} 个附件")
            else:
                print("该记录没有操作日志")
        else:
            print("响应数据中没有result字段")
            
    except requests.exceptions.RequestException as e:
        print(f"请求API失败: {e}")
    except json.JSONDecodeError as e:
        print(f"解析JSON失败: {e}")
    except Exception as e:
        print(f"发生未知错误: {e}")

if __name__ == "__main__":
    import time
    main()
