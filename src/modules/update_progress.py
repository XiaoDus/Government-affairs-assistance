# 上传研判记录

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import requests
import pandas as pd
import os
import json

# 抑制urllib3未验证HTTPS请求的警告
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 配置文件路径
CONFIG_FILE = 'config.json'

# 显示进度条
def show_progress_bar(total, title="处理进度"):
    """
    显示进度条窗口
    :param total: 总任务数
    :param title: 窗口标题
    :return: 窗口对象、更新进度函数、关闭窗口函数
    """
    progress_window = tk.Toplevel()
    progress_window.title(title)
    progress_window.geometry("400x120")
    progress_window.resizable(False, False)
    progress_window.transient()
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(
        progress_window,
        variable=progress_var,
        maximum=100,
        length=350,
        mode='determinate'
    )
    progress_bar.pack(pady=20, padx=20)
    
    status_label = tk.Label(progress_window, text="准备开始...", font=("微软雅黑", 10))
    status_label.pack(pady=10)
    
    progress_window.update_idletasks()
    width = progress_window.winfo_width()
    height = progress_window.winfo_height()
    x = (progress_window.winfo_screenwidth() // 2) - (width // 2)
    y = (progress_window.winfo_screenheight() // 2) - (height // 2)
    progress_window.geometry(f'{width}x{height}+{x}+{y}')
    
    def update_progress_bar(current, total, status):
        percent = (current / total) * 100
        progress_var.set(percent)
        status_label.config(text=status)
        progress_window.update()
    
    def close():
        progress_window.destroy()
    
    return progress_window, update_progress_bar, close

# 加载配置
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

# 保存配置
def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

# 获取服务器基础URL
def get_base_url():
    """从配置文件获取服务器基础URL"""
    config = load_config()
    host = config.get('server_host', 'your_server_ip_here')
    port = config.get('server_port', 'your_server_port_here')
    return f'https://{host}:{port}'

# 文件类型映射
def get_file_type(file_path):
    """
    根据文件扩展名获取MIME类型
    :param file_path: 文件路径
    :return: MIME类型
    """
    ext = os.path.splitext(file_path)[1].lower()
    mime_types = {
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

def upload_file(file_path, token):
    """
    上传文件到服务器
    :param file_path: 要上传的文件路径
    :param token: 访问令牌
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
            # 发送POST请求
            response = requests.post(url, headers=headers, files=files, verify=False)
        
        # 检查响应状态
        response.raise_for_status()
        
        # 解析响应数据
        result = response.json()
        return result
    except Exception as error:
        print(f'上传错误: {error}')
        raise

def update_progress(issue_id, circulation_id, opinion, attachment_list, token):
    """
    更新纠纷进度
    :param issue_id: 问题ID
    :param circulation_id: 流转ID
    :param opinion: 处理意见
    :param attachment_list: 附件列表
    :param token: 访问令牌
    :return: 更新结果
    """
    # API URL
    url = f'{get_base_url()}/event-center/api/circulation/updateProgress'
    
    # 准备请求头
    headers = {
        'x-access-token': token,
        'Content-Type': 'application/json'
    }
    
    # 准备请求数据
    data = {
        "issueId": issue_id,
        "circulationId": circulation_id,
        "operationType": 100009,  # 固定值
        "mediationType": 2,  # 固定值
        "opinion": opinion,
        "attachmentList": attachment_list
    }
    
    try:
        # 发送POST请求
        response = requests.post(url, headers=headers, json=data, verify=False)
        
        # 检查响应状态
        response.raise_for_status()
        
        # 解析响应数据
        result = response.json()
        return result
    except Exception as error:
        print(f'更新进度错误: {error}')
        raise

def process_excel():
    """
    处理Excel文件并更新纠纷进度
    """
    # 加载配置
    config = load_config()
    saved_token = config.get('token', '')
    
    # 创建主窗口
    root = tk.Tk()
    root.title("选择Excel文件")
    root.geometry("700x250")
    root.resizable(False, False)
    
    # 存储选择的文件路径和令牌
    excel_path_var = tk.StringVar()
    judgment_excel_path_var = tk.StringVar()
    token_var = tk.StringVar(value=saved_token)
    
    # 结果变量
    result = {'excel_path': '', 'judgment_excel_path': '', 'token': '', 'cancelled': True}
    
    def browse_excel():
        file_path = filedialog.askopenfilename(
            title="选择主Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            excel_path_var.set(file_path)
    
    def browse_judgment_excel():
        file_path = filedialog.askopenfilename(
            title="选择研判记录表",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            judgment_excel_path_var.set(file_path)
    
    def confirm():
        result['excel_path'] = excel_path_var.get()
        result['judgment_excel_path'] = judgment_excel_path_var.get()
        result['token'] = token_var.get()
        result['cancelled'] = False
        # 保存令牌
        config['token'] = result['token']
        save_config(config)
        root.destroy()
    
    def cancel():
        result['cancelled'] = True
        root.destroy()
    
    # 处理窗口关闭事件
    def on_closing():
        result['cancelled'] = True
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 创建界面元素
    # 主Excel文件选择
    excel_frame = tk.Frame(root)
    excel_frame.pack(pady=10, fill=tk.X, padx=20)
    tk.Label(excel_frame, text="主Excel文件：", font=("微软雅黑", 10), width=15, anchor=tk.W).pack(side=tk.LEFT, padx=5)
    excel_entry_frame = tk.Frame(excel_frame)
    excel_entry_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Entry(excel_entry_frame, textvariable=excel_path_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    tk.Button(excel_entry_frame, text="浏览...", command=browse_excel, relief=tk.FLAT, bg="#f0f0f0", padx=10, pady=2).pack(side=tk.LEFT)
    
    # 研判记录表选择
    judgment_frame = tk.Frame(root)
    judgment_frame.pack(pady=10, fill=tk.X, padx=20)
    tk.Label(judgment_frame, text="研判记录表：", font=("微软雅黑", 10), width=15, anchor=tk.W).pack(side=tk.LEFT, padx=5)
    judgment_entry_frame = tk.Frame(judgment_frame)
    judgment_entry_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Entry(judgment_entry_frame, textvariable=judgment_excel_path_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    tk.Button(judgment_entry_frame, text="浏览...", command=browse_judgment_excel, relief=tk.FLAT, bg="#f0f0f0", padx=10, pady=2).pack(side=tk.LEFT)
    
    # 令牌输入框
    token_frame = tk.Frame(root)
    token_frame.pack(pady=10, fill=tk.X, padx=20)
    tk.Label(token_frame, text="用户令牌：", font=("微软雅黑", 10), width=15, anchor=tk.W).pack(side=tk.LEFT, padx=5)
    tk.Entry(token_frame, textvariable=token_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
    
    # 按钮区域
    button_frame = tk.Frame(root)
    button_frame.pack(pady=20)
    
    tk.Button(button_frame, text="确定", command=confirm, width=12, bg="#4CAF50", fg="white", relief=tk.FLAT, padx=15, pady=5).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="取消", command=cancel, width=12, relief=tk.FLAT, bg="#f0f0f0", padx=15, pady=5).pack(side=tk.LEFT, padx=10)
    
    # 设置默认按钮为回车键
    root.bind('<Return>', lambda event: confirm())
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()
    
    # 检查是否取消
    if result['cancelled']:
        print("操作取消")
        return
    
    excel_path = result['excel_path']
    judgment_excel_path = result['judgment_excel_path']
    token = result['token']
    
    if not excel_path:
        print("未选择主Excel文件")
        return
    
    if not judgment_excel_path:
        print("未选择研判记录表")
        return
    
    if not token:
        print("未输入用户令牌")
        return
    
    # 读取研判记录表，建立单号到链接和opinion的映射
    try:
        judgment_df = pd.read_excel(judgment_excel_path)
        
        # 单号固定在A列（索引0）
        order_col = 0
        
        # 自动查找opinion列和链接列
        opinion_col = None
        link_col = None
        
        # 遍历列名，查找可能的列
        for col_idx, col_name in enumerate(judgment_df.columns):
            col_name_str = str(col_name).lower()
            if '研判' in col_name_str or '处置' in col_name_str or '意见' in col_name_str:
                opinion_col = col_idx
            elif '链接' in col_name_str or 'url' in col_name_str or 'message' in col_name_str:
                link_col = col_idx
        
        # 如果没有找到，使用默认列
        if opinion_col is None:
            opinion_col = 3  # 默认第四列
        if link_col is None:
            link_col = 4  # 默认第五列
        
        print(f"识别列: 单号列=1 (A列), 研判列={opinion_col+1}, 链接列={link_col+1}")
        
        # 建立映射
        order_info_map = {}
        for idx, row in judgment_df.iterrows():
            # 检查单元格是否为NaN
            if pd.notna(row.iloc[order_col]):
                order_no = str(row.iloc[order_col]).strip()
                # 只处理有单号的行
                if order_no:
                    opinion = str(row.iloc[opinion_col]) if len(row) > opinion_col and pd.notna(row.iloc[opinion_col]) else ''
                    link = str(row.iloc[link_col]) if len(row) > link_col and pd.notna(row.iloc[link_col]) else ''
                    order_info_map[order_no] = {
                        'opinion': opinion,
                        'link': link
                    }
        print(f"读取研判记录表成功，共 {len(order_info_map)} 条记录")
    except Exception as e:
        print(f"读取研判记录表失败: {e}")
        return
    
    # 读取Excel文件
    try:
        df = pd.read_excel(excel_path, header=None)
        print(f"读取Excel成功，共 {len(df)} 行数据")
    except Exception as e:
        print(f"读取Excel失败: {e}")
        return
    
    # 记录统计信息
    success_count = 0
    failure_count = 0
    not_found_order_nos = []
    failed_order_nos = []

    # 获取总行数（减去标题行）
    total_rows = len(df) - 1

    # 显示进度条
    progress_window, update_progress_bar, close_progress = show_progress_bar(total_rows, title="上传进度")

    try:
        # 处理每一行数据
        for index, row in df.iterrows():
            current_row = index  # 当前处理的行（从0开始）
            if index == 0:  # 跳过标题行
                continue
            
            try:
                # 获取数据
                issue_id = str(row[0]) if pd.notna(row[0]) else ''  # A列
                circulation_id = str(row[1]) if pd.notna(row[1]) else ''  # B列
                order_no = str(row[3]) if len(row) > 3 and pd.notna(row[3]) else ''  # D列（单号）
                
                # 更新进度条
                update_progress_bar(current_row - 1, total_rows, f"处理第 {current_row} 行: 单号 {order_no}")
                
                if not issue_id or not circulation_id or not order_no:
                    print(f"第 {index+1} 行缺少必要数据，跳过")
                    failure_count += 1
                    failed_order_nos.append(order_no)
                    continue
            except Exception as e:
                print(f"处理第 {index+1} 行失败: {e}")
                failure_count += 1
                failed_order_nos.append(order_no)
                continue
            
            # 从研判记录表中获取opinion和链接
            if order_no in order_info_map:
                opinion = order_info_map[order_no]['opinion']
                url = order_info_map[order_no]['link']
                
                if not url:
                    print(f"单号 {order_no} 对应链接为空，跳过")
                    not_found_order_nos.append(order_no)
                    failure_count += 1
                    failed_order_nos.append(order_no)
                    continue
                
                print(f"处理第 {index+1} 行: issueId={issue_id}, circulationId={circulation_id}, orderNo={order_no}")
                
                # 准备附件列表
                attachment_list = []
                # 从链接中提取文件名
                file_name = os.path.basename(url)
                # 根据链接后缀获取文件类型
                file_type = get_file_type(url)
                
                attachment = {
                    "fileName": file_name,
                    "url": url,
                    "fileType": file_type
                }
                attachment_list.append(attachment)
                print(f"准备附件: {attachment}")
            else:
                print(f"未找到单号 {order_no} 对应的信息")
                not_found_order_nos.append(order_no)
                failure_count += 1
                failed_order_nos.append(order_no)
                continue
            
            # 更新进度
            try:
                result = update_progress(issue_id, circulation_id, opinion, attachment_list, token)
                print(f"更新结果: {result}")
                # 检查success字段
                if result.get('success') == True:
                    print(f"上传单号 {order_no} 成功！")
                    success_count += 1
                else:
                    print(f"上传单号 {order_no} 失败: {result.get('message', '未知错误')}")
                    failure_count += 1
                    failed_order_nos.append(order_no)
            except Exception as e:
                # 检查是否是权限不足的错误
                error_msg = str(e)
                if "权限不足" in error_msg or "401" in error_msg or "unauthorized" in error_msg.lower():
                    print(f"权限不足，需要重新输入令牌: {error_msg}")
                    
                    # 弹出令牌输入窗口
                    def get_new_token():
                        new_token = tk.StringVar()
                        token_window = tk.Toplevel()
                        token_window.title("权限不足")
                        token_window.geometry("500x200")
                        token_window.resizable(False, False)
                        token_window.grab_set()
                        
                        result = {"token": "", "cancelled": True}
                        
                        def confirm_token():
                            result["token"] = new_token.get()
                            result["cancelled"] = False
                            token_window.destroy()
                        
                        def cancel_token():
                            result["cancelled"] = True
                            token_window.destroy()
                        
                        tk.Label(token_window, text="权限不足，请重新输入用户令牌：", font=("微软雅黑", 10)).pack(pady=20)
                        
                        token_frame = tk.Frame(token_window)
                        token_frame.pack(pady=10)
                        tk.Entry(token_frame, textvariable=new_token, width=50).pack(side=tk.LEFT, padx=5)
                        
                        button_frame = tk.Frame(token_window)
                        button_frame.pack(pady=20)
                        
                        tk.Button(button_frame, text="确定", command=confirm_token, width=12, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=10)
                        tk.Button(button_frame, text="取消", command=cancel_token, width=12).pack(side=tk.LEFT, padx=10)
                        
                        # 居中显示窗口
                        token_window.update_idletasks()
                        width = token_window.winfo_width()
                        height = token_window.winfo_height()
                        x = (token_window.winfo_screenwidth() // 2) - (width // 2)
                        y = (token_window.winfo_screenheight() // 2) - (height // 2)
                        token_window.geometry(f'{width}x{height}+{x}+{y}')
                        
                        token_window.mainloop()
                        return result
                    
                    # 获取新令牌
                    token_result = get_new_token()
                    if token_result["cancelled"]:
                        print("用户取消输入令牌")
                        failure_count += 1
                        failed_order_nos.append(order_no)
                        continue
                    
                    # 更新令牌
                    token = token_result["token"]
                    # 保存新令牌
                    config['token'] = token
                    save_config(config)
                    
                    # 重新尝试更新进度
                    try:
                        result = update_progress(issue_id, circulation_id, opinion, attachment_list, token)
                        print(f"使用新令牌更新成功: {result}")
                        success_count += 1
                    except Exception as e2:
                        print(f"使用新令牌更新失败: {e2}")
                        failure_count += 1
                        failed_order_nos.append(order_no)
                        continue
                else:
                    print(f"处理第 {index+1} 行失败: {e}")
                    failure_count += 1
                    failed_order_nos.append(order_no)
                    continue
    
    finally:
        # 关闭进度条
        update_progress_bar(total_rows, total_rows, "任务完成！")
        import time
        time.sleep(0.5)
        close_progress()

    print("处理完成")
    
    # 显示处理完成的综合统计弹窗
    total_count = success_count + failure_count
    message = f"处理完成！\n\n总处理条数: {total_count}\n成功条数: {success_count}\n失败条数: {failure_count}"
    
    if failed_order_nos:
        failed_str = "\n".join(failed_order_nos)
        message += f"\n\n未成功的单号:\n{failed_str}"
    
    tk.messagebox.showinfo("处理完成", message)
    


if __name__ == "__main__":
    process_excel()
