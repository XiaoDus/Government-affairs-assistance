# 政务协助系统 / Government Affairs Assistance System

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## 📋 项目简介

本项目是一个**政务协助系统**，主要用于**矛盾纠纷数据的处理、导出和管理**。系统提供了多种实用工具，帮助政务工作人员高效地处理综治信息平台中的纠纷数据。

## ✨ 功能特性

- 🏗️ **模块化架构**：代码结构清晰，易于维护和扩展
- ⚙️ **统一配置管理**：集中式配置，支持配置验证
- 🔧 **API客户端封装**：统一的HTTP请求处理
- 🛠️ **工具函数库**：可复用的文件和UI工具
- 🎨 **主界面入口**：统一的应用程序启动器
- 📦 **依赖管理**：requirements.txt 规范依赖版本
- 🔒 **安全设计**：敏感信息与代码分离

## 🎯 主要功能

| 功能模块 | 说明 |
|---------|------|
| 📊 综合数据处理器 | 集成化数据处理中心，支持数据导出、研判记录上传 |
| 📤 纠纷详情导出 | 导出纠纷详情数据，支持批量操作 |
| 📥 研判记录上传 | 批量上传研判记录，支持进度显示 |
| 📥 批量下载附件 | 从Excel读取ID，批量下载事件附件 |
| 📥 单个下载附件 | 下载指定纠纷的附件文件 |
| 🏢 回访部门管理 | 管理回访部门信息 |
| 🔄 数据合并 | 合并多个数据源 |
| 🌐 Web数据台账 | 基于Web的数据展示和导出 |

## 🛠️ 技术栈

- **Python 3.8+**: 主要开发语言
- **Tkinter**: GUI图形界面
- **Pandas**: 数据处理和分析
- **OpenPyXL**: Excel文件操作
- **python-docx**: Word文档生成
- **Requests**: HTTP请求处理
- **HTML/CSS/JavaScript**: Web界面
- **Bootstrap 5**: 前端UI框架
- **Chart.js**: 数据可视化

## 📁 项目结构

```
Government-affairs-assistance/
├── main.py                          # 🎯 主入口程序（推荐使用）
├── requirements.txt                 # 依赖列表
├── README.md                        # 项目说明
├── .gitignore                       # Git忽略配置
│
├── src/                             # 源代码目录
│   ├── __init__.py
│   ├── core/                        # 核心模块
│   │   ├── __init__.py
│   │   ├── config.py               # 配置管理类
│   │   └── api_client.py           # API客户端封装
│   │
│   ├── utils/                       # 工具模块
│   │   ├── __init__.py
│   │   ├── file_utils.py           # 文件操作工具
│   │   └── ui_utils.py             # UI工具函数
│   │
│   └── modules/                     # 功能模块
│       ├── __init__.py
│       ├── integrated_processor.py  # 综合数据处理器
│       ├── export_issue_detail.py   # 导出纠纷详情
│       ├── update_progress.py       # 上传研判记录
│       ├── add_return_visit_department.py  # 回访部门管理
│       ├── merged.py                # 数据合并工具
│       └── update_file.py           # 文件更新工具
│
├── config/                          # 配置目录
│   └── config.json                  # ⚠️ 本地配置文件（不上传）
│
├── web/                             # Web界面
│   ├── 矛盾纠纷数据台账.html        # 数据台账页面
│   └── js/
│       ├── index.html               # Web主页面
│       └── getResult.js            # 数据获取脚本
│
├── js/                              # 前端资源
│   ├── index.html
│   └── getResult.js
│
└── assets/                          # 资源文件
    ├── icons/                       # 图标目录
    └── templates/                   # 模板目录
```

## 🚀 快速开始

### 1. 环境要求

- Python 3.8+
- Windows操作系统

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 配置项目

1. 编辑 `config/config.json`，填入实际值：
   ```json
   {
     "server_host": "your_server_ip_here",
     "server_port": "your_server_port_here",
     "token": "your_token_here",
     "default_save_path": "your_save_path_here"
   }
   ```

### 4. 运行程序

#### 方式一：使用主界面（推荐）

```bash
python main.py
```

这将打开统一的应用程序界面，包含所有功能的入口。

#### 方式二：直接运行单个模块

```bash
# 综合数据处理器
python integrated_processor.py

# 导出纠纷详情
python export_issue_detail.py

# 上传研判记录
python update_progress.py

# 批量下载附件
python batch_download_attachments.py

# 单个下载附件
python download_attachments.py

# 回访部门管理
python add_return_visit_department.py

# 数据合并
python merged.py
```

## ⚙️ 配置说明

配置文件位于 `config/config.json`，包含以下字段：

| 字段 | 说明 | 占位符 |
|-----|------|--------|
| `server_host` | 服务器IP地址 | `your_server_ip_here` |
| `server_port` | 服务器端口 | `your_server_port_here` |
| `token` | API访问令牌 | `your_token_here` |
| `default_save_path` | 默认保存路径 | `your_save_path_here` |

## 🔒 安全说明

- ✅ 所有敏感信息（IP、端口、Token、路径）已从代码中移除
- ✅ 使用占位符 `your_*_here`，需在配置文件中设置实际值
- ✅ `config/config.json` 已在 `.gitignore` 中排除，不会上传到代码仓库
- ✅ 支持HTTPS请求

## 📝 使用说明

### 导出纠纷详情

1. 运行 `python export_issue_detail.py`
2. 输入有效的Token
3. 选择保存路径
4. 点击导出按钮

### 上传研判记录

1. 运行 `python update_progress.py`
2. 选择Excel文件（一级/二级研判记录表）
3. 系统将自动处理并上传数据
4. 查看进度条了解处理进度

### 批量下载附件

1. 运行 `python batch_download_attachments.py`
2. 选择包含纠纷ID的Excel文件
3. 选择保存位置
4. 输入Token（可选保存）
5. 系统自动下载所有附件

### 数据台账

1. 在浏览器中打开 `矛盾纠纷数据台账.html` 或 `web/矛盾纠纷数据台账.html`
2. 点击按钮获取数据
3. 支持图表展示和数据导出

## 🏗️ 架构说明

### 核心模块 (`src/core/`)

- **config.py**: 配置管理类，支持单例模式
- **api_client.py**: 封装HTTP请求，统一API调用

### 工具模块 (`src/utils/`)

- **file_utils.py**: 文件操作工具（文件名清理、MIME类型等）
- **ui_utils.py**: UI工具（窗口居中、对话框等）

### 功能模块 (`src/modules/`)

各业务功能的具体实现，通过导入核心模块和工具模块复用代码。

## 📦 依赖说明

```
requests>=2.28.0          # HTTP请求
urllib3>=1.26.0           # URL处理
pandas>=1.5.0             # 数据处理
openpyxl>=3.0.10          # Excel操作
python-docx>=0.8.11       # Word文档
```

## 🤝 贡献指南

欢迎提交Issue和Pull Request来改进项目。

### 代码规范

- 遵循 PEP 8 代码风格
- 使用类型注解
- 添加文档字符串
- 保持函数单一职责

## 📄 许可证

本项目采用 MIT 许可证 - 详见 LICENSE 文件

## 👨‍💻 开发者

- **维护者**: XiaoDus
- **项目地址**: https://github.com/XiaoDus/Government-affairs-assistance

## 📞 联系方式

如有问题或建议，请通过GitHub Issue联系。

---

**注意**: 本项目仅供学习和参考使用，请遵守相关法律法规。
