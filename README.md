# 政务协助系统 / Government Affairs Assistance System

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Code Style](https://img.shields.io/badge/code%20style-pep8-orange.svg)](https://www.python.org/dev/peps/pep-0008/)

## 📋 项目简介

本项目是一个**政务协助系统**，主要用于**矛盾纠纷数据的处理、导出和管理**。系统提供了多种实用工具，帮助政务工作人员高效地处理综治信息平台中的纠纷数据。

## ✨ 新特性（v1.0 重构版）

- 🏗️ **模块化架构**：代码结构清晰，易于维护和扩展
- ⚙️ **统一配置管理**：集中式配置，支持配置验证
- 🔧 **API客户端封装**：统一的HTTP请求处理
- 🛠️ **工具函数库**：可复用的文件和UI工具
- 🎨 **主界面入口**：统一的应用程序启动器
- 📦 **依赖管理**：requirements.txt 规范依赖版本

## 🎯 主要功能

| 功能模块 | 说明 | 文件 |
|---------|------|------|
| 📊 综合数据处理器 | 集成化数据处理中心 | `src/modules/integrated_processor.py` |
| 📤 纠纷详情导出 | 导出纠纷详情数据 | `src/modules/export_issue_detail.py` |
| 📥 研判记录上传 | 批量上传研判记录 | `src/modules/update_progress.py` |
| 🏢 回访部门管理 | 管理回访部门信息 | `src/modules/add_return_visit_department.py` |
| 🌐 Web数据台账 | 基于Web的数据展示 | `web/矛盾纠纷数据台账.html` |

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
├── main.py                          # 主入口程序
├── requirements.txt                 # 依赖列表
├── README.md                        # 项目说明
├── .gitignore                       # Git忽略配置
│
├── src/                             # 源代码目录
│   ├── __init__.py
│   ├── core/                        # 核心模块
│   │   ├── __init__.py
│   │   ├── config.py               # 配置管理
│   │   └── api_client.py           # API客户端
│   │
│   ├── utils/                       # 工具模块
│   │   ├── __init__.py
│   │   ├── file_utils.py           # 文件工具
│   │   └── ui_utils.py             # UI工具
│   │
│   └── modules/                     # 功能模块
│       ├── __init__.py
│       ├── integrated_processor.py  # 综合数据处理器
│       ├── export_issue_detail.py   # 导出纠纷详情
│       ├── update_progress.py       # 上传研判记录
│       ├── add_return_visit_department.py  # 回访部门管理
│       ├── update_file.py           # 文件更新工具
│       └── merged.py                # 数据合并工具
│
├── config/                          # 配置目录
│   ├── config.json                  # 主配置文件（需自行创建）
│   └── config.example.json          # 配置示例
│
├── web/                             # Web界面
│   ├── 矛盾纠纷数据台账.html
│   └── js/
│       ├── index.html
│       └── getResult.js
│
├── assets/                          # 资源文件
│   ├── icons/                       # 图标
│   └── templates/                   # 模板文件
│
└── docs/                            # 文档（可选）
```

## 🚀 快速开始

### 1. 环境要求

- Python 3.8+
- Windows操作系统（部分功能依赖Windows路径）

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 配置项目

1. 复制配置示例文件：
   ```bash
   copy config/config.example.json config/config.json
   ```

2. 编辑 `config/config.json`，填入实际值：
   ```json
   {
     "server_host": "192.168.1.1",
     "server_port": "39999",
     "token": "your_access_token_here",
     "default_save_path": "D:/导出数据"
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
python -c "from src.modules.integrated_processor import main; main()"

# 导出纠纷详情
python -c "from src.modules.export_issue_detail import main; main()"

# 上传研判记录
python -c "from src.modules.update_progress import main; main()"

# 回访部门管理
python -c "from src.modules.add_return_visit_department import main; main()"
```

## ⚙️ 配置说明

配置文件位于 `config/config.json`，包含以下字段：

| 字段 | 说明 | 示例 |
|-----|------|------|
| `server_host` | 服务器IP地址 | `"192.168.1.1"` |
| `server_port` | 服务器端口 | `"39999"` |
| `token` | API访问令牌 | `"eyJ0eXAiOiJKV1Qi..."` |
| `default_save_path` | 默认保存路径 | `"D:/导出数据"` |

## 🔒 安全说明

- ✅ 所有敏感信息（IP、端口、Token、路径）已从代码中移除
- ✅ 使用占位符 `your_*_here`，需在配置文件中设置实际值
- ✅ 支持HTTPS请求（已禁用证书验证警告）
- ⚠️ 请勿将包含真实配置的 `config.json` 提交到代码仓库

## 📝 使用说明

### 导出纠纷详情

1. 运行主程序 `python main.py`
2. 点击"📤 导出纠纷详情"按钮
3. 输入有效的Token
4. 选择保存路径
5. 点击导出按钮

### 上传研判记录

1. 运行主程序 `python main.py`
2. 点击"📥 上传研判记录"按钮
3. 选择Excel文件
4. 系统将自动处理并上传数据
5. 查看进度条了解处理进度

### 数据台账查看

1. 运行主程序 `python main.py`
2. 点击"🌐 打开数据台账"按钮
3. 或直接在浏览器中打开 `web/矛盾纠纷数据台账.html`

## 🏗️ 架构说明

### 核心模块 (`src/core/`)

- **config.py**: 配置管理类，支持单例模式
- **api_client.py**: 封装HTTP请求，统一API调用

### 工具模块 (`src/utils/`)

- **file_utils.py**: 文件操作工具（文件名清理、MIME类型等）
- **ui_utils.py**: UI工具（窗口居中、对话框等）

### 功能模块 (`src/modules/`)

各业务功能的具体实现，通过导入核心模块和工具模块复用代码。

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
