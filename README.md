# 政务协助系统 / Government Affairs Assistance System

[![Python](https://img.shields.io/badge/Python-3.x-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

## 📋 项目简介

本项目是一个**政务协助系统**，主要用于**矛盾纠纷数据的处理、导出和管理**。系统提供了多种实用工具，帮助政务工作人员高效地处理综治信息平台中的纠纷数据，包括数据导出、研判记录上传、回访部门管理等功能。

## 🎯 主要功能

### 1. 综合数据处理器 (integrated_processor.py)
- 集成化的数据处理中心
- 支持从综治信息平台API获取数据
- Excel和Word文档生成
- 自动化的数据整理和归档

### 2. 纠纷详情导出 (export_issue_detail.py)
- 导出纠纷详情数据
- 支持自定义保存路径
- Token身份验证管理
- 图形化操作界面

### 3. 研判记录上传 (update_progress.py / update_file.py)
- 批量上传研判记录
- 进度条显示
- 支持一级、二级研判记录表
- 自动化的数据处理

### 4. 回访部门管理 (add_return_visit_department.py)
- 管理回访部门信息
- 支持多个社区/部门
- 自动分配和记录

### 5. 数据合并工具 (merged.py)
- 合并多个数据源
- 数据去重和整理
- 统一格式输出

### 6. Web数据台账 (矛盾纠纷数据台账.html)
- 基于Web的数据展示界面
- 使用Bootstrap美化界面
- Chart.js数据可视化
- 支持Excel导出功能

## 🛠️ 技术栈

- **Python 3.x**: 主要开发语言
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
├── README.md                           # 项目说明文档
├── .gitignore                          # Git忽略配置
├── config.json                         # 配置文件
├── dispute_data.json                   # 纠纷数据存储
│
├── integrated_processor.py             # 综合数据处理器
├── export_issue_detail.py              # 纠纷详情导出
├── add_return_visit_department.py      # 回访部门管理
├── update_progress.py                  # 上传研判记录
├── update_file.py                      # 文件更新工具
├── merged.py                           # 数据合并工具
├── 1.py, 2.py                          # 辅助脚本
│
├── js/                                 # Web前端文件
│   ├── index.html                      # 主页面
│   ├── getResult.js                    # 数据获取逻辑
│   └── result.json                     # 结果数据
│
├── app.ico                             # 应用程序图标
├── update_progress.ico                 # 进度更新图标
│
├── 一级研判记录表.xlsx                  # 一级研判记录模板
├── 二级研判记录表.xlsx                  # 二级研判记录模板
└── 矛盾纠纷数据台账.html                # Web数据台账页面
```

## 🚀 快速开始

### 环境要求

- Python 3.8+
- Windows操作系统（部分功能依赖Windows路径）

### 安装依赖

```bash
pip install pandas openpyxl python-docx requests
```

### 运行程序

1. **综合数据处理器**
   ```bash
   python integrated_processor.py
   ```

2. **导出纠纷详情**
   ```bash
   python export_issue_detail.py
   ```

3. **上传研判记录**
   ```bash
   python update_progress.py
   ```

4. **Web数据台账**
   直接在浏览器中打开 `矛盾纠纷数据台账.html`

## ⚙️ 配置说明

项目使用 `config.json` 进行配置，主要包括：

- `token`: API访问令牌
- `save_path`: 默认保存路径
- 其他自定义配置项

## 🔒 安全说明

- 系统使用Token进行身份验证
- 支持HTTPS请求（已禁用证书验证警告）
- 请勿将敏感信息（如token）提交到代码仓库

## 📝 使用说明

### 导出纠纷详情
1. 运行 `export_issue_detail.py`
2. 输入有效的Token
3. 选择保存路径
4. 点击导出按钮

### 上传研判记录
1. 运行 `update_progress.py`
2. 选择Excel文件
3. 系统将自动处理并上传数据
4. 查看进度条了解处理进度

### 数据台账查看
1. 打开 `矛盾纠纷数据台账.html`
2. 系统自动加载数据
3. 支持图表展示和数据导出

## 🤝 贡献指南

欢迎提交Issue和Pull Request来改进项目。

## 📄 许可证

本项目采用 MIT 许可证 - 详见 LICENSE 文件

## 👨‍💻 开发者

- **维护者**: XiaoDus
- **项目地址**: https://github.com/XiaoDus/Government-affairs-assistance

## 📞 联系方式

如有问题或建议，请通过GitHub Issue联系。

---

**注意**: 本项目仅供学习和参考使用，请遵守相关法律法规。