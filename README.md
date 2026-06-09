# 5G通信基建数智化交付系统

RK-001~RK-010 合规审查 · 12 项标准 BOM · 16 项站点字段 · Word/Excel 双格式交付

## 功能模块

- **合规审查**：基于标准 BOM 的自动化合规校验，输出审查报告
- **站点管理**：16 项站点字段的增删改查与批量导入
- **BOM 管理**：12 项标准物料清单的维护与版本控制
- **文档交付**：支持 Word / Excel 双格式自动生成与导出
- **AI 引擎**：内置标准 AI 引擎（开箱即用），支持接入外部大模型（OpenAI 兼容 API）

## 技术栈

| 组件 | 说明 |
|------|------|
| 框架 | Streamlit |
| 数据处理 | pandas, openpyxl |
| 文档生成 | python-docx |
| AI 接口 | openai SDK（兼容任意 OpenAI 格式 API） |
| 数据存储 | JSON（task_logs） |

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动应用
streamlit run app_v3.py
```

启动后浏览器访问 `http://localhost:8501`。

> 默认使用内置 AI 引擎，无需配置 API Key 即可使用合规审查功能。如需启用全量 AI 生成，在界面中配置 API Key 并测试连接。

## 项目结构

```
SZGC/
├── app_v3.py          # 主入口，Streamlit 应用
├── ai_client.py       # AI 客户端封装
├── requirements.txt   # Python 依赖
├── task_logs.json     # 任务日志数据
└── .gitignore
```

## 环境要求

- Python 3.8+
- Windows / macOS / Linux