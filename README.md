# 5G通信基建数智化交付系统 (app_v3)

基于 Streamlit 的 5G 通信基站工程交付全流程管理系统。

## 功能模块

- **工程概览**：站点信息仪表盘，项目状态与进度可视化
- **AI生成施工资料**：智能生成施工BOM、纤芯分配表、交付报告
- **合规审查**：RK-001~RK-010 十项合规检查规则，自动问题整改台账
- **任务中心**：施工任务分配、跟踪与闭环管理
- **交付输出**：Excel 交付包 + Word 交付报告，双格式一键导出
- **AI 助手**：内置 AI 模式 + 外部大模型 API 双模式，支持工程问答与辅助分析

## 技术栈

- **前端**：Streamlit (Wide Layout)
- **数据处理**：pandas
- **Excel 生成**：openpyxl（含样式、边框、条件格式）
- **Word 生成**：python-docx
- **AI 接口**：OpenAI SDK（兼容第三方 API）

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动应用
streamlit run main.py
```

内置 AI 模式无需 API Key 即可运行。外部大模型模式需在侧边栏配置 API Key。

## 项目结构

```
5g-delivery-system/
├── main.py            # 主程序入口 (app_v3)
├── requirements.txt   # Python 依赖
├── .gitignore         # Git 忽略规则
└── README.md          # 项目说明
```

## License

MIT
