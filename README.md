# 🏗️ 5G通信基建数智化交付系统（Demo0.1）


## ✨ 功能

上传基站设计元数据表（CSV/Excel），AI 一键生成：

- 📦 **施工BOM清单** — 物料汇总，相同物料自动合并，数量累加
- 📝 **资源需求清单** — 工具、人员、工时配置
- 🔧 **工序指导书** — 8道标准工序，含工艺要求与验收标准
- 🔌 **纤芯分配表** — 端口对应、纤芯颜色标识
- ⚠️ **风险提示与注意事项** — 施工风险、合规建议

支持查看全部站点汇总或单站点详细，一键下载 Excel。

## 🚀 快速开始

```bash
git clone https://github.com/你的用户名/5g-base-station-delivery.git
cd 5g-base-station-delivery
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
