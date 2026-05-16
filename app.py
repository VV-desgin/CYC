import os
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["PYTHONUTF8"] = "1"

import streamlit as st
import pandas as pd
import time
from io import BytesIO
from openai import OpenAI

st.set_page_config(page_title="5G通信基建数智化交付系统", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")

DEMO_DATA = {
    "站点编号": ["SZ-BS-001", "SZ-BS-002", "GZ-BS-001", "BJ-BS-001", "SH-BS-001"],
    "站点名称": ["深圳大学城1号站", "深圳科技园2号站", "广州天河3号站", "北京海淀4号站", "上海浦东5号站"],
    "站点类型": ["宏站", "楼面站", "室分站", "微站", "宏站"],
    "线缆距离(米)": [200, 350, 180, 420, 280],
    "取电方式": ["市电直供", "交流配电箱", "弱电井取电", "路灯杆取电", "市电直供"],
    "光缆芯数(芯)": [4, 6, 4, 8, 4],
    "起点": ["1号机房", "2号机房", "3号机房", "4号机房", "5号机房"],
    "终点": ["2号机房", "3号机房", "4号机房", "5号机房", "6号机房"],
    "AAU型号": ["AAU5613", "AAU5636", "AAU5339", "AAU5639", "AAU5613"],
    "BBU型号": ["BBU5900", "BBU5900", "BBU3910", "BBU5900", "BBU5900"],
}
DEMO_DF = pd.DataFrame(DEMO_DATA)
REQUIRED_COLS = ["站点编号", "站点名称", "站点类型", "线缆距离(米)", "取电方式", "光缆芯数(芯)", "起点", "终点"]

PLATFORM_PRESETS = {
    "硅基流动": {"base_url": "https://api.siliconflow.cn/v1", "model": "deepseek-ai/DeepSeek-V3"},
    "DeepSeek官方": {"base_url": "https://api.deepseek.com", "model": "deepseek-chat"},
    "Groq": {"base_url": "https://api.groq.com/openai/v1", "model": "llama-3.3-70b-versatile"},
    "自定义/本地": {"base_url": "http://localhost:11434/v1", "model": "qwen2.5:7b"},
}

AI_STEPS = [
    {"key": "bom", "label": "📦 施工BOM", "system": "汇总全部站点生成一份施工物料清单。相同物料合并，数量累加。表格列：物料类别、物料名称、规格型号、数量、单位、适用站点。直接输出表格，不要开场白结尾语。"},
    {"key": "bor", "label": "📝 资源需求清单", "system": "汇总全部站点生成一份资源需求清单。相同工具合并，数量累加。表格列：工具名称、规格、数量、单位、适用站点、备注。直接输出表格，不要开场白结尾语。"},
    {"key": "bop", "label": "🔧 工序指导书", "system": "输出8道工序指导书。表格列：工序编号、工序名称、工艺要求、验收标准。直接输出表格，不要开场白结尾语。"},
    {"key": "fiber", "label": "🔌 纤芯分配表", "system": "汇总全部站点生成纤芯分配表。表格列：站点编号、光缆编号、纤芯序号、纤芯颜色、起始端子、终止端子、业务类型。直接输出表格，不要开场白结尾语。"},
    {"key": "risk", "label": "⚠️ 风险提示与注意事项", "system": "依据YD/T 5264-2021和GB/T 51431-2020，分析全部站点数据，输出：1.施工风险提示（至少5条，含严重程度） 2.注意事项（至少5条） 3.合规建议。直接输出，不要开场白结尾语。"},
]

for k, v in {
    "uploaded_files": [{"name": "样例数据", "df": DEMO_DF}],
    "current_idx": 0, "result_df": None, "pending_upload": None,
    "ai_data": {}, "fix_solutions": {}, "ai_running": False,
    "ai_generation_done": False, "ai_step_index": 0, "ai_start_time": 0,
    "ai_step_times": {}
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

def read_file(f):
    try: return pd.read_csv(f) if f.name.endswith(".csv") else pd.read_excel(f)
    except: return None

def get_client(url, key):
    keys = [k.strip() for k in key.split(",") if k.strip()]
    return OpenAI(base_url=url, api_key=keys[0] if keys else key)

# 侧边栏
with st.sidebar:
    st.title("⚙️ 配置与控制台")
    st.subheader("📁 上传基站设计元数据表")
    def on_upload():
        if st.session_state.file_uploader: st.session_state.pending_upload = st.session_state.file_uploader
    st.file_uploader("支持 CSV/Excel", type=["csv","xlsx","xls"], label_visibility="collapsed", key="file_uploader", on_change=on_upload)
    if st.session_state.pending_upload:
        f = st.session_state.pending_upload
        names = [x["name"] for x in st.session_state.uploaded_files]
        new_name = f"📄 {f.name}"
        if new_name not in names:
            df = read_file(f)
            if df is not None:
                st.session_state.uploaded_files.append({"name": new_name, "df": df})
                st.session_state.current_idx = len(st.session_state.uploaded_files) - 1
                st.session_state.ai_generation_done = False
        else: st.session_state.current_idx = names.index(new_name)
        st.session_state.pending_upload = None
        st.rerun()

    st.subheader("📂 已缓存文件")
    file_list = [x["name"] for x in st.session_state.uploaded_files]
    if st.session_state.current_idx >= len(file_list): st.session_state.current_idx = 0
    sel = st.selectbox("选择数据源", file_list, index=st.session_state.current_idx, label_visibility="collapsed")
    if sel in file_list: st.session_state.current_idx = file_list.index(sel)
    current_df = st.session_state.uploaded_files[st.session_state.current_idx]["df"]
    if "样例数据" not in sel and st.button("🗑️ 删除当前文件", use_container_width=True):
        del st.session_state.uploaded_files[file_list.index(sel)]
        st.session_state.current_idx = 0
        st.rerun()

    st.markdown("---")
    st.subheader("🤖 AI 配置")
    platform = st.selectbox("AI 平台", list(PLATFORM_PRESETS.keys()), index=0)
    preset = PLATFORM_PRESETS[platform]
    is_custom = (platform == "自定义/本地")
    base_url = st.text_input("API 地址", value=preset["base_url"], disabled=not is_custom)
    model = st.text_input("模型名称", value=preset["model"])
    api_key = st.text_input("API Key", type="password", placeholder="sk-xxxx")

    st.markdown("---")
    if not st.session_state.ai_running and st.button("🚀 启动AI数智化指令转化", type="primary", use_container_width=True):
        if not base_url: st.error("请填写 API 地址")
        else:
            existing = [c for c in REQUIRED_COLS if c in current_df.columns]
            missing = [c for c in REQUIRED_COLS if c not in current_df.columns]
            if missing: st.error(f"缺少字段: {', '.join(missing)}")
            else:
                st.session_state.result_df = current_df.copy()
                st.session_state.ai_running = True
                st.session_state.ai_generation_done = False
                st.session_state.ai_step_index = 0
                st.session_state.ai_start_time = time.time()
                st.session_state.ai_step_times = {}
                st.session_state.ai_data = {}
                st.session_state.fix_solutions = {}
                st.rerun()

    if st.session_state.ai_running:
        elapsed = time.time() - st.session_state.ai_start_time
        st.warning(f"⏳ 运行中 ({int(elapsed)}s)")
        si = st.session_state.ai_step_index
        if si < len(AI_STEPS): st.caption(f"当前: {AI_STEPS[si]['label']} ({si+1}/5)")
        for s in AI_STEPS[:si]:
            t = st.session_state.ai_step_times.get(s["key"], 0)
            st.caption(f"✅ {s['label']} ({t:.1f}s)" if t else f"✅ {s['label']}")
        for s in AI_STEPS[si:]: st.caption(f"⏳ {s['label']}")
        if st.button("⏹️ 取消", use_container_width=True):
            st.session_state.ai_running = False
            st.session_state.ai_generation_done = True
            st.rerun()

    if st.session_state.ai_generation_done and not st.session_state.ai_running:
        total = st.session_state.ai_step_times.get("total", 0)
        st.success(f"✅ 总耗时 {total:.1f}s")
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as w:
            for s in AI_STEPS:
                c = st.session_state.ai_data.get(s["key"], "暂无数据")
                pd.DataFrame({"内容": c.split("\n")}).to_excel(w, index=False, sheet_name=s["key"][:31])
        st.download_button("💾 下载 Excel", output.getvalue(), "5G基站AI交付结果.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# 主页面
st.title("🏗️ 5G通信基建数智化交付系统（Demo0.1）")
st.caption("上传设计元数据表 → AI一键生成施工BOM、资源清单、工序指导书、纤芯分配表及风险提示")
st.markdown("---")
tab1, tab2, tab3 = st.tabs(["📊 生成", "📋 资料", "🔍 审查结果"])

with tab1:
    st.subheader("📊 原始数据")
    st.dataframe(current_df, use_container_width=True, hide_index=True)
    st.markdown("---")
    st.subheader("🔍 校验")
    existing = [c for c in REQUIRED_COLS if c in current_df.columns]
    missing = [c for c in REQUIRED_COLS if c not in current_df.columns]
    if not missing: st.success(f"✅ 通过 | {len(current_df)}站点 | {len(current_df.columns)}字段")
    else: st.error(f"❌ 缺少: {', '.join(missing)}")
    st.markdown("---")

    if st.session_state.ai_running:
        sites = st.session_state.result_df
        si = st.session_state.ai_step_index
        if si < len(AI_STEPS):
            step = AI_STEPS[si]
            st.subheader(f"⏳ {step['label']} ({si+1}/5)")
            ph = st.empty()
            try:
                client = get_client(base_url, api_key)
                t0 = time.time()
                rv = (step["key"] == "risk")
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role":"system","content":step["system"]}, {"role":"user","content":f"站点数据（{len(sites)}个）：\n{sites.to_string()}"}],
                    timeout=90 if rv else 180, temperature=0.3, max_tokens=2048 if rv else 4096
                )
                content = resp.choices[0].message.content
                elapsed = time.time() - t0
                st.session_state.ai_data[step["key"]] = content
                st.session_state.ai_step_times[step["key"]] = elapsed
                ph.success(f"✅ {step['label']} ({elapsed:.1f}s)")
            except Exception as e:
                st.session_state.ai_data[step["key"]] = f"❌ {str(e)[:200]}"
                ph.error(str(e)[:200])
            st.session_state.ai_step_index = si + 1
            time.sleep(0.3)
            st.rerun()
        else:
            st.session_state.ai_step_times["total"] = time.time() - st.session_state.ai_start_time
            st.session_state.ai_running = False
            st.session_state.ai_generation_done = True
            st.rerun()

    if st.session_state.ai_generation_done and not st.session_state.ai_running:
        st.success(f"🎉 完成！总耗时 {st.session_state.ai_step_times.get('total',0):.1f}s")

with tab2:
    st.subheader("📋 AI生成施工资料")
    if not st.session_state.ai_generation_done:
        st.info("👈 请在左侧点击启动按钮")
    else:
        all_sites = list(st.session_state.result_df["站点编号"]) if st.session_state.result_df is not None and "站点编号" in st.session_state.result_df.columns else []
        view_mode = st.radio("查看模式", ["全部站点汇总", "单站点详细"], horizontal=True)
        selected_site = st.selectbox("选择站点", all_sites) if view_mode == "单站点详细" and all_sites else None

        tabs = st.tabs(["📦 施工BOM", "📝 资源清单", "🔧 工序指导书", "🔌 纤芯分配表", "⚠️ 风险提示"])
        keys = ["bom", "bor", "bop", "fiber", "risk"]
        for i, t in enumerate(tabs):
            with t:
                content = st.session_state.ai_data.get(keys[i], "暂无数据")
                if selected_site and content and not content.startswith("❌"):
                    lines = content.split("\n")
                    filtered = [l for l in lines if selected_site in l or not any(s in l for s in all_sites if s != selected_site)]
                    st.markdown("\n".join(filtered) if filtered else content)
                else:
                    st.markdown(content)

with tab3:
    st.subheader("🔍 审查结果")
    st.info("审查结果模块预留，暂未接入AI生成。")