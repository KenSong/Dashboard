import html
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
import plotly.express as px
import streamlit as st

# Streamlit 要求尽早调用 set_page_config
st.set_page_config(
    page_title="数据看板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


def _result_csv_path() -> Path:
    """与 app.py 同目录下的 result.csv（相对脚本所在文件夹，不依赖当前工作目录）。"""
    return Path(__file__).resolve().parent / "result.csv"


RESULT_CSV = _result_csv_path()

# --------------------------
# 侧边栏：刷新数据
# --------------------------
st.sidebar.title("⚙️ 控制面板")
if st.sidebar.button("🔄 刷新最新数据！"):
    st.cache_data.clear()
    st.success("✅ 缓存已清除，已重新读取 result.csv！")


@st.cache_data(show_spinner=True)
def load_result_csv(path_str: str, _mtime: float) -> pd.DataFrame:
    """_mtime 参与缓存键：result.csv 变更后自动重新读取。"""
    path = Path(path_str)
    if not path.is_file():
        return pd.DataFrame()
    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    if "日期" not in df.columns:
        df["日期"] = ""
    if "子平台" not in df.columns:
        df["子平台"] = ""
    df["日期"] = df["日期"].fillna("").astype(str).str.strip()
    df["子平台"] = df["子平台"].fillna("").astype(str).str.strip()
    df["部门"] = df["部门"].fillna("").astype(str).str.strip()
    df["平台"] = df["平台"].fillna("").astype(str).str.strip()
    df["成交金额"] = pd.to_numeric(df["成交金额"], errors="coerce").fillna(0.0)
    return df


def max_business_date_label(df: pd.DataFrame) -> str:
    """取「日期」列中可解析的最大日期，格式 YYYY-MM-DD；无有效值时返回 —。"""
    if df.empty or "日期" not in df.columns:
        return "—"
    s = df["日期"].astype(str).str.strip()
    s = s[s.ne("") & s.str.lower().ne("nan")]
    if s.empty:
        return "—"
    parsed = pd.to_datetime(s, errors="coerce")
    tmax = parsed.max()
    if pd.isna(tmax):
        return "—"
    return tmax.strftime("%Y-%m-%d")


try:
    _result_mtime = RESULT_CSV.stat().st_mtime
except OSError:
    _result_mtime = 0.0

df_all = load_result_csv(str(RESULT_CSV), _result_mtime)

if df_all.empty:
    st.error(f"❌ 未读取到数据。尝试路径：`{RESULT_CSV}`（文件不存在或无法解析）。")
    st.markdown(
        "请将 `result.csv` 与 `app.py` 放在同一目录下（本地或 **Streamlit Community Cloud** 仓库中路径一致即可）。"
    )
    st.stop()

# --------------------------
# 侧边栏筛选
# --------------------------
st.sidebar.header("🔍 筛选条件")

date_list = sorted([d for d in df_all["日期"].unique() if d])
date_options = ["全部"] + date_list
if "selected_date_radio" not in st.session_state:
    st.session_state["selected_date_radio"] = "全部"
selected_date_radio = st.sidebar.radio("日期", date_options, key="selected_date_radio")
selected_dates = date_list if selected_date_radio == "全部" else [selected_date_radio]

dept_list = sorted([d for d in df_all["部门"].unique() if d])
dept_options = ["全部"] + dept_list
if "selected_dept_radio" not in st.session_state:
    st.session_state["selected_dept_radio"] = "全部"
selected_dept_radio = st.sidebar.radio("部门", dept_options, key="selected_dept_radio")
selected_dept = dept_list if selected_dept_radio == "全部" else [selected_dept_radio]

plat_list = sorted([p for p in df_all["平台"].unique() if p])
plat_options = ["全部"] + plat_list
if "selected_plat_radio" not in st.session_state:
    st.session_state["selected_plat_radio"] = "全部"
selected_plat_radio = st.sidebar.radio("平台", plat_options, key="selected_plat_radio")
selected_plat = plat_list if selected_plat_radio == "全部" else [selected_plat_radio]


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if selected_dates:
        out = out[out["日期"].isin(selected_dates)]
    if selected_dept:
        out = out[out["部门"].isin(selected_dept)]
    if selected_plat:
        out = out[out["平台"].isin(selected_plat)]
    return out


def render_detail_table_vertical_merge(df: pd.DataFrame, columns: list) -> str:
    """HTML 表格：各列内连续相同取值合并为 rowspan（纵向合并）。"""
    if df.empty:
        return "<p>无明细数据</p>"
    view = df[columns].reset_index(drop=True)
    n = len(view)
    rowspan = {col: [0] * n for col in columns}
    for col in columns:
        i = 0
        while i < n:
            j = i + 1
            v = view.iloc[i][col]
            while j < n and view.iloc[j][col] == v:
                j += 1
            rowspan[col][i] = j - i
            for k in range(i + 1, j):
                rowspan[col][k] = 0
            i = j

    body_rows = []
    for i in range(n):
        cells = []
        for col in columns:
            rs = rowspan[col][i]
            if rs <= 0:
                continue
            val = view.iloc[i][col]
            if pd.isna(val):
                text = ""
            elif col == "成交金额(万元)":
                text = f"{float(val):,.2f}"
            else:
                text = "" if val == "" else str(val)
            cells.append(f'<td rowspan="{rs}">{html.escape(text)}</td>')
        body_rows.append("<tr>" + "".join(cells) + "</tr>")

    header = "<tr>" + "".join(f"<th>{html.escape(str(c))}</th>" for c in columns) + "</tr>"
    style = """
    <style>
        .detail-merge-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }
        .detail-merge-table th, .detail-merge-table td {
            border: 1px solid #ddd;
            padding: 10px 12px;
            text-align: center;
            vertical-align: middle;
        }
        .detail-merge-table th {
            background-color: #f0f2f6;
            font-weight: bold;
        }
        .detail-merge-table td {
            background-color: #fff;
        }
    </style>
    """
    return (
        style
        + '<table class="detail-merge-table"><thead>'
        + header
        + "</thead><tbody>"
        + "".join(body_rows)
        + "</tbody></table>"
    )


df_filtered = apply_filters(df_all)

# --------------------------
# 主界面
# --------------------------
last_date_display = max_business_date_label(df_all)

st.markdown(
    f"<h1>📊 数据看板 <span style='font-size: 16px; color: #666; font-weight: normal;'>"
    f"数据源: result.csv · 数据最后日期: {last_date_display}</span></h1>",
    unsafe_allow_html=True,
)
st.markdown("---")

total_amount = df_filtered["成交金额"].sum() / 10000.0

metric_html = """
<style>
    .custom-metric {
        background: white;
        border-radius: 8px;
        padding: 16px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1);
    }
    .metric-label { font-size: 14px; color: #666; margin-bottom: 8px; }
    .metric-value { font-size: 24px; font-weight: bold; color: #333; }
</style>
"""
st.markdown(metric_html, unsafe_allow_html=True)

st.markdown(
    f"<div class='custom-metric'><div class='metric-label'>💰 成交金额合计（万元）</div>"
    f"<div class='metric-value'>{total_amount:,.2f}</div></div>",
    unsafe_allow_html=True,
)

st.markdown("---")

# 汇总透视：部门 × 平台
st.subheader("📋 部门 × 平台 汇总（万元）")
pivot = (
    df_filtered.pivot_table(
        index="部门",
        columns="平台",
        values="成交金额",
        aggfunc="sum",
        fill_value=0.0,
    )
    / 10000.0
)
if not pivot.empty:
    pivot = pivot.assign(合计=pivot.sum(axis=1))
st.dataframe(pivot.style.format("{:,.2f}"), use_container_width=True)

st.markdown("---")
st.subheader("📋 明细数据")
display_df = df_filtered.copy()
display_df["成交金额(万元)"] = (display_df["成交金额"] / 10000.0).round(2)
detail_cols = ["部门", "平台", "子平台", "成交金额(万元)"]
st.markdown(
    render_detail_table_vertical_merge(display_df, detail_cols),
    unsafe_allow_html=True,
)

st.markdown("---")
st.subheader("📊 成交金额分布（按部门）")
pie_data = df_filtered.groupby("部门", as_index=False)["成交金额"].sum()
if not pie_data.empty and pie_data["成交金额"].sum() > 0:
    fig_pie = px.pie(
        pie_data,
        values="成交金额",
        names="部门",
        title="各部门成交金额占比",
        hole=0.3,
        color_discrete_map={
            "常温": "#1f77b4",
            "低温": "#ff7f0e",
            "奶粉": "#2ca02c",
            "八喜": "#9467bd",
        },
    )
    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
    st.plotly_chart(fig_pie, use_container_width=True)
else:
    st.info("当前筛选下无成交金额数据。")

st.markdown("---")
st.subheader("📊 成交金额（按部门堆叠 · 平台）")
bar_src = (
    df_filtered.groupby(["部门", "平台"], as_index=False)["成交金额"]
    .sum()
    .assign(成交金额万元=lambda d: d["成交金额"] / 10000.0)
)
if not bar_src.empty:
    fig_bar = px.bar(
        bar_src,
        x="部门",
        y="成交金额万元",
        color="平台",
        barmode="stack",
        title="各部门成交金额（万元），按平台堆叠",
    )
    fig_bar.update_layout(yaxis_title="成交金额（万元）", xaxis_title="部门", height=420)
    fig_bar.update_traces(texttemplate="%{y:.1f}", textposition="inside")
    st.plotly_chart(fig_bar, use_container_width=True)
else:
    st.info("当前筛选下无数据。")
