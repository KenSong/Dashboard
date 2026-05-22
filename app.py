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

# 日期范围选择
date_list = sorted([d for d in df_all["日期"].unique() if d])
if date_list:
    # 将日期字符串转换为 datetime 对象以便选择
    date_objects = pd.to_datetime(date_list, errors="coerce")
    valid_dates = [(d_str, d_obj) for d_str, d_obj in zip(date_list, date_objects) if pd.notna(d_obj)]
    valid_dates.sort(key=lambda x: x[1])
    
    if valid_dates:
        min_date = valid_dates[0][1].date()
        max_date = valid_dates[-1][1].date()
        
        # 使用日期范围选择器
        date_range = st.sidebar.date_input(
            "日期范围",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            format="YYYY-MM-DD"
        )
        
        # 处理日期范围选择
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start_date, end_date = date_range
            # 筛选在日期范围内的数据
            selected_dates = [
                d_str for d_str, d_obj in valid_dates
                if start_date <= d_obj.date() <= end_date
            ]
        else:
            selected_dates = date_list
    else:
        selected_dates = date_list
else:
    selected_dates = []

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
    """HTML 表格：按天为单位合并单元格（当天相同值才合并），成交金额不合并。"""
    if df.empty:
        return "<p>无明细数据</p>"
    view = df[columns].reset_index(drop=True)
    n = len(view)
    
    # 需要合并的列（按优先级顺序）
    mergeable_cols = ["日期", "部门", "平台", "子平台"]
    no_merge_col = "成交金额(万元)"
    
    # 每列的合并信息
    rowspan = {col: [0] * n for col in columns}
    
    # 按天合并：同一天内，相同值才合并
    i = 0
    while i < n:
        # 获取当前行的日期，作为合并的边界
        current_date = view.iloc[i]["日期"]
        
        # 同一天内的结束位置
        day_end = i + 1
        while day_end < n and view.iloc[day_end]["日期"] == current_date:
            day_end += 1
        
        # 在同一天内处理合并
        for col in mergeable_cols:
            j = i
            while j < day_end:
                col_span = 1
                v = view.iloc[j][col]
                k = j + 1
                # 只在同一天内检查相同值
                while k < day_end and view.iloc[k][col] == v:
                    col_span += 1
                    k += 1
                # 设置该列的合并跨度
                rowspan[col][j] = col_span
                for m in range(j + 1, j + col_span):
                    rowspan[col][m] = 0
                j = k
        
        # 成交金额列不合并，每行独立显示
        for row_idx in range(i, day_end):
            rowspan[no_merge_col][row_idx] = 1
        
        # 移动到下一天
        i = day_end

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

# 按「日期」汇总成交金额趋势（与当前侧边栏筛选一致）
_trend = df_filtered.assign(
    _日期解析=pd.to_datetime(
        df_filtered["日期"].astype(str).str.strip(),
        errors="coerce",
    )
).dropna(subset=["_日期解析"])
if not _trend.empty:
    # 按日期汇总总金额（按天汇总）
    trend_sum = _trend.copy()
    trend_sum["_日期"] = trend_sum["_日期解析"].dt.date
    trend_sum = trend_sum.groupby("_日期", as_index=False)["成交金额"].sum()
    
    # 生成连续的日期序列
    min_date = trend_sum["_日期"].min()
    max_date = trend_sum["_日期"].max()
    all_dates = pd.date_range(start=min_date, end=max_date, freq='D')
    all_dates_df = pd.DataFrame({"_日期": all_dates.date})
    
    # 合并数据，缺失日期的成交金额设为0
    trend_sum_full = pd.merge(all_dates_df, trend_sum, on="_日期", how="left").fillna(0)
    
    st.subheader("📊 成交金额按日期趋势")
    fig_trend = px.bar(
        trend_sum_full,
        x="_日期",
        y="成交金额",
        title="",
    )
    fig_trend.update_layout(
        xaxis_title="日期",
        yaxis_title="成交金额（元）",
        height=400,
        margin=dict(l=10, r=10, t=10, b=10),
        showlegend=False,
    )
    fig_trend.update_xaxes(tickformat="%m-%d", hoverformat="%Y-%m-%d", tickmode='linear')
    fig_trend.update_traces(
        texttemplate="%{y:,.2f} 元",
        textposition="outside",
        hovertemplate="%{x|%Y-%m-%d}<br>成交金额：%{y:,.2f} 元<extra></extra>"
    )
    st.plotly_chart(fig_trend, use_container_width=True)
else:
    st.info("当前筛选下无有效「日期」数据，无法绘制成交金额趋势图。")

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
    # 计算日均，需要检查日期变量是否已定义
    if 'start_date' in locals() and 'end_date' in locals():
        date_days = (end_date - start_date).days + 1
        pivot = pivot.assign(日均=pivot["合计"] / date_days)
    else:
        # 如果日期范围未定义，使用数据中的日期数量计算
        date_count = len(df_filtered["日期"].unique())
        if date_count > 0:
            pivot = pivot.assign(日均=pivot["合计"] / date_count)
st.dataframe(pivot.style.format("{:,.2f}"), use_container_width=True)

st.markdown("---")
st.subheader("📋 明细数据")

# 获取所有不重复的日期并转换为datetime对象
all_dates = [d for d in df_filtered["日期"].unique() if d]
if all_dates:
    # 转换为datetime对象以便使用日历选择器
    date_objects = pd.to_datetime(all_dates, errors="coerce")
    valid_date_tuples = [(d_str, d_obj) for d_str, d_obj in zip(all_dates, date_objects) if pd.notna(d_obj)]
    valid_date_tuples.sort(key=lambda x: x[1])
    
    if valid_date_tuples:
        # 获取最早和最晚日期
        min_date = valid_date_tuples[0][1].date()
        max_date = valid_date_tuples[-1][1].date()
        
        # 使用日历选择器，默认选中最新日期
        selected_date_obj = st.date_input(
            "选择日期",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="detail_date_selector"
        )
        
        # 将筛选数据和选择的日期都转换为datetime对象进行比较（兼容各种日期格式）
        df_with_date = df_filtered.copy()
        df_with_date["_日期解析"] = pd.to_datetime(df_filtered["日期"], errors="coerce")
        selected_date_dt = pd.to_datetime(selected_date_obj)
        
        # 根据选择的日期筛选数据
        display_df = df_with_date[df_with_date["_日期解析"].dt.date == selected_date_obj].copy()
        # 删除临时列
        display_df = display_df.drop("_日期解析", axis=1)
    else:
        display_df = df_filtered.copy()
else:
    display_df = df_filtered.copy()

display_df["成交金额(万元)"] = (display_df["成交金额"] / 10000.0).round(2)
detail_cols = ["日期", "部门", "平台", "子平台", "成交金额(万元)"]
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
