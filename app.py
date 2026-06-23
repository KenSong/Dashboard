import html
import urllib.parse
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
import plotly.express as px
import streamlit as st

# --------------------------
# 报告页面单独处理
# --------------------------
report_param = st.query_params.get("report_path")
if report_param:
    decoded_path = urllib.parse.unquote(report_param)
    if Path(decoded_path).exists():
        with open(decoded_path, "r", encoding="utf-8") as f:
            report_content = f.read()
        st.set_page_config(page_title="618报告", page_icon="📊", layout="wide")
        st.markdown(report_content, unsafe_allow_html=True)
    else:
        st.error("报告文件不存在")
    st.stop()

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
    # 确保目标金额列存在
    if "目标金额" not in df.columns:
        df["目标金额"] = 0.0
    df["目标金额"] = pd.to_numeric(df["目标金额"], errors="coerce").fillna(0.0)
    
    # 按日期排序，确保同一日期的数据显示在一起
    df["日期_sort"] = pd.to_datetime(df["日期"], errors="coerce")
    
    # 自定义平台排序：让小程序及其他在同一部门内排在最后
    platform_order = {"京东": 1, "天猫": 2, "拼多多": 3, "抖音": 4, "新零售": 5, "多多买菜": 6, "小程序及其他": 99}
    df["平台_sort"] = df["平台"].map(platform_order).fillna(100)
    
    df = df.sort_values(by=["日期_sort", "部门", "平台_sort"], na_position="last").drop(["日期_sort", "平台_sort"], axis=1)
    
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
        
        # 618日期范围
        start_date_618 = pd.to_datetime("2026-05-13").date()
        end_date_618 = pd.to_datetime("2026-06-20").date()
        
        # 添加618数据和618报告链接（并排显示）
        col_btn1, col_btn2 = st.sidebar.columns(2)
        with col_btn1:
            st.markdown(f'<a href="/?show_618_data=true" target="_self">📊 618数据</a>', unsafe_allow_html=True)
        with col_btn2:
            report_url = str(Path(__file__).resolve().parent / "618完成百分比报告.md")
            encoded_url = urllib.parse.quote(report_url)
            st.markdown(f'<a href="/?report_path={encoded_url}" target="_blank">📄 618报告</a>', unsafe_allow_html=True)

        # 处理618数据点击
        if st.query_params.get("show_618_data") == "true":
            st.session_state["date_range_selector"] = (start_date_618, end_date_618)
            st.session_state["is_618_mode"] = True
            st.session_state["user_manually_selected_date"] = True
            # 移除URL参数以避免重复执行
            st.query_params.clear()
            st.rerun()

        # 显示618报告内容（通过URL参数传递）
        report_param = st.query_params.get("report_path")
        if report_param:
            decoded_path = urllib.parse.unquote(report_param)
            if Path(decoded_path).exists():
                with open(decoded_path, "r", encoding="utf-8") as f:
                    report_content = f.read()
                st.markdown(report_content, unsafe_allow_html=True)
        
        # 初始化用户是否手动选择日期范围的标志
        if "user_manually_selected_date" not in st.session_state:
            st.session_state["user_manually_selected_date"] = False
        
        # 计算默认日期范围：最近10天（每次都重新计算，确保数据更新后能显示最新日期）
        if len(valid_dates) > 10:
            default_start_date = valid_dates[-10][1].date()
            default_end_date = valid_dates[-1][1].date()
        else:
            default_start_date = min_date
            default_end_date = max_date
        
        # 初始化日期范围选择器的默认值（只在第一次加载时设置）
        if "date_range_selector" not in st.session_state:
            st.session_state["date_range_selector"] = (default_start_date, default_end_date)
        
        # 使用日期范围选择器，使用key来管理状态
        date_range = st.sidebar.date_input(
            "日期范围",
            value=st.session_state["date_range_selector"],
            min_value=min_date,
            max_value=max_date,
            format="YYYY-MM-DD",
            key="date_range_selector"
        )
        
        # 保存原始的日期范围值
        original_date_range = date_range
        
        # 检测日期范围是否有效
        if not isinstance(date_range, tuple) or len(date_range) != 2:
            # 用户可能正在选择日期，暂时使用默认日期范围，不显示警告
            date_range = (default_start_date, default_end_date)
        
        # 检测用户是否手动修改了日期范围（使用原始值进行比较）
        if isinstance(original_date_range, tuple) and len(original_date_range) == 2:
            if original_date_range != (default_start_date, default_end_date):
                st.session_state["user_manually_selected_date"] = True
            else:
                st.session_state["user_manually_selected_date"] = False
        
        # 检查是否在618模式
        is_618_mode = st.session_state.get("is_618_mode", False)
        # 如果日期范围不是618日期范围，退出618模式
        if date_range != (start_date_618, end_date_618):
            st.session_state["is_618_mode"] = False
            is_618_mode = False
        
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

plat_list = [p for p in df_all["平台"].unique() if p]

# 自定义平台排序：将"小程序及其他"放在"新零售"之后
def plat_sort_key(plat):
    # 定义期望的顺序
    order = ["京东", "天猫", "拼多多", "抖音", "新零售", "小程序及其他"]
    if plat in order:
        return order.index(plat)
    return len(order) + ord(plat[0])

plat_list = sorted(plat_list, key=plat_sort_key)
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
    f"<div style='display: flex; justify-content: space-between; align-items: flex-start;'>"
    f"<h1>📊 数据看板</h1>"
    f"<div style='font-size: 13px; color: #666; text-align: right;'>"
    f"<strong>数据源：</strong><br>"
    f"<em>京东</em>，<em>天猫</em>，<em>拼多多</em>，<em>抖音</em> 来自平台<br>"
    f"<em>新零售</em> 来自ERP里销售金额（低温*1.2，常温不变）<br>"
    f"<em>多多买菜</em>和<em>小程序及其他</em> 来自常温<br>"
    f"<strong>数据最后日期: {last_date_display}</strong>"
    f"</div>"
    f"</div>",
    unsafe_allow_html=True,
)
st.markdown("---")

# 计算成交金额合计
is_618_mode = st.session_state.get("is_618_mode", False)
user_manually_selected = st.session_state.get("user_manually_selected_date", False)

if is_618_mode:
    # 618模式，计算5/13-6/20范围内的成交金额
    total_amount = df_filtered["成交金额"].sum()
elif user_manually_selected:
    # 非618模式，用户手动选择了日期范围，计算该范围内的成交金额
    total_amount = df_filtered["成交金额"].sum()
else:
    # 非618模式，用户没有手动选择日期范围，计算最近10天的成交金额
    df_dates = df_filtered.assign(
        _日期解析=pd.to_datetime(df_filtered["日期"].astype(str).str.strip(), errors="coerce")
    ).dropna(subset=["_日期解析"])
    
    if not df_dates.empty:
        # 获取最近10天的日期
        recent_dates = df_dates["_日期解析"].dt.date.unique()
        if len(recent_dates) > 10:
            recent_dates = sorted(recent_dates)[-10:]
        # 筛选最近10天的数据
        df_recent = df_dates[df_dates["_日期解析"].dt.date.isin(recent_dates)]
        total_amount = df_recent["成交金额"].sum()
    else:
        total_amount = df_filtered["成交金额"].sum()

total_goal = df_filtered["目标金额"].sum()
achievement_rate = (total_amount / total_goal * 100) if total_goal > 0 else 0

metric_html = """
<style>
    .custom-metric-row {
        display: flex;
        gap: 16px;
    }
    .custom-metric {
        background: white;
        border-radius: 8px;
        padding: 16px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1);
        flex: 1;
    }
    .metric-label { font-size: 14px; color: #666; margin-bottom: 8px; }
    .metric-value { font-size: 24px; font-weight: bold; color: #333; }
</style>
"""
st.markdown(metric_html, unsafe_allow_html=True)

# 第一行指标
st.markdown(
    f"<div class='custom-metric-row'>"
    f"<div class='custom-metric'><div class='metric-label'>💰 成交金额合计（万元）</div>"
    f"<div class='metric-value'>{total_amount:,.2f}</div></div>"
    f"</div>",
    unsafe_allow_html=True,
)

# 按「日期」汇总成交金额和目标金额趋势（与当前侧边栏筛选一致）
_trend = df_filtered.assign(
    _日期解析=pd.to_datetime(
        df_filtered["日期"].astype(str).str.strip(),
        errors="coerce",
    )
).dropna(subset=["_日期解析"])
if not _trend.empty:
    # 按日期汇总成交金额和目标金额（按天汇总）
    trend_sum = _trend.copy()
    trend_sum["_日期"] = trend_sum["_日期解析"].dt.date
    trend_sum = trend_sum.groupby("_日期", as_index=False).agg(
        成交金额=("成交金额", "sum"),
        目标金额=("目标金额", "sum")
    )
    
    # 只显示有数据的日期，不生成连续日期序列
    trend_sum_full = trend_sum.sort_values("_日期")
    
    # 检查是否在618模式
    is_618_mode = st.session_state.get("is_618_mode", False)
    # 检查用户是否手动选择了日期范围
    user_manually_selected = st.session_state.get("user_manually_selected_date", False)
    
    # 非618模式下，如果用户没有手动选择日期范围，默认只显示最近10天的数据
    if not is_618_mode and not user_manually_selected and len(trend_sum_full) > 10:
        trend_sum_full = trend_sum_full.tail(10)
    
    if is_618_mode:
        st.subheader("📊 成交金额与目标金额趋势")
    else:
        st.subheader("📊 成交金额趋势")
    
    fig_trend = px.bar(
        trend_sum_full,
        x="_日期",
        y="成交金额",
        title="",
        labels={'成交金额': '成交金额'},
    )
    
    # 只有在618模式下才显示目标金额折线图
    if is_618_mode:
        fig_trend.add_scatter(
            x=trend_sum_full["_日期"],
            y=trend_sum_full["目标金额"],
            mode='lines+markers',
            name='目标金额',
            line=dict(width=3, color='red'),
            marker=dict(size=8),
            hovertemplate="%{x|%Y-%m-%d}<br>目标金额：%{y:,.2f} 万元<extra></extra>"
        )
    
    fig_trend.update_layout(
        xaxis_title="日期",
        yaxis_title="金额（万元）",
        height=400,
        margin=dict(l=10, r=10, t=10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig_trend.update_xaxes(
        tickformat="%m-%d", 
        hoverformat="%Y-%m-%d", 
        tickmode='array',
        tickvals=trend_sum_full["_日期"].tolist(),
        ticktext=[d.strftime("%m-%d") for d in trend_sum_full["_日期"]]
    )
    fig_trend.update_traces(
        texttemplate="%{y:,.2f}",
        textposition="outside",
        hovertemplate="%{x|%Y-%m-%d}<br>成交金额：%{y:,.2f} 万元<extra></extra>",
        selector=dict(type='bar')
    )
    st.plotly_chart(fig_trend, use_container_width=True)
else:
    st.info("当前筛选下无有效「日期」数据，无法绘制趋势图。")

st.markdown("---")

# 汇总透视：部门 × 平台
st.subheader("📋 部门 × 平台 汇总（万元）")
pivot = df_filtered.pivot_table(
    index="部门",
    columns="平台",
    values="成交金额",
    aggfunc="sum",
    fill_value=0.0,
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

display_df["成交金额(万元)"] = display_df["成交金额"].round(2)
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
    .assign(成交金额万元=lambda d: d["成交金额"])
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
