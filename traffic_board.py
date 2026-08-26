"""爆链看板：由 app.py 路由进入，读取 traffic_source_summary.csv。"""
from pathlib import Path
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
CSV_FILENAME = "traffic_source_summary.csv"
GOAL_CSV_FILENAME = "traffic_goal_summary.csv"
# 已知部门排序；CSV 中出现的新部门会追加在末尾
DEPARTMENT_ORDER = ["常温", "低温", "八喜", "奶粉"]
CHANNEL_ORDER = ["品牌", "站内", "站外"]
CHANNEL_COLORS = {
    "品牌": "#1e4976",
    "站内": "#1aabb8",
    "站外": "#e67e22",
}
NUMERIC_COLUMNS = ("访客数(UV)", "UV价值", "客单价", "成交转化率", "CTR", "成交金额")
def _csv_path() -> Path:
    return Path(__file__).resolve().parent / CSV_FILENAME
def _department_sort_key(name: str) -> int:
    if name in DEPARTMENT_ORDER:
        return DEPARTMENT_ORDER.index(name)
    return len(DEPARTMENT_ORDER) + ord(name[0] if name else "z")
def _channel_sort_key(name: str) -> int:
    if name in CHANNEL_ORDER:
        return CHANNEL_ORDER.index(name)
    return len(CHANNEL_ORDER) + ord(name[0] if name else "z")
@st.cache_data(show_spinner=True)
def load_traffic_csv(path_str: str, _mtime: float) -> pd.DataFrame:
    path = Path(path_str)
    if not path.is_file():
        return pd.DataFrame()
    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    for col in ("日期", "部门", "商品", "渠道"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    if "日期" in df.columns:
        df["_日期解析"] = pd.to_datetime(df["日期"], errors="coerce")
        df = df.sort_values(
            by=["_日期解析", "部门", "商品", "渠道"],
            na_position="last",
        )
    return df

def _goal_csv_path() -> Path:
    return Path(__file__).resolve().parent / GOAL_CSV_FILENAME

@st.cache_data(show_spinner=False)
def load_goal_csv(path_str: str, _mtime: float) -> pd.DataFrame:
    """读取目标数据CSV，列结构与源数据一致（日期/部门/商品 + 5个指标）。"""
    path = Path(path_str)
    if not path.is_file():
        return pd.DataFrame()
    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    for col in ("日期", "部门", "商品"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    if "日期" in df.columns:
        df["_日期解析"] = pd.to_datetime(df["日期"], errors="coerce")
    return df

def max_business_date_label(df: pd.DataFrame) -> str:
    if df.empty or "_日期解析" not in df.columns:
        return "—"
    tmax = df["_日期解析"].max()
    if pd.isna(tmax):
        return "—"
    return tmax.strftime("%Y-%m-%d")
def departments_in_data(df: pd.DataFrame) -> list[str]:
    if df.empty or "部门" not in df.columns:
        return []
    names = sorted({d for d in df["部门"].unique() if d}, key=_department_sort_key)
    return names
def apply_filters(
    df: pd.DataFrame,
    date_range: tuple,
    selected_products: list[str] | None = None,
    selected_channels: list[str] | None = None,
) -> pd.DataFrame:
    out = df.copy()
    if date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
        start_date = pd.Timestamp(date_range[0])
        end_date = pd.Timestamp(date_range[1])
        out = out[(out["_日期解析"] >= start_date) & (out["_日期解析"] <= end_date)]
    if selected_products:
        out = out[out["商品"].isin(selected_products)]
    if selected_channels:
        out = out[out["渠道"].isin(selected_channels)]
    return out
def aggregate_daily(df: pd.DataFrame) -> pd.DataFrame:
    """按日汇总；转化率/客单价按 UV 加权。"""
    if df.empty:
        return pd.DataFrame()
    tmp = df.dropna(subset=["_日期解析"]).copy()
    if tmp.empty:
        return pd.DataFrame()
    def _weighted(group: pd.DataFrame, col: str) -> float:
        uv = group["访客数(UV)"].sum()
        if uv <= 0:
            return 0.0
        return float((group[col] * group["访客数(UV)"]).sum() / uv)
    rows = []
    for day, group in tmp.groupby(tmp["_日期解析"].dt.date):
        uv = float(group["访客数(UV)"].sum())
        amount = float(group["成交金额"].sum())
        rows.append(
            {
                "日期": str(day),
                "_日期解析": pd.Timestamp(day),
                "访客数(UV)": uv,
                "成交金额": amount,
                "UV价值": amount / uv if uv > 0 else 0.0,
                "客单价": _weighted(group, "客单价"),
                "成交转化率": _weighted(group, "成交转化率"),
                "CTR": _weighted(group, "CTR"),
            }
        )
    return pd.DataFrame(rows).sort_values("_日期解析")
def aggregate_daily_by_channel(df: pd.DataFrame) -> pd.DataFrame:
    """按日+渠道汇总；转化率/客单价按 UV 加权。"""
    if df.empty:
        return pd.DataFrame()
    tmp = df.dropna(subset=["_日期解析"]).copy()
    if tmp.empty:
        return pd.DataFrame()
    def _weighted(group: pd.DataFrame, col: str) -> float:
        uv = group["访客数(UV)"].sum()
        if uv <= 0:
            return 0.0
        return float((group[col] * group["访客数(UV)"]).sum() / uv)
    rows = []
    for (day, channel), group in tmp.groupby([tmp["_日期解析"].dt.date, "渠道"]):
        uv = float(group["访客数(UV)"].sum())
        amount = float(group["成交金额"].sum())
        rows.append(
            {
                "日期": str(day),
                "_日期解析": pd.Timestamp(day),
                "渠道": channel,
                "访客数(UV)": uv,
                "成交金额": amount,
                "UV价值": amount / uv if uv > 0 else 0.0,
                "客单价": _weighted(group, "客单价"),
                "成交转化率": _weighted(group, "成交转化率"),
                "CTR": _weighted(group, "CTR"),
            }
        )
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out["渠道"] = pd.Categorical(
        out["渠道"],
        categories=[c for c in CHANNEL_ORDER if c in out["渠道"].unique()]
        + sorted(set(out["渠道"]) - set(CHANNEL_ORDER)),
        ordered=True,
    )
    return out.sort_values(["_日期解析", "渠道"])
def _pct_delta(current: float, previous: float) -> float | None:
    if previous == 0:
        return None
    return (current - previous) / previous * 100
def _metric_card(label: str, value: str, delta_pct: float | None, delta_hint: str) -> None:
    if delta_pct is None:
        st.metric(label, value, delta=delta_hint, delta_color="off")
    else:
        st.metric(label, value, delta=f"{delta_pct:+.1f}% {delta_hint}")
# 图表配色（参考电商核心数据看板）
COLOR_AMOUNT = "#1e4976"
COLOR_UV = "#1e4976"
COLOR_UV_VALUE = "#1aabb8"
COLOR_CONV = "#e67e22"
COLOR_AOV = "#1aabb8"
CHART_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color="#333", size=12),
    hovermode="closest",
)
def _prepare_chart_daily(daily: pd.DataFrame, goal_df: pd.DataFrame | None = None) -> pd.DataFrame:
    out = daily.copy().reset_index(drop=True)
    out["天序"] = [f"第{i + 1}天" for i in range(len(out))]
    out["成交金额_万"] = out["成交金额"] / 10000
    out["UV_万"] = out["访客数(UV)"] / 10000
    out["转化率_pct"] = out["成交转化率"] * 100
    out["日期_label"] = out["_日期解析"].apply(
        lambda t: t.strftime("%m-%d") if pd.notna(t) else ""
    )
    # 合并目标数据（按日期），目标列统一加 "_目标" 后缀
    if goal_df is not None and not goal_df.empty and "_日期解析" in goal_df.columns:
        goal_cols = [c for c in NUMERIC_COLUMNS if c in goal_df.columns]
        if goal_cols:
            g = goal_df[["_日期解析"] + goal_cols].copy()
            g = g.rename(columns={c: f"{c}_目标" for c in goal_cols})
            out = out.merge(g, on="_日期解析", how="left")
            # 派生图表单位的目标列
            if "成交金额_目标" in out.columns:
                out["成交金额_万_目标"] = out["成交金额_目标"] / 10000
            if "访客数(UV)_目标" in out.columns:
                out["UV_万_目标"] = out["访客数(UV)_目标"] / 10000
            if "成交转化率_目标" in out.columns:
                out["转化率_pct_目标"] = out["成交转化率_目标"] * 100
    return out
def _wan_label(value: float) -> str:
    if abs(value) >= 1:
        return f"{value:.1f}万"
    if abs(value) >= 0.01:
        return f"{value:.2f}万"
    return f"{value * 10000:,.0f}"
def _fig_daily_metric(
    chart_df: pd.DataFrame,
    y_col: str,
    title: str,
    color: str,
    chart_type: str = "bar",
    y_title: str = "",
    hover_col: str | None = None,
    hover_fmt: str = "{:,.2f}",
    text_fn=None,
    goal_col: str | None = None,
    goal_hover_col: str | None = None,
    goal_hover_fmt: str | None = None,
) -> go.Figure:
    """按天显示单个指标：柱状图或折线图；若提供 goal_col 则叠加目标虚线。"""
    x = chart_df["日期_label"]
    y = chart_df[y_col]
    hover_src = chart_df[hover_col] if hover_col else y
    customdata = [hover_fmt.format(v) for v in hover_src]
    text_labels = [text_fn(v) for v in y] if text_fn else None
    hover_tpl = "%{x}<br>" + title + ": %{customdata}<extra></extra>"
    if chart_type == "bar":
        trace = go.Bar(
            x=x,
            y=y,
            marker_color=color,
            name=title,
            customdata=customdata,
            hovertemplate=hover_tpl,
            text=text_labels,
            textposition="outside",
            textfont=dict(size=10, color="#FFFFFF"),
        )
    else:
        trace = go.Scatter(
            x=x,
            y=y,
            mode="lines+markers+text",  # 强制开启文字显示
            line=dict(color=color, width=2.5),
            marker=dict(size=7, color=color),
            name=title,
            customdata=customdata,
            hovertemplate=hover_tpl,
            text=text_labels,
            textposition="top center",
            textfont=dict(size=10, color="#FFFFFF"),
        )
    fig = go.Figure(trace)
    # 叠加目标虚线
    has_goal = goal_col is not None and goal_col in chart_df.columns
    if has_goal:
        goal_y = chart_df[goal_col]
        goal_hover_src = chart_df[goal_hover_col] if goal_hover_col and goal_hover_col in chart_df.columns else goal_y
        goal_fmt = goal_hover_fmt or hover_fmt
        goal_customdata = [goal_fmt.format(v) if pd.notna(v) else "—" for v in goal_hover_src]
        goal_hover_tpl = "%{x}<br>目标: %{customdata}<extra></extra>"
        fig.add_trace(
            go.Scatter(
                x=x,
                y=goal_y,
                mode="lines",
                line=dict(color="#e74c3c", width=2, dash="dash"),
                name="目标",
                customdata=goal_customdata,
                hovertemplate=goal_hover_tpl,
            )
        )
    fig.update_layout(
        **CHART_LAYOUT,
        height=320,
        margin=dict(l=40, r=20, t=40, b=40),
        title=dict(text=title, x=0, font=dict(size=15)),
        yaxis=dict(title=y_title, gridcolor="rgba(0,0,0,0.06)", zeroline=False),
        xaxis=dict(title="", gridcolor="rgba(0,0,0,0)"),
        showlegend=has_goal,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=11),
        ) if has_goal else None,
    )
    return fig
def _render_metric_tabs(daily_df: pd.DataFrame, goal_df: pd.DataFrame | None = None) -> None:
    """按天显示 5 个指标的 Tab 页：成交金额/UV价值/访客数/客单价/成交转化率。"""
    chart_df = _prepare_chart_daily(daily_df, goal_df)
    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["成交金额", "UV价值", "访客数", "客单价", "成交转化率"]
    )
    with tab1:
        st.plotly_chart(
            _fig_daily_metric(
                chart_df, "成交金额_万", "成交金额", COLOR_AMOUNT,
                chart_type="bar", y_title="金额（万）",
                hover_col="成交金额", hover_fmt="¥{:,.2f}",
                text_fn=_wan_label,
                goal_col="成交金额_万_目标",
                goal_hover_col="成交金额_目标",
                goal_hover_fmt="¥{:,.2f}",
            ),
            width="stretch",
        )
    with tab2:
        st.plotly_chart(
            _fig_daily_metric(
                chart_df, "UV价值", "UV价值", COLOR_UV_VALUE,
                chart_type="line", y_title="UV价值（¥）",
                hover_fmt="¥{:,.2f}", text_fn=lambda v: f"¥{v:.2f}",
                goal_col="UV价值_目标",
                goal_hover_fmt="¥{:,.2f}",
            ),
            width="stretch",
        )
    with tab3:
        st.plotly_chart(
            _fig_daily_metric(
                chart_df, "UV_万", "访客数", COLOR_UV,
                chart_type="line", y_title="访客数（万）",
                hover_col="访客数(UV)", hover_fmt="{:,.0f}",
                text_fn=_wan_label,
                goal_col="UV_万_目标",
                goal_hover_col="访客数(UV)_目标",
                goal_hover_fmt="{:,.0f}",
            ),
            width="stretch",
        )
    with tab4:
        st.plotly_chart(
            _fig_daily_metric(
                chart_df, "客单价", "客单价", COLOR_AOV,
                chart_type="line", y_title="客单价（¥）",
                hover_fmt="¥{:,.2f}", text_fn=lambda v: f"¥{v:.0f}",
                goal_col="客单价_目标",
                goal_hover_fmt="¥{:,.2f}",
            ),
            width="stretch",
        )
    with tab5:
        st.plotly_chart(
            _fig_daily_metric(
                chart_df, "转化率_pct", "成交转化率", COLOR_CONV,
                chart_type="line", y_title="转化率（%）",
                hover_col="成交转化率", hover_fmt="{:.2%}",
                text_fn=lambda v: f"{v:.1f}%",
                goal_col="转化率_pct_目标",
                goal_hover_col="成交转化率_目标",
                goal_hover_fmt="{:.2%}",
            ),
            width="stretch",
        )
def _prepare_chart_daily_by_channel(by_channel: pd.DataFrame) -> pd.DataFrame:
    """为按渠道拆分的日数据添加图表列。"""
    if by_channel.empty:
        return pd.DataFrame()
    out = by_channel.copy().reset_index(drop=True)
    out["日期_label"] = out["_日期解析"].apply(
        lambda t: t.strftime("%m-%d") if pd.notna(t) else ""
    )
    out["成交金额_万"] = out["成交金额"] / 10000
    out["UV_万"] = out["访客数(UV)"] / 10000
    out["转化率_pct"] = out["成交转化率"] * 100
    out["CTR_pct"] = out["CTR"] * 100
    daily_total = out.groupby("_日期解析")["成交金额"].transform("sum")
    out["成交金额_pct"] = (out["成交金额"] / daily_total * 100).round(2)
    return out
def _fig_channel_amount_pct(chart_df: pd.DataFrame) -> go.Figure:
    """成交金额按渠道百分比堆叠柱状图。"""
    fig = go.Figure()
    for channel in CHANNEL_ORDER:
        sub = chart_df[chart_df["渠道"] == channel]
        if sub.empty:
            continue
        fig.add_trace(
            go.Bar(
                x=sub["日期_label"],
                y=sub["成交金额_pct"],
                name=channel,
                marker_color=CHANNEL_COLORS.get(channel, COLOR_AMOUNT),
                hovertemplate="%{x}<br>%{fullData.name}: %{y}%<extra></extra>",
                text=[f"{v:.1f}%" for v in sub["成交金额_pct"]],
                textposition="inside",
                textfont=dict(size=10, color="#fff"),
            )
        )
    fig.update_layout(
        **CHART_LAYOUT,
        height=320,
        margin=dict(l=40, r=20, t=40, b=40),
        title=dict(text="成交金额占比（按渠道）", x=0, font=dict(size=15)),
        barmode="stack",
        yaxis=dict(
            title="占比（%）",
            gridcolor="rgba(0,0,0,0.06)",
            zeroline=False,
            ticksuffix="%",
        ),
        xaxis=dict(title="", gridcolor="rgba(0,0,0,0)"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=11),
        ),
    )
    return fig
def _fig_channel_line(
    chart_df: pd.DataFrame,
    y_col: str,
    title: str,
    y_title: str = "",
    hover_fmt: str = "{:,.2f}",
) -> go.Figure:
    """按渠道拆分的多折线图（3 条线）。"""
    fig = go.Figure()
    for channel in CHANNEL_ORDER:
        sub = chart_df[chart_df["渠道"] == channel]
        if sub.empty:
            continue
        color = CHANNEL_COLORS.get(channel, COLOR_UV)
        customdata = [hover_fmt.format(v) for v in sub[y_col]]
        # 生成图上显示的文本
        text_labels = [hover_fmt.format(v) for v in sub[y_col]]
        fig.add_trace(
            go.Scatter(
                x=sub["日期_label"],
                y=sub[y_col],
                mode="lines+markers+text",  # 开启文字
                name=channel,
                line=dict(color=color, width=2.5),
                marker=dict(size=7, color=color),
                customdata=customdata,
                text=text_labels,
                textposition="top center",
                textfont=dict(size=9, color="#FFFFFF"), # 强制纯白文字
                hovertemplate="%{x}<br>" + title + "（%{fullData.name}）: %{customdata}<extra></extra>",
            )
        )
    fig.update_layout(
        **CHART_LAYOUT,
        height=320,
        margin=dict(l=40, r=20, t=40, b=40),
        title=dict(text=title, x=0, font=dict(size=15)),
        yaxis=dict(title=y_title, gridcolor="rgba(0,0,0,0.06)", zeroline=False),
        xaxis=dict(title="", gridcolor="rgba(0,0,0,0)"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=11),
        ),
    )
    return fig
def _render_metric_tabs_by_channel(by_channel_df: pd.DataFrame) -> None:
    """按渠道拆分的 5 指标 Tab 页：成交金额(占比堆叠)/UV价值/访客数/客单价/成交转化率。"""
    chart_df = _prepare_chart_daily_by_channel(by_channel_df)
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        ["成交金额", "UV价值", "访客数", "客单价", "成交转化率", "CTR"]
    )
    with tab1:
        st.plotly_chart(_fig_channel_amount_pct(chart_df), width="stretch")
    with tab2:
        st.plotly_chart(
            _fig_channel_line(chart_df, "UV价值", "UV价值（按渠道）", "UV价值（¥）", "¥{:,.2f}"),
            width="stretch",
        )
    with tab3:
        # 修复：使用原始访客数字段，不再用UV_万，展示真实几百/几千数值
        st.plotly_chart(
            _fig_channel_line(chart_df, "访客数(UV)", "访客数（按渠道）", "访客数", "{:,.0f}"),
            width="stretch",
        )
    with tab4:
        st.plotly_chart(
            _fig_channel_line(chart_df, "客单价", "客单价（按渠道）", "客单价（¥）", "¥{:,.2f}"),
            width="stretch",
        )
    with tab5:
        st.plotly_chart(
            _fig_channel_line(chart_df, "转化率_pct", "成交转化率（按渠道）", "转化率（%）", "{:.2f}%"),
            width="stretch",
        )
    with tab6:
        # CTR不分渠道，按日聚合为商品级单条线（同一天各渠道CTR值相同）
        if not chart_df.empty and "CTR" in chart_df.columns:
            ctr_daily = chart_df.groupby(
                ["_日期解析", "日期_label"], as_index=False
            ).agg({"CTR": "mean", "CTR_pct": "mean"})
            st.plotly_chart(
                _fig_daily_metric(
                    ctr_daily, "CTR_pct", "CTR", "#27ae60",
                    chart_type="line", y_title="CTR（%）",
                    hover_col="CTR", hover_fmt="{:.2%}",
                    text_fn=lambda v: f"{v:.1f}%",
                ),
                width="stretch",
            )
def render() -> None:
    st.sidebar.title("⚙️ 控制面板")
    csv_path = _csv_path()
    try:
        mtime = csv_path.stat().st_mtime
    except OSError:
        mtime = 0.0
    df_all = load_traffic_csv(str(csv_path), mtime)
    # 加载目标数据
    goal_path = _goal_csv_path()
    try:
        goal_mtime = goal_path.stat().st_mtime
    except OSError:
        goal_mtime = 0.0
    goal_all = load_goal_csv(str(goal_path), goal_mtime)
    dept_options = departments_in_data(df_all)
    if not dept_options:
        dept_options = DEPARTMENT_ORDER[:1]
    selected_dept = st.sidebar.radio(
        "部门",
        dept_options,
        index=0,
        key="traffic_dept",
    )
    df_dept = df_all[df_all["部门"] == selected_dept].copy() if not df_all.empty else df_all
    last_date = max_business_date_label(df_dept)
    st.markdown(
        f"<div style='display:flex;justify-content:space-between;align-items:flex-start;'>"
        f"<h1>🔗 {selected_dept}爆链看板</h1>"
        f"<div style='font-size:13px;color:#666;text-align:right;'>"
        f"<strong>数据源：</strong> {CSV_FILENAME}<br>"
        f"<strong>数据最后日期：</strong> {last_date}"
        f"</div></div>",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    if df_dept.empty:
        st.error(f"❌ 未读取到数据。请将 `{CSV_FILENAME}` 放在：`{csv_path.parent}`")
        st.info("期望列：日期、部门、商品、渠道、访客数(UV)、UV价值、客单价、成交转化率、CTR、成交金额")
        return
    st.sidebar.header("🔍 筛选条件")
    valid_dates = df_dept["_日期解析"].dropna()
    min_date = valid_dates.min().date() if not valid_dates.empty else pd.Timestamp.now().date()
    max_date = valid_dates.max().date() if not valid_dates.empty else pd.Timestamp.now().date()
    default_start = max(min_date, (pd.Timestamp(max_date) - pd.Timedelta(days=13)).date())
    date_range = st.sidebar.date_input(
        "统计日期范围",
        value=(default_start, max_date),
        min_value=min_date,
        max_value=max_date,
        key=f"traffic_date_range_{selected_dept}",
    )
    # 规范化：st.date_input 在起止日期相同时可能返回单个 date 对象，需转为 (date, date)
    if not isinstance(date_range, tuple):
        date_range = (date_range, date_range)
    elif len(date_range) == 1:
        date_range = (date_range[0], date_range[0])
    df = apply_filters(df_dept, date_range, None, None)
    # 指标卡片：取商品=all 的数据
    df_all_product_full = apply_filters(df_dept, date_range, None, None)
    if not df_all_product_full.empty and "商品" in df_all_product_full.columns:
        df_all_product_full = df_all_product_full[df_all_product_full["商品"] == "all"]
    daily = aggregate_daily(df_all_product_full) if not df_all_product_full.empty else pd.DataFrame()
    if daily.empty:
        st.info("当前筛选条件下无商品=all 的数据。")
        return
    # 最新日 vs 前一日（商品=all 全量日期范围）
    df_for_dod = apply_filters(df_dept, (min_date, max_date), None, None)
    if not df_for_dod.empty and "商品" in df_for_dod.columns:
        df_for_dod = df_for_dod[df_for_dod["商品"] == "all"]
    daily_all = aggregate_daily(df_for_dod) if not df_for_dod.empty else pd.DataFrame()
    latest_row = daily_all.iloc[-1]
    prev_row = daily_all.iloc[-2] if len(daily_all) >= 2 else None
    # 查找最新业务日对应的目标数据（同部门 + 商品=all）
    goal_row = None
    if not goal_all.empty:
        goal_dept = goal_all[goal_all["部门"] == selected_dept].copy()
        if not goal_dept.empty and "商品" in goal_dept.columns:
            goal_dept = goal_dept[goal_dept["商品"] == "all"]
        if not goal_dept.empty and "_日期解析" in goal_dept.columns:
            latest_ts = latest_row["_日期解析"]
            matched = goal_dept[goal_dept["_日期解析"] == latest_ts]
            if not matched.empty:
                goal_row = matched.iloc[0]
    range_uv = float(daily["访客数(UV)"].sum())
    range_amount = float(daily["成交金额"].sum())
    range_uv_value = range_amount / range_uv if range_uv > 0 else 0.0
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        if goal_row is not None:
            _metric_card(
                "访客数(UV)",
                f"{latest_row['访客数(UV)']:,.0f}",
                _pct_delta(latest_row['访客数(UV)'], goal_row["访客数(UV)"]),
                "较目标",
            )
        else:
            st.metric("访客数(UV)", f"{latest_row['访客数(UV)']:,.0f}")
    with c2:
        if goal_row is not None:
            _metric_card(
                "成交金额",
                f"¥{latest_row['成交金额']:,.0f}",
                _pct_delta(latest_row["成交金额"], goal_row["成交金额"]),
                "较目标",
            )
        else:
            st.metric("成交金额", f"¥{latest_row['成交金额']:,.0f}")
    with c3:
        conv_pct = latest_row["成交转化率"] * 100
        if goal_row is not None:
            delta_pp = (latest_row["成交转化率"] - goal_row["成交转化率"]) * 100
            st.metric("成交转化率", f"{conv_pct:.2f}%", delta=f"{delta_pp:+.2f}pp 较目标")
        else:
            st.metric("成交转化率", f"{conv_pct:.2f}%")
    with c4:
        uv_val = latest_row["UV价值"]
        if goal_row is not None:
            delta_uv = uv_val - goal_row["UV价值"]
            st.metric("UV价值", f"¥{uv_val:.2f}", delta=f"{delta_uv:+.2f} ¥ 较目标")
        else:
            st.metric("UV价值", f"¥{uv_val:.2f}")
    with c5:
        aov_val = latest_row["客单价"]
        if goal_row is not None:
            delta_aov = aov_val - goal_row["客单价"]
            st.metric("客单价", f"¥{aov_val:.2f}", delta=f"{delta_aov:+.2f} ¥ 较目标")
        else:
            st.metric("客单价", f"¥{aov_val:.2f}")
    st.caption(f"指标卡片为最新业务日 **{latest_row['日期']}** 店铺数据，对比当日目标")
    def _daily_for_product(product: str) -> pd.DataFrame:
        """按商品过滤后按天聚合；仅尊重日期范围。"""
        sub = apply_filters(df_dept, date_range, None, None)
        if not sub.empty and "商品" in sub.columns:
            sub = sub[sub["商品"] == product]
        if sub.empty:
            return pd.DataFrame()
        return aggregate_daily(sub)
    def _daily_by_channel_for_product(product: str) -> pd.DataFrame:
        """按商品过滤后按天+渠道聚合。"""
        sub = apply_filters(df_dept, date_range, None, None)
        if not sub.empty and "商品" in sub.columns:
            sub = sub[sub["商品"] == product]
        if sub.empty:
            return pd.DataFrame()
        return aggregate_daily_by_channel(sub)
    all_daily = _daily_for_product("all")
    # 过滤目标数据：同部门 + 商品=all + 日期范围
    goal_daily = pd.DataFrame()
    if not goal_all.empty:
        goal_daily = goal_all[goal_all["部门"] == selected_dept].copy()
        if not goal_daily.empty and "商品" in goal_daily.columns:
            goal_daily = goal_daily[goal_daily["商品"] == "all"]
        if not goal_daily.empty and date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
            start_ts = pd.Timestamp(date_range[0])
            end_ts = pd.Timestamp(date_range[1])
            goal_daily = goal_daily[
                (goal_daily["_日期解析"] >= start_ts) & (goal_daily["_日期解析"] <= end_ts)
            ]
    st.markdown("---")
    st.subheader("📊 店铺总览按天趋势")
    if all_daily.empty:
        st.info("当前筛选条件下无商品=all 的数据。")
    else:
        st.caption("仅统计商品=all 的数据，点击标签切换指标；红色虚线为目标值")
        _render_metric_tabs(all_daily, goal_daily if not goal_daily.empty else None)
    # 动态获取当前部门下的细分商品（排除 all），逐个生成分析区域
    sub_products = sorted({p for p in df_dept["商品"].unique() if p and p != "all"})
    for product in sub_products:
        by_channel = _daily_by_channel_for_product(product)
        if by_channel.empty:
            continue
        st.markdown("---")
        st.subheader(f"📈 {product}分析")
        st.caption(f"仅统计商品={product} 的数据，按渠道拆分，点击标签切换指标")
        _render_metric_tabs_by_channel(by_channel)
    st.markdown("---")
    product_options = ["全部"] + sorted({p for p in df["商品"].unique() if p})
    channel_options = ["全部"] + sorted(
        {c for c in df["渠道"].unique() if c},
        key=_channel_sort_key,
    )
    col_h, col_p, col_c = st.columns([2, 1.5, 1.5])
    with col_h:
        st.subheader("📋 明细数据")
    with col_p:
        detail_product = st.selectbox(
            "商品",
            product_options,
            index=0,
            key=f"traffic_products_{selected_dept}",
        )
    with col_c:
        detail_channel = st.selectbox(
            "渠道",
            channel_options,
            index=0,
            key=f"traffic_channels_{selected_dept}",
        )
    detail_df = df
    if detail_product != "全部":
        detail_df = detail_df[detail_df["商品"] == detail_product]
    if detail_channel != "全部":
        detail_df = detail_df[detail_df["渠道"] == detail_channel]
    display = detail_df.drop(columns=["_日期解析"], errors="ignore").copy()
    display = display.sort_values(["日期", "商品", "渠道"], ascending=[False, True, True])
    format_dict = {
        "访客数(UV)": "{:,.0f}",
        "UV价值": "¥{:,.2f}",
        "客单价": "¥{:,.2f}",
        "成交转化率": "{:.2%}",
        "CTR": "{:.2%}",
        "成交金额": "¥{:,.2f}",
    }
    st.dataframe(
        display.style.format({k: v for k, v in format_dict.items() if k in display.columns}),
        width="stretch",
        hide_index=True,
    )