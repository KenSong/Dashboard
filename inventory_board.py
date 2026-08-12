"""常温销量看板：由 app.py 路由进入，读取 normal_sales_amount.csv。"""

from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

PLATFORM_ORDER = ["京东", "天猫", "拼多多", "抖音", "新零售", "多多买菜", "小程序及其他"]

DEPARTMENTS = ["常温", "低温"]

DEPARTMENT_CSV_MAP = {
    "常温": "normal_sales_amount.csv",
    "低温": "low_sales_amount.csv",
    "奶粉": "milk_powder_sales_amount.csv",
    "八喜": "baxi_sales_amount.csv",
}

CSV_2025_MAP = {
    "常温": "normal_sales_amount_2025.csv",
    "低温": "low_sales_amount_2025.csv",
}

DAMAGE_CSV_MAP = {
    "常温": "normal_damage.csv",
    "低温": "low_damage.csv",
}


def _sales_csv_path(department: str) -> Path:
    filename = DEPARTMENT_CSV_MAP.get(department, "normal_sales_amount.csv")
    return Path(__file__).resolve().parent / filename


def _sales_csv_path_2025(department: str) -> Path:
    filename = CSV_2025_MAP.get(department, "")
    if not filename:
        return Path()
    return Path(__file__).resolve().parent / filename


def _damage_csv_path(department: str) -> Path:
    filename = DAMAGE_CSV_MAP.get(department, "normal_damage.csv")
    return Path(__file__).resolve().parent / filename


def _platform_sort_key(name: str) -> int:
    if name in PLATFORM_ORDER:
        return PLATFORM_ORDER.index(name)
    return len(PLATFORM_ORDER) + ord(name[0] if name else "z")


@st.cache_data(show_spinner=True)
def load_sales_csv(path_str: str, _mtime: float) -> pd.DataFrame:
    path = Path(path_str)
    if not path.is_file():
        return pd.DataFrame()
    df = pd.read_csv(path, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]
    rename = {
        "统计日期": "日期",
        "产品名称": "产品",
        "销售金额": "销售成本",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    for col in ("日期", "平台", "子平台", "产品"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    if "销售数量" in df.columns:
        df["销售数量"] = pd.to_numeric(df["销售数量"], errors="coerce").fillna(0.0)
    if "销售成本" in df.columns:
        df["销售成本"] = pd.to_numeric(df["销售成本"], errors="coerce").fillna(0.0)
    if "日期" in df.columns:
        df["_日期解析"] = pd.to_datetime(df["日期"], errors="coerce")
        df = df.sort_values(
            by=["_日期解析", "平台", "子平台", "产品"],
            na_position="last",
        ).drop(columns=["_日期解析"])
    return df


def max_business_date_label(df: pd.DataFrame) -> str:
    if df.empty or "日期" not in df.columns:
        return "—"
    s = df["日期"].astype(str).str.strip()
    s = s[s.ne("") & s.str.lower().ne("nan")]
    if s.empty:
        return "—"
    tmax = pd.to_datetime(s, errors="coerce").max()
    if pd.isna(tmax):
        return "—"
    return tmax.strftime("%Y-%m-%d")


def apply_filters(
    df: pd.DataFrame,
    date_range: tuple,
    selected_platforms: list = None,
    selected_sub_platforms: list = None,
) -> pd.DataFrame:
    out = df.copy()
    if date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
        start_date = pd.Timestamp(date_range[0])
        end_date = pd.Timestamp(date_range[1])
        out = out.assign(_日期解析=pd.to_datetime(out["日期"], errors="coerce"))
        out = out[(out["_日期解析"] >= start_date) & (out["_日期解析"] <= end_date)].drop(columns=["_日期解析"])
    if selected_platforms and len(selected_platforms) > 0:
        out = out[out["平台"].isin(selected_platforms)]
    if selected_sub_platforms and len(selected_sub_platforms) > 0:
        out = out[out["子平台"].isin(selected_sub_platforms)]
    return out


def apply_damage_filters(
    df: pd.DataFrame,
    date_range: tuple,
    selected_platforms: list = None,
    selected_sub_platforms: list = None,
) -> pd.DataFrame:
    """退损数据按月份过滤：日期范围覆盖的所有月份都纳入统计。

    退损 CSV 为按月汇总数据（每月一条记录），故按所选日期范围覆盖的
    月份进行过滤，而非精确到日期。
    """
    out = df.copy()
    if date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
        start_month = pd.Timestamp(date_range[0]).to_period("M")
        end_month = pd.Timestamp(date_range[1]).to_period("M")
        out = out.assign(_日期解析=pd.to_datetime(out["日期"], errors="coerce"))
        out["_月份"] = out["_日期解析"].dt.to_period("M")
        out = out[(out["_月份"] >= start_month) & (out["_月份"] <= end_month)]
        out = out.drop(columns=["_日期解析", "_月份"])
    if selected_platforms and len(selected_platforms) > 0:
        out = out[out["平台"].isin(selected_platforms)]
    if selected_sub_platforms and len(selected_sub_platforms) > 0:
        out = out[out["子平台"].isin(selected_sub_platforms)]
    return out


def render() -> None:
    st.sidebar.title("⚙️ 控制面板")

    selected_dept = st.sidebar.radio(
        "部门",
        DEPARTMENTS,
        index=DEPARTMENTS.index("常温"),
        key="inv_dept",
    )

    csv_path = _sales_csv_path(selected_dept)
    try:
        mtime = csv_path.stat().st_mtime
    except OSError:
        mtime = 0.0

    df_all = load_sales_csv(str(csv_path), mtime)
    last_date = max_business_date_label(df_all)

    # 2025年对比数据
    df_2025 = pd.DataFrame()
    csv_2025_path = _sales_csv_path_2025(selected_dept)
    if csv_2025_path.is_file():
        try:
            mtime_2025 = csv_2025_path.stat().st_mtime
            df_2025 = load_sales_csv(str(csv_2025_path), mtime_2025)
        except Exception:
            df_2025 = pd.DataFrame()

    # 损坏数据
    df_damage_all = pd.DataFrame()
    damage_path = _damage_csv_path(selected_dept)
    if damage_path.is_file():
        try:
            mtime_damage = damage_path.stat().st_mtime
            df_damage_all = load_sales_csv(str(damage_path), mtime_damage)
        except Exception:
            df_damage_all = pd.DataFrame()

    st.markdown(
        f"<div style='display:flex;justify-content:space-between;align-items:flex-start;'>"
        f"<h1>📦 {selected_dept}销量看板</h1>"
        f"<div style='font-size:13px;color:#666;text-align:right;'>"
        f"<strong>数据源：</strong> 电商管家及Oracle<br>"
        f"<strong>数据最后日期：</strong> {last_date}"
        f"</div></div>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    if df_all.empty:
        st.error(f"❌ 未读取到数据。请将 `{csv_path.name}` 放在：`{csv_path.parent}`")
        st.info("期望列：统计日期、平台、子平台、产品名称、销售数量")
        return

    st.sidebar.header("🔍 筛选条件")

    date_list = sorted(
        [d for d in df_all["日期"].unique() if d],
        key=lambda x: pd.to_datetime(x, errors="coerce"),
    )
    min_date = pd.to_datetime(date_list[0]) if date_list else pd.Timestamp.now() - pd.Timedelta(days=30)
    max_date = pd.to_datetime(date_list[-1]) if date_list else pd.Timestamp.now()

    default_start_date = max(min_date, max_date - pd.Timedelta(days=6))
    date_range = st.sidebar.date_input(
        "统计日期范围",
        value=(default_start_date, max_date),
        min_value=min_date,
        max_value=max_date,
        key=f"inv_date_range_{selected_dept}",
    )

    # 启用2025年数据对比
    show_2025 = st.sidebar.checkbox(
        "显示2025年同期对比",
        value=False,
        key=f"inv_show_2025_{selected_dept}",
        disabled=df_2025.empty,
    )

    df_date_filtered = apply_filters(df_all, date_range)

    platform_options = sorted(
        [p for p in df_date_filtered["平台"].unique() if p],
        key=_platform_sort_key,
    )
    platform_options_with_all = ["全部"] + platform_options
    selected_platform = st.sidebar.selectbox(
        "平台",
        platform_options_with_all,
        index=0,
        key=f"inv_platform_{selected_dept}",
    )
    selected_platforms = [] if selected_platform == "全部" else [selected_platform]

    sub_platform_options = sorted(
        [s for s in df_date_filtered[df_date_filtered["平台"].isin(selected_platforms)]["子平台"].unique() if s]
        if selected_platforms
        else [s for s in df_date_filtered["子平台"].unique() if s]
    )
    sub_platform_options_with_all = ["全部"] + sub_platform_options
    selected_sub_platform = st.sidebar.selectbox(
        "子平台",
        sub_platform_options_with_all,
        index=0,
        key=f"inv_sub_platform_{selected_dept}",
    )
    selected_sub_platforms = [] if selected_sub_platform == "全部" else [selected_sub_platform]

    df = apply_filters(df_all, date_range, selected_platforms, selected_sub_platforms)

    # 2025年同期数据处理（日期对齐：年份减1）
    df_2025_filtered = pd.DataFrame()
    if show_2025 and not df_2025.empty and date_range and len(date_range) == 2:
        start_date_2025 = pd.Timestamp(date_range[0]) - pd.Timedelta(days=365)
        end_date_2025 = pd.Timestamp(date_range[1]) - pd.Timedelta(days=365)
        df_2025_filtered = apply_filters(df_2025, (start_date_2025.date(), end_date_2025.date()), selected_platforms, selected_sub_platforms)

    # 损坏数据筛选（按月份覆盖统计）
    df_damage = pd.DataFrame()
    if not df_damage_all.empty:
        df_damage = apply_damage_filters(df_damage_all, date_range, selected_platforms, selected_sub_platforms)

    total_qty = float(df["销售数量"].sum()) if not df.empty else 0.0
    total_amount = float(df["销售成本"].sum()) if not df.empty and "销售成本" in df.columns else 0.0
    product_cnt = df["产品"].nunique() if not df.empty else 0
    total_qty_2025 = float(df_2025_filtered["销售数量"].sum()) if not df_2025_filtered.empty else 0.0
    yoy_qty_growth = ((total_qty - total_qty_2025) / total_qty_2025 * 100) if total_qty_2025 > 0 else 0.0

    total_amount_m = total_amount / 1000000

    total_damage = float(abs(df_damage["销售数量"]).sum()) if not df_damage.empty else 0.0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("销售数量合计", f"{total_qty:,.0f}")
    c2.metric("产品销售成本(百万)", f"{total_amount_m:,.2f}")
    c3.metric("产品数", f"{product_cnt:,}")
    c4.metric(
        "数量同比",
        f"{yoy_qty_growth:+.1f}%",
        delta=f"较2025年 {total_qty_2025:,.0f}",
        delta_color="inverse"
    )
    c5.metric("退损数量", f"{total_damage:,.0f}")

    st.markdown("---")

    if not df.empty:
        tmp = df.assign(_日期=pd.to_datetime(df["日期"], errors="coerce")).dropna(subset=["_日期"])
        if not tmp.empty:
            st.subheader("📊 销售趋势（按日）")

            has_amount = "销售成本" in df.columns

            # 2026年数据 - 按日汇总数量和金额
            daily_2026 = tmp.groupby(tmp["_日期"].dt.date).agg(
                销售数量=("销售数量", "sum"),
                销售成本=("销售成本", "sum") if has_amount else ("销售数量", "sum"),
            ).reset_index()
            daily_2026 = daily_2026.rename(columns={"_日期": "日期"})
            daily_2026["日期"] = daily_2026["日期"].astype(str)

            # 2025年对比数据
            daily_2025 = None
            if show_2025 and not df_2025_filtered.empty:
                tmp_2025 = df_2025_filtered.assign(_日期=pd.to_datetime(df_2025_filtered["日期"], errors="coerce")).dropna(subset=["_日期"])
                if not tmp_2025.empty:
                    agg_dict = {"销售数量_2025": ("销售数量", "sum")}
                    if "销售成本" in df_2025_filtered.columns:
                        agg_dict["销售成本_2025"] = ("销售成本", "sum")
                    daily_2025 = tmp_2025.groupby(tmp_2025["_日期"].dt.date).agg(**agg_dict).reset_index()
                    daily_2025 = daily_2025.rename(columns={"_日期": "日期"})
                    daily_2025["日期"] = daily_2025["日期"].astype(str)
                    daily_2025["日期"] = daily_2025["日期"].str.replace("2025", "2026")

            # 合并数据
            daily = daily_2026.copy()
            if daily_2025 is not None:
                daily = pd.merge(daily, daily_2025, on="日期", how="outer").fillna(0)
                daily = daily.sort_values("日期")

            # 构建组合图：数量为柱状图，金额为折线图（仅在未对比2025年时显示金额）
            fig_trend = go.Figure()

            # 柱状图：销售数量
            bar_cols = [c for c in ["销售数量", "销售数量_2025"] if c in daily.columns]
            bar_colors = {"销售数量": "#636EFA", "销售数量_2025": "#B4A2F9"}
            bar_labels = {"销售数量": "数量(2026)", "销售数量_2025": "数量(2025)"}

            for col in bar_cols:
                fig_trend.add_trace(go.Bar(
                    x=daily["日期"],
                    y=daily[col],
                    name=bar_labels.get(col, col),
                    marker_color=bar_colors.get(col),
                    text=daily[col].apply(lambda v: f"{v:,.0f}"),
                    textposition="outside",
                ))

            # 折线图：销售成本（仅在未对比2025年时显示）
            if not show_2025:
                line_cols = [c for c in ["销售成本"] if c in daily.columns]
                line_colors = {"销售成本": "#EF553B"}
                line_labels = {"销售成本": "销售成本"}

                for col in line_cols:
                    fig_trend.add_trace(go.Scatter(
                        x=daily["日期"],
                        y=daily[col],
                        name=line_labels.get(col, col),
                        mode="lines+markers",
                        line=dict(color=line_colors.get(col), width=2),
                        marker=dict(size=5),
                        yaxis="y2",
                        text=[f"{v/1000000:,.2f}M" for v in daily[col]],
                        textposition="top center",
                        hovertemplate="%{text}<extra></extra>",
                    ))

            layout_kwargs = dict(
                height=380,
                margin=dict(l=60, r=60, t=10, b=10),
                xaxis_title="日期",
                xaxis=dict(type="category"),
                yaxis=dict(title="销售数量"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                barmode="group",
            )
            if not show_2025 and "销售成本" in daily.columns:
                layout_kwargs["yaxis2"] = dict(title="销售成本 (百万)", overlaying="y", side="right", showgrid=False)
            fig_trend.update_layout(**layout_kwargs)
            st.plotly_chart(fig_trend, width="stretch")

            st.markdown("---")
            st.subheader("📊 平台销售分布")

            has_amount = "销售成本" in df.columns

            # 按平台汇总（数量和金额）
            plat_grouped = df.groupby("平台", as_index=False).agg(
                销售数量=("销售数量", "sum"),
                销售成本=("销售成本", "sum") if has_amount else ("销售数量", "sum"),
            )
            plat_grouped = plat_grouped.sort_values("销售数量", ascending=False)

            # 2025年对比数据
            plat_2025 = None
            if show_2025 and not df_2025_filtered.empty:
                plat_agg = {"销售数量_2025": ("销售数量", "sum")}
                if "销售成本" in df_2025_filtered.columns:
                    plat_agg["销售成本_2025"] = ("销售成本", "sum")
                plat_2025 = df_2025_filtered.groupby("平台", as_index=False).agg(**plat_agg)

            if plat_2025 is not None:
                plat_compare = pd.merge(plat_grouped, plat_2025, on="平台", how="outer").fillna(0)
                plat_compare = plat_compare.sort_values("平台", key=lambda s: s.map(_platform_sort_key))

                # 组合图：数量柱状图（对比时只显示数量）
                fig_plat = go.Figure()

                bar_cols = [c for c in ["销售数量", "销售数量_2025"] if c in plat_compare.columns]
                bar_colors = {"销售数量": "#636EFA", "销售数量_2025": "#B4A2F9"}
                bar_labels = {"销售数量": "数量(2026)", "销售数量_2025": "数量(2025)"}

                for col in bar_cols:
                    fig_plat.add_trace(go.Bar(
                        x=plat_compare["平台"],
                        y=plat_compare[col],
                        name=bar_labels.get(col, col),
                        marker_color=bar_colors.get(col),
                        text=plat_compare[col].apply(lambda v: f"{v:,.0f}"),
                        textposition="outside",
                    ))

                fig_plat.update_layout(
                    height=380,
                    margin=dict(l=60, r=60, t=40, b=10),
                    yaxis=dict(title="销售数量"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    barmode="group",
                )
                st.plotly_chart(fig_plat, width="stretch")
            else:
                col_pie, col_bar = st.columns(2)
                with col_pie:
                    fig_pie = px.pie(
                        plat_grouped,
                        values="销售数量",
                        names="平台",
                        hole=0.35,
                        title="各平台数量占比",
                    )
                    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                    st.plotly_chart(fig_pie, width="stretch")
                with col_bar:
                    # 组合图：数量柱状图 + 金额折线图
                    plat_sorted = plat_grouped.sort_values("平台", key=lambda s: s.map(_platform_sort_key))

                    fig_plat = go.Figure()

                    fig_plat.add_trace(go.Bar(
                        x=plat_sorted["平台"],
                        y=plat_sorted["销售数量"],
                        name="销售数量",
                        marker_color="#636EFA",
                        text=plat_sorted["销售数量"].apply(lambda v: f"{v:,.0f}"),
                        textposition="outside",
                    ))

                    if "销售成本" in plat_sorted.columns:
                        fig_plat.add_trace(go.Scatter(
                            x=plat_sorted["平台"],
                            y=plat_sorted["销售成本"],
                            name="销售成本",
                            mode="lines+markers",
                            line=dict(color="#EF553B", width=2),
                            marker=dict(size=6),
                            yaxis="y2",
                            text=[f"{v/1000000:,.2f}M" for v in plat_sorted["销售成本"]],
                            textposition="top center",
                            hovertemplate="%{text}<extra></extra>",
                        ))

                    fig_plat.update_layout(
                        height=380,
                        margin=dict(l=60, r=60, t=40, b=10),
                        yaxis=dict(title="销售数量"),
                        yaxis2=dict(title="销售成本 (百万)", overlaying="y", side="right", showgrid=False),
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    )
                    st.plotly_chart(fig_plat, width="stretch")

            # 平台损坏产品分布
            st.markdown("---")
            st.subheader("📊 平台损坏产品分布")

            if not df_damage.empty:
                damage_abs = df_damage.copy()
                damage_abs["销售数量"] = damage_abs["销售数量"].abs()

                # 判断日期范围是否跨多个月
                is_multi_month = False
                if date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
                    start_month = pd.Timestamp(date_range[0]).to_period("M")
                    end_month = pd.Timestamp(date_range[1]).to_period("M")
                    is_multi_month = end_month > start_month

                if is_multi_month:
                    # 跨月：按月+平台堆叠柱状图
                    damage_abs = damage_abs.assign(
                        _月份=pd.to_datetime(damage_abs["日期"], errors="coerce").dt.to_period("M").astype(str)
                    )
                    month_platform = damage_abs.groupby(["_月份", "平台"], as_index=False).agg(
                        损坏数量=("销售数量", "sum"),
                    )
                    pivot = month_platform.pivot(
                        index="_月份", columns="平台", values="损坏数量"
                    ).fillna(0).sort_index()
                    pivot = pivot[sorted(pivot.columns, key=_platform_sort_key)]

                    fig_damage_stack = go.Figure()
                    for platform in pivot.columns:
                        fig_damage_stack.add_trace(go.Bar(
                            x=pivot.index,
                            y=pivot[platform],
                            name=platform,
                            text=pivot[platform].apply(lambda v: f"{v:,.0f}" if v > 0 else ""),
                            textposition="inside",
                        ))
                    fig_damage_stack.update_layout(
                        height=380,
                        margin=dict(l=60, r=60, t=40, b=10),
                        yaxis=dict(title="损坏数量"),
                        xaxis=dict(title="月份", type="category"),
                        barmode="stack",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    )
                    st.plotly_chart(fig_damage_stack, width="stretch")
                else:
                    # 单月：保持现有饼图+柱状图
                    plat_damage = damage_abs.groupby("平台", as_index=False).agg(
                        损坏数量=("销售数量", "sum"),
                    )
                    plat_damage = plat_damage.sort_values("损坏数量", ascending=False)
                    plat_damage = plat_damage.sort_values("平台", key=lambda s: s.map(_platform_sort_key))

                    col_pie_dmg, col_bar_dmg = st.columns(2)
                    with col_pie_dmg:
                        fig_damage_pie = px.pie(
                            plat_damage,
                            values="损坏数量",
                            names="平台",
                            hole=0.35,
                            title="各平台损坏占比",
                        )
                        fig_damage_pie.update_traces(textposition="inside", textinfo="percent+label")
                        st.plotly_chart(fig_damage_pie, width="stretch")
                    with col_bar_dmg:
                        fig_damage_bar = go.Figure()
                        fig_damage_bar.add_trace(go.Bar(
                            x=plat_damage["平台"],
                            y=plat_damage["损坏数量"],
                            name="损坏数量",
                            marker_color="#EF553B",
                            text=plat_damage["损坏数量"].apply(lambda v: f"{v:,.0f}"),
                            textposition="outside",
                        ))
                        fig_damage_bar.update_layout(
                            height=380,
                            margin=dict(l=60, r=60, t=40, b=10),
                            yaxis=dict(title="损坏数量"),
                        )
                        st.plotly_chart(fig_damage_bar, width="stretch")
            else:
                st.info("当前筛选条件下无损坏数据。")

            st.markdown("---")
            st.subheader("🏆 产品 Top 15（按销售数量）")

            top_prod = (
                df.groupby("产品", as_index=False)["销售数量"]
                .sum()
                .sort_values("销售数量", ascending=False)
                .head(15)
            )

            if show_2025 and not df_2025_filtered.empty:
                top_prod_2025 = (
                    df_2025_filtered.groupby("产品", as_index=False)["销售数量"]
                    .sum()
                    .sort_values("销售数量", ascending=False)
                )
                top_prod_2025.columns = ["产品", "销售数量_2025"]
                top_compare = pd.merge(top_prod, top_prod_2025, on="产品", how="left").fillna(0)

                if not top_compare.empty:
                    fig_top = px.bar(
                        top_compare,
                        x=["销售数量", "销售数量_2025"],
                        y="产品",
                        orientation="h",
                        barmode="group",
                        title="",
                    )
                    fig_top.for_each_trace(lambda t: t.update(name="2026年" if t.name == "销售数量" else "2025年"))
                    fig_top.update_layout(
                        height=480,
                        margin=dict(l=10, r=10, t=10, b=10),
                        yaxis={"categoryorder": "total ascending"},
                    )
                    fig_top.update_traces(
                        texttemplate="%{x:,.0f}",
                        textposition="outside",
                    )
                    st.plotly_chart(fig_top, width="stretch")
            elif not top_prod.empty:
                fig_top = px.bar(
                    top_prod,
                    x="销售数量",
                    y="产品",
                    orientation="h",
                    title="",
                )
                fig_top.update_layout(
                    height=480,
                    margin=dict(l=10, r=10, t=10, b=10),
                    yaxis={"categoryorder": "total ascending"},
                )
                fig_top.update_traces(
                    texttemplate="%{x:,.0f}",
                    textposition="outside",
                )
                st.plotly_chart(fig_top, width="stretch")

            st.markdown("---")
            st.subheader("📋 明细数据")

            display_cols = ["日期", "平台", "子平台", "产品", "销售数量"]
            if "销售成本" in df.columns:
                display_cols.append("销售成本")

            display = df[display_cols].copy()
            display["销售数量"] = display["销售数量"].round(0)
            if "销售成本" in display.columns:
                display["销售成本"] = display["销售成本"].round(2)

            if show_2025 and not df_2025_filtered.empty:
                display_2025_cols = ["日期", "平台", "子平台", "产品", "销售数量"]
                if "销售成本" in df_2025_filtered.columns:
                    display_2025_cols.append("销售成本")

                display_2025 = df_2025_filtered[display_2025_cols].copy()
                rename_map = {"销售数量": "销售数量_2025"}
                if "销售成本" in display_2025.columns:
                    rename_map["销售成本"] = "销售成本_2025"
                display_2025 = display_2025.rename(columns=rename_map)
                display_2025["日期"] = display_2025["日期"].str.replace("2025", "2026")
                display_2025["销售数量_2025"] = display_2025["销售数量_2025"].round(0)
                if "销售成本_2025" in display_2025.columns:
                    display_2025["销售成本_2025"] = display_2025["销售成本_2025"].round(2)

                display_compare = pd.merge(display, display_2025, on=["日期", "平台", "子平台", "产品"], how="outer").fillna(0)

                format_dict = {"销售数量": "{:,.0f}", "销售数量_2025": "{:,.0f}"}
                if "销售成本" in display_compare.columns:
                    format_dict["销售成本"] = "¥{:,.2f}"
                if "销售成本_2025" in display_compare.columns:
                    format_dict["销售成本_2025"] = "¥{:,.2f}"

                st.dataframe(
                    display_compare.style.format(format_dict),
                    width="stretch",
                    hide_index=True,
                )
            else:
                format_dict = {"销售数量": "{:,.0f}"}
                if "销售成本" in display.columns:
                    format_dict["销售成本"] = "¥{:,.2f}"
                st.dataframe(
                    display.style.format(format_dict),
                    width="stretch",
                    hide_index=True,
                )
        else:
            st.info("当前筛选下无有效日期数据。")
    else:
        st.info("当前筛选条件下无数据。")
