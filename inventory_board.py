"""常温销量看板：由 app.py 路由进入，读取 normal_sales_amount.csv。"""

from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

PLATFORM_ORDER = ["京东", "天猫", "拼多多", "抖音", "新零售", "多多买菜", "小程序及其他"]

DEPARTMENTS = ["常温"]

DEPARTMENT_CSV_MAP = {
    "常温": "normal_sales_amount.csv",
    "低温": "low_temp_sales_amount.csv",
    "奶粉": "milk_powder_sales_amount.csv",
    "八喜": "baxi_sales_amount.csv",
}

CSV_2025_MAP = {
    "常温": "normal_sales_amount_2025.csv",
}


def _sales_csv_path(department: str) -> Path:
    filename = DEPARTMENT_CSV_MAP.get(department, "normal_sales_amount.csv")
    return Path(__file__).resolve().parent / filename


def _sales_csv_path_2025(department: str) -> Path:
    filename = CSV_2025_MAP.get(department, "")
    if not filename:
        return Path()
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
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    for col in ("日期", "平台", "子平台", "产品"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    if "销售数量" in df.columns:
        df["销售数量"] = pd.to_numeric(df["销售数量"], errors="coerce").fillna(0.0)
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

    date_range = st.sidebar.date_input(
        "统计日期范围",
        value=(max_date, max_date),
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

    total_qty = float(df["销售数量"].sum()) if not df.empty else 0.0
    product_cnt = df["产品"].nunique() if not df.empty else 0
    total_qty_2025 = float(df_2025_filtered["销售数量"].sum()) if not df_2025_filtered.empty else 0.0
    yoy_growth = ((total_qty - total_qty_2025) / total_qty_2025 * 100) if total_qty_2025 > 0 else 0.0

    c1, c2, c3 = st.columns(3)
    c1.metric("销售数量合计", f"{total_qty:,.0f}")
    c2.metric("产品数", f"{product_cnt:,}")
    c3.metric(
        "同比增长",
        f"{yoy_growth:+.1f}%",
        delta=f"较2025年同期 {total_qty_2025:,.0f}",
        delta_color="inverse"
    )

    st.markdown("---")

    if not df.empty:
        tmp = df.assign(_日期=pd.to_datetime(df["日期"], errors="coerce")).dropna(subset=["_日期"])
        if not tmp.empty:
            st.subheader("📊 销售数量趋势（按日）")
            
            # 2026年数据
            daily_2026 = tmp.groupby(tmp["_日期"].dt.date, as_index=False)["销售数量"].sum()
            daily_2026.columns = ["日期", "销售数量"]
            daily_2026["日期"] = daily_2026["日期"].astype(str)
            
            # 准备对比数据
            if show_2025 and not df_2025_filtered.empty:
                tmp_2025 = df_2025_filtered.assign(_日期=pd.to_datetime(df_2025_filtered["日期"], errors="coerce")).dropna(subset=["_日期"])
                if not tmp_2025.empty:
                    daily_2025 = tmp_2025.groupby(tmp_2025["_日期"].dt.date, as_index=False)["销售数量"].sum()
                    daily_2025.columns = ["日期", "销售数量_2025"]
                    daily_2025["日期"] = daily_2025["日期"].astype(str)
                    # 转换为2026年日期以便对比显示
                    daily_2025["日期"] = daily_2025["日期"].str.replace("2025", "2026")
                    daily = pd.merge(daily_2026, daily_2025, on="日期", how="outer").fillna(0)
                    daily = daily.sort_values("日期")
                    
                    fig_trend = px.bar(
                        daily,
                        x="日期",
                        y=["销售数量", "销售数量_2025"],
                        labels={"value": "销售数量", "variable": "年份"},
                        barmode="group",
                    )
                    fig_trend.for_each_trace(lambda t: t.update(name="2026年" if t.name == "销售数量" else "2025年"))
                else:
                    daily = daily_2026
                    fig_trend = px.bar(daily, x="日期", y="销售数量", labels={"销售数量": "销售数量"})
            else:
                daily = daily_2026
                fig_trend = px.bar(daily, x="日期", y="销售数量", labels={"销售数量": "销售数量"})
            
            fig_trend.update_layout(
                height=380,
                margin=dict(l=10, r=10, t=10, b=10),
                xaxis_title="日期",
                yaxis_title="销售数量",
                xaxis=dict(
                    type="category",
                ),
            )
            fig_trend.update_traces(
                texttemplate="%{y:,.0f}",
                textposition="outside",
            )
            st.plotly_chart(fig_trend, width="stretch")

            st.markdown("---")
            st.subheader("📊 平台销售分布")
            plat_sum = (
                df.groupby("平台", as_index=False)["销售数量"]
                .sum()
                .sort_values("销售数量", ascending=False)
            )
            
            if show_2025 and not df_2025_filtered.empty:
                plat_sum_2025 = (
                    df_2025_filtered.groupby("平台", as_index=False)["销售数量"]
                    .sum()
                    .sort_values("销售数量", ascending=False)
                )
                plat_sum_2025.columns = ["平台", "销售数量_2025"]
                plat_compare = pd.merge(plat_sum, plat_sum_2025, on="平台", how="outer").fillna(0)
                plat_compare = plat_compare.sort_values("平台", key=lambda s: s.map(_platform_sort_key))
                
                col_bar = st.columns(1)
                with col_bar[0]:
                    fig_plat = px.bar(
                        plat_compare,
                        x="平台",
                        y=["销售数量", "销售数量_2025"],
                        barmode="group",
                        title="各平台销售数量对比",
                    )
                    fig_plat.for_each_trace(lambda t: t.update(name="2026年" if t.name == "销售数量" else "2025年"))
                    fig_plat.update_layout(height=380, margin=dict(l=10, r=10, t=40, b=10))
                    fig_plat.update_traces(
                        texttemplate="%{y:,.0f}",
                        textposition="outside",
                    )
                    st.plotly_chart(fig_plat, width="stretch")
            else:
                col_pie, col_bar = st.columns(2)
                with col_pie:
                    fig_pie = px.pie(
                        plat_sum,
                        values="销售数量",
                        names="平台",
                        hole=0.35,
                        title="各平台占比",
                    )
                    fig_pie.update_traces(textposition="inside", textinfo="percent+label")
                    st.plotly_chart(fig_pie, width="stretch")
                with col_bar:
                    fig_plat = px.bar(
                        plat_sum.sort_values("平台", key=lambda s: s.map(_platform_sort_key)),
                        x="平台",
                        y="销售数量",
                        title="各平台销售数量",
                    )
                    fig_plat.update_layout(height=380, margin=dict(l=10, r=10, t=40, b=10))
                    fig_plat.update_traces(
                        texttemplate="%{y:,.0f}",
                        textposition="outside",
                    )
                    st.plotly_chart(fig_plat, width="stretch")

            st.markdown("---")
            st.subheader("🏆 产品 Top 15")
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
                st.plotly_chart(fig_top, width="stretch")

            st.markdown("---")
            st.subheader("📋 明细数据")
            display = df[["日期", "平台", "子平台", "产品", "销售数量"]].copy()
            display["销售数量"] = display["销售数量"].round(0)
            
            if show_2025 and not df_2025_filtered.empty:
                display_2025 = df_2025_filtered[["日期", "平台", "子平台", "产品", "销售数量"]].copy()
                display_2025.columns = ["日期", "平台", "子平台", "产品", "销售数量_2025"]
                display_2025["日期"] = display_2025["日期"].str.replace("2025", "2026")
                display_2025["销售数量_2025"] = display_2025["销售数量_2025"].round(0)
                display_compare = pd.merge(display, display_2025, on=["日期", "平台", "子平台", "产品"], how="outer").fillna(0)
                display_compare["同比增长"] = display_compare.apply(
                    lambda row: ((row["销售数量"] - row["销售数量_2025"]) / row["销售数量_2025"] * 100) 
                    if row["销售数量_2025"] > 0 else 0,
                    axis=1
                )
                st.dataframe(
                    display_compare.style.format({"销售数量": "{:,.0f}", "销售数量_2025": "{:,.0f}", "同比增长": "{:+.1f}%"}),
                    width="stretch",
                    hide_index=True,
                )
            else:
                st.dataframe(
                    display.style.format({"销售数量": "{:,.0f}"}),
                    width="stretch",
                    hide_index=True,
                )
        else:
            st.info("当前筛选下无有效日期数据。")
    else:
        st.info("当前筛选条件下无数据。")
