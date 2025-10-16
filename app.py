# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ==================== 页面基础设置 ====================
st.set_page_config(page_title="客服时效分析报告", layout="wide")

# -------- 页面 CSS 样式 --------
st.markdown("""
<style>
    .main { background-color: #F5F6FA; }
    h1 { color: #2B3A67; text-align: center; padding: 0.5rem 0; border-bottom: 3px solid #5B8FF9; }
    h2, h3 { color: #2B3A67; margin-top: 1.5rem; }
    .stDataFrame { background-color: white; border-radius: 12px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
    div.stButton > button:first-child {
        background-color: #5B8FF9; color: white; border: none; border-radius: 8px;
        padding: 0.4rem 1.2rem;
    }
    div.stButton > button:hover { background-color: #3A6CE5; color: white; }
</style>
""", unsafe_allow_html=True)

st.title("客服时效分析报告")

# ==================== 上传文件 ====================
uploaded_files = st.file_uploader(
    "📂 上传一个或多个数据文件（支持 Excel 或 CSV）",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

if uploaded_files:
    all_dfs = []
    for i, file in enumerate(uploaded_files, 1):
        try:
            st.write(f"📘 正在读取第 {i} 个文件：**{file.name}** ...")
            if file.name.endswith("xlsx"):
                df = pd.read_excel(file)
            else:
                df = pd.read_csv(file)

            # 去掉最后一行、空行并重置索引
            df = df.iloc[:-1, :].dropna(how="all").reset_index(drop=True)
            all_dfs.append(df)
        except Exception as e:
            st.warning(f"⚠️ 文件 {file.name} 读取失败：{e}")

    # 合并所有文件
    if all_dfs:
        df = pd.concat(all_dfs, ignore_index=True)
        st.success(f"✅ 共成功导入 {len(all_dfs)} 个文件，合并后共 {len(df)} 行数据")
        st.dataframe(df.head(), use_container_width=True)
    else:
        st.error("❌ 未能成功读取任何文件，请检查格式")
        st.stop()

    # === 列名清理 ===
    df.columns = df.columns.astype(str).str.strip()

    # === 时间列识别 ===
    created_col = next((c for c in df.columns if "ticket_created" in c.lower()), None)
    if created_col is None:
        st.error("❌ 未找到创建时间列（应包含 ticket_created 关键字）")
        st.stop()

    df["ticket_created_datetime"] = pd.to_datetime(df[created_col], errors="coerce")
    df["month"] = df["ticket_created_datetime"].dt.to_period("M").astype(str)

    # === 清洗 "-" 空值等 ===
    def clean_numeric_column(s: pd.Series) -> pd.Series:
        s = s.astype(str).str.strip()
        s = s.replace(
            {r"^[-‐-‒–—―−]+$": None, r"^(null|None|nan|NaN)$": None, r"^\s*$": None},
            regex=True,
        )
        s = s.str.replace(",", "", regex=False)
        return pd.to_numeric(s, errors="coerce")

    for col in ["首次响应时长", "处理时长", "message_count"]:
        if col in df.columns:
            df[col] = clean_numeric_column(df[col])

    # ==================== 🎯 渠道筛选入口 ====================
    if "ticket_channel" in df.columns:
        all_channels = sorted(df["ticket_channel"].dropna().unique().tolist())
        selected_channels = st.multiselect(
            "💌请选择要分析的渠道（可多选）",
            options=all_channels,
            default=all_channels,
        )
        if selected_channels:
            df = df[df["ticket_channel"].isin(selected_channels)]
            st.info(f"当前筛选渠道：{', '.join(selected_channels)}，共 {len(df)} 条记录")
        else:
            st.warning("⚠️ 未选择任何渠道，将不显示后续分析结果。")
            st.stop()
    else:
        st.warning("⚠️ 数据中未找到渠道字段（ticket_channel），跳过渠道筛选。")

    # === 子集 ===
    df_reply = df.query("rn == 1")
    df_close = df.query("rn == 1")

    # ==================== Ⅰ. 整体分析 ====================
    st.header("📅 每月整体表现")

    reply_stats = df_reply.groupby("month", as_index=False).agg(
        message_count_median=("message_count", "median"),
        message_count_p90=("message_count", lambda x: x.quantile(0.9)),
    )
    df_resp = df_close[df_close["首次响应时长"].notna()]
    resp_stats = df_resp.groupby("month", as_index=False).agg(
        response_median=("首次响应时长", "median"),
        response_p90=("首次响应时长", lambda x: x.quantile(0.9)),
    )
    df_handle = df_close[df_close["处理时长"].notna()]
    handle_stats = df_handle.groupby("month", as_index=False).agg(
        handle_median=("处理时长", "median"),
        handle_p90=("处理时长", lambda x: x.quantile(0.9)),
    )

    overall = (
        reply_stats.merge(resp_stats, on="month", how="outer")
        .merge(handle_stats, on="month", how="outer")
        .sort_values("month")
    )

    overall = overall.rename(columns={
        "month": "月份",
        "message_count_median": "回复次数-中位数",
        "message_count_p90": "回复次数-P90",
        "response_median": "首次响应时长h-中位数",
        "response_p90": "首次响应时长h-P90",
        "handle_median": "处理时长d-中位数",
        "handle_p90": "处理时长d-P90",
    })
    st.dataframe(overall, use_container_width=True)

    metric_all = st.selectbox("请选择要查看的整体指标", ["回复次数-P90", "首次响应时长h-P90", "处理时长d-P90"], index=2)
    if overall[metric_all].notna().any():
        fig = px.line(
            overall,
            x="月份", y=metric_all,
            title=f"整体 {metric_all} 趋势",
            markers=True,
            line_shape="spline",
            color_discrete_sequence=["#5B8FF9"],
        )
        fig.update_traces(marker=dict(size=8, opacity=0.8))
        st.plotly_chart(fig, use_container_width=True)

    # ==================== Ⅱ. 品牌线分析 ====================
    if "business_line" in df.columns:
        st.header("🏷️ 各贸易品牌线表现")

        reply_line = df_reply.groupby(["month", "business_line"], as_index=False).agg(
            message_count_median=("message_count", "median"),
            message_count_p90=("message_count", lambda x: x.quantile(0.9)),
        )
        df_resp_line = df_close[df_close["首次响应时长"].notna()]
        resp_line = df_resp_line.groupby(["month", "business_line"], as_index=False).agg(
            response_median=("首次响应时长", "median"),
            response_p90=("首次响应时长", lambda x: x.quantile(0.9)),
        )
        df_handle_line = df_close[df_close["处理时长"].notna()]
        handle_line = df_handle_line.groupby(["month", "business_line"], as_index=False).agg(
            handle_median=("处理时长", "median"),
            handle_p90=("处理时长", lambda x: x.quantile(0.9)),
        )

        line_stats = (
            reply_line.merge(resp_line, on=["month", "business_line"], how="outer")
            .merge(handle_line, on=["month", "business_line"], how="outer")
            .sort_values(["month", "business_line"])
        )
        line_stats = line_stats.rename(columns={
            "month": "月份", "business_line": "品牌线",
            "message_count_median": "回复次数-中位数",
            "message_count_p90": "回复次数-P90",
            "response_median": "首次响应时长h-中位数",
            "response_p90": "首次响应时长h-P90",
            "handle_median": "处理时长d-中位数",
            "handle_p90": "处理时长d-P90",
        })
        st.dataframe(line_stats, use_container_width=True)

        metric_line = st.selectbox("请选择要查看的品牌线指标", ["回复次数-P90", "首次响应时长h-P90", "处理时长d-P90"], index=2)
        if line_stats[metric_line].notna().any():
            fig = px.line(
                line_stats,
                x="月份", y=metric_line, color="品牌线",
                title=f"各品牌线 {metric_line} 趋势",
                markers=True,
                line_shape="spline",
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig.update_traces(marker=dict(size=7, opacity=0.8))
            st.plotly_chart(fig, use_container_width=True)

    # ==================== Ⅲ. 国家分析 ====================
    if "site_code" in df.columns:
        st.header("🌍 各国家表现")

        reply_site = df_reply.groupby(["month", "site_code"], as_index=False).agg(
            message_count_median=("message_count", "median"),
            message_count_p90=("message_count", lambda x: x.quantile(0.9)),
        )
        df_resp_site = df_close[df_close["首次响应时长"].notna()]
        resp_site = df_resp_site.groupby(["month", "site_code"], as_index=False).agg(
            response_median=("首次响应时长", "median"),
            response_p90=("首次响应时长", lambda x: x.quantile(0.9)),
        )
        df_handle_site = df_close[df_close["处理时长"].notna()]
        handle_site = df_handle_site.groupby(["month", "site_code"], as_index=False).agg(
            handle_median=("处理时长", "median"),
            handle_p90=("处理时长", lambda x: x.quantile(0.9)),
        )

        site_stats = (
            reply_site.merge(resp_site, on=["month", "site_code"], how="outer")
            .merge(handle_site, on=["month", "site_code"], how="outer")
            .sort_values(["month", "site_code"])
        )
        site_stats = site_stats.rename(columns={
            "month": "月份", "site_code": "国家",
            "message_count_median": "回复次数-中位数",
            "message_count_p90": "回复次数-P90",
            "response_median": "首次响应时长h-中位数",
            "response_p90": "首次响应时长h-P90",
            "handle_median": "处理时长d-中位数",
            "handle_p90": "处理时长d-P90",
        })
        st.dataframe(site_stats, use_container_width=True)

        metric_site = st.selectbox("请选择要查看的国家指标", ["回复次数-P90", "首次响应时长h-P90", "处理时长d-P90"], index=2)
        if site_stats[metric_site].notna().any():
            fig = px.line(
                site_stats,
                x="月份", y=metric_site, color="国家",
                title=f"各国家 {metric_site} 趋势",
                markers=True,
                line_shape="spline",
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig.update_traces(marker=dict(size=7, opacity=0.8))
            st.plotly_chart(fig, use_container_width=True)
    # ==================== Ⅲ. 渠道分析 ====================
    if "ticket_channel" in df.columns:
        st.header("💬 各渠道表现")

        reply_channel = df_reply.groupby(["month", "ticket_channel"], as_index=False).agg(
            message_count_median=("message_count", "median"),
            message_count_p90=("message_count", lambda x: x.quantile(0.9)),
        )
        df_resp_channel = df_close[df_close["首次响应时长"].notna()]
        resp_channel = df_resp_channel.groupby(["month", "ticket_channel"], as_index=False).agg(
            response_median=("首次响应时长", "median"),
            response_p90=("首次响应时长", lambda x: x.quantile(0.9)),
        )
        df_handle_channel = df_close[df_close["处理时长"].notna()]
        handle_channel = df_handle_channel.groupby(["month", "ticket_channel"], as_index=False).agg(
            handle_median=("处理时长", "median"),
            handle_p90=("处理时长", lambda x: x.quantile(0.9)),
        )

        channel_stats = (
            reply_channel.merge(resp_channel, on=["month", "ticket_channel"], how="outer")
            .merge(handle_channel, on=["month", "ticket_channel"], how="outer")
            .sort_values(["month", "ticket_channel"])
        )

        channel_stats = channel_stats.rename(columns={
            "month": "月份", "ticket_channel": "渠道",
            "message_count_median": "回复次数-中位数",
            "message_count_p90": "回复次数-P90",
            "response_median": "首次响应时长h-中位数",
            "response_p90": "首次响应时长h-P90",
            "handle_median": "处理时长d-中位数",
            "handle_p90": "处理时长d-P90",
        })

        st.dataframe(channel_stats, use_container_width=True)

        metric_channel = st.selectbox("请选择要查看的渠道指标", ["回复次数-P90", "首次响应时长h-P90", "处理时长d-P90"], index=2)
        if channel_stats[metric_channel].notna().any():
            fig = px.line(
                channel_stats,
                x="月份", y=metric_channel, color="渠道",
                title=f"各渠道 {metric_channel} 趋势",
                markers=True,
                line_shape="spline",
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig.update_traces(marker=dict(size=7, opacity=0.8))
            st.plotly_chart(fig, use_container_width=True)

    # ==================== 📤 导出 Excel 报告 ====================
    st.header("📤 导出分析报告")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        overall.to_excel(writer, index=False, sheet_name="整体表现")
        if "business_line" in df.columns:
            line_stats.to_excel(writer, index=False, sheet_name="品牌线表现")
        if "site_code" in df.columns:
            site_stats.to_excel(writer, index=False, sheet_name="国家表现")
        if "ticket_channel" in df.columns:
            channel_stats.to_excel(writer, index=False, sheet_name="渠道表现")
    buffer.seek(0)

    st.download_button(
        label="📥 下载完整分析报告（Excel）",
        data=buffer,
        file_name="客服时效分析报告.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.success("✅ 报告生成完毕，可在页面或导出文件中查看。")
