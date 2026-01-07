# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

# ==================== 页面基础设置 ====================
st.set_page_config(page_title="客服时效分析报告", layout="wide")

st.markdown("""
<style>
    .main { background-color: #F5F6FA; }
    h1 { color: #2B3A67; text-align: center; padding: 0.5rem 0; border-bottom: 3px solid #5B8FF9; }
    h2, h3 { color: #2B3A67; margin-top: 1.5rem; }
    .stDataFrame { background-color: white; border-radius: 12px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
</style>
""", unsafe_allow_html=True)

st.title("客服时效分析报告")

# ==================== 上传文件 ====================
uploaded_files = st.file_uploader(
    "📂 上传一个或多个数据文件（支持 Excel / CSV）",
    type=["xlsx", "csv"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.stop()

# ==================== 读取 & 合并 ====================
dfs = []
for f in uploaded_files:
    df_tmp = pd.read_excel(f) if f.name.endswith("xlsx") else pd.read_csv(f)
    df_tmp = df_tmp.iloc[:-1, :].dropna(how="all")
    dfs.append(df_tmp)

df = pd.concat(dfs, ignore_index=True)
df.columns = df.columns.astype(str).str.strip()

# ==================== 时间字段 ====================
created_col = next(c for c in df.columns if "ticket_created" in c.lower())
df["ticket_created_datetime"] = pd.to_datetime(df[created_col], errors="coerce")
df["month"] = df["ticket_created_datetime"].dt.to_period("M").astype(str)
df["year"] = df["ticket_created_datetime"].dt.year

# ==================== 数值清洗 ====================
def clean_numeric(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False),
        errors="coerce"
    )

for c in ["message_count", "首次响应时长", "处理时长"]:
    if c in df.columns:
        df[c] = clean_numeric(df[c])

# =====================================================
# 🔒【全局口径锁死】只分析 closed 工单
# =====================================================
df_base = df[df["ticket_status"] == "closed"].copy()

# =====================================================
# 🔁 各分析口径去重数据集
# =====================================================
# 整体（ticket 级）
df_all = df_base.drop_duplicates(subset=["ticket_id"])

# 品牌线
df_bl = (
    df_base.drop_duplicates(subset=["ticket_id", "business_line"])
    if "business_line" in df_base.columns else None
)

# 国家
df_site = (
    df_base.drop_duplicates(subset=["ticket_id", "site_code"])
    if "site_code" in df_base.columns else None
)

# 渠道
df_channel = (
    df_base.drop_duplicates(subset=["ticket_id", "ticket_channel"])
    if "ticket_channel" in df_base.columns else None
)

# 一级问题
df_cls1 = (
    df_base.drop_duplicates(subset=["ticket_id", "class_one"])
    if "class_one" in df_base.columns else None
)

# 二级问题
df_cls2 = (
    df_base.drop_duplicates(subset=["ticket_id", "class_one", "class_two"])
    if {"class_one", "class_two"}.issubset(df_base.columns) else None
)

# ==================== 公共函数：环比 ====================
def add_mom(df, group_cols=None):
    out = df.copy()
    metrics = [c for c in out.columns if any(k in c for k in ["回复次数", "响应时长", "处理时长"])]
    for m in metrics:
        if group_cols:
            out[f"{m}-环比"] = out.groupby(group_cols)[m].pct_change().apply(
                lambda x: f"{x*100:.1f}%" if pd.notnull(x) else "-"
            )
        else:
            out[f"{m}-环比"] = out[m].pct_change().apply(
                lambda x: f"{x*100:.1f}%" if pd.notnull(x) else "-"
            )
    return out

# =====================================================
# Ⅰ. 每月整体表现（先整体去重）
# =====================================================
st.header("📅 每月整体表现")

reply_m = df_all.groupby("month", as_index=False).agg(
    工单量=("ticket_id", "count"), 
    回复次数_平均数=("message_count", "mean"),
    回复次数_中位数=("message_count", "median"),
    回复次数_P90=("message_count", lambda x: x.quantile(0.9)),
)

resp_m = df_all.groupby("month", as_index=False).agg(
    首次响应时长h_中位数=("首次响应时长", "median"),
    首次响应时长h_P90=("首次响应时长", lambda x: x.quantile(0.9)),
)

handle_m = df_all.groupby("month", as_index=False).agg(
    处理时长d_中位数=("处理时长", "median"),
    处理时长d_P90=("处理时长", lambda x: x.quantile(0.9)),
)

overall_m = (
    reply_m.merge(resp_m, on="month")
           .merge(handle_m, on="month")
           .rename(columns={"month": "月份"})
           .sort_values("月份")
)

overall_m = add_mom(overall_m)
st.dataframe(overall_m, use_container_width=True)

# =====================================================
# Ⅰ-2. 每年整体表现
# =====================================================
st.header("📆 每年整体表现")

reply_y = df_all.groupby("year", as_index=False).agg(
    工单量=("ticket_id", "count"), 
    回复次数_平均数=("message_count", "mean"),
    回复次数_中位数=("message_count", "median"),
    回复次数_P90=("message_count", lambda x: x.quantile(0.9)),
)

resp_y = df_all.groupby("year", as_index=False).agg(
    首次响应时长h_中位数=("首次响应时长", "median"),
    首次响应时长h_P90=("首次响应时长", lambda x: x.quantile(0.9)),
)

handle_y = df_all.groupby("year", as_index=False).agg(
    处理时长d_中位数=("处理时长", "median"),
    处理时长d_P90=("处理时长", lambda x: x.quantile(0.9)),
)

overall_y = (
    reply_y.merge(resp_y, on="year")
           .merge(handle_y, on="year")
           .rename(columns={"year": "年份"})
           .sort_values("年份")
)

overall_y = add_mom(overall_y)
st.dataframe(overall_y, use_container_width=True)

# =====================================================
# Ⅱ. 品牌线
# =====================================================
if df_bl is not None:
    st.header("🏷️ 品牌线表现")
    bl_stats = (
        df_bl.groupby(["month", "business_line"], as_index=False)
        .agg(
            回复次数_P90=("message_count", lambda x: x.quantile(0.9)),
            首次响应时长h_P90=("首次响应时长", lambda x: x.quantile(0.9)),
            处理时长d_P90=("处理时长", lambda x: x.quantile(0.9)),
        )
        .rename(columns={"month": "月份", "business_line": "品牌线"})
        .sort_values(["月份", "品牌线"])
    )
    bl_stats = add_mom(bl_stats, ["品牌线"])
    st.dataframe(bl_stats, use_container_width=True)

# =====================================================
# Ⅲ. 国家
# =====================================================
if df_site is not None:
    st.header("🌍 国家表现")
    site_stats = (
        df_site.groupby(["month", "site_code"], as_index=False)
        .agg(
            回复次数_P90=("message_count", lambda x: x.quantile(0.9)),
            首次响应时长h_P90=("首次响应时长", lambda x: x.quantile(0.9)),
            处理时长d_P90=("处理时长", lambda x: x.quantile(0.9)),
        )
        .rename(columns={"month": "月份", "site_code": "国家"})
        .sort_values(["月份", "国家"])
    )
    site_stats = add_mom(site_stats, ["国家"])
    st.dataframe(site_stats, use_container_width=True)

# =====================================================
# Ⅳ. 渠道
# =====================================================
if df_channel is not None:
    st.header("💬 渠道表现")
    ch_stats = (
        df_channel.groupby(["month", "ticket_channel"], as_index=False)
        .agg(
            回复次数_P90=("message_count", lambda x: x.quantile(0.9)),
            首次响应时长h_P90=("首次响应时长", lambda x: x.quantile(0.9)),
            处理时长d_P90=("处理时长", lambda x: x.quantile(0.9)),
        )
        .rename(columns={"month": "月份", "ticket_channel": "渠道"})
        .sort_values(["月份", "渠道"])
    )
    ch_stats = add_mom(ch_stats, ["渠道"])
    st.dataframe(ch_stats, use_container_width=True)

# =====================================================
# Ⅴ. 问题分类（年）
# =====================================================
st.header("🧩 问题分类年均回复次数分析")

if df_cls1 is not None:
    st.subheader("① 一级问题")
    class_one_stats = (
        df_cls1.groupby(["year", "class_one"], as_index=False)
        .agg(
            回复次数_平均数=("message_count", "mean"),
            回复次数_中位数=("message_count", "median"),
            回复次数_P90=("message_count", lambda x: x.quantile(0.9)),
            工单量=("ticket_id", "count"),
        )
        .sort_values(["year", "回复次数_P90"], ascending=[True, False])
    )
    st.dataframe(class_one_stats, use_container_width=True)

if df_cls2 is not None:
    st.subheader("② 二级问题（含一级）")
    class_two_stats = (
        df_cls2.groupby(["year", "class_one", "class_two"], as_index=False)
        .agg(
            回复次数_平均数=("message_count", "mean"),
            回复次数_中位数=("message_count", "median"),
            回复次数_P90=("message_count", lambda x: x.quantile(0.9)),
            工单量=("ticket_id", "count"),
        )
        .sort_values(["year", "class_one", "回复次数_P90"], ascending=[True, True, False])
    )
    st.dataframe(class_two_stats, use_container_width=True)

# =====================================================
# 📤 导出 Excel
# =====================================================
st.header("📤 导出分析报告")

buffer = BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    overall_m.to_excel(writer, index=False, sheet_name="每月整体表现")
    overall_y.to_excel(writer, index=False, sheet_name="每年整体表现")
    if df_bl is not None:
        bl_stats.to_excel(writer, index=False, sheet_name="品牌线表现")
    if df_site is not None:
        site_stats.to_excel(writer, index=False, sheet_name="国家表现")
    if df_channel is not None:
        ch_stats.to_excel(writer, index=False, sheet_name="渠道表现")
    if df_cls1 is not None:
        class_one_stats.to_excel(writer, index=False, sheet_name="一级问题_年统计")
    if df_cls2 is not None:
        class_two_stats.to_excel(writer, index=False, sheet_name="二级问题_年统计")

buffer.seek(0)

st.download_button(
    "📥 下载完整分析报告（Excel）",
    data=buffer,
    file_name="客服时效分析报告.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("✅ 报告生成完毕")
