import streamlit as st
import pandas as pd
import io

# ===================== 核心解析逻辑 =====================

def extract_title_from_excel(df):
    """
    从Excel的B1/C1合并单元格中提取文章标题
    """
    try:
        if len(df.columns) > 1 and pd.notna(df.iloc[0, 1]):
            title = str(df.iloc[0, 1]).strip()
            if title:
                return title
        if len(df.columns) > 2 and pd.notna(df.iloc[0, 2]):
            title = str(df.iloc[0, 2]).strip()
            if title:
                return title
    except Exception:
        pass
    return "未知标题"


def find_row_by_keyword(df, keyword, start_row=0, col_index=1):
    """
    在指定列中查找包含关键词的行
    返回: 行索引，如果没找到返回None
    """
    for idx in range(start_row, len(df)):
        if len(df.columns) > col_index and pd.notna(df.iloc[idx, col_index]):
            if keyword in str(df.iloc[idx, col_index]):
                return idx
    return None


def parse_gender_distribution(df, start_row):
    """
    解析性别分布，按表格内原值输出，不做转换
    """
    gender_data = {
        "性别_女": 0,
        "性别_男": 0,
        "性别_未知": 0
    }

    if start_row is None:
        return gender_data

    for idx in range(start_row, min(start_row + 10, len(df))):
        if len(df.columns) < 3:
            break

        b_col_val = df.iloc[idx, 1] if pd.notna(df.iloc[idx, 1]) else ""
        c_col_val = df.iloc[idx, 2] if len(df.columns) > 2 and pd.notna(df.iloc[idx, 2]) else 0

        if pd.isna(b_col_val):
            continue

        b_col_str = str(b_col_val).strip()

        if "女" in b_col_str and "男" not in b_col_str:
            gender_data["性别_女"] = c_col_val
        elif "男" in b_col_str:
            gender_data["性别_男"] = c_col_val
        elif "未知" in b_col_str:
            gender_data["性别_未知"] = c_col_val

        if b_col_str == "" and idx > start_row:
            break

    return gender_data


def parse_age_distribution(df, start_row):
    """
    解析年龄分布，按表格内原值输出，不做转换
    分类：18岁以下、18-25岁、26-35岁、36-45岁、46-60岁、60岁以上、未知
    """
    age_data = {
        "年龄_18岁以下": 0,
        "年龄_18-25岁": 0,
        "年龄_26-35岁": 0,
        "年龄_36-45岁": 0,
        "年龄_46-60岁": 0,
        "年龄_60岁以上": 0,
        "年龄_未知": 0
    }

    if start_row is None:
        return age_data

    for idx in range(start_row, min(start_row + 15, len(df))):
        if len(df.columns) < 3:
            break

        b_col_val = df.iloc[idx, 1] if pd.notna(df.iloc[idx, 1]) else ""
        c_col_val = df.iloc[idx, 2] if len(df.columns) > 2 and pd.notna(df.iloc[idx, 2]) else 0

        if pd.isna(b_col_val):
            continue

        b_col_str = str(b_col_val).strip()

        if "未知" in b_col_str and "年龄" not in b_col_str:
            age_data["年龄_未知"] = c_col_val
        elif "0-17" in b_col_str or "0~17" in b_col_str or "18岁以下" in b_col_str or "18以下" in b_col_str:
            age_data["年龄_18岁以下"] = c_col_val
        elif "18-25" in b_col_str or "18~25" in b_col_str:
            age_data["年龄_18-25岁"] = c_col_val
        elif "26-35" in b_col_str or "26~35" in b_col_str:
            age_data["年龄_26-35岁"] = c_col_val
        elif "36-45" in b_col_str or "36~45" in b_col_str:
            age_data["年龄_36-45岁"] = c_col_val
        elif "46-60" in b_col_str or "46~60" in b_col_str or "46-55" in b_col_str or "46~55" in b_col_str:
            age_data["年龄_46-60岁"] = c_col_val
        elif "56-65" in b_col_str or "56~65" in b_col_str or "60以上" in b_col_str or "60岁以上" in b_col_str or "65以上" in b_col_str:
            age_data["年龄_60岁以上"] = c_col_val

        if b_col_str == "" and idx > start_row + 2:
            break

    return age_data


def parse_wx_excel(file_bytes, filename):
    """
    解析单个公众号后台导出的Excel文件，返回结果字典和日志列表
    """
    logs = []
    logs.append(f"📄 正在处理：**{filename}**")

    try:
        if filename.endswith('.xls'):
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='xlrd')
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')
    except Exception as e:
        logs.append(f"❌ 文件读取失败：{e}")
        return None, logs

    logs.append(f"  文件维度：{len(df)}行 × {len(df.columns)}列")

    article_title = extract_title_from_excel(df)
    logs.append(f"  标题：{article_title}")

    result = {
        "文章标题": article_title,
        "全部阅读人数": 0,
        "送达人数": 0,
        "首次分享人数": 0,
        "分享产生的阅读人数": 0,
        "推荐阅读人数": 0,
        "其他阅读人数": 0,
        "公众号消息阅读人数": 0,
        "公众号主页阅读人数": 0,
        "搜一搜阅读人数": 0,
        "朋友圈阅读人数": 0,
        "聊天会话阅读人数": 0,
        "阅读后关注人数": 0,
        "点赞人数": 0,
        "评论次数": 0,
        "总分享人数": 0,
        "在看人数": 0,
        "收藏人数": 0,
        "停留时间": 0,
        "完读率": 0,
        # 性别分布
        "性别_女": 0,
        "性别_男": 0,
        "性别_未知": 0,
        # 年龄分布
        "年龄_18岁以下": 0,
        "年龄_18-25岁": 0,
        "年龄_26-35岁": 0,
        "年龄_36-45岁": 0,
        "年龄_46-60岁": 0,
        "年龄_60岁以上": 0,
        "年龄_未知": 0,
    }

    # ========== 1. 从"数据概况"区域提取指标 ==========
    data_overview_row = find_row_by_keyword(df, "数据概况", start_row=0, col_index=1)

    if data_overview_row is not None:
        logs.append(f"  ✅ 找到'数据概况'，起始行：{data_overview_row + 1}")
        start_idx = data_overview_row + 1
        metrics_mapping = {
            "阅读(人)": "全部阅读人数",
            "阅读人数": "全部阅读人数",
            "阅读后关注": "阅读后关注人数",
            "点赞": "点赞人数",
            "评论": "评论次数",
            "分享(人)": "总分享人数",
            "分享人数": "总分享人数",
            "在看": "在看人数",
            "收藏": "收藏人数",
            "平均停留时长": "停留时间",
            "完读率": "完读率",
        }
        matched_count = 0
        for idx in range(start_idx, min(start_idx + 30, len(df))):
            if len(df.columns) < 3:
                break
            b_col_val = df.iloc[idx, 1] if pd.notna(df.iloc[idx, 1]) else ""
            c_col_val = df.iloc[idx, 2] if len(df.columns) > 2 and pd.notna(df.iloc[idx, 2]) else 0
            if pd.isna(b_col_val):
                continue
            b_col_str = str(b_col_val).strip()
            if b_col_str == "":
                if matched_count > 0:
                    break
                continue
            matched = False
            for keyword, field_name in metrics_mapping.items():
                if keyword in b_col_str:
                    if field_name in ("停留时间", "完读率"):
                        result[field_name] = float(c_col_val) if pd.notna(c_col_val) else 0
                    else:
                        result[field_name] = int(c_col_val) if pd.notna(c_col_val) else 0
                    matched_count += 1
                    matched = True
                    break
            if not matched and ("阅读转化" in b_col_str or "数据趋势" in b_col_str):
                break
    else:
        logs.append("  ⚠️ 未找到'数据概况'区域")

    # ========== 2. 从"阅读转化"区域提取送达人数、首次分享人数、分享产生的阅读人数 ==========
    conversion_row = find_row_by_keyword(df, "阅读转化", start_row=0, col_index=1)

    if conversion_row is not None:
        logs.append(f"  ✅ 找到'阅读转化'，起始行：{conversion_row + 1}")
        start_idx = conversion_row + 1
        for idx in range(start_idx, min(start_idx + 20, len(df))):
            if len(df.columns) < 3:
                break
            b_col_val = df.iloc[idx, 1] if pd.notna(df.iloc[idx, 1]) else ""
            c_col_val = df.iloc[idx, 2] if len(df.columns) > 2 and pd.notna(df.iloc[idx, 2]) else 0
            if pd.isna(b_col_val):
                continue
            b_col_str = str(b_col_val).strip()
            if "送达人数" in b_col_str:
                result["送达人数"] = int(c_col_val) if pd.notna(c_col_val) else 0
            elif "首次分享人数" in b_col_str or "首次分享" in b_col_str:
                result["首次分享人数"] = int(c_col_val) if pd.notna(c_col_val) else 0
            elif "分享产生的阅读人数" in b_col_str or "分享产生阅读" in b_col_str:
                result["分享产生的阅读人数"] = int(c_col_val) if pd.notna(c_col_val) else 0
    else:
        logs.append("  ⚠️ 未找到'阅读转化'区域")

    # ========== 3. 从"数据趋势明细"区域提取各渠道阅读人数 ==========
    detail_row = find_row_by_keyword(df, "数据趋势明细", start_row=0, col_index=1)

    if detail_row is not None:
        logs.append(f"  ✅ 找到'数据趋势明细'，起始行：{detail_row + 1}")
        start_idx = detail_row + 2
        channel_mapping = {
            "推荐": "推荐阅读人数",
            "其他": "其他阅读人数",
            "公众号消息": "公众号消息阅读人数",
            "公众号主页": "公众号主页阅读人数",
            "搜一搜": "搜一搜阅读人数",
            "朋友圈": "朋友圈阅读人数",
            "聊天会话": "聊天会话阅读人数",
        }
        row_count = 0
        for idx in range(start_idx, len(df)):
            if len(df.columns) < 4:
                break
            date_val = df.iloc[idx, 1] if len(df.columns) > 1 else None
            channel_val = df.iloc[idx, 2] if len(df.columns) > 2 else None
            read_val = df.iloc[idx, 3] if len(df.columns) > 3 else 0
            if pd.isna(date_val) and pd.isna(channel_val):
                break
            if pd.isna(date_val) or pd.isna(channel_val):
                continue
            channel = str(channel_val).strip()
            if "日期" in channel or "传播渠道" in channel or "渠道" in channel:
                continue
            read_num = int(read_val) if pd.notna(read_val) else 0
            for keyword, field_name in channel_mapping.items():
                if keyword in channel:
                    result[field_name] += read_num
                    break
            row_count += 1
            if row_count > 500:
                break
        logs.append(f"  共读取 {row_count} 行渠道数据")
    else:
        logs.append("  ⚠️ 未找到'数据趋势明细'区域")

    # ========== 4. 解析性别分布 ==========
    gender_row = find_row_by_keyword(df, "性别分布", start_row=0, col_index=1)
    if gender_row is not None:
        logs.append(f"  ✅ 找到'性别分布'，起始行：{gender_row + 1}")
        gender_data = parse_gender_distribution(df, gender_row + 2)
        result["性别_女"] = gender_data["性别_女"]
        result["性别_男"] = gender_data["性别_男"]
        result["性别_未知"] = gender_data["性别_未知"]
    else:
        logs.append("  ⚠️ 未找到'性别分布'区域")

    # ========== 5. 解析年龄分布 ==========
    age_row = find_row_by_keyword(df, "年龄分布", start_row=0, col_index=1)
    if age_row is not None:
        logs.append(f"  ✅ 找到'年龄分布'，起始行：{age_row + 1}")
        age_data = parse_age_distribution(df, age_row + 2)
        result["年龄_18岁以下"] = age_data["年龄_18岁以下"]
        result["年龄_18-25岁"] = age_data["年龄_18-25岁"]
        result["年龄_26-35岁"] = age_data["年龄_26-35岁"]
        result["年龄_36-45岁"] = age_data["年龄_36-45岁"]
        result["年龄_46-60岁"] = age_data["年龄_46-60岁"]
        result["年龄_60岁以上"] = age_data["年龄_60岁以上"]
        result["年龄_未知"] = age_data["年龄_未知"]
    else:
        logs.append("  ⚠️ 未找到'年龄分布'区域")

    logs.append("  ✅ 解析完成")
    return result, logs


# ===================== 辅助函数 =====================

def pct_to_float(s):
    """把 '48.83%' 或数字 转成 48.83，解析失败返回 0"""
    try:
        if isinstance(s, (int, float)):
            return float(s)
        return float(str(s).replace('%', '').strip())
    except Exception:
        return 0.0


# 列顺序
COLUMN_ORDER = [
    "文章标题", "全部阅读人数", "送达人数", "首次分享人数", "分享产生的阅读人数",
    "推荐阅读人数", "其他阅读人数",
    "公众号消息阅读人数", "公众号主页阅读人数", "搜一搜阅读人数",
    "朋友圈阅读人数", "聊天会话阅读人数", "阅读后关注人数", "点赞人数",
    "评论次数", "总分享人数", "在看人数", "收藏人数", "停留时间", "完读率",
    # 性别分布
    "性别_女", "性别_男", "性别_未知",
    # 年龄分布
    "年龄_18岁以下", "年龄_18-25岁", "年龄_26-35岁", "年龄_36-45岁",
    "年龄_46-60岁", "年龄_60岁以上", "年龄_未知",
]

# 年龄展示顺序
AGE_ORDER = [
    "年龄_18岁以下", "年龄_18-25岁", "年龄_26-35岁", "年龄_36-45岁",
    "年龄_46-60岁", "年龄_60岁以上", "年龄_未知",
]

GENDER_KEYS = ["性别_女", "性别_男", "性别_未知"]


# ===================== Streamlit UI =====================

st.set_page_config(
    page_title="公众号数据解析系统",
    page_icon="📊",
    layout="wide",
)

# 顶部标题
st.markdown("""
<div style="background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
            padding: 2rem 2rem 1.5rem; border-radius: 12px; margin-bottom: 1.5rem;">
    <h1 style="color:white; margin:0; font-size:2rem;">📊 公众号数据解析系统</h1>
    <p style="color:#90caf9; margin:0.4rem 0 0; font-size:1rem;">
        上传公众号后台导出的 Excel / XLS 文件，一键解析，批量汇总，立即下载
    </p>
</div>
""", unsafe_allow_html=True)

# 使用说明
with st.expander("📖 使用说明", expanded=False):
    st.markdown("""
    1. **导出数据**：在微信公众平台后台 → 数据 → 文章数据 → 导出单篇文章Excel
    2. **批量上传**：点击下方上传框，可同时选择多个 `.xls` / `.xlsx` 文件
    3. **解析汇总**：系统自动提取核心指标 + 性别/年龄分布，所有数据整合在一张表中
    4. **下载结果**：点击"下载 Excel 结果"按钮，保存到本地

    > 💡 下载的 Excel 列顺序：文章标题 → 各项指标 → 完读率 → 性别分布 → 年龄分布
    """)

# 文件上传区域
st.markdown("### 📁 上传文件")
uploaded_files = st.file_uploader(
    "拖拽文件到此处，或点击选择文件（支持多选）",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    help="支持微信公众平台导出的 .xls 和 .xlsx 格式"
)

if uploaded_files:
    st.markdown(f"**已选择 {len(uploaded_files)} 个文件**")

    if st.button("🚀 开始解析", type="primary", use_container_width=True):
        all_results = []
        all_logs = []
        failed_files = []

        progress = st.progress(0)
        status_text = st.empty()

        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"正在解析 {uploaded_file.name} ({i+1}/{len(uploaded_files)})...")
            progress.progress((i + 1) / len(uploaded_files))

            file_bytes = uploaded_file.read()
            result, logs = parse_wx_excel(file_bytes, uploaded_file.name)

            all_logs.extend(logs)
            all_logs.append("")

            if result:
                all_results.append(result)
            else:
                failed_files.append(uploaded_file.name)

        progress.empty()
        status_text.empty()

        if all_results:
            # ── 构建宽表 ──
            df_output = pd.DataFrame(all_results)
            for col in COLUMN_ORDER:
                if col not in df_output.columns:
                    df_output[col] = 0
            df_output = df_output[COLUMN_ORDER]

            # ── 顶部汇总指标 ──
            col1, col2, col3 = st.columns(3)
            col1.metric("✅ 成功解析", f"{len(all_results)} 篇")
            col2.metric("❌ 解析失败", f"{len(failed_files)} 篇")
            total_read = sum(r.get("全部阅读人数", 0) for r in all_results)
            col3.metric("📈 总阅读人数", f"{total_read:,}")

            st.markdown("---")

            # ── 数据总览 ──
            st.markdown("### 📈 数据总览")
            m1, m2, m3, m4, m5, m6 = st.columns(6)
            total_send = sum(r.get("送达人数", 0) for r in all_results)
            total_like = sum(r.get("点赞人数", 0) for r in all_results)
            total_share = sum(r.get("总分享人数", 0) for r in all_results)
            total_fav = sum(r.get("收藏人数", 0) for r in all_results)
            avg_stay = sum(r.get("停留时间", 0) for r in all_results) / len(all_results)
            avg_cr_raw = sum(r.get("完读率", 0) for r in all_results) / len(all_results)
            m1.metric("总送达人数", f"{total_send:,}")
            m2.metric("总点赞人数", f"{total_like:,}")
            m3.metric("总分享人数", f"{total_share:,}")
            m4.metric("总收藏人数", f"{total_fav:,}")
            m5.metric("平均停留时间", f"{avg_stay:.1f}秒")
            if avg_cr_raw <= 1:
                m6.metric("平均完读率", f"{avg_cr_raw:.1%}")
            else:
                m6.metric("平均完读率", f"{avg_cr_raw:.1f}%")

            st.markdown("---")

            # ── 解析结果明细 ──
            st.markdown("### 📋 解析结果明细（含性别/年龄分布）")
            st.dataframe(df_output, use_container_width=True, hide_index=True)

            # ── 渠道分布图 ──
            st.markdown("---")
            st.markdown("### 📡 各渠道阅读人数分布")
            channel_keys = [
                "推荐阅读人数", "其他阅读人数", "公众号消息阅读人数",
                "公众号主页阅读人数", "搜一搜阅读人数", "朋友圈阅读人数", "聊天会话阅读人数"
            ]
            channel_totals = {}
            for ck in channel_keys:
                channel_totals[ck.replace("阅读人数", "")] = sum(r.get(ck, 0) for r in all_results)
            channel_df = pd.DataFrame({
                "渠道": list(channel_totals.keys()),
                "阅读人数": list(channel_totals.values())
            }).sort_values("阅读人数", ascending=False)
            channel_df = channel_df[channel_df["阅读人数"] > 0]
            if not channel_df.empty:
                st.bar_chart(channel_df.set_index("渠道"), use_container_width=True)
            else:
                st.info("暂无渠道数据")

            # ── 受众画像（加权合并后的可视化）──
            st.markdown("---")
            st.markdown("### 👥 受众画像分布（多篇加权均值）")

            # 加权合并性别/年龄分布
            gender_merged = {k: 0.0 for k in GENDER_KEYS}
            age_merged = {k: 0.0 for k in AGE_ORDER}
            total_readers = sum(r.get("全部阅读人数", 1) or 1 for r in all_results)

            for r in all_results:
                readers = r.get("全部阅读人数", 1) or 1
                weight = readers / total_readers if total_readers > 0 else 1 / len(all_results)
                for gk in GENDER_KEYS:
                    gender_merged[gk] += pct_to_float(r.get(gk, 0)) * weight
                for ak in AGE_ORDER:
                    age_merged[ak] += pct_to_float(r.get(ak, 0)) * weight

            tab1, tab2 = st.tabs(["⚥ 性别分布", "🎂 年龄分布"])

            with tab1:
                gdf = pd.DataFrame({
                    "性别": [k.replace("性别_", "") for k in GENDER_KEYS],
                    "占比(%)": [round(gender_merged[k], 2) for k in GENDER_KEYS]
                })
                gdf = gdf[gdf["占比(%)"] > 0]
                if not gdf.empty:
                    gcols = st.columns(len(gdf))
                    for ci, (_, row) in enumerate(gdf.iterrows()):
                        gcols[ci].metric(row["性别"], f"{row['占比(%)']:.2f}%")
                    st.bar_chart(gdf.set_index("性别"), use_container_width=True)
                else:
                    st.info("未解析到性别分布数据")

            with tab2:
                adf = pd.DataFrame({
                    "年龄段": [k.replace("年龄_", "") for k in AGE_ORDER],
                    "占比(%)": [round(age_merged[k], 2) for k in AGE_ORDER]
                })
                adf = adf[adf["占比(%)"] > 0]
                if not adf.empty:
                    st.bar_chart(adf.set_index("年龄段"), use_container_width=True)
                    st.dataframe(adf, use_container_width=True, hide_index=True)
                else:
                    st.info("未解析到年龄分布数据")

            # ── 下载结果 ──
            st.markdown("---")
            st.markdown("### 💾 下载结果")

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_output.to_excel(writer, index=False, sheet_name='数据汇总')

            excel_buffer.seek(0)

            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label="📥 下载 Excel（一张表含全部数据）",
                    data=excel_buffer,
                    file_name="公众号数据汇总.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col_dl2:
                csv_buffer = io.StringIO()
                df_output.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 下载 CSV（全部数据）",
                    data=csv_buffer.getvalue().encode('utf-8-sig'),
                    file_name="公众号数据汇总.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            # ── 解析日志 ──
            with st.expander("🔍 查看解析详情日志"):
                st.markdown("\n".join(all_logs))

            if failed_files:
                st.warning(f"以下文件解析失败：{', '.join(failed_files)}")

        else:
            st.error("所有文件解析失败，请检查文件格式是否正确。")
            with st.expander("查看错误日志"):
                st.markdown("\n".join(all_logs))

else:
    st.markdown("""
    <div style="text-align:center; padding:3rem; background:#f8f9fa; border-radius:12px;
                border: 2px dashed #dee2e6; color:#6c757d;">
        <div style="font-size:3rem; margin-bottom:1rem;">📂</div>
        <div style="font-size:1.1rem; font-weight:500;">请先上传文件开始解析</div>
        <div style="font-size:0.9rem; margin-top:0.5rem;">
            支持微信公众平台导出的 .xls / .xlsx 格式，可批量上传
        </div>
    </div>
    """, unsafe_allow_html=True)

# 底部版权
st.markdown("""
<hr style="margin-top:3rem; border-color:#e0e0e0;">
<div style="text-align:center; color:#9e9e9e; font-size:0.85rem; padding-bottom:1rem;">
    公众号数据解析系统 · 数据仅在本地浏览器处理，不上传至任何服务器，安全可靠
</div>
""", unsafe_allow_html=True)
