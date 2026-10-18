import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from pathlib import Path
import openpyxl
import re
import html
from datetime import datetime
from difflib import SequenceMatcher
from supabase import create_client

# ============================================================
# Supabase 資料庫連線
# ============================================================

supabase = create_client(
    st.secrets["SUPABASE_URL"],
    st.secrets["SUPABASE_KEY"]
)
# ============================================================
# 基本設定
# ============================================================

st.set_page_config(
    page_title="客服作業系統",
    page_icon="🐱",
    layout="wide"
)

EXCEL_PATH = Path("DATA/知識庫實驗品.xlsx")

# ============================================================
# 日系貓咪 UI｜全站統一樣式
# 僅調整外觀，不更動資料處理與功能邏輯
# ============================================================

st.html("""
<style>
/* ---------- 1. 全站基礎色彩與背景 ---------- */
:root {
    --cat-bg: #F7F4EE;
    --cat-panel: #FFFDF9;
    --cat-sidebar: #EEE8DD;
    --cat-border: #E4D9CB;
    --cat-border-strong: #D8C9B8;
    --cat-brown: #8A7863;
    --cat-terracotta: #B9785E;
    --cat-text: #514C45;
    --cat-muted: #898176;
    --cat-shadow: 0 3px 14px rgba(90, 70, 50, 0.055);
    --cat-radius: 16px;
}

.stApp {
    background: var(--cat-bg);
    color: var(--cat-text);
}

[data-testid="stHeader"] {
    background: rgba(247, 244, 238, 0.88) !important;
}

[data-testid="stToolbar"] {
    right: 1rem;
}

.block-container {
    max-width: 1450px;
    padding-top: 2.6rem;
    padding-bottom: 3.5rem;
}

/* ---------- 2. 標題與文字層級 ---------- */
h1, h2, h3, h4 {
    color: var(--cat-text) !important;
    letter-spacing: 0.02em;
}

h1 {
    font-weight: 750 !important;
    letter-spacing: 0.035em;
}

h2 {
    font-weight: 700 !important;
}

h3, h4 {
    font-weight: 650 !important;
}

p, label, .stMarkdown {
    color: var(--cat-text);
}

small, [data-testid="stCaptionContainer"] {
    color: var(--cat-muted) !important;
}

hr {
    border-color: var(--cat-border) !important;
    opacity: 0.85;
}

/* ---------- 3. 低調貓掌裝飾 ---------- */
[data-testid="stAppViewContainer"] {
    position: relative;
}

[data-testid="stAppViewContainer"]::after {
    content: "🐾";
    position: fixed;
    right: 24px;
    top: 82px;
    font-size: 25px;
    opacity: 0.12;
    pointer-events: none;
    z-index: 0;
}

/* ---------- 4. 側邊欄 ---------- */
[data-testid="stSidebar"] {
    background: var(--cat-sidebar);
    border-right: 1px solid var(--cat-border);
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 1.2rem;
}

[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: #554F48 !important;
}

[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
    line-height: 1.65;
}

[data-testid="stSidebar"] [data-testid="stRadio"] label,
[data-testid="stSidebar"] [data-testid="stCheckbox"] label {
    border-radius: 10px;
}

/* ---------- 5. 一般按鈕 ---------- */
.stButton > button,
.stDownloadButton > button,
[data-testid="stFormSubmitButton"] > button {
    min-height: 42px;
    padding: 0.48rem 1.05rem;
    border-radius: 11px;
    border: 1px solid var(--cat-border-strong);
    background: var(--cat-panel);
    color: var(--cat-text);
    font-weight: 600;
    box-shadow: 0 1px 3px rgba(90, 70, 50, 0.035);
    transition: background 0.18s ease,
                border-color 0.18s ease,
                transform 0.18s ease,
                box-shadow 0.18s ease;
}

.stButton > button:hover,
.stDownloadButton > button:hover,
[data-testid="stFormSubmitButton"] > button:hover {
    background: #F6EDE3;
    border-color: var(--cat-terracotta);
    color: #85553F;
    box-shadow: 0 3px 9px rgba(90, 70, 50, 0.08);
}

.stButton > button:active {
    transform: translateY(1px);
}

.stButton > button:focus:not(:active),
.stDownloadButton > button:focus:not(:active) {
    border-color: var(--cat-terracotta) !important;
    box-shadow: 0 0 0 3px rgba(185, 120, 94, 0.14) !important;
}

/* ---------- 6. 文字輸入框 ---------- */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stTextArea"] textarea {
    border: 1px solid var(--cat-border-strong) !important;
    border-radius: 11px !important;
    background: var(--cat-panel) !important;
    color: var(--cat-text) !important;
    transition: border-color 0.18s ease,
                box-shadow 0.18s ease;
}

[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {
    min-height: 43px;
    padding: 0 13px !important;
}

[data-testid="stTextArea"] textarea {
    padding: 12px 13px !important;
    line-height: 1.65;
}

[data-testid="stTextInput"] input:focus,
[data-testid="stNumberInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {
    border-color: var(--cat-terracotta) !important;
    box-shadow: 0 0 0 3px rgba(185, 120, 94, 0.11) !important;
}

input::placeholder,
textarea::placeholder {
    color: #AAA196 !important;
}

/* ---------- 7. 下拉選單與日期選擇器 ---------- */
[data-baseweb="select"] > div {
    border-color: var(--cat-border-strong) !important;
    border-radius: 11px !important;
    background: var(--cat-panel) !important;
}

[data-testid="stDateInput"] input,
[data-testid="stTimeInput"] input {
    border-radius: 11px !important;
    border-color: var(--cat-border-strong) !important;
    background: var(--cat-panel) !important;
}

/* ---------- 8. 頁籤 ---------- */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    gap: 7px;
    border-bottom: 1px solid var(--cat-border);
}

[data-testid="stTabs"] button[data-baseweb="tab"] {
    min-height: 43px;
    padding: 0.6rem 1rem;
    border-radius: 10px 10px 0 0;
    color: var(--cat-muted);
    font-weight: 600;
}

[data-testid="stTabs"] button[data-baseweb="tab"]:hover {
    background: rgba(255, 253, 249, 0.7);
    color: var(--cat-terracotta);
}

[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
    color: #895A45 !important;
    border-bottom: 3px solid var(--cat-terracotta) !important;
    background: rgba(255, 253, 249, 0.7);
}

/* ---------- 9. 容器卡片 ---------- */
[data-testid="stVerticalBlockBorderWrapper"] {
    border-color: var(--cat-border) !important;
    border-radius: var(--cat-radius) !important;
    background: rgba(255, 253, 249, 0.72);
    box-shadow: var(--cat-shadow);
}

/* ---------- 10. 資料表格 ---------- */
[data-testid="stDataFrame"],
[data-testid="stTable"] {
    border: 1px solid var(--cat-border);
    border-radius: 12px;
    overflow: hidden;
    background: var(--cat-panel);
}

/* ---------- 11. 提示訊息 ---------- */
[data-testid="stAlert"] {
    border-radius: 12px;
    border: 1px solid var(--cat-border);
}

/* ---------- 12. 展開區塊 ---------- */
[data-testid="stExpander"] {
    border: 1px solid var(--cat-border) !important;
    border-radius: 13px !important;
    background: rgba(255, 253, 249, 0.72);
    overflow: hidden;
}

[data-testid="stExpander"] summary {
    font-weight: 600;
}

/* ---------- 13. 表單 ---------- */
[data-testid="stForm"] {
    border: 1px solid var(--cat-border);
    border-radius: 15px;
    padding: 1.2rem;
    background: rgba(255, 253, 249, 0.72);
}

/* ---------- 14. 勾選框與單選按鈕 ---------- */
[data-testid="stCheckbox"] label,
[data-testid="stRadio"] label {
    color: var(--cat-text);
}

/* ---------- 15. 分隔線與載入狀態 ---------- */
[data-testid="stProgress"] > div > div {
    background-color: var(--cat-terracotta) !important;
}

[data-testid="stSpinner"] {
    color: var(--cat-brown);
}

/* ---------- 16. 右下角貓咪小幫手 ---------- */
[data-testid="stPopover"] {
    position: fixed !important;
    right: 22px !important;
    bottom: 22px !important;
    z-index: 999999 !important;
    width: 58px !important;
    height: 58px !important;
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin: 0 !important;
}

[data-testid="stPopover"] > button {
    width: 58px !important;
    height: 58px !important;
    min-width: 58px !important;
    min-height: 58px !important;
    padding: 0 !important;
    margin: 0 !important;
    border-radius: 50% !important;
    background: var(--cat-panel) !important;
    border: 1px solid var(--cat-border-strong) !important;
    box-shadow: 0 4px 15px rgba(90, 70, 50, 0.18) !important;
    font-size: 27px !important;
    transition: transform 0.18s ease,
                box-shadow 0.18s ease;
}

[data-testid="stPopover"] > button:hover {
    transform: scale(1.07);
    border-color: var(--cat-terracotta) !important;
    box-shadow: 0 6px 18px rgba(90, 70, 50, 0.2) !important;
}

[data-testid="stPopover"] > button:focus {
    outline: none !important;
}

/* ---------- 17. 手機與窄螢幕 ---------- */
@media (max-width: 768px) {
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    [data-testid="stAppViewContainer"]::after {
        right: 12px;
        top: 65px;
        font-size: 20px;
    }

    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 2px;
    }

    [data-testid="stTabs"] button[data-baseweb="tab"] {
        padding: 0.45rem 0.6rem;
    }
}
</style>
""")

# ============================================================
# 共用工具
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_search_text(value):
    value = clean_text(value).lower()
    value = value.replace("台", "臺")
    value = re.sub(r"\s+", "", value)
    return value


def clean_announcement_content(text):
    if text is None or pd.isna(text):
        return ""

    text = html.unescape(str(text))

    text = re.sub(
        r"```(?:html|HTML|text|TEXT)?",
        "",
        text
    )

    text = text.replace("```", "")

    text = re.sub(
        r"<svg[\s\S]*?</svg>",
        "",
        text,
        flags=re.I
    )

    text = re.sub(
        r"\bsvg\b",
        "",
        text,
        flags=re.I
    )

    text = re.sub(
        r"<br\s*/?>",
        "\n",
        text,
        flags=re.I
    )

    text = re.sub(
        r"</(?:p|div|h[1-6]|li|tr)>",
        "\n",
        text,
        flags=re.I
    )

    text = re.sub(
        r"<li[^>]*>",
        "• ",
        text,
        flags=re.I
    )

    text = re.sub(
        r"<[^>]+>",
        "",
        text
    )

    text = re.sub(
        r"\b(?:id|class)\s*=\s*['\"][^'\"]*['\"]",
        "",
        text,
        flags=re.I
    )

    text = text.replace("\\~", "~").replace("\\", "")

    text = re.sub(
        r"\n\s*\n\s*\n+",
        "\n\n",
        text
    )

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    return text.strip()


def clean_announcement_title(text):
    text = clean_announcement_content(text)
    text = text.replace("📌", "").strip()

    if text.startswith("【") and text.endswith("】"):
        text = text[1:-1].strip()

    return text


# ============================================================
# 公告欄位判斷
# ============================================================

def get_announcement_columns(df):

    cols = list(df.columns)

    title_col = next(
        (
            c for c in cols
            if "標題" in str(c)
        ),
        cols[0] if cols else None
    )

    content_col = next(
        (
            c for c in cols
            if any(
                x in str(c)
                for x in ["內容", "說明"]
            )
        ),
        None
    )

    if content_col is None and len(cols) > 1:
        content_col = cols[1]

    date_col = next(
        (
            c for c in cols
            if "日期" in str(c)
            and "活動" not in str(c)
            and "點數" not in str(c)
        ),
        None
    )

    activity_date_col = next(
        (
            c for c in cols
            if "活動日期" in str(c)
        ),
        None
    )

    if activity_date_col is None:
        activity_date_col = next(
            (
                c for c in cols
                if "活動期限" in str(c)
            ),
            None
        )

    point_date_col = next(
        (
            c for c in cols
            if "點數期限" in str(c)
        ),
        None
    )

    return {
        "title": title_col,
        "content": content_col,
        "date": date_col,
        "activity_date": activity_date_col,
        "point_date": point_date_col
    }


def get_announcement_values(row, columns):

    title = ""
    content = ""
    date = ""
    activity_date = ""
    point_date = ""

    if columns.get("title") is not None:
        title = clean_announcement_title(
            row.get(columns["title"], "")
        )

    if columns.get("content") is not None:
        content = clean_announcement_content(
            row.get(columns["content"], "")
        )

    if columns.get("date") is not None:
        date = clean_text(
            row.get(columns["date"], "")
        )

    if columns.get("activity_date") is not None:
        activity_date = clean_text(
            row.get(columns["activity_date"], "")
        )

    if columns.get("point_date") is not None:
        point_date = clean_text(
            row.get(columns["point_date"], "")
        )

    return (
        title,
        content,
        date,
        activity_date,
        point_date
    )


def render_announcement(row, columns):

    (
        title,
        content,
        date,
        activity_date,
        point_date
    ) = get_announcement_values(
        row,
        columns
    )
    status = clean_text(
        row.get(
            "顯示狀態",
            ""
        )
    )

    with st.container(border=True):

        if title:
            st.subheader(
                "📌 " + title
            )

        if date:
            try:
                date = pd.to_datetime(
                    date
                ).strftime("%Y/%m/%d")
            except:
                date = str(date)

            st.caption(
                "📅 公告日期：" + date
            )

        if activity_date:
            try:
                activity_date = pd.to_datetime(
                    activity_date
                ).strftime("%Y/%m/%d")
            except:
                activity_date = str(activity_date)

            st.caption(
                "🎉 活動日期：" + activity_date
            )

        if point_date:
            try:
                point_date = pd.to_datetime(
                    point_date
                ).strftime("%Y/%m/%d")
            except:
                point_date = str(point_date)

            st.caption(
                "🎁 點數期限：" + point_date
            )
        if content:
            st.write(content)

def render_announcement_admin(
    row,
    columns,
    index
):
    (
        title,
        content,
        date,
        activity_date,
        point_date
    ) = get_announcement_values(
        row,
        columns
    )

    status = clean_text(
        row.get("顯示狀態", "")
    )

    announcement_id = clean_text(
        row.get("公告編號", "")
    )

    with st.container(border=True):

        # 標題
        if title:
            st.markdown(f"### 📌 {title}")
        else:
            st.markdown("### 📌 未命名公告")

        # 顯示狀態
        if status == "是":
            st.caption("🟢 顯示中")
        elif status == "否":
            st.caption("⚪ 已隱藏")

        # 日期資訊
        info = []

        if date:
            try:
                date = pd.to_datetime(date).strftime("%Y/%m/%d")
            except Exception:
                date = str(date)

            info.append(f"📅 公告日期：{date}")

        if activity_date:
            try:
                activity_date = pd.to_datetime(
                    activity_date
                ).strftime("%Y/%m/%d")
            except Exception:
                activity_date = str(activity_date)

            info.append(f"🎉 活動日期：{activity_date}")

        if point_date:
            try:
                point_date = pd.to_datetime(
                    point_date
                ).strftime("%Y/%m/%d")
            except Exception:
                point_date = str(point_date)

            info.append(f"🎁 點數期限：{point_date}")

        if info:
            st.caption("　｜　".join(info))

        # 內容預覽
        if content:
            preview = str(content).replace("\n", " ")

            if len(preview) > 100:
                preview = preview[:100] + "……"

            st.write(preview)

        # 編輯與刪除按鈕
        edit_col, delete_col = st.columns(2)

        with edit_col:
            if st.button(
                "✏️ 編輯",
                key=f"edit_announcement_{index}",
                use_container_width=True
            ):
                st.session_state[
                    "editing_announcement"
                ] = index

                st.session_state.pop(
                    "deleting_announcement",
                    None
                )

                st.rerun()

        with delete_col:
            if st.button(
                "🗑️ 刪除",
                key=f"delete_announcement_{index}",
                use_container_width=True
            ):
                st.session_state[
                    "deleting_announcement"
                ] = index

                st.rerun()

        # 刪除確認區
        if (
            st.session_state.get(
                "deleting_announcement"
            ) == index
        ):

            st.divider()

            st.warning(
                f"即將刪除公告「{title or '未命名公告'}」。"
                "請確認公告內容及編號後再繼續。"
            )

            if not announcement_id:
                st.error(
                    "找不到公告編號，為避免誤刪，已停止此操作。"
                )

            else:
                delete_editor_name = st.text_input(
                    "👤 刪除人姓名（必填）",
                    key=f"delete_editor_name_{index}",
                    placeholder="請輸入實際操作人的姓名"
                )

                confirm_col, cancel_col = st.columns(2)

                with confirm_col:
                    confirm_delete = st.button(
                        "⚠️ 確認刪除",
                        type="primary",
                        key=f"confirm_delete_{index}",
                        use_container_width=True,
                        disabled=not delete_editor_name.strip()
                    )

                with cancel_col:
                    cancel_delete = st.button(
                        "取消",
                        key=f"cancel_delete_{index}",
                        use_container_width=True
                    )

                if cancel_delete:
                    st.session_state.pop(
                        "deleting_announcement",
                        None
                    )

                    st.rerun()

                if confirm_delete:

                    # 保存刪除前的公告內容
                    old_data = {
                        "title": str(title or ""),
                        "content": str(content or ""),
                        "announcement_date": str(date or ""),
                        "event_date": str(activity_date or ""),
                        "points_deadline": str(point_date or ""),
                        "display_status": status
                    }

                    # 第一步：先寫入刪除歷程
                    history_saved = save_announcement_history(
                        action="刪除",
                        announcement_id=announcement_id,
                        title=title,
                        editor_name=delete_editor_name.strip(),
                        old_data=old_data
                    )

                    if not history_saved:
                        st.error(
                            "刪除歷程保存失敗，公告尚未刪除。"
                            "請檢查 Supabase 後再試。"
                        )

                    else:
                        # 第二步：歷程保存成功後才刪除公告
                        try:
                            response = (
                                supabase
                                .table("announcements")
                                .delete()
                                .eq(
                                    "announcement_id",
                                    announcement_id
                                )
                                .select("announcement_id")
                                .execute()
                            )

                            if response.data:

                                load_announcements.clear()

                                st.session_state.pop(
                                    "deleting_announcement",
                                    None
                                )

                                if (
                                    st.session_state.get(
                                        "editing_announcement"
                                    ) == index
                                ):
                                    st.session_state.pop(
                                        "editing_announcement",
                                        None
                                    )

                                st.success(
                                    "公告已刪除，刪除歷程已保存。"
                                )

                                st.rerun()

                            else:
                                st.error(
                                    "沒有刪除任何公告。"
                                    "請確認公告編號是否正確；"
                                    "刪除操作歷程已保存。"
                                )

                        except Exception as e:
                            st.error(
                                f"公告刪除失敗：{e}。"
                                "刪除操作歷程已保存，請確認公告是否仍存在。"
                            )
# ============================================================
# Excel 載入
# ============================================================

@st.cache_data(ttl=60)
def get_excel_sheets():

    if not EXCEL_PATH.exists():
        return []

    return pd.ExcelFile(
        EXCEL_PATH
    ).sheet_names


def find_sheet(candidates, fallback_index=None):

    sheets = get_excel_sheets()

    for name in candidates:

        if name in sheets:
            return name

    if (
        fallback_index is not None
        and len(sheets) > fallback_index
    ):
        return sheets[fallback_index]

    return None


@st.cache_data(ttl=60)
def load_data():

    sheet = find_sheet(
        ["知識庫", "問題集", "Sheet1"],
        0
    )

    if not sheet:
        return pd.DataFrame()

    
    return pd.read_excel(
        EXCEL_PATH,
        sheet_name=sheet,
        dtype=str
    ).fillna("")


@st.cache_data(ttl=60)
def load_synonyms():

    sheet = find_sheet(
        ["搜尋詞對照", "關鍵字對照"]
    )

    if not sheet:
        return pd.DataFrame()

    return pd.read_excel(
        EXCEL_PATH,
        sheet_name=sheet,
        dtype=str
    ).fillna("")


@st.cache_data(ttl=60)
def load_announcements():

    try:

        # 從 Supabase 讀取公告資料
        response = (
            supabase
            .table("announcements")
            .select("*")
            .execute()
        )

        df = pd.DataFrame(
            response.data or []
        )

        # 將 Supabase 英文欄位名稱轉回原本的中文欄位
        column_mapping = {
            "announcement_id": "公告編號",
            "title": "公告標題",
            "content": "公告內容",
            "announcement_date": "公告日期",
            "event_date": "活動日期",
            "points_deadline": "點數期限",
            "display_status": "顯示狀態"
        }

        if df.empty:

            return pd.DataFrame(
                columns=list(column_mapping.values())
            )

        df = df.rename(
            columns=column_mapping
        )

        # 確保必要欄位存在
        for column in column_mapping.values():

            if column not in df.columns:
                df[column] = ""

        df = df.fillna("")

        # 清理公告內容
        df["公告內容"] = df["公告內容"].map(
            clean_announcement_content
        )

        # 清理公告標題
        df["公告標題"] = df["公告標題"].map(
            clean_announcement_title
        )

        return df

    except Exception as e:

        st.warning(
            f"無法讀取 Supabase 公告資料：{e}"
        )

        return pd.DataFrame(
            columns=[
                "公告編號",
                "公告標題",
                "公告內容",
                "公告日期",
                "活動日期",
                "點數期限",
                "顯示狀態"
            ]
        )

# =========================================================
# 新人訓練中心
# =========================================================

TRAINING_OVERVIEW_COLUMNS = [
    "課程ID",
    "DAY",
    "課程大項",
    "Level",
]

TRAINING_CONTENT_COLUMNS = [
    "課程ID",
    "內容ID",
    "課程大項",
    "課程細項",
    "注意事項",
    "Level",
]

TRAINING_VIDEO_COLUMNS = [
    "課程ID",
    "影片URL",
    "影片名稱",
    "影片說明",
]
@st.cache_data(ttl=60)
def load_training_videos():

    try:

        sheet = find_sheet(
            ["訓練影片", "課程影片", "影片"]
        )

        if not sheet:
            return pd.DataFrame(
                columns=TRAINING_VIDEO_COLUMNS
            )

        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet
        )

        if df.empty:
            return pd.DataFrame(
                columns=TRAINING_VIDEO_COLUMNS
            )

        for col in TRAINING_VIDEO_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        df = df[
            TRAINING_VIDEO_COLUMNS
        ].copy()

        df["課程ID"] = (
            df["課程ID"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        return df

    except Exception as e:

        st.warning(
            f"無法讀取訓練影片：{e}"
        )

        return pd.DataFrame(
            columns=TRAINING_VIDEO_COLUMNS
        )

TRAINING_QUIZ_COLUMNS = [
    "題目ID",
    "課程ID",
    "題型",
    "題目",
    "選項A",
    "選項B",
    "選項C",
    "選項D",
    "正確答案",
    "解說",
    "Level",
]

@st.cache_data(ttl=60)
def load_training_overview():

    try:

        sheet = find_sheet(
            ["課程總表"]
        )

        if not sheet:

            return pd.DataFrame(
                columns=TRAINING_OVERVIEW_COLUMNS
            )

        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet
        )

        for col in TRAINING_OVERVIEW_COLUMNS:

            if col not in df.columns:
                df[col] = ""

        df = df[
            TRAINING_OVERVIEW_COLUMNS
        ].copy()

        return df

    except Exception as e:

        st.warning(
            f"無法讀取課程總表：{e}"
        )

        return pd.DataFrame(
            columns=TRAINING_OVERVIEW_COLUMNS
        )


@st.cache_data(ttl=60)
def load_training_content():

    try:

        # 從 Supabase 讀取課程細項
        response = (
            supabase
            .table("course_contents")
            .select("*")
            .execute()
        )

        # 將 Supabase 資料轉成 DataFrame
        df = pd.DataFrame(
            response.data or []
        )

        # Supabase 欄位對應原本程式使用的中文欄位
        column_mapping = {
            "course_id": "課程ID",
            "content_id": "內容ID",
            "course_name": "課程大項",
            "content_title": "課程細項",
            "notes": "注意事項",
            "level": "Level"
        }

        # 沒有資料時，回傳原本需要的欄位
        if df.empty:
            return pd.DataFrame(
                columns=TRAINING_CONTENT_COLUMNS
            )

        # 轉換欄位名稱
        df = df.rename(
            columns=column_mapping
        )

        # 補齊原本程式需要的欄位
        for col in TRAINING_CONTENT_COLUMNS:

            if col not in df.columns:
                df[col] = ""

        # 保留原本欄位順序
        df = df[
            TRAINING_CONTENT_COLUMNS
        ].copy()

        # 清理 Level 欄位
        df["Level"] = (
            df["Level"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        # 保留原本 Level 1 的篩選規則
        df = df[
            df["Level"].isin(["", "1"])
        ].copy()

        # 統一課程 ID 格式，避免比對時出現空白
        df["課程ID"] = (
            df["課程ID"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        
        st.write("課程細項讀取筆數：", len(df))
        st.dataframe(df, use_container_width=True)
        
        return df

    except Exception as e:

        st.warning(
            f"無法讀取 Supabase 課程內容：{e}"
        )

        return pd.DataFrame(
            columns=TRAINING_CONTENT_COLUMNS
        )



@st.cache_data(ttl=60)
def load_employee_data():

    try:

        sheet = find_sheet(["員工資料"])

        if not sheet:
            return pd.DataFrame(
                columns=[
                    "員工編號",
                    "姓名",
                    "職稱",
                    "部門",
                    "角色",
                    "到職日"
                ]
            )

        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet
        )

        if df.empty:
            return pd.DataFrame(
                columns=[
                    "員工編號",
                    "姓名",
                    "職稱",
                    "部門",
                    "角色",
                    "到職日"
                ]
            )

        return df

    except Exception as e:

        st.warning(
            f"無法讀取員工資料：{e}"
        )

        return pd.DataFrame()
@st.cache_data(ttl=60)
def load_training_records():

    try:

        sheet = find_sheet(["新人學習紀錄"])

        if not sheet:
            return pd.DataFrame(
                columns=[
                    "員工編號",
                    "姓名",
                    "課程ID",
                    "課程名稱",
                    "DAY",
                    "開始日期",
                    "學習狀態",
                    "測驗狀態",
                    "實作",
                    "完成日期",
                    "教育者備註"
                ]
            )

        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet
        )

        if df.empty:
            return pd.DataFrame(
                columns=[
                    "員工編號",
                    "姓名",
                    "課程ID",
                    "課程名稱",
                    "DAY",
                    "開始日期",
                    "學習狀態",
                    "測驗狀態",
                    "實作",
                    "完成日期",
                    "教育者備註"
                ]
            )

        return df

    except Exception as e:

        st.warning(
            f"無法讀取新人學習紀錄：{e}"
        )

        return pd.DataFrame()
        
def save_training_record(
    employee_id,
    employee_name,
    course_id,
    course_name,
    day,
    start_date,
    learning_status,
    quiz_status,
    completed_date
):
    try:
        sheet = find_sheet(["新人學習紀錄"])

        if not sheet:
            st.error("找不到「新人學習紀錄」工作表。")
            return False

        records = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet
        )

        # 確保必要欄位存在
        required_columns = [
            "員工編號", "姓名", "課程ID", "課程名稱",
            "DAY", "開始日期", "學習狀態", "測驗狀態",
            "實作", "完成日期", "教育者備註"
        ]

        for column in required_columns:
            if column not in records.columns:
                records[column] = ""

        # 比對同一位員工的同一堂課
        employee_ids = (
            records["員工編號"].astype(str).str.strip()
        )
        course_ids = (
            records["課程ID"].astype(str).str.strip()
        )

        matched_indices = records.index[
            (employee_ids == str(employee_id).strip())
            & (course_ids == str(course_id).strip())
        ].tolist()

        if matched_indices:
            # 已有紀錄：更新原本那一列
            idx = matched_indices[0]

            records.at[idx, "姓名"] = employee_name
            records.at[idx, "課程名稱"] = course_name
            records.at[idx, "DAY"] = day
            records.at[idx, "學習狀態"] = learning_status
            records.at[idx, "測驗狀態"] = quiz_status

            # 有提供新日期才更新，避免清除原日期
            if str(start_date).strip():
                records.at[idx, "開始日期"] = start_date

            if str(completed_date).strip():
                records.at[idx, "完成日期"] = completed_date

        else:
            # 沒有紀錄：新增一列
            new_record = {
                "員工編號": employee_id,
                "姓名": employee_name,
                "課程ID": course_id,
                "課程名稱": course_name,
                "DAY": day,
                "開始日期": start_date,
                "學習狀態": learning_status,
                "測驗狀態": quiz_status,
                "實作": "",
                "完成日期": completed_date,
                "教育者備註": ""
            }

            records = pd.concat(
                [records, pd.DataFrame([new_record])],
                ignore_index=True
            )

        # 寫回 Excel
        with pd.ExcelWriter(
            EXCEL_PATH,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            records.to_excel(
                writer,
                sheet_name=sheet,
                index=False
            )

        # 清除讀取快取，讓下次讀取取得最新紀錄
        load_training_records.clear()

        return True

    except Exception as e:
        st.error(f"儲存新人學習紀錄失敗：{e}")
        return False

def load_training_quiz():

    try:

        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="測驗題庫"
        )

        df.columns = [
            str(column).strip()
            for column in df.columns
        ]

        for column in TRAINING_QUIZ_COLUMNS:

            if column not in df.columns:
                df[column] = ""

        return df[
            TRAINING_QUIZ_COLUMNS
        ].fillna("")

    except Exception as e:

        st.error(
            f"讀取「測驗題庫」失敗：{e}"
        )

        return pd.DataFrame(
            columns=TRAINING_QUIZ_COLUMNS
        )
        
training_overview = load_training_overview()
training_content = load_training_content()
training_videos = load_training_videos()
training_records = load_training_records()
employee_data = load_employee_data()
training_quiz = load_training_quiz()

# ============================================================
# 待審核問題
# ============================================================

PENDING_COLUMNS = [
    "提出時間",
    "問題",
    "建議分類",
    "狀態",
    "審核備註"
]


@st.cache_data(ttl=60)
def load_pending_questions():

    if not EXCEL_PATH.exists():

        return pd.DataFrame(
            columns=PENDING_COLUMNS
        )

    try:

        sheets = get_excel_sheets()

        if "待審核問題" not in sheets:

            return pd.DataFrame(
                columns=PENDING_COLUMNS
            )

        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name="待審核問題",
            dtype=str
        ).fillna("")

        for col in PENDING_COLUMNS:

            if col not in df.columns:
                df[col] = ""

        return df[
            PENDING_COLUMNS
        ].copy()

    except Exception as e:

        st.error(
            f"讀取待審核問題失敗：{e}"
        )

        return pd.DataFrame(
            columns=PENDING_COLUMNS
        )


# ============================================================
# 修改歷程
# ============================================================

HISTORY_COLUMNS = [
    "修改時間",
    "修改人",
    "資料區域",
    "動作",
    "資料識別",
    "欄位名稱",
    "修改前",
    "修改後"
]


def get_history_identifier(
    row,
    index
):

    identifier_fields = [
        "序號",
        "商品編號",
        "問題摘要",
        "問題",
        "標題",
        "名稱",
        "工具名稱"
    ]

    for field in identifier_fields:

        if field in row.index:

            value = clean_text(
                row.get(
                    field,
                    ""
                )
            )

            if value:
                return value

    return f"第 {index + 1} 筆"


def build_change_history(
    old_df,
    new_df,
    data_area,
    editor_name
):

    history = []

    old_df = old_df.copy()
    new_df = new_df.copy()

    all_columns = list(
        dict.fromkeys(
            list(old_df.columns)
            + list(new_df.columns)
        )
    )

    for col in all_columns:

        if col not in old_df.columns:
            old_df[col] = ""

        if col not in new_df.columns:
            new_df[col] = ""

    old_df = old_df[
        all_columns
    ].fillna("")

    new_df = new_df[
        all_columns
    ].fillna("")

    max_rows = max(
        len(old_df),
        len(new_df)
    )

    now = datetime.now().strftime(
        "%Y/%m/%d %H:%M:%S"
    )

    for i in range(max_rows):

        old_exists = i < len(old_df)
        new_exists = i < len(new_df)

        if not old_exists and new_exists:

            new_row = new_df.iloc[i]

            history.append({
                "修改時間": now,
                "修改人": editor_name,
                "資料區域": data_area,
                "動作": "新增",
                "資料識別": get_history_identifier(
                    new_row,
                    i
                ),
                "欄位名稱": "整筆資料",
                "修改前": "",
                "修改後": "新增資料"
            })

            continue

        if old_exists and not new_exists:

            old_row = old_df.iloc[i]

            history.append({
                "修改時間": now,
                "修改人": editor_name,
                "資料區域": data_area,
                "動作": "刪除",
                "資料識別": get_history_identifier(
                    old_row,
                    i
                ),
                "欄位名稱": "整筆資料",
                "修改前": "原有資料",
                "修改後": ""
            })

            continue

        old_row = old_df.iloc[i]
        new_row = new_df.iloc[i]

        identifier = get_history_identifier(
            new_row,
            i
        )

        if identifier.startswith("第 "):

            identifier = get_history_identifier(
                old_row,
                i
            )

        for col in all_columns:

            old_value = clean_text(
                old_row.get(
                    col,
                    ""
                )
            )

            new_value = clean_text(
                new_row.get(
                    col,
                    ""
                )
            )

            if old_value != new_value:

                history.append({
                    "修改時間": now,
                    "修改人": editor_name,
                    "資料區域": data_area,
                    "動作": "修改",
                    "資料識別": identifier,
                    "欄位名稱": col,
                    "修改前": old_value,
                    "修改後": new_value
                })

    return pd.DataFrame(
        history,
        columns=HISTORY_COLUMNS
    )


def save_change_history(
    history_df
):

    if history_df.empty:
        return True

    if not EXCEL_PATH.exists():
        return False

    try:

        wb = openpyxl.load_workbook(
            EXCEL_PATH
        )

        sheet_name = "修改歷程"

        if sheet_name not in wb.sheetnames:

            ws = wb.create_sheet(
                sheet_name
            )

            for c, col in enumerate(
                HISTORY_COLUMNS,
                1
            ):

                ws.cell(
                    1,
                    c,
                    col
                )

        else:

            ws = wb[
                sheet_name
            ]

            for c, col in enumerate(
                HISTORY_COLUMNS,
                1
            ):

                ws.cell(
                    1,
                    c,
                    col
                )

        for _, row in history_df.iterrows():

            ws.append([
                clean_text(
                    row.get(
                        col,
                        ""
                    )
                )
                for col in HISTORY_COLUMNS
            ])

        wb.save(
            EXCEL_PATH
        )

        return True

    except PermissionError:

        st.error(
            "❌ Excel 目前可能正在被開啟，無法寫入修改歷程。"
        )

        return False

    except Exception as e:

        st.error(
            f"❌ 修改歷程儲存失敗：{e}"
        )

        return False
        
def save_announcement_history(
    action,
    announcement_id,
    title,
    editor_name,
    old_data=None,
    new_data=None
):
    """將公告新增、修改、刪除紀錄寫入 Supabase。"""

    history = []

    base_data = {
        "announcement_id": str(announcement_id),
        "announcement_title": str(title or ""),
        "action": action,
        "editor_name": str(editor_name or "").strip()
    }

    if action == "新增":
        history.append({
            **base_data,
            "field_name": "整筆資料",
            "old_value": "",
            "new_value": "新增公告"
        })

    elif action in ("修改", "刪除"):
        old_data = old_data or {}
        new_data = new_data or {}

        fields = [
            ("title", "公告標題"),
            ("content", "公告內容"),
            ("announcement_date", "公告日期"),
            ("event_date", "活動日期"),
            ("points_deadline", "點數期限"),
            ("display_status", "顯示狀態")
        ]

        for field, field_label in fields:
            old_value = str(old_data.get(field, "") or "").strip()

            if action == "刪除":
                if old_value:
                    history.append({
                        **base_data,
                        "field_name": field_label,
                        "old_value": old_value,
                        "new_value": ""
                    })

            else:
                new_value = str(
                    new_data.get(field, "") or ""
                ).strip()

                if old_value != new_value:
                    history.append({
                        **base_data,
                        "field_name": field_label,
                        "old_value": old_value,
                        "new_value": new_value
                    })

    if not history:
        return True

    try:
        response = (
            supabase
            .table("announcement_history")
            .insert(history)
            .execute()
        )

        return bool(response.data)

    except Exception as e:
        print(f"公告歷程寫入失敗：{e}")
        return False



# ============================================================
# 儲存 Excel
# ============================================================

def save_sheet(
    sheet_name,
    data,
    editor_name=""
):

    if not EXCEL_PATH.exists():

        st.error(
            "找不到 Excel：DATA/知識庫實驗品.xlsx"
        )

        return False

    if not clean_text(
        editor_name
    ):

        st.error(
            "❌ 請先填寫修改人姓名。"
        )

        return False

    try:

        old_df = pd.DataFrame()

        if (
            sheet_name
            and sheet_name in get_excel_sheets()
        ):

            try:

                old_df = pd.read_excel(
                    EXCEL_PATH,
                    sheet_name=sheet_name,
                    dtype=str
                ).fillna("")

            except Exception:

                old_df = pd.DataFrame()

        history_df = build_change_history(
            old_df,
            data,
            sheet_name,
            clean_text(editor_name)
        )

        wb = openpyxl.load_workbook(
            EXCEL_PATH
        )

        if sheet_name in wb.sheetnames:

            wb.remove(
                wb[sheet_name]
            )

        ws = wb.create_sheet(
            sheet_name
        )

        for c, col in enumerate(
            data.columns,
            1
        ):

            ws.cell(
                1,
                c,
                str(col)
            )

        for r, row in enumerate(
            data.itertuples(
                index=False
            ),
            2
        ):

            for c, value in enumerate(
                row,
                1
            ):

                if pd.isna(value):
                    value = ""

                ws.cell(
                    r,
                    c,
                    value
                )

        wb.save(
            EXCEL_PATH
        )

        if not history_df.empty:

            if not save_change_history(
                history_df
            ):

                st.warning(
                    "⚠️ 資料已儲存，但修改歷程寫入失敗。"
                )

        st.cache_data.clear()

        return True

    except PermissionError:

        st.error(
            "❌ Excel 目前可能正在被開啟，請先關閉「知識庫實驗品.xlsx」後再試一次。"
        )

        return False

    except Exception as e:

        st.error(
            f"❌ 儲存失敗：{e}"
        )

        return False


def save_knowledge_data(
    data,
    editor_name=""
):

    sheet = find_sheet(
        [
            "知識庫",
            "問題集",
            "Sheet1"
        ],
        0
    )

    if sheet:

        return save_sheet(
            sheet,
            data,
            editor_name
        )

    return False


# ============================================================
# 新增待審核問題
# ============================================================

def add_pending_question(
    question
):

    question = clean_text(
        question
    )

    if not question:
        return False

    if not EXCEL_PATH.exists():

        st.error(
            "❌ 找不到 Excel 檔案。"
        )

        return False

    try:

        wb = openpyxl.load_workbook(
            EXCEL_PATH
        )

        sheet_name = "待審核問題"

        if sheet_name not in wb.sheetnames:

            ws = wb.create_sheet(
                sheet_name
            )

            for c, header in enumerate(
                PENDING_COLUMNS,
                1
            ):

                ws.cell(
                    1,
                    c,
                    header
                )

        else:

            ws = wb[
                sheet_name
            ]

            for index, header in enumerate(
                PENDING_COLUMNS,
                1
            ):

                if (
                    index > ws.max_column
                    or not clean_text(
                        ws.cell(
                            1,
                            index
                        ).value
                    )
                ):

                    ws.cell(
                        1,
                        index,
                        header
                    )

        existing_questions = []

        if ws.max_row >= 2:

            problem_column = None

            for c in range(
                1,
                ws.max_column + 1
            ):

                header = clean_text(
                    ws.cell(
                        1,
                        c
                    ).value
                )

                if header == "問題":

                    problem_column = c
                    break

            if problem_column is None:
                problem_column = 2

            for r in range(
                2,
                ws.max_row + 1
            ):

                value = clean_text(
                    ws.cell(
                        r,
                        problem_column
                    ).value
                )

                if value:
                    existing_questions.append(
                        value
                    )

        if question in existing_questions:

            wb.save(
                EXCEL_PATH
            )

            st.cache_data.clear()

            return True

        ws.append([
            datetime.now().strftime(
                "%Y/%m/%d %H:%M"
            ),
            question,
            "",
            "確認中",
            ""
        ])

        wb.save(
            EXCEL_PATH
        )

        wb_check = openpyxl.load_workbook(
            EXCEL_PATH,
            read_only=True,
            data_only=True
        )

        if sheet_name not in wb_check.sheetnames:

            wb_check.close()

            st.error(
                "❌ 儲存後找不到「待審核問題」工作表。"
            )

            return False

        ws_check = wb_check[
            sheet_name
        ]

        saved = False

        for row in ws_check.iter_rows(
            min_row=2,
            values_only=True
        ):

            if (
                len(row) >= 2
                and clean_text(row[1])
                == question
            ):

                saved = True
                break

        wb_check.close()

        st.cache_data.clear()

        if saved:
            return True

        st.error(
            "❌ 已執行儲存，但重新讀取 Excel 時沒有找到這筆問題。"
        )

        return False

    except PermissionError:

        st.error(
            "❌ Excel 目前正在被開啟，請先關閉「知識庫實驗品.xlsx」再提交。"
        )

        return False

    except Exception as e:

        st.error(
            f"❌ 待審核問題儲存失敗：{e}"
        )

        return False


# ============================================================
# 客訴資料
# ============================================================

@st.cache_data(ttl=60)
def load_complaints():

    sheet = find_sheet(
        ["客訴資料"]
    )

    if not sheet:
        return pd.DataFrame()

    df = pd.read_excel(
        EXCEL_PATH,
        sheet_name=sheet,
        dtype=str
    ).fillna("")

    expected_columns = [
        "序號",
        "日期",
        "月份",
        "購買通路",
        "商品編號",
        "商品名稱",
        "客訴類型",
        "客訴細項",
        "客訴內容",
        "處理方式",
        "處理結果",
        "備註"
    ]

    for col in expected_columns:

        if col not in df.columns:
            df[col] = ""

    df["_日期"] = pd.to_datetime(
        df["日期"],
        errors="coerce"
    )

    df["_月份"] = df[
        "_日期"
    ].dt.strftime(
        "%Y/%m"
    )

    df["_月份"] = df[
        "_月份"
    ].fillna(
        df["月份"].apply(
            clean_text
        )
    )

    return df


# ============================================================
# 搜尋詞擴充
# ============================================================

def expand_search_word(
    keyword
):

    keyword = clean_text(
        keyword
    )

    if not keyword:
        return ""

    synonym_df = load_synonyms()

    words = [keyword]

    if (
        not synonym_df.empty
        and len(synonym_df.columns) >= 2
    ):

        a, b = synonym_df.columns[:2]

        for _, row in synonym_df.iterrows():

            left = clean_text(
                row[a]
            )

            right = clean_text(
                row[b]
            )

            if not left and not right:
                continue

            keyword_normalized = normalize_search_text(
                keyword
            )

            left_normalized = normalize_search_text(
                left
            )

            right_normalized = normalize_search_text(
                right
            )

            if (
                left_normalized
                and keyword_normalized
                in left_normalized
            ):

                words.extend(
                    re.split(
                        r"[,，、/／\s]+",
                        right
                    )
                )

            elif (
                right_normalized
                and keyword_normalized
                in right_normalized
            ):

                words.append(
                    left
                )

    return " ".join(
        dict.fromkeys([
            x for x in words
            if clean_text(x)
        ])
    )


# ============================================================
# 搜尋
# ============================================================

def split_search_terms(
    keyword
):

    keyword = clean_text(
        keyword
    )

    if not keyword:
        return []

    stop_words = {
        "請問",
        "想問",
        "想詢問",
        "請教",
        "麻煩問一下",
        "可以請問",
        "可以",
        "嗎",
        "呢",
        "如何",
        "怎麼",
        "怎樣",
        "什麼",
        "是否",
        "有沒有",
        "有無",
        "我要",
        "我想",
        "需要",
        "一下",
        "問題",
        "請",
        "幫我",
        "告訴我",
        "想知道",
        "請幫忙",
        "謝謝",
        "謝謝您"
    }

    parts = re.split(
        r"[\s,，、/／+＋&＆|｜;；]+",
        keyword
    )

    terms = []

    for part in parts:

        part = clean_text(
            part
        )

        if not part:
            continue

        normalized = normalize_search_text(
            part
        )

        if not normalized:
            continue

        for stop_word in stop_words:

            normalized_stop = normalize_search_text(
                stop_word
            )

            if normalized == normalized_stop:

                normalized = ""
                break

        if normalized:
            terms.append(
                normalized
            )

    if not terms:

        normalized = normalize_search_text(
            keyword
        )

        if normalized:
            terms.append(
                normalized
            )

    return list(
        dict.fromkeys(terms)
    )


def build_search_terms(
    keyword
):

    original = clean_text(
        keyword
    )

    if not original:
        return []

    terms = split_search_terms(
        original
    )

    expanded_terms = []
    expanded_terms.extend(
        terms
    )

    for term in terms:

        expanded = expand_search_word(
            term
        )

        if not expanded:
            continue

        expanded_parts = re.split(
            r"[\s,，、/／+＋&＆|｜;；]+",
            expanded
        )

        for part in expanded_parts:

            normalized = normalize_search_text(
                part
            )

            if normalized:
                expanded_terms.append(
                    normalized
                )

    normalized_original = normalize_search_text(
        original
    )

    if normalized_original:
        expanded_terms.append(
            normalized_original
        )

    return list(
        dict.fromkeys([
            x for x in expanded_terms
            if x
        ])
    )


def calculate_fuzzy_similarity(
    term,
    value
):

    term = normalize_search_text(
        term
    )

    value = normalize_search_text(
        value
    )

    if not term or not value:
        return 0.0

    if term in value:
        return 1.0

    if len(term) < 2:
        return 0.0

    if len(value) <= len(term) + 3:

        return SequenceMatcher(
            None,
            term,
            value
        ).ratio()

    window_size = len(term)
    best_score = 0.0

    max_start = (
        len(value)
        - window_size
        + 1
    )

    if max_start > 120:

        positions = list(
            range(
                0,
                max_start,
                max(
                    1,
                    max_start // 120
                )
            )
        )

        if (
            positions
            and positions[-1]
            != max_start - 1
        ):

            positions.append(
                max_start - 1
            )

    else:

        positions = range(
            max_start
        )

    for i in positions:

        piece = value[
            i:i + window_size
        ]

        score = SequenceMatcher(
            None,
            term,
            piece
        ).ratio()

        if score > best_score:
            best_score = score

        if best_score >= 0.92:
            break

    return best_score


def search_knowledge(
    keyword,
    top_n=10
):

    df = load_data()

    if df.empty:
        return pd.DataFrame()

    fields = [
        "問題類型",
        "問題細項",
        "商品分類代碼",
        "商品名稱",
        "商品系列",
        "商品編號",
        "問題摘要",
        "標準回覆",
        "判斷依據",
        "資料來源",
        "關鍵字"
    ]

    fields = [
        c for c in fields
        if c in df.columns
    ]

    if not fields:
        return pd.DataFrame()

    original_keyword = clean_text(
        keyword
    )

    if not original_keyword:

        return df.head(
            top_n
        ).copy()

    search_terms = build_search_terms(
        original_keyword
    )

    if not search_terms:

        return df.head(
            top_n
        ).copy()

    normalized_query = normalize_search_text(
        original_keyword
    )

    expanded_query = normalize_search_text(
        expand_search_word(
            original_keyword
        )
    )

    weights = {
        "問題摘要": 100,
        "關鍵字": 90,
        "商品名稱": 80,
        "商品系列": 70,
        "商品編號": 70,
        "商品分類代碼": 60,
        "問題類型": 40,
        "問題細項": 40,
        "標準回覆": 30,
        "判斷依據": 30,
        "資料來源": 20
    }

    corpus = []

    for _, row in df.iterrows():

        parts = []

        for field in fields:

            value = normalize_search_text(
                row.get(
                    field,
                    ""
                )
            )

            if not value:
                continue

            weight = weights.get(
                field,
                20
            )

            repeat_count = max(
                1,
                int(weight / 20)
            )

            parts.extend(
                [value] * repeat_count
            )

        corpus.append(
            " ".join(parts)
        )

    vectorizer = TfidfVectorizer(
        analyzer="char",
        ngram_range=(1, 3),
        min_df=1
    )

    query_for_tfidf = (
        expanded_query
        if expanded_query
        else normalized_query
    )

    matrix = vectorizer.fit_transform(
        corpus + [query_for_tfidf]
    )

    scores = cosine_similarity(
        matrix[-1],
        matrix[:-1]
    ).ravel()

    result = df.copy()

    result["_score"] = scores

    meaningful_terms = []

    for term in search_terms:

        if len(term) > 12:
            continue

        if term not in meaningful_terms:

            meaningful_terms.append(
                term
            )

    if not meaningful_terms:

        meaningful_terms = [
            normalized_query
        ]

    term_coverage_scores = []
    term_quality_scores = []
    fuzzy_scores = []
    direct_scores = []

    for _, row in result.iterrows():

        row_term_scores = []
        exact_hit_count = 0
        fuzzy_total = 0.0
        direct_score = 0.0

        for term in meaningful_terms:

            if not term:
                continue

            best_exact_field_score = 0.0
            best_fuzzy_field_score = 0.0

            for field in fields:

                value = normalize_search_text(
                    row.get(
                        field,
                        ""
                    )
                )

                if not value:
                    continue

                field_weight = weights.get(
                    field,
                    20
                )

                if term in value:

                    weighted_exact = (
                        field_weight / 100
                    )

                    if (
                        weighted_exact
                        > best_exact_field_score
                    ):

                        best_exact_field_score = (
                            weighted_exact
                        )

                fuzzy = calculate_fuzzy_similarity(
                    term,
                    value
                )

                if fuzzy >= 0.62:

                    weighted_fuzzy = (
                        fuzzy
                        * (
                            0.5
                            + (
                                field_weight
                                / 100
                            ) * 0.5
                        )
                    )

                    if (
                        weighted_fuzzy
                        > best_fuzzy_field_score
                    ):

                        best_fuzzy_field_score = (
                            weighted_fuzzy
                        )

            best_term_score = max(
                best_exact_field_score,
                best_fuzzy_field_score
            )

            row_term_scores.append(
                best_term_score
            )

            if best_exact_field_score > 0:

                exact_hit_count += 1

                direct_score += (
                    best_exact_field_score
                )

            fuzzy_total += best_term_score

        total_terms = max(
            1,
            len(row_term_scores)
        )

        matched_terms = sum(
            1
            for score in row_term_scores
            if score >= 0.25
        )

        coverage = (
            matched_terms
            / total_terms
        )

        term_quality = (
            sum(row_term_scores)
            / total_terms
        )

        fuzzy_average = (
            fuzzy_total
            / total_terms
        )

        multi_exact_bonus = min(
            0.20,
            exact_hit_count * 0.05
        )

        term_coverage_scores.append(
            min(
                1.0,
                coverage
                + multi_exact_bonus
            )
        )

        term_quality_scores.append(
            min(
                1.0,
                term_quality
            )
        )

        fuzzy_scores.append(
            min(
                1.0,
                fuzzy_average
            )
        )

        direct_scores.append(
            min(
                1.0,
                direct_score
                / max(
                    1,
                    len(
                        meaningful_terms
                    )
                )
            )
        )

    result["_term_coverage"] = (
        term_coverage_scores
    )

    result["_term_quality"] = (
        term_quality_scores
    )

    result["_fuzzy_score"] = (
        fuzzy_scores
    )

    result["_direct_score"] = (
        direct_scores
    )

    full_query_scores = []

    for _, row in result.iterrows():

        full_match_score = 0.0

        for field in fields:

            value = normalize_search_text(
                row.get(
                    field,
                    ""
                )
            )

            if not value:
                continue

            field_weight = weights.get(
                field,
                20
            )

            if (
                normalized_query
                and normalized_query in value
            ):

                full_match_score = max(
                    full_match_score,
                    field_weight / 100
                )

            if (
                expanded_query
                and expanded_query in value
            ):

                full_match_score = max(
                    full_match_score,
                    field_weight / 100
                )

        full_query_scores.append(
            min(
                1.0,
                full_match_score
            )
        )

    result["_full_query_score"] = (
        full_query_scores
    )

    result["_final_score"] = (
        result["_score"] * 0.40
        + result["_term_coverage"] * 0.25
        + result["_term_quality"] * 0.15
        + result["_direct_score"] * 0.10
        + result["_fuzzy_score"] * 0.10
    )

    result["_final_score"] = (
        result["_final_score"]
        + result["_full_query_score"] * 0.10
    )

    result["_final_score"] = (
        result["_final_score"]
        .clip(upper=1.0)
    )

    result = result.sort_values(
        [
            "_final_score",
            "_term_coverage",
            "_score"
        ],
        ascending=False
    )

    return result.head(
        top_n
    )


# ============================================================
# 客服小幫手
# ============================================================

def ask_customer_assistant(
    question
):

    result = search_knowledge(
        question,
        top_n=5
    )

    if result.empty:
        return False, result

    best_score = float(
        result.iloc[0].get(
            "_final_score",
            0
        )
    )

    MIN_SCORE = 0.12

    if best_score < MIN_SCORE:
        return False, result

    return True, result


# ============================================================
# 側邊欄
# ============================================================

with st.sidebar:

    # --------------------------------------------------------
    # 品牌標題
    # --------------------------------------------------------
    st.markdown(
        """
        <div style="
            padding: 12px 8px 22px 8px;
            text-align: center;
        ">
            <div style="font-size: 42px; margin-bottom: 5px;">
                🐱
            </div>
            <div style="
                color: #514C45;
                font-size: 21px;
                font-weight: 700;
                letter-spacing: 1px;
            ">
                客服作業系統
            </div>
            <div style="
                color: #9A8F84;
                font-size: 11px;
                margin-top: 7px;
                letter-spacing: 1px;
            ">
                NITORI · CUSTOMER SUPPORT
            </div>
            <div style="
                width: 42px;
                height: 3px;
                background: #B9785E;
                border-radius: 5px;
                margin: 15px auto 0 auto;
            "></div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div style="
            color: #9A8F84;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: 2px;
            padding: 0 8px 8px 8px;
        ">
            MENU / 功能選單
        </div>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # 功能選單
    # --------------------------------------------------------
    menu_items = [
        "🏠 首頁",
        "🔍 知識搜尋",
        "📊 客訴資料統計",
        "🗺️ 門市地圖",
        "🎓 新人訓練中心",
        "📢 公告",
        "🛠️ 資料維護"
    ]

    if "current_page" not in st.session_state:
        st.session_state["current_page"] = "🏠 首頁"

    for item in menu_items:

        is_current = (
            st.session_state["current_page"] == item
        )

        if is_current:

            st.markdown(
                f"""
                <div style="
                    background: linear-gradient(
                        100deg, #EAD6C9 0%, #F5EAE1 100%
                    );
                    border: 1px solid #D9B6A3;
                    border-left: 4px solid #B9785E;
                    border-radius: 12px;
                    padding: 12px;
                    margin: 5px 0;
                    color: #754D3D;
                    font-weight: 700;
                    font-size: 14px;
                    box-shadow: 0 2px 5px rgba(90,70,50,0.04);
                ">
                    {item}
                    <span style="
                        float: right;
                        color: #B9785E;
                        font-size: 12px;
                    ">●</span>
                </div>
                """,
                unsafe_allow_html=True
            )

        else:

            if st.button(
                item,
                key=f"menu_{item}",
                use_container_width=True
            ):
                st.session_state["current_page"] = item
                st.rerun()

    # --------------------------------------------------------
    # 側邊欄頁尾
    # --------------------------------------------------------
    st.markdown(
        """
        <div style="
            margin-top: 22px;
            border-top: 1px solid #D8CFC2;
            padding-top: 15px;
            text-align: center;
            color: #93877B;
            font-size: 12px;
            line-height: 1.9;
        ">
            <div style="font-size: 20px;">🐾</div>
            讓小貓陪你一起找資料
            <div style="font-size: 10px; color: #AAA094;">
                A LITTLE HELP, EVERY DAY
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ------------------------------------------------------------
# 目前頁面
# ------------------------------------------------------------
page = st.session_state["current_page"]

# ============================================================
# 修改歷程顯示
# ============================================================

def render_change_history(
    data_area,
    key_prefix
):

    st.divider()

    with st.expander(
        f"📜 查看「{data_area}」修改歷程",
        expanded=False
    ):

        st.caption(
            "只顯示此資料區域的新增、修改與刪除紀錄。"
        )

        if not EXCEL_PATH.exists():

            st.info(
                "找不到 Excel 檔案。"
            )

            return

        try:

            sheets = get_excel_sheets()

            if "修改歷程" not in sheets:

                st.info(
                    "目前還沒有修改歷程。第一次修改並儲存後會自動建立。"
                )

                return

            history = pd.read_excel(
                EXCEL_PATH,
                sheet_name="修改歷程",
                dtype=str
            ).fillna("")

            for col in HISTORY_COLUMNS:

                if col not in history.columns:
                    history[col] = ""

            history = history[
                HISTORY_COLUMNS
            ].copy()

            history = history[
                history["資料區域"]
                .astype(str)
                .map(clean_text)
                == data_area
            ].copy()

            if history.empty:

                st.info(
                    f"「{data_area}」目前還沒有修改紀錄。"
                )

                return

            filter_col1, filter_col2 = st.columns(2)

            with filter_col1:

                action_options = sorted([
                    x
                    for x in history["動作"].unique()
                    if clean_text(x)
                ])

                selected_action = st.selectbox(
                    "動作",
                    ["全部"] + action_options,
                    key=f"{key_prefix}_action_filter"
                )

            with filter_col2:

                editor_options = sorted([
                    x
                    for x in history["修改人"].unique()
                    if clean_text(x)
                ])

                selected_editor = st.selectbox(
                    "修改人",
                    ["全部"] + editor_options,
                    key=f"{key_prefix}_editor_filter"
                )

            search_history = st.text_input(
                "🔎 搜尋此區域的修改紀錄",
                placeholder="可搜尋商品編號、問題、欄位名稱或修改內容……",
                key=f"{key_prefix}_search"
            )

            display_history = history.copy()

            if selected_action != "全部":

                display_history = display_history[
                    display_history["動作"]
                    == selected_action
                ]

            if selected_editor != "全部":

                display_history = display_history[
                    display_history["修改人"]
                    == selected_editor
                ]

            if search_history.strip():

                search_value = normalize_search_text(
                    search_history
                )

                mask = pd.Series(
                    False,
                    index=display_history.index
                )

                for col in HISTORY_COLUMNS:

                    mask = (
                        mask
                        | display_history[col]
                        .astype(str)
                        .map(
                            normalize_search_text
                        )
                        .str.contains(
                            search_value,
                            na=False
                        )
                    )

                display_history = display_history[
                    mask
                ]

            display_history = (
                display_history
                .iloc[::-1]
                .reset_index(drop=True)
            )

            st.caption(
                f"目前顯示 {len(display_history)} 筆「{data_area}」修改紀錄"
            )

            if display_history.empty:

                st.info(
                    "沒有符合條件的修改紀錄。"
                )

            else:

                st.dataframe(
                    display_history,
                    use_container_width=True,
                    hide_index=True,
                    height=450
                )

        except Exception as e:

            st.error(
                f"讀取「{data_area}」修改歷程失敗：{e}"
            )
def render_announcement_history():
    """從 Supabase 讀取公告修改歷程。"""

    st.markdown("### 📝 公告修改歷程")

    try:
        response = (
            supabase
            .table("announcement_history")
            .select("*")
            .order("changed_at", desc=True)
            .execute()
        )

        records = response.data or []

        if not records:
            st.info("目前尚無公告修改紀錄。")
            return

        history_df = pd.DataFrame(records)

        column_mapping = {
            "changed_at": "修改時間",
            "editor_name": "修改人",
            "action": "動作",
            "announcement_id": "公告編號",
            "announcement_title": "公告標題",
            "field_name": "欄位名稱",
            "old_value": "修改前",
            "new_value": "修改後"
        }

        history_df = history_df.rename(
            columns=column_mapping
        )

        display_columns = [
            "修改時間",
            "修改人",
            "動作",
            "公告編號",
            "公告標題",
            "欄位名稱",
            "修改前",
            "修改後"
        ]

        display_columns = [
            col for col in display_columns
            if col in history_df.columns
        ]

        st.dataframe(
            history_df[display_columns],
            use_container_width=True,
            hide_index=True
        )

    except Exception as e:
        st.error(f"讀取公告修改歷程失敗：{e}")


# ============================================================
# 首頁
# ============================================================

if page == "🏠 首頁":

    # --------------------------------------------------------
    # 依照目前時間顯示問候語
    # --------------------------------------------------------
    current_hour = datetime.now().hour

    if current_hour < 12:
        greeting = "🌅 早安，今天也辛苦了！"
        greeting_subtitle = "新的一天，慢慢開始今天的工作吧。"

    elif current_hour < 18:
        greeting = "☀️ 午安，辛苦了一半，繼續加油！"
        greeting_subtitle = "喝口水、整理一下思緒，繼續完成今天的任務。"

    else:
        greeting = "🌙 辛苦了，今天的工作也接近尾聲囉！"
        greeting_subtitle = "整理好今天的工作，明天也能更順利地開始。"

    # --------------------------------------------------------
    # 首頁歡迎區
    # --------------------------------------------------------
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(120deg, #FFFDF9 0%, #F0E4D7 100%);
            border: 1px solid #E4D9CB;
            border-radius: 20px;
            padding: 30px 30px 26px 30px;
            margin-bottom: 24px;
            box-shadow: 0 4px 16px rgba(90,70,50,0.04);
        ">
            <div style="
                font-size: 12px;
                color: #9A8F84;
                letter-spacing: 2px;
                margin-bottom: 10px;
            ">
                NITORI · CUSTOMER SUPPORT
            </div>
            <div style="
                font-size: 27px;
                font-weight: 700;
                color: #514C45;
                line-height: 1.5;
                margin-bottom: 8px;
            ">
                {greeting}
            </div>
            <div style="
                font-size: 15px;
                line-height: 1.8;
                color: #898176;
            ">
                {greeting_subtitle}<br>
                歡迎回到客服作業系統，讓小貓陪你一起處理今天的工作 🐾
            </div>
            <div style="
                margin-top: 18px;
                color: #B9785E;
                font-size: 13px;
                font-weight: 600;
            ">
                ♡ 細心聆聽・正確查詢・溫暖回覆
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # 每日客服小提醒
    # --------------------------------------------------------
    daily_tips = [
        "遇到不確定的問題，先確認資料，再回覆顧客。",
        "先理解顧客的需求，再確認需要查詢的資訊。",
        "回覆顧客前，記得確認商品編號與商品名稱。",
        "遇到無法立即確認的問題，可以先整理資訊，再進行內部確認。",
        "客服交接時，資訊越完整，下一位同仁就越容易接手。",
        "顧客情緒較強烈時，先傾聽並確認問題，再說明處理方式。",
        "遇到規則不確定的案件，請先確認最新公告或內部資訊。",
        "回覆顧客時，除了說明目前狀況，也可以告知顧客下一步該怎麼做。",
        "商品問題除了確認商品本身，也別忘了確認購買通路與購買時間。",
        "電話結束前，確認顧客是否還有其他需要協助的事項。"
    ]

    today_index = (
        datetime.now().timetuple().tm_yday
        % len(daily_tips)
    )

    st.markdown(
        """
        <div style="
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 8px 0 12px 0;
        ">
            <span style="font-size: 23px;">🌱</span>
            <span style="
                font-size: 20px;
                font-weight: 700;
                color: #514C45;
            ">每日客服小提醒</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div style="
            background: #F2E8DB;
            border-left: 4px solid #B9785E;
            border-radius: 12px;
            padding: 17px 20px;
            margin-bottom: 28px;
            color: #6B5B4E;
            font-size: 15px;
            line-height: 1.8;
        ">
            🐾 {daily_tips[today_index]}
        </div>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # 最新公告
    # --------------------------------------------------------
    st.markdown(
        """
        <div style="
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 8px 0 14px 0;
        ">
            <span style="font-size: 23px;">📢</span>
            <span style="
                font-size: 20px;
                font-weight: 700;
                color: #514C45;
            ">最新公告</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    ann = load_announcements()

    if ann.empty:
        st.info("目前沒有公告資料。")

    else:
        announcement_columns = get_announcement_columns(ann)

        for _, row in ann.head(5).iterrows():
            render_announcement(
                row,
                announcement_columns
            )

# ============================================================
# 知識搜尋
# ============================================================

elif page == "🔍 知識搜尋":

    st.title(
        "🔍 知識庫搜尋"
    )

    st.caption(
        "輸入顧客詢問的關鍵字，小貓會幫你找出相關資料。"
    )

    st.markdown(
        "### 找找看顧客問題"
    )

    keyword = st.text_input(
        "搜尋關鍵字",
        placeholder="輸入顧客問題，例如：床墊有味道、床包尺寸、退貨方式……",
        label_visibility="collapsed"
    )

    top_n = st.slider(
        "顯示筆數",
        5,
        30,
        10
    )

    if keyword.strip():

        result = search_knowledge(
            keyword,
            top_n
        )

        if result.empty:

            st.warning(
                "目前沒有找到相關資料。"
            )

        else:

            st.caption(
                f"共找到 {len(result)} 筆相關資料"
            )

            for _, row in result.iterrows():

                title = clean_text(
                    row.get(
                        "問題摘要",
                        "未命名問題"
                    )
                )

                score = float(
                    row.get(
                        "_final_score",
                        0
                    )
                )

                meta_parts = []

                for c in [
                    "問題類型",
                    "問題細項",
                    "商品名稱",
                    "商品系列",
                    "商品編號"
                ]:

                    value = clean_text(
                        row.get(
                            c,
                            ""
                        )
                    )

                    if value:
                        meta_parts.append(
                            value
                        )

                meta = "　｜　".join(
                    meta_parts
                )

                answer = clean_text(
                    row.get(
                        "標準回覆",
                        ""
                    )
                )

                judgment = clean_text(
                    row.get(
                        "判斷依據",
                        ""
                    )
                )

                source = clean_text(
                    row.get(
                        "資料來源",
                        ""
                    )
                )

                with st.container(
                    border=True
                ):

                    st.subheader(
                        title
                    )

                    if meta:

                        st.caption(
                            meta
                            + f"　｜　相關度 {score:.1%}"
                        )

                    else:

                        st.caption(
                            f"相關度 {score:.1%}"
                        )

                    st.markdown(
                        "**💬 標準回覆**"
                    )

                    st.write(
                        answer
                    )

                    if judgment:

                        st.markdown(
                            "**🔎 判斷依據**"
                        )

                        st.write(
                            judgment
                        )

                    if source:

                        st.markdown(
                            "**📚 資料來源**"
                        )

                        st.write(
                            source
                        )

    else:

        st.info(
            "🐾 輸入關鍵字開始搜尋"
        )


# ============================================================
# 客訴資料統計
# ============================================================

elif page == "📊 客訴資料統計":

    st.title(
        "📊 客訴資料統計"
    )

    st.caption(
        "整理客訴資料，查看客訴趨勢、類型與細項分布。"
    )

    complaints = load_complaints()

    if complaints.empty:

        st.info(
            "目前尚無客訴資料，請確認 Excel 中是否已建立「客訴資料」工作表。"
        )

    else:

        st.subheader(
            "🔎 資料篩選"
        )

        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)

        month_options = sorted(
            [
                x
                for x in complaints["_月份"].dropna().unique()
                if clean_text(x)
            ],
            reverse=True
        )

        with filter_col1:

            selected_month = st.selectbox(
                "月份",
                ["全部"] + month_options
            )

        channel_options = sorted(
            [
                x
                for x in complaints["購買通路"].dropna().unique()
                if clean_text(x)
            ]
        )

        with filter_col2:

            selected_channel = st.selectbox(
                "購買通路",
                ["全部"] + channel_options
            )

        type_options = sorted(
            [
                x
                for x in complaints["客訴類型"].dropna().unique()
                if clean_text(x)
            ]
        )

        with filter_col3:

            selected_type = st.selectbox(
                "客訴類型",
                ["全部"] + type_options
            )

        detail_options = sorted(
            [
                x
                for x in complaints["客訴細項"].dropna().unique()
                if clean_text(x)
            ]
        )

        with filter_col4:

            selected_detail = st.selectbox(
                "客訴細項",
                ["全部"] + detail_options
            )

        filtered = complaints.copy()

        if selected_month != "全部":

            filtered = filtered[
                filtered["_月份"]
                == selected_month
            ]

        if selected_channel != "全部":

            filtered = filtered[
                filtered["購買通路"]
                == selected_channel
            ]

        if selected_type != "全部":

            filtered = filtered[
                filtered["客訴類型"]
                == selected_type
            ]

        if selected_detail != "全部":

            filtered = filtered[
                filtered["客訴細項"]
                == selected_detail
            ]

        st.divider()

        total_count = len(
            filtered
        )

        current_month = datetime.now().strftime(
            "%Y/%m"
        )

        current_month_count = len(
            complaints[
                complaints["_月份"]
                == current_month
            ]
        )

        product_count = len(
            filtered[
                filtered["客訴類型"]
                .astype(str)
                .str.contains(
                    "商品",
                    na=False
                )
            ]
        )

        other_count = (
            total_count
            - product_count
        )

        card1, card2, card3, card4 = st.columns(4)

        with card1:

            st.metric(
                "📋 篩選後客訴",
                f"{total_count} 件"
            )

        with card2:

            st.metric(
                "📅 本月客訴",
                f"{current_month_count} 件"
            )

        with card3:

            st.metric(
                "📦 商品相關",
                f"{product_count} 件"
            )

        with card4:

            st.metric(
                "📝 其他類型",
                f"{other_count} 件"
            )

        st.divider()

        st.subheader(
            "📈 每月客訴趨勢"
        )

        monthly_data = (
            filtered[
                filtered["_月份"]
                .astype(str)
                .str.strip()
                != ""
            ]
            .groupby("_月份")
            .size()
            .reset_index(
                name="客訴件數"
            )
        )

        if monthly_data.empty:

            st.info(
                "目前沒有足夠的月份資料可以統計。"
            )

        else:

            monthly_data = monthly_data.sort_values(
                "_月份"
            )

            monthly_chart = (
                monthly_data.set_index(
                    "_月份"
                )
            )

            st.line_chart(
                monthly_chart["客訴件數"],
                use_container_width=True
            )

        st.divider()

        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:

            st.subheader(
                "📊 客訴類型"
            )

            type_data = (
                filtered[
                    filtered["客訴類型"]
                    .astype(str)
                    .str.strip()
                    != ""
                ]
                .groupby("客訴類型")
                .size()
                .reset_index(
                    name="件數"
                )
                .sort_values(
                    "件數",
                    ascending=False
                )
            )

            if type_data.empty:

                st.info(
                    "目前沒有客訴類型資料。"
                )

            else:

                type_chart = (
                    type_data.set_index(
                        "客訴類型"
                    )
                )

                st.bar_chart(
                    type_chart["件數"],
                    use_container_width=True
                )

        with chart_col2:

            st.subheader(
                "📊 客訴細項"
            )

            detail_data = (
                filtered[
                    filtered["客訴細項"]
                    .astype(str)
                    .str.strip()
                    != ""
                ]
                .groupby("客訴細項")
                .size()
                .reset_index(
                    name="件數"
                )
                .sort_values(
                    "件數",
                    ascending=False
                )
            )

            if detail_data.empty:

                st.info(
                    "目前沒有客訴細項資料。"
                )

            else:

                detail_chart = (
                    detail_data.set_index(
                        "客訴細項"
                    )
                )

                st.bar_chart(
                    detail_chart["件數"],
                    use_container_width=True
                )

        st.divider()

        st.subheader(
            "🛒 購買通路"
        )

        channel_data = (
            filtered[
                filtered["購買通路"]
                .astype(str)
                .str.strip()
                != ""
            ]
            .groupby("購買通路")
            .size()
            .reset_index(
                name="件數"
            )
            .sort_values(
                "件數",
                ascending=False
            )
        )

        if channel_data.empty:

            st.info(
                "目前沒有購買通路資料。"
            )

        else:

            channel_chart = (
                channel_data.set_index(
                    "購買通路"
                )
            )

            st.bar_chart(
                channel_chart["件數"],
                use_container_width=True
            )

        st.divider()

        st.subheader(
            "📋 客訴資料明細"
        )

        st.caption(
            f"目前顯示 {len(filtered)} 筆資料"
        )

        display_columns = [
            "序號",
            "日期",
            "月份",
            "購買通路",
            "商品編號",
            "商品名稱",
            "客訴類型",
            "客訴細項",
            "客訴內容",
            "處理方式",
            "處理結果",
            "備註"
        ]

        display_columns = [
            c
            for c in display_columns
            if c in filtered.columns
        ]

        st.dataframe(
            filtered[display_columns],
            use_container_width=True,
            hide_index=True,
            height=500
        )


# ============================================================
# 門市地圖
# ============================================================

elif page == "🗺️ 門市地圖":

    st.title(
        "🗺️ 宜得利小地圖"
    )

    st.caption(
        "可使用路名、捷運站、車站名稱、門市名稱或店號搜尋。"
    )

    POWER_BI_URL = (
        "https://app.powerbi.com/reportEmbed?"
        "reportId=851d09f1-f169-4049-9090-640f09cdc385"
        "&autoAuth=true"
        "&ctid=94d67acc-8c17-4d3b-8463-9829cf08ef55"
    )

    components.iframe(
        POWER_BI_URL,
        height=760,
        scrolling=True
    )


# ============================================================
# 新人訓練中心
# ============================================================

elif page == "🎓 新人訓練中心":
    # ========================================================
    # 身分選擇
    # ========================================================

    role = st.radio(
        "請選擇訓練身分",
        [
            "👤 新人",
            "👩‍🏫 小老師"
        ],
        horizontal=True,
        key="training_role"
    )
    # ========================================================
    # 小老師管理
    # ========================================================

    if role == "👩‍🏫 小老師":

        TEACHER_PASSWORD = "nitori123"

        if not st.session_state.get(
            "teacher_authenticated",
            False
        ):

            st.subheader(
                "🔒 小老師管理"
            )

            st.write(
                "此頁面為小老師管理專用，請輸入密碼。"
            )

            teacher_password = st.text_input(
                "請輸入密碼",
                type="password",
                key="teacher_password_input"
            )

            if st.button(
                "🔓 登入",
                key="teacher_login_button"
            ):

                if teacher_password == TEACHER_PASSWORD:

                    st.session_state[
                        "teacher_authenticated"
                    ] = True

                    st.rerun()

                else:

                    st.error(
                        "❌ 密碼錯誤，請重新輸入。"
                    )

            st.stop()

        st.markdown("""
        <div style="
            background: #F5F1E9;
            padding: 24px 28px;
            border-radius: 12px;
            border-left: 5px solid #B49A78;
            margin-bottom: 20px;
        ">
            <div style="
                color: #8A7863;
                font-size: 12px;
                letter-spacing: 3px;
                margin-bottom: 8px;
            ">NITORI · TRAINING</div>
            <div style="
                color: #403B34;
                font-size: 27px;
                font-weight: 600;
                margin-bottom: 8px;
            ">新人訓練管理</div>
            <div style="
                color: #777067;
                font-size: 14px;
            ">掌握學習進度，陪伴每位新人穩定成長。</div>
        </div>
        """, unsafe_allow_html=True)


        if st.button(
            "🚪 登出小老師",
            key="teacher_logout_button"
        ):

            st.session_state[
                "teacher_authenticated"
            ] = False

            st.rerun()

        st.write(
            "新人訓練管理"
        )
        # ========================================
        # 📊 訓練總覽統計
        # ========================================

        st.divider()
        st.markdown("### 📊 訓練總覽")

        if employee_data.empty:
            dashboard_staff = employee_data.copy()
        else:
            dashboard_staff = employee_data[
                employee_data["角色"]
                .astype(str)
                .str.strip()
                == "新人"
            ].copy()

        dashboard_records = load_training_records()

        dashboard_course_ids = [
            "DAY1_BASIC",
            "DAY1_SERVICE",
            "DAY2_PRODUCT",
            "DAY2_SYSTEM",
            "DAY3_PHONE",
            "DAY3_PRACTICE",
        ]

        total_staff = len(dashboard_staff)
        not_started_count = 0
        in_progress_count = 0
        completed_count = 0

        for _, employee in dashboard_staff.iterrows():

            employee_id = str(
                employee["員工編號"]
            ).strip()

            if dashboard_records.empty:
                employee_records = pd.DataFrame()
            else:
                employee_records = dashboard_records[
                    dashboard_records["員工編號"]
                    .astype(str)
                    .str.strip()
                    == employee_id
                ]

            statuses = []

            for course_id in dashboard_course_ids:

                course_record = employee_records[
                    employee_records["課程ID"]
                    .astype(str)
                    .str.strip()
                    == course_id
                ] if not employee_records.empty else pd.DataFrame()

                if course_record.empty:
                    status = "尚未開始"
                else:
                    status = str(
                        course_record.iloc[-1]["學習狀態"]
                    ).strip()

                statuses.append(status)

            if all(
                status == "已完成"
                for status in statuses
            ):
                completed_count += 1

            elif any(
                status in ["學習中", "已完成"]
                for status in statuses
            ):
                in_progress_count += 1

            else:
                not_started_count += 1

        stat1, stat2, stat3, stat4 = st.columns(4, gap="medium")

        stat_cards = [
            (stat1, "新人總數", total_staff, "位新人", "#F5F1E9", "#A18A6B"),
            (stat2, "尚未開始", not_started_count, "位新人", "#F3F0EC", "#9A9186"),
            (stat3, "訓練中", in_progress_count, "位新人", "#F5EBDD", "#B18A57"),
            (stat4, "已完成", completed_count, "位新人", "#E9EFE7", "#71866B"),
        ]

        for column, label, value, unit, bg_color, accent_color in stat_cards:
            with column:
                st.markdown(
                    f"""
                    <div style="
                        background: {bg_color};
                        border-radius: 12px;
                        padding: 20px 18px;
                        min-height: 118px;
                        border: 1px solid #E8E1D7;
                    ">
                        <div style="
                            color: #746D63;
                            font-size: 13px;
                            margin-bottom: 14px;
                        ">{label}</div>
                        <div style="
                            color: #403B34;
                            font-size: 30px;
                            font-weight: 600;
                            line-height: 1.2;
                        ">{value}
                            <span style="
                                color: {accent_color};
                                font-size: 13px;
                                font-weight: 400;
                            ">{unit}</span>
                        </div>
                        <div style="
                            background: {accent_color};
                            width: 28px;
                            height: 3px;
                            border-radius: 3px;
                            margin-top: 14px;
                        "></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        if employee_data.empty:

            st.info(
                "目前尚未建立員工資料。"
            )

        else:

            new_staff = employee_data[
                employee_data["角色"]
                .astype(str)
                .str.strip()
                == "新人"
            ].copy()

            if new_staff.empty:

                st.info(
                    "目前尚未建立新人資料。"
                )

            else:

                display_columns = [
                    "員工編號",
                    "姓名",
                    "職稱",
                    "部門",
                    "到職日"
                ]

                display_columns = [
                    column
                    for column in display_columns
                    if column in new_staff.columns
                ]

                # ========================================
                # ✨ 新人卡片式名單與學習進度
                # ========================================

                training_records = load_training_records()

                course_names = {
                    "DAY1_BASIC": "🏠 宜得利基礎",
                    "DAY1_SERVICE": "💬 接客基礎",
                    "DAY2_PRODUCT": "🛋 商品基礎",
                    "DAY2_SYSTEM": "💻 系統操作",
                    "DAY3_PHONE": "☎️ 電話應對",
                    "DAY3_PRACTICE": "📝 實務練習",
                }
                st.markdown("""
                <div style="margin: 26px 0 12px 0;">
                    <div style="
                        color: #8A7863;
                        font-size: 11px;
                        letter-spacing: 2px;
                    ">MEMBER DIRECTORY</div>
                    <div style="
                        color: #403B34;
                        font-size: 21px;
                        font-weight: 600;
                        margin-top: 4px;
                    ">新人學習名單</div>
                    <div style="
                        color: #888077;
                        font-size: 13px;
                        margin-top: 4px;
                    ">搜尋新人資料，快速掌握目前的訓練狀態。</div>
                </div>
                """, unsafe_allow_html=True)

                search_name = st.text_input(
                    "🔍 搜尋新人",
                    placeholder="輸入姓名或員工編號",
                    key="training_staff_search"
                )

                status_filter = st.selectbox(
                    "📋 訓練狀態",
                    ["全部", "尚未開始", "訓練中", "已完成訓練"],
                    key="training_staff_status_filter"
                )

                filtered_staff = new_staff.copy()

                if search_name.strip():
                    keyword = search_name.strip()
                    filtered_staff = filtered_staff[
                        filtered_staff["姓名"].astype(str).str.contains(
                            keyword, case=False, na=False
                        )
                        | filtered_staff["員工編號"].astype(str).str.contains(
                            keyword, case=False, na=False
                        )
                    ]

                if status_filter != "全部":
                    matched_employee_ids = []

                    for _, filter_employee in filtered_staff.iterrows():
                        filter_id = str(filter_employee["員工編號"]).strip()

                        if training_records.empty:
                            filter_records = pd.DataFrame()
                        else:
                            filter_records = training_records[
                                training_records["員工編號"].astype(str).str.strip()
                                == filter_id
                            ]

                        completed = 0
                        in_progress = 0

                        for course_id in course_names:
                            course_record = filter_records[
                                filter_records["課程ID"].astype(str).str.strip()
                                == course_id
                            ] if not filter_records.empty else pd.DataFrame()

                            if course_record.empty:
                                course_status = "尚未開始"
                            else:
                                course_status = str(
                                    course_record.iloc[-1]["學習狀態"]
                                ).strip()

                            if course_status == "已完成":
                                completed += 1
                            elif course_status == "學習中":
                                in_progress += 1

                        if completed == len(course_names):
                            overall_status = "已完成訓練"
                        elif completed > 0 or in_progress > 0:
                            overall_status = "訓練中"
                        else:
                            overall_status = "尚未開始"

                        if overall_status == status_filter:
                            matched_employee_ids.append(filter_id)

                    filtered_staff = filtered_staff[
                        filtered_staff["員工編號"].astype(str).str.strip().isin(
                            matched_employee_ids
                        )
                    ]

                st.caption(f"目前顯示 {len(filtered_staff)} 位新人")

                if filtered_staff.empty:
                    st.info("找不到符合條件的新人，請調整搜尋關鍵字或篩選條件。")
                else:
                    for _, employee in filtered_staff.iterrows():
                        employee_id = str(
                        employee["員工編號"]
                    ).strip()

                    employee_name = str(
                        employee["姓名"]
                    ).strip()

                    department = str(
                        employee.get("部門", "未設定")
                    ).strip()

                    position = str(
                        employee.get("職稱", "未設定")
                    ).strip()

                    hire_date = str(
                        employee.get("到職日", "未設定")
                    ).strip()

                    if not department or department == "nan":
                        department = "未設定"

                    if not position or position == "nan":
                        position = "未設定"

                    if not hire_date or hire_date == "nan":
                        hire_date = "未設定"

                    # 取得這位新人的學習紀錄
                    if training_records.empty:
                        employee_records = pd.DataFrame()
                    else:
                        employee_records = training_records[
                            training_records["員工編號"]
                            .astype(str)
                            .str.strip()
                            .eq(employee_id)
                        ].copy()

                    course_details = []
                    completed_count = 0
                    in_progress_count = 0

                    for course_id, course_name in course_names.items():

                        course_record = employee_records[
                            employee_records["課程ID"]
                            .astype(str)
                            .str.strip()
                            .eq(course_id)
                        ] if not employee_records.empty else pd.DataFrame()

                        if course_record.empty:
                            status = "尚未開始"
                            quiz_status = "尚未測驗"
                            completed_date = "—"
                            educator_note = "—"
                        else:
                            latest = course_record.iloc[-1]

                            status = str(
                                latest.get("學習狀態", "尚未開始")
                            ).strip()

                            quiz_status = str(
                                latest.get("測驗狀態", "尚未測驗")
                            ).strip()

                            completed_date = str(
                                latest.get("完成日期", "")
                            ).strip()

                            educator_note = str(
                                latest.get("教育者備註", "")
                            ).strip()

                            if not quiz_status or quiz_status == "nan":
                                quiz_status = "尚未測驗"

                            if not completed_date or completed_date == "nan":
                                completed_date = "—"

                            if not educator_note or educator_note == "nan":
                                educator_note = "—"

                        if status == "已完成":
                            completed_count += 1
                        elif status == "學習中":
                            in_progress_count += 1

                        course_details.append({
                            "課程": course_name,
                            "學習狀態": status,
                            "測驗狀態": quiz_status,
                            "完成日期": completed_date,
                            "教育者備註": educator_note,
                        })

                    progress = int(
                        completed_count / len(course_names) * 100
                    )

                    if completed_count == len(course_names):
                        overall_status = "已完成訓練"
                    elif completed_count > 0 or in_progress_count > 0:
                        overall_status = "訓練中"
                    else:
                        overall_status = "尚未開始"

                    with st.container(border=True):

                        st.markdown(
                            """
                            <div style="
                                height: 4px;
                                background: #C5AD8C;
                                border-radius: 4px;
                                margin-bottom: 16px;
                            "></div>
                            """,
                            unsafe_allow_html=True
                        )

                        card_col1, card_col2 = st.columns([4, 1])

                        with card_col1:
                            st.markdown(
                                f"<div style='font-size:22px;"
                                f"font-weight:600;color:#403B34;'>"
                                f"👤 {employee_name}</div>",
                                unsafe_allow_html=True
                            )

                            st.markdown(
                                f"<div style='color:#81786D;"
                                f"font-size:13px;margin-top:8px;'>"
                                f"員工編號　{employee_id}</div>"
                                f"<div style='color:#81786D;"
                                f"font-size:13px;margin-top:5px;'>"
                                f"{department}　／　{position}</div>"
                                f"<div style='color:#81786D;"
                                f"font-size:13px;margin-top:5px;'>"
                                f"到職日　{hire_date}</div>",
                                unsafe_allow_html=True
                            )

                        with card_col2:
                            if overall_status == "已完成訓練":
                                status_bg = "#E8EFE5"
                                status_color = "#607B59"
                            elif overall_status == "訓練中":
                                status_bg = "#F5EBDD"
                                status_color = "#9A7446"
                            else:
                                status_bg = "#F0EEEA"
                                status_color = "#827B72"

                            st.markdown(
                                f"<div style='background:{status_bg};"
                                f"color:{status_color};padding:7px 10px;"
                                f"border-radius:20px;text-align:center;"
                                f"font-size:12px;font-weight:600;"
                                f"margin-top:4px;'>"
                                f"{overall_status}</div>",
                                unsafe_allow_html=True
                            )

                        st.markdown(
                            "<div style='border-top:1px solid #EAE5DD;"
                            "margin:20px 0 14px 0;'></div>",
                            unsafe_allow_html=True
                        )

                        progress_col1, progress_col2 = st.columns([4, 1])

                        with progress_col1:
                            st.markdown(
                                "<div style='color:#6F675D;"
                                "font-size:13px;'>整體訓練進度</div>",
                                unsafe_allow_html=True
                            )

                        with progress_col2:
                            st.markdown(
                                f"<div style='text-align:right;"
                                f"font-size:20px;font-weight:600;"
                                f"color:#8A7863;'>{progress}%</div>",
                                unsafe_allow_html=True
                            )

                        st.progress(progress)

                        st.caption(
                            f"已完成 {completed_count} / 6 堂課"
                        )

                        with st.expander("📚 查看六堂課詳細進度"):

                            for detail in course_details:

                                course_progress = (
                                    100 if detail["學習狀態"] == "已完成"
                                    else 50 if detail["學習狀態"] == "學習中"
                                    else 0
                                )

                                if detail["學習狀態"] == "已完成":
                                    course_bg = "#E8EFE5"
                                    course_color = "#607B59"
                                elif detail["學習狀態"] == "學習中":
                                    course_bg = "#F5EBDD"
                                    course_color = "#9A7446"
                                else:
                                    course_bg = "#F0EEEA"
                                    course_color = "#827B72"

                                st.markdown(
                                    f"""
                                    <div style="
                                        background:#F8F6F1;
                                        border:1px solid #EAE5DD;
                                        border-radius:10px;
                                        padding:14px 16px;
                                        margin:8px 0 10px 0;
                                    ">
                                        <div style="
                                            color:#403B34;
                                            font-size:15px;
                                            font-weight:600;
                                            margin-bottom:10px;
                                        ">{detail['課程']}</div>
                                        <span style="
                                            background:{course_bg};
                                            color:{course_color};
                                            padding:5px 10px;
                                            border-radius:16px;
                                            font-size:12px;
                                        ">{detail['學習狀態']}</span>
                                    </div>
                                    """,
                                    unsafe_allow_html=True
                                )

                                st.progress(course_progress)

                                detail_col1, detail_col2 = st.columns(2)

                                with detail_col1:
                                    st.caption(
                                        f"📝 測驗狀態：{detail['測驗狀態']}"
                                    )

                                with detail_col2:
                                    st.caption(
                                        f"📅 完成日期：{detail['完成日期']}"
                                    )

                                if detail["教育者備註"] != "—":
                                    st.info(
                                        f"💬 教育者備註：{detail['教育者備註']}"
                                    )

                                st.markdown(
                                    "<div style='border-bottom:1px solid "
                                    "#EAE5DD;margin:12px 0;'></div>",
                                    unsafe_allow_html=True
                                )
        st.stop()
    # ========================================================
    # 新人登入
    # ========================================================

    if role == "👤 新人":

        if not st.session_state.get(
            "newcomer_authenticated",
            False
        ):

            st.subheader(
                "👤 新人登入"
            )

            st.write(
                "請輸入您的員工編號"
            )

            employee_id_input = st.text_input(
                "員工編號",
                key="newcomer_employee_id"
            ).strip()

            if st.button(
                "🔓 登入",
                key="newcomer_login_button"
            ):

                employee_match = employee_data[
                    employee_data["員工編號"]
                    .astype(str)
                    .str.strip()
                    == employee_id_input
                ]

                if employee_match.empty:

                    st.error(
                        "❌ 查無此員工編號，請確認後重新輸入。"
                    )

                else:

                    selected_employee = (
                        employee_match.iloc[0]
                    )

                    if str(
                        selected_employee["角色"]
                    ).strip() != "新人":

                        st.error(
                            "此員工編號不是新人身分。"
                        )

                    else:

                        st.session_state[
                            "newcomer_authenticated"
                        ] = True

                        st.session_state[
                            "training_employee_id"
                        ] = str(
                            selected_employee["員工編號"]
                        ).strip()

                        st.session_state[
                            "training_employee_name"
                        ] = str(
                            selected_employee["姓名"]
                        ).strip()
                        # 自動建立新人的六堂課程紀錄
                        current_employee_id = str(
                            st.session_state["training_employee_id"]
                        ).strip()

                        current_employee_name = str(
                            st.session_state["training_employee_name"]
                        ).strip()

                        course_names = {
                            "DAY1_BASIC": "🏠 宜得利基礎",
                            "DAY1_SERVICE": "💬 接客基礎",
                            "DAY2_PRODUCT": "🛋 商品基礎",
                            "DAY2_SYSTEM": "💻 系統操作",
                            "DAY3_PHONE": "☎️ 電話應對",
                            "DAY3_PRACTICE": "📝 實務練習"
                        }

                        existing_records = load_training_records()

                        for course_id, course_name in course_names.items():

                            already_exists = False

                            if not existing_records.empty:
                                already_exists = (
                                    (
                                        existing_records["員工編號"]
                                        .astype(str)
                                        .str.strip()
                                        == current_employee_id
                                    )
                                    &
                                    (
                                        existing_records["課程ID"]
                                        .astype(str)
                                        .str.strip()
                                        == course_id
                                    )
                                ).any()

                            if not already_exists:
                                day = course_id[:4]

                                save_training_record(
                                    employee_id=current_employee_id,
                                    employee_name=current_employee_name,
                                    course_id=course_id,
                                    course_name=course_name,
                                    day=day,
                                    start_date="",
                                    learning_status="尚未開始",
                                    quiz_status="尚未測驗",
                                    completed_date=""
                                )

                                existing_records = load_training_records()

                        st.rerun()

            st.stop()
            

    # --------------------------------------------------------
    # 頁面狀態
    # --------------------------------------------------------

    if "training_page" not in st.session_state:
        st.session_state["training_page"] = "home"

    if "training_step" not in st.session_state:
        st.session_state["training_step"] = "subtopics"

    if "training_course_id" not in st.session_state:
        st.session_state["training_course_id"] = ""

    if "training_content_id" not in st.session_state:
        st.session_state["training_content_id"] = ""

    if "training_quiz_passed" not in st.session_state:
        st.session_state["training_quiz_passed"] = False

    # --------------------------------------------------------
    # 標題＋右上角快捷鍵
    # --------------------------------------------------------

    title_col, home_col, record_col, keypoint_col, systems_col = st.columns(
        [4, 1.2, 1.2, 1.2, 1.2]
    )

    with title_col:

        st.title("🎓 新人訓練中心")

    with home_col:

        if st.button(
            "🏠 新人首頁",
            use_container_width=True,
            key="training_top_home"
        ):

            st.session_state["training_page"] = "home"

            st.session_state["training_step"] = "video"

            st.session_state["training_course_id"] = ""

            st.session_state["training_quiz_passed"] = False

            st.rerun()

    with record_col:

        if st.button(
            "📊 學習紀錄",
            use_container_width=True,
            key="training_top_records"
        ):

            st.session_state["training_page"] = "records"

            st.rerun()

    with keypoint_col:

        if st.button(
            "📚 重點整理",
            use_container_width=True,
            key="training_top_keypoints"
        ):

            st.session_state["training_page"] = "keypoints"
    with systems_col:

        if st.button(
            "💻 系統索引",
            use_container_width=True,
            key="training_top_systems"
        ):

            st.session_state["training_page"] = "systems"

            st.session_state["training_step"] = "video"

            st.session_state["training_course_id"] = ""

            st.session_state["training_quiz_passed"] = False

            st.rerun()

    st.caption(
        "新人訓練 Level 1｜共 3 天"
    )
    
    if role == "👤 新人" and st.session_state.get(
        "newcomer_authenticated",
        False
    ):

        st.success(
            f"👋 歡迎 {st.session_state.get('training_employee_name', '')}！"
        )

        st.caption(
            f"員工編號：{st.session_state.get('training_employee_id', '')}"
        )
        
    st.divider()

    # ========================================================
    # 🏠 新人首頁
    # ========================================================

    if st.session_state["training_page"] == "home":

        st.markdown("""
        <div style="
            background:#F5F1E9;
            padding:26px 28px;
            border-radius:12px;
            border-left:5px solid #B49A78;
            margin-bottom:20px;
        ">
            <div style="
                color:#8A7863;
                font-size:12px;
                letter-spacing:3px;
                margin-bottom:8px;
            ">NITORI · TRAINING</div>
            <div style="
                color:#403B34;
                font-size:27px;
                font-weight:600;
                margin-bottom:8px;
            ">我的新人訓練</div>
            <div style="
                color:#777067;
                font-size:14px;
            ">循序漸進，完成每一堂課程。</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div style="
            color:#403B34;
            font-size:21px;
            font-weight:600;
            margin:20px 0 6px 0;
        ">📚 課程選擇</div>
        <div style="
            color:#81786D;
            font-size:13px;
            margin-bottom:18px;
        ">依照 DAY 1 至 DAY 3 的順序，完成影片、課程內容與測驗。</div>
        """, unsafe_allow_html=True)

        course_groups = [
            (
                "DAY 1",
                "宜得利基礎",
                [
                    (
                        "DAY1_BASIC",
                        "🏠 宜得利基礎",
                        "了解宜得利基本資訊、企業理念與新人需要掌握的基礎知識。"
                    ),
                    (
                        "DAY1_SERVICE",
                        "💬 接客基礎",
                        "學習接客基本概念與顧客應對方式。"
                    ),
                ],
            ),
            (
                "DAY 2",
                "商品與系統",
                [
                    (
                        "DAY2_PRODUCT",
                        "🛋 商品基礎",
                        "認識家具、家飾及常見商品資訊。"
                    ),
                    (
                        "DAY2_SYSTEM",
                        "💻 系統操作",
                        "認識客服工作中常用的系統及基本操作。"
                    ),
                ],
            ),
            (
                "DAY 3",
                "實務訓練",
                [
                    (
                        "DAY3_PHONE",
                        "☎️ 電話應對",
                        "學習顧客來電時的基本接聽流程與應對方式。"
                    ),
                    (
                        "DAY3_PRACTICE",
                        "📝 實務練習",
                        "透過實際案例進行新人訓練與確認。"
                    ),
                ],
            ),
        ]

        for day_label, day_title, courses in course_groups:

            st.markdown(
                f"""
                <div style="
                    color:#8A7863;
                    font-size:12px;
                    letter-spacing:2px;
                    margin:26px 0 12px 0;
                ">{day_label}</div>
                <div style="
                    color:#403B34;
                    font-size:18px;
                    font-weight:600;
                    margin-bottom:14px;
                ">{day_title}</div>
                """,
                unsafe_allow_html=True
            )

            course_col1, course_col2 = st.columns(2, gap="medium")

            for index, course in enumerate(courses):

                course_id, course_name, description = course

                current_col = (
                    course_col1 if index == 0 else course_col2
                )

                with current_col:

                    with st.container(border=True):

                        st.markdown(
                            f"""
                            <div style="
                                height:4px;
                                background:#C5AD8C;
                                border-radius:4px;
                                margin-bottom:16px;
                            "></div>
                            <div style="
                                color:#403B34;
                                font-size:18px;
                                font-weight:600;
                                margin-bottom:10px;
                            ">{course_name}</div>
                            <div style="
                                color:#81786D;
                                font-size:13px;
                                line-height:1.8;
                                min-height:72px;
                                margin-bottom:12px;
                            ">{description}</div>
                            <div style="
                                display:inline-block;
                                background:#F5F1E9;
                                color:#8A7863;
                                padding:5px 10px;
                                border-radius:16px;
                                font-size:12px;
                            ">{day_label} 課程</div>
                            """,
                            unsafe_allow_html=True
                        )

                        if st.button(
                            "開始課程 →",
                            key=f"training_card_{course_id}",
                            use_container_width=True
                        ):

                            st.session_state[
                                "training_page"
                            ] = "course"

                            st.session_state[
                                "training_course_id"
                            ] = course_id
                            
                            # 點擊大項後，先顯示該大項的細項清單
                            st.session_state[
                                "training_step"
                           ] = "subtopics"

                            st.session_state[
                                "training_quiz_passed"
                            ] = False

                            # 清除前一次選取的細項，避免沿用舊內容
                            st.session_state[
                                "training_content_id"
                            ] = ""
                            
                            if course_id == "DAY1_BASIC":

                                current_employee_id = str(
                                    st.session_state.get(
                                        "training_employee_id", ""
                                    )
                                ).strip()

                                current_employee_name = str(
                                    st.session_state.get(
                                        "training_employee_name", ""
                                    )
                                ).strip()

                                save_training_record(
                                    employee_id=current_employee_id,
                                    employee_name=current_employee_name,
                                    course_id=course_id,
                                    course_name=course_name,
                                    day="DAY1",
                                    start_date=str(
                                        pd.Timestamp.now().date()
                                    ),
                                    learning_status="學習中",
                                    quiz_status="尚未測驗",
                                    completed_date=""
                                )

                            st.rerun()

    # ========================================================
    # 📊 學習紀錄｜日系簡約版
    # ========================================================

    elif st.session_state["training_page"] == "records":

        current_employee_id = str(
            st.session_state.get(
                "training_employee_id", ""
            )
        ).strip()

        current_employee_name = st.session_state.get(
            "training_employee_name", ""
        )

        # ----------------------------------------------------
        # 頁面標題
        # ----------------------------------------------------

        st.markdown("""
        <div style="
            background:#F5F1E9;
            padding:24px 26px;
            border-radius:12px;
            border-left:5px solid #B49A78;
            margin-bottom:20px;
        ">
            <div style="
                font-size:12px;
                letter-spacing:2px;
                color:#81786D;
                margin-bottom:8px;
            ">
                NITORI · MY LEARNING
            </div>
            <div style="
                font-size:25px;
                font-weight:600;
                color:#403B34;
                margin-bottom:8px;
            ">
                我的學習進度
            </div>
            <div style="
                font-size:14px;
                color:#81786D;
            ">
                按照課程查看學習狀態，逐步完成三天的新人訓練。
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.caption(
            f"👤 {current_employee_name}　｜　員工編號：{current_employee_id}"
        )

        # ----------------------------------------------------
        # 六堂課程設定
        # ----------------------------------------------------

        course_list = [
            {
                "id": "DAY1_BASIC",
                "name": "🏠 宜得利基礎",
                "day": "DAY 1",
                "day_name": "宜得利與企業基礎"
            },
            {
                "id": "DAY1_SERVICE",
                "name": "💬 接客基礎",
                "day": "DAY 1",
                "day_name": "宜得利與企業基礎"
            },
            {
                "id": "DAY2_PRODUCT",
                "name": "🛋 商品基礎",
                "day": "DAY 2",
                "day_name": "商品與系統操作"
            },
            {
                "id": "DAY2_SYSTEM",
                "name": "💻 系統操作",
                "day": "DAY 2",
                "day_name": "商品與系統操作"
            },
            {
                "id": "DAY3_PHONE",
                "name": "☎️ 電話應對",
                "day": "DAY 3",
                "day_name": "電話與實務訓練"
            },
            {
                "id": "DAY3_PRACTICE",
                "name": "📝 實務練習",
                "day": "DAY 3",
                "day_name": "電話與實務訓練"
            }
        ]

        # ----------------------------------------------------
        # 取得目前登入新人的學習紀錄
        # ----------------------------------------------------

        employee_records = training_records.iloc[0:0].copy()

        if not training_records.empty:
            employee_records = training_records[
                training_records["員工編號"]
                .astype(str)
                .str.strip()
                .eq(current_employee_id)
            ].copy()

        course_progress = []

        for course in course_list:

            course_record = employee_records[
                employee_records["課程ID"]
                .astype(str)
                .str.strip()
                .eq(course["id"])
            ]

            if course_record.empty:

                status = "尚未開始"
                progress = 0

            else:

                learning_status = str(
                    course_record.iloc[-1].get(
                        "學習狀態", ""
                    )
                ).strip()

                if learning_status == "已完成":
                    status = "✓ 已完成"
                    progress = 100

                elif learning_status == "學習中":
                    status = "◐ 學習中"
                    progress = 50

                else:
                    status = "○ 尚未開始"
                    progress = 0

            course_progress.append({
                "id": course["id"],
                "name": course["name"],
                "day": course["day"],
                "day_name": course["day_name"],
                "status": status,
                "progress": progress
            })

        # ----------------------------------------------------
        # 整體學習進度
        # ----------------------------------------------------

        total_courses = len(course_progress)

        completed_courses = sum(
            1
            for record in course_progress
            if record["progress"] == 100
        )

        learning_count = sum(
            1
            for record in course_progress
            if record["status"] == "◐ 學習中"
        )

        overall_progress = (
            sum(
                record["progress"]
                for record in course_progress
            )
            / (total_courses * 100)
        ) if total_courses else 0

        st.markdown("""
        <div style="
            font-size:19px;
            font-weight:600;
            color:#403B34;
            margin-top:24px;
            margin-bottom:12px;
        ">
            整體訓練進度
        </div>
        """, unsafe_allow_html=True)

        progress_col1, progress_col2, progress_col3 = st.columns(3)

        with progress_col1:
            st.metric(
                "整體進度",
                f"{overall_progress:.0%}"
            )

        with progress_col2:
            st.metric(
                "已完成課程",
                f"{completed_courses} / {total_courses}"
            )

        with progress_col3:
            st.metric(
                "學習中",
                learning_count
            )

        st.progress(overall_progress)

        st.caption(
            "每堂課完成後，系統會依照目前儲存的學習狀態更新進度。"
        )

        if training_records.empty:
            st.warning(
                "目前讀取不到學習紀錄，請確認 Excel 的「新人學習紀錄」工作表。"
            )

        elif employee_records.empty:
            st.info(
                "目前尚無學習紀錄，請從新人首頁開始課程。"
            )

        st.divider()

        # ----------------------------------------------------
        # 依照 DAY 顯示課程卡片
        # ----------------------------------------------------

        for day in ["DAY 1", "DAY 2", "DAY 3"]:

            day_records = [
                record
                for record in course_progress
                if record["day"] == day
            ]

            if not day_records:
                continue

            day_name = day_records[0]["day_name"]

            st.markdown(
                f"""
                <div style="
                    background:#F5F1E9;
                    padding:12px 16px;
                    border-radius:8px;
                    margin-top:14px;
                    margin-bottom:12px;
                ">
                    <span style="
                        color:#A58A68;
                        font-size:12px;
                        letter-spacing:1px;
                    ">{day}</span>
                    <span style="
                        color:#403B34;
                        font-size:16px;
                        font-weight:600;
                        margin-left:10px;
                    ">{day_name}</span>
                </div>
                """,
                unsafe_allow_html=True
            )

            course_col1, course_col2 = st.columns(2)

            for index, record in enumerate(day_records):

                current_col = (
                    course_col1
                    if index % 2 == 0
                    else course_col2
                )

                with current_col:

                    with st.container(border=True):

                        st.markdown(
                            f"**{record['name']}**"
                        )

                        st.caption(
                            f"課程編號：{record['id']}"
                        )

                        st.write(
                            record["status"]
                        )

                        st.progress(
                            record["progress"] / 100
                        )

                        st.caption(
                            f"完成度：{record['progress']}%"
                        )

        st.divider()

        # ----------------------------------------------------
        # 返回首頁
        # ----------------------------------------------------

        if st.button(
            "← 回到新人首頁",
            key="training_records_back_home",
            use_container_width=True
        ):
            st.session_state["training_page"] = "home"
            st.rerun()

    elif st.session_state["training_page"] == "systems":

        st.subheader("💻 系統索引")

        st.caption(
            "不知道要用哪個系統？先找你要完成的工作。"
        )

        system_keyword = st.text_input(
            "搜尋功能、工作內容或系統名稱",
            placeholder="例如：庫存、商品資訊、訂單、配送",
            key="training_system_index_search"
        )

        system_index_demo = [
            {
                "分類": "📦 商品相關",
                "工作": "商品資訊查詢",
                "系統": "2-F1",
                "說明": "查詢商品基本資料、商品資訊等內容。",
                "關鍵字": "商品、商品資訊、品番、規格"
            },
            {
                "分類": "📦 商品相關",
                "工作": "庫存查詢",
                "系統": "2-F5",
                "說明": "查詢門市及商品庫存相關資訊。",
                "關鍵字": "庫存、門市庫存、存貨"
            },
            {
                "分類": "🧾 訂單相關",
                "工作": "訂單查詢",
                "系統": "待建立",
                "說明": "查詢顧客訂單及訂單處理狀態。",
                "關鍵字": "訂單、訂單編號、購買"
            },
            {
                "分類": "🚚 配送相關",
                "工作": "配送資訊查詢",
                "系統": "待建立",
                "說明": "查詢配送相關資訊及配送狀態。",
                "關鍵字": "配送、送貨、配送日期"
            }
        ]

        keyword = system_keyword.strip().lower()

        filtered_systems = [
            item
            for item in system_index_demo
            if not keyword
            or keyword in item["工作"].lower()
            or keyword in item["系統"].lower()
            or keyword in item["說明"].lower()
            or keyword in item["關鍵字"].lower()
        ]

        if not filtered_systems:

            st.info(
                "找不到符合的系統或功能，請換其他關鍵字搜尋。"
            )

        else:

            for item in filtered_systems:

                with st.container(border=True):

                    st.markdown(
                        f"### {item['分類']}｜{item['工作']}"
                    )

                    st.write(
                        f"**使用系統：{item['系統']}**"
                    )

                    st.write(item["說明"])

                    st.caption(
                        f"關鍵字：{item['關鍵字']}"
                    )

        st.divider()

        if st.button(
            "← 回到新人首頁",
            key="training_systems_back_home",
            use_container_width=True
        ):

            st.session_state["training_page"] = "home"
            st.rerun()

    # ========================================================
    # 新人訓練：課程頁面
    # ========================================================

    elif st.session_state["training_page"] == "course":

        course_id = st.session_state.get(
            "training_course_id", ""
        )

        training_step = st.session_state.get(
            "training_step", "subtopics"
        )

        course_names = {
            "DAY1_BASIC": "🏠 宜得利基礎",
            "DAY1_SERVICE": "💬 接客基礎",
            "DAY2_PRODUCT": "🛋 商品基礎",
            "DAY2_SYSTEM": "💻 系統操作",
            "DAY3_PHONE": "☎️ 電話應對",
            "DAY3_PRACTICE": "📝 實務練習",
        }

        course_name = course_names.get(
            course_id, "新人訓練課程"
        )

        course_id_mapping = {
            "DAY1_BASIC": "D1-01",
        }

        supabase_course_id = course_id_mapping.get(
            course_id
        )

        st.subheader(course_name)

        st.caption(f"課程ID：{course_id}")

        st.divider()

        # ====================================================
        # ① 課程細項清單
        # ====================================================

        if training_step == "subtopics":

            st.markdown("### 📚 課程細項清單")
            st.write("請選擇要學習的細項。")

            if supabase_course_id is None:

                st.info(
                    "這堂課的 Supabase 課程 ID 尚未設定，"
                    "目前先保留頁面架構。"
                )

            elif training_content.empty:

                st.warning(
                    "目前沒有讀取到課程細項資料，請稍後再試。"
                )

            else:

                course_details = training_content[
                    training_content["課程ID"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .eq(supabase_course_id)
                ].copy()

                course_details = course_details.reset_index(
                    drop=True
                )

                if course_details.empty:

                    st.info(
                        f"目前找不到課程 ID：{supabase_course_id} 的細項。"
                    )

                else:

                    completed_key = (
                        f"completed_{course_id}"
                    )

                    if completed_key not in st.session_state:
                        st.session_state[completed_key] = []

                    completed_ids = st.session_state[
                        completed_key
                    ]

                    completed_count = sum(
                        str(cid).strip() in completed_ids
                        for cid in course_details["內容ID"]
                    )

                    st.progress(
                        completed_count / len(course_details),
                        text=(
                            f"已完成 {completed_count} / "
                            f"{len(course_details)} 個細項"
                        )
                    )

                    for _, detail in course_details.iterrows():

                        content_id = str(
                            detail["內容ID"]
                        ).strip()

                        content_title = str(
                            detail["課程細項"]
                        ).strip()

                        is_completed = (
                            content_id in completed_ids
                        )

                        with st.container(border=True):

                            st.markdown(
                                f"{'✅' if is_completed else '📖'} "
                                f"**{content_title}**"
                            )

                            st.caption(
                                "已完成" if is_completed
                                else "尚未完成"
                            )

                            if st.button(
                                "查看細項 →",
                                key=(
                                    f"open_{course_id}_"
                                    f"{content_id}"
                                ),
                                use_container_width=True
                            ):

                                st.session_state[
                                    "training_content_id"
                                ] = content_id

                                st.session_state[
                                    "training_step"
                                ] = "detail"

                                st.rerun()

                    st.divider()

                    all_completed = (
                        completed_count == len(course_details)
                    )

                    if all_completed:

                        if st.button(
                            "📘 進入重點整理",
                            key=f"summary_{course_id}",
                            use_container_width=True
                        ):

                            st.session_state[
                                "training_step"
                            ] = "summary"

                            st.rerun()

                    else:

                        st.info(
                            "完成全部細項後，才會開放重點整理。"
                        )

            if st.button(
                "⬅️ 返回新人首頁",
                key=f"back_home_{course_id}"
            ):

                st.session_state["training_page"] = "home"
                st.session_state["training_step"] = "subtopics"
                st.rerun()

        # ====================================================
        # ② 單一細項學習流程
        # 影片 → 重點整理 → 測驗
        # ====================================================

        elif training_step in [
            "detail",
            "video",
            "summary",
            "quiz"
        ]:

            selected_content_id = str(
                st.session_state.get(
                    "training_content_id", ""
                )
            ).strip()

            if training_content.empty:

                st.warning("目前無法讀取課程細項。")

            else:

                selected_detail = training_content[
                    training_content["內容ID"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .eq(selected_content_id)
                    &
                    training_content["課程ID"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .eq(supabase_course_id or "")
                ].copy()

                if selected_detail.empty:

                    st.warning(
                        "找不到這個細項，請返回清單重新選擇。"
                    )

                    if st.button(
                        "⬅️ 返回細項清單",
                        key=f"missing_detail_back_{course_id}"
                    ):
                        st.session_state[
                            "training_step"
                        ] = "subtopics"

                        st.rerun()

                else:

                    detail = selected_detail.iloc[0]

                    st.markdown("### 📖 細項學習")

                    st.markdown(
                        f"#### {detail['課程細項']}"
                    )

                    st.caption(
                        f"細項編號：{selected_content_id}"
                    )

                    st.progress(
                        {
                            "detail": 0,
                            "video": 0,
                            "summary": 0.33,
                            "quiz": 0.66
                        }.get(training_step, 0)
                    )

                    st.divider()

                    # ------------------------------------------------
                    # 影片階段
                    # ------------------------------------------------

                    if training_step in ["detail", "video"]:

                        st.markdown("### ① 🎬 訓練影片")

                        st.info(
                            "此處預留目前細項的專屬影片播放區，"
                            "待影片資料整理完成後再串接。"
                        )

                        if st.button(
                            "下一步：重點整理 ➡️",
                            key=(
                                f"video_next_{course_id}_"
                                f"{selected_content_id}"
                            ),
                            use_container_width=True
                        ):

                            st.session_state[
                                "training_step"
                            ] = "summary"

                            st.rerun()

                    # ------------------------------------------------
                    # 重點整理階段
                    # ------------------------------------------------

                    elif training_step == "summary":

                        st.markdown("### ② 📝 重點整理")

                        notes = str(
                            detail.get("注意事項", "")
                        ).strip()

                        if notes and notes.lower() != "nan":

                            st.write(notes)

                        else:

                            st.info(
                                "此細項的重點整理尚待建立。"
                            )

                        if st.button(
                            "下一步：進入測驗 ➡️",
                            key=(
                                f"summary_next_{course_id}_"
                                f"{selected_content_id}"
                            ),
                            use_container_width=True
                        ):

                            st.session_state[
                                "training_step"
                            ] = "quiz"

                            st.session_state[
                                "training_quiz_passed"
                            ] = False

                            st.rerun()

                    # ------------------------------------------------
                    # 測驗階段
                    # ------------------------------------------------

                    elif training_step == "quiz":

                        st.markdown("### ③ 🧠 課程測驗")

                        st.warning(
                            "目前先建立測驗流程。"
                            "正式題目與答題判分功能接回後，"
                            "必須確認測驗通過才能完成此細項。"
                        )

                        st.caption(
                            "目前尚未接入正式測驗，"
                            "因此不會將此細項標記為完成。"
                        )

                    st.divider()

                    # ------------------------------------------------
                    # 返回細項清單
                    # ------------------------------------------------
                    # 只允許離開，不會直接標記完成。
                    # 正式測驗通過後，才會自動返回清單。

                    if st.button(
                        "⬅️ 返回細項清單",
                        key=(
                            f"learning_back_{course_id}_"
                            f"{selected_content_id}"
                        ),
                        use_container_width=True
                    ):

                        st.session_state[
                            "training_step"
                        ] = "subtopics"

                        st.rerun()
        # ====================================================
        # ③ 大項重點整理
        # ====================================================

        elif training_step == "summary":

            st.markdown("### 📘 課程重點整理")

            # 進入重點整理前，再確認所有細項都已完成
            if supabase_course_id is None or training_content.empty:

                st.warning(
                    "目前無法確認課程細項的完成狀態。"
                )

            else:

                course_details = training_content[
                    training_content["課程ID"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .eq(supabase_course_id)
                ].copy()

                completed_ids = st.session_state.get(
                    f"completed_{course_id}", []
                )

                all_completed = (
                    not course_details.empty
                    and all(
                        str(cid).strip() in completed_ids
                        for cid in course_details["內容ID"]
                    )
                )

                if not all_completed:

                    st.warning(
                        "請先完成全部課程細項，再進入重點整理。"
                    )

                    if st.button(
                        "⬅️ 返回細項清單",
                        key=f"summary_locked_{course_id}"
                    ):

                        st.session_state[
                            "training_step"
                        ] = "subtopics"

                        st.rerun()

                else:

                    st.info(
                        "重點整理內容待補上。"
                        "之後會在這裡統整本堂課的學習重點。"
                    )

                    if st.button(
                        "📝 前往測驗",
                        key=f"summary_to_quiz_{course_id}",
                        use_container_width=True
                    ):

                        st.session_state[
                            "training_step"
                        ] = "quiz"

                        st.session_state[
                            "training_quiz_passed"
                        ] = False

                        st.rerun()

                    if st.button(
                        "⬅️ 返回細項清單",
                        key=f"summary_back_{course_id}"
                    ):

                        st.session_state[
                            "training_step"
                        ] = "subtopics"

                        st.rerun()

        # ====================================================
        # ④ 測驗頁面
        # ====================================================

        elif training_step == "quiz":

            st.markdown("### 📝 課程測驗")

            st.info(
                "目前先保留測驗頁面的入口。"
                "請將原本的測驗題目與作答程式接回此區塊，"
                "避免覆蓋原有的測驗內容。"
            )

            if st.button(
                "⬅️ 返回重點整理",
                key=f"quiz_back_summary_{course_id}"
            ):

                st.session_state[
                    "training_step"
                ] = "summary"

                st.rerun()

        # ====================================================
        # 未知階段：自動回到課程細項清單
        # ====================================================

        else:

            st.session_state["training_step"] = "subtopics"
            st.rerun()
# ============================================================
# 公告
# ============================================================

elif page == "📢 公告":

    st.title(
        "📢 內部公告"
    )
    # --------------------------------------------------------
    # 內部公告
    # --------------------------------------------------------
    ann = load_announcements()

    if ann.empty:

        st.info(
            "目前沒有公告資料。"
        )

    else:

        announcement_columns = (
            get_announcement_columns(
                ann
            )
        )

        for _, row in ann.iterrows():

            render_announcement(
                row,
                announcement_columns
            )
# ============================================================
# 資料維護
# ============================================================

elif page == "🛠️ 資料維護":

    # --------------------------------------------------------
    # 資料維護頁面標題區
    # --------------------------------------------------------
    st.markdown(
        """
        <div style="
            background: linear-gradient(120deg, #FFFDF9 0%, #F0E4D7 100%);
            border: 1px solid #E4D9CB;
            border-radius: 18px;
            padding: 24px 28px;
            margin-bottom: 24px;
            box-shadow: 0 4px 14px rgba(90,70,50,0.04);
        ">
            <div style="
                font-size: 12px;
                color: #9A8F84;
                letter-spacing: 2px;
                margin-bottom: 8px;
            ">
                DATA MANAGEMENT
            </div>
            <div style="
                font-size: 26px;
                font-weight: 700;
                color: #514C45;
                line-height: 1.5;
            ">
                🛠️ 資料維護中心
            </div>
            <div style="
                color: #898176;
                font-size: 14px;
                line-height: 1.8;
                margin-top: 8px;
            ">
                集中管理客服知識、公告、待審核問題與客訴資料。
                <br>
                讓每一筆資料都清楚、完整，方便後續查詢與維護。 🐾
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # 資料統計
    # --------------------------------------------------------

    knowledge_df = load_data()
    announcement_df = load_announcements()
    pending_df = load_pending_questions()
    complaint_df = load_complaints()

    knowledge_count = len(knowledge_df)
    announcement_count = len(announcement_df)
    pending_count = len(pending_df)
    complaint_count = len(complaint_df)

    stat1, stat2, stat3, stat4 = st.columns(4)

    with stat1:

        st.metric(
            "📚 知識庫",
            f"{knowledge_count:,}",
            "筆資料"
        )

    with stat2:

        st.metric(
            "📢 公告",
            f"{announcement_count:,}",
            "筆資料"
        )

    with stat3:

        st.metric(
            "📝 待審核",
            f"{pending_count:,}",
            "筆資料"
        )

    with stat4:

        st.metric(
            "📊 客訴資料",
            f"{complaint_count:,}",
            "筆資料"
        )

    st.divider()

    # --------------------------------------------------------
    # 資料管理分類
    # --------------------------------------------------------

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "📚 知識庫",
            "📢 公告",
            "📝 待審核問題",
            "📊 客訴資料",
            "💻 系統索引"
        ]
    )

    # ========================================================
    # 知識庫
    # ========================================================

    with tab1:

        st.subheader(
            "📚 知識庫管理"
        )

        st.caption(
            "集中查看、編輯及新增客服標準回答。"
        )

        df = load_data()

        if df.empty:

            st.warning(
                "找不到知識庫資料。"
            )

        else:

            # ------------------------------------------------
            # 知識庫功能切換
            # ------------------------------------------------

            knowledge_mode = st.radio(
                "功能",
                [
                    "📋 總表",
                    "✏️ 編輯／新增知識"
                ],
                horizontal=True,
                label_visibility="collapsed",
                key="knowledge_mode"
            )

            st.divider()
        # =================================================
        # 📋 總表
        # =================================================

        if knowledge_mode == "📋 總表":

            st.markdown(
                "### 📋 知識庫總覽"
            )

            st.caption(
                "可查看目前所有客服知識資料；如需修改，請至「編輯知識」。"
            )

            search_col, count_col = st.columns(
                [4, 1]
            )

            with search_col:

                overview_search = st.text_input(
                    "🔎 搜尋知識",
                    placeholder="輸入問題類型、商品編號、問題摘要或關鍵字...",
                    key="knowledge_overview_search"
                )

            with count_col:

                st.metric(
                    "總筆數",
                    f"{len(df):,}"
                )

            overview_df = df.copy()

            if overview_search.strip():

                keyword = overview_search.strip()

                search_mask = (
                    overview_df
                    .astype(str)
                    .apply(
                        lambda column:
                        column.str.contains(
                            keyword,
                            case=False,
                            na=False
                        )
                    )
                    .any(axis=1)
                )

                overview_df = overview_df[
                    search_mask
                ]

            st.caption(
                f"目前顯示 {len(overview_df):,} 筆資料"
            )

            # --------------------------------------------
            # 總表只讀，不使用 data_editor
            # --------------------------------------------

            show_columns = [
                "序號",
                "問題類型",
                "商品編號",
                "問題摘要",
                "關鍵字",
                "資料來源",
                "更新人",
                "更新日期"
            ]

            show_columns = [
                column
                for column in show_columns
                if column in overview_df.columns
            ]

            st.dataframe(
                overview_df[show_columns],
                use_container_width=True,
                hide_index=True,
                height=550
            )

        # =================================================
        # ✏️ 編輯知識
        # =================================================

        elif knowledge_mode == "✏️ 編輯／新增知識":

            st.caption(
                "先選擇要進行的操作。"
            )

            st.markdown(
                "#### 📌 我要做什麼？"
            )

            knowledge_action = st.radio(
                "操作",
                [
                    "✏️ 編輯既有知識",
                    "➕ 新增客服知識"
                ],
                horizontal=True,
                label_visibility="collapsed",
                key="knowledge_action"
            )
            st.divider()
            
            if knowledge_action == "✏️ 編輯既有知識":

                search_col, count_col = st.columns(
                    [4, 1]
                )

                with search_col:

                    knowledge_search = st.text_input(
                        "🔎 搜尋知識",
                        placeholder="輸入問題、商品編號、關鍵字或問題摘要...",
                        key="knowledge_edit_search"
                    )

                with count_col:

                    st.metric(
                        "知識筆數",
                        f"{len(df):,}"
                    )

                display_df = df.copy()

                # --------------------------------------------
                # 搜尋
                # --------------------------------------------

                if knowledge_search.strip():

                    keyword = knowledge_search.strip()

                    search_mask = (
                        display_df
                        .astype(str)
                        .apply(
                            lambda column:
                            column.str.contains(
                                keyword,
                                case=False,
                                na=False
                            )
                        )
                        .any(axis=1)
                    )

                    display_df = display_df[
                        search_mask
                    ]

                st.caption(
                    f"🔎 找到 {len(display_df):,} 筆符合資料"
                )

                if display_df.empty:

                    st.info(
                        "找不到符合的知識內容，請嘗試其他關鍵字。"
                    )

                else:

                    # ----------------------------------------
                    # 建立選擇清單
                    # ----------------------------------------

                    display_df = display_df.copy()

                    display_df["_original_index"] = (
                        display_df.index
                    )

                    display_df["顯示名稱"] = (
                        display_df["問題摘要"]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                    )

                    if "商品編號" in display_df.columns:

                        display_df["顯示名稱"] = (
                            display_df["商品編號"]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            + "｜"
                            + display_df["顯示名稱"]
                        )

                    if "問題類型" in display_df.columns:

                        display_df["顯示名稱"] = (
                            display_df["問題類型"]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            + "｜"
                            + display_df["顯示名稱"]
                        )

                    selected_display = st.selectbox(
                        "📌 選擇要編輯的知識",
                        display_df["顯示名稱"].tolist(),
                        key="selected_knowledge"
                    )

                    selected_row = display_df[
                        display_df["顯示名稱"]
                        == selected_display
                    ].iloc[0]

                    original_index = selected_row[
                        "_original_index"
                    ]

                    # ----------------------------------------
                    # 編輯區
                    # ----------------------------------------

                    st.divider()

                    st.markdown(
                        "### 📝 知識內容"
                    )

                    left_col, right_col = st.columns(
                        [1, 2]
                    )

                    with left_col:

                        st.markdown(
                            "#### 📌 基本資料"
                        )

                        problem_type = st.text_input(
                            "問題類型",
                            value=str(
                                df.at[
                                    original_index,
                                    "問題類型"
                                ]
                            ),
                            key=f"knowledge_problem_type_{original_index}"
                        )

                        product_no = st.text_input(
                            "商品編號",
                            value=str(
                                df.at[
                                    original_index,
                                    "商品編號"
                                ]
                            ),
                            key=f"knowledge_product_no_{original_index}"
                        )

                        keywords = st.text_input(
                            "關鍵字",
                            value=str(
                                df.at[
                                    original_index,
                                    "關鍵字"
                                ]
                            ),
                            key=f"knowledge_keywords_{original_index}"
                        )

                        source = st.text_input(
                            "資料來源",
                            value=str(
                                df.at[
                                    original_index,
                                    "資料來源"
                                ]
                            ),
                            key=f"knowledge_source_{original_index}"
                        )

                    with right_col:

                        st.markdown(
                            "#### 💬 客服內容"
                        )

                        question_summary = st.text_area(
                            "問題摘要",
                            value=str(
                                df.at[
                                    original_index,
                                    "問題摘要"
                                ]
                            ),
                            height=100,
                            key=f"knowledge_question_{original_index}"
                        )

                        standard_reply = st.text_area(
                            "標準回覆",
                            value=str(
                                df.at[
                                    original_index,
                                    "標準回覆"
                                ]
                            ),
                            height=220,
                            key=f"knowledge_reply_{original_index}"
                        )

                        judgment_basis = st.text_area(
                            "判斷依據",
                            value=str(
                                df.at[
                                    original_index,
                                    "判斷依據"
                                ]
                            ),
                            height=150,
                            key=f"knowledge_basis_{original_index}"
                        )

                    # ----------------------------------------
                    # 系統紀錄
                    # ----------------------------------------

                    st.markdown(
                        "#### 🕘 系統紀錄"
                    )

                    info_col1, info_col2, info_col3, info_col4 = st.columns(
                        4
                    )

                    with info_col1:

                        st.caption(
                            f"序號：{df.at[original_index, '序號']}"
                        )

                    with info_col2:

                        st.caption(
                            f"建立人：{df.at[original_index, '建立人']}"
                        )

                    with info_col3:

                        st.caption(
                            f"更新人：{df.at[original_index, '更新人']}"
                        )

                    with info_col4:

                        update_date = df.at[
                            original_index,
                            "更新日期"
                        ]

                        if pd.notna(update_date):

                            try:

                                update_date = pd.to_datetime(
                                    update_date
                                )

                                if pd.isna(update_date):

                                    update_date = ""

                                else:

                                    update_date = update_date.strftime(
                                        "%Y/%m/%d"
                                    )

                            except:

                                update_date = ""

                        else:

                            update_date = ""
                        st.caption(
                            f"更新日期：{update_date}"
                        )

                    st.divider()
                    editor_name_knowledge = st.text_input(
                        "👤 本次修改人姓名（必填）",
                        placeholder="例如：皮皮、66",
                        key="editor_name_knowledge",
                        help="儲存時會更新「更新人」及「更新日期」。"
                    )

                    if st.button(
                        "💾 儲存這筆知識",
                        type="primary",
                        use_container_width=True,
                        key="save_single_knowledge",
                        disabled=not editor_name_knowledge.strip()
                    ):

                        save_df = df.copy()

                        save_df.at[
                            original_index,
                            "問題類型"
                        ] = problem_type

                        save_df.at[
                            original_index,
                            "商品編號"
                        ] = product_no

                        save_df.at[
                            original_index,
                            "問題摘要"
                        ] = question_summary

                        save_df.at[
                            original_index,
                            "標準回覆"
                        ] = standard_reply

                        save_df.at[
                            original_index,
                            "判斷依據"
                        ] = judgment_basis

                        save_df.at[
                            original_index,
                            "資料來源"
                        ] = source

                        save_df.at[
                            original_index,
                            "關鍵字"
                        ] = keywords

                        save_df.at[
                            original_index,
                            "更新人"
                        ] = editor_name_knowledge

                        save_df.at[
                            original_index,
                            "更新日期"
                        ] = pd.Timestamp.now().strftime(
                            "%Y/%m/%d"
                        )

                        if save_knowledge_data(
                            save_df,
                            editor_name_knowledge
                        ):

                            st.success(
                                "✅ 知識已更新！"
                            )

                            st.rerun()

                    st.divider()

                    with st.expander(
                        "🕘 查看知識庫修改歷程"
                    ):

                        render_change_history(
                            find_sheet(
                                [
                                    "知識庫",
                                    "問題集",
                                    "Sheet1"
                                ],
                                0
                            ) or "知識庫",
                            "knowledge_history"
                        )
        # =================================================
        # ➕ 新增知識
        # =================================================

        elif knowledge_action == "➕ 新增客服知識":

            st.divider()
       
            st.markdown(
                "### ➕ 新增客服知識"
            )

            st.caption(
                "🆕 建立新的客服標準回答，完成後會新增至知識庫。"
            )

            left_col, right_col = st.columns(
                [1, 2]
            )

            with left_col:

                st.markdown(
                    "#### 📌 基本資料"
                )

                new_problem_type = st.text_input(
                    "問題類型",
                    placeholder="例如：配送、退換貨、商品資訊",
                    key="new_problem_type"
                )

                new_product_no = st.text_input(
                    "商品編號",
                    placeholder="沒有商品編號可留空",
                    key="new_product_no"
                )

                new_keywords = st.text_input(
                    "關鍵字",
                    placeholder="例如：配送、提早、缺貨",
                    key="new_keywords"
                )

                new_source = st.text_input(
                    "資料來源",
                    placeholder="例如：商品部、物流、官網",
                    key="new_source"
                )

            with right_col:

                st.markdown(
                    "#### 💬 客服內容"
                )

                new_question_summary = st.text_area(
                    "問題摘要",
                    placeholder="整理顧客常見問題...",
                    height=100,
                    key="new_question_summary"
                )

                new_standard_reply = st.text_area(
                    "標準回覆",
                    placeholder="輸入客服正式回覆內容...",
                    height=220,
                    key="new_standard_reply"
                )

                new_judgment_basis = st.text_area(
                    "判斷依據",
                    placeholder="說明此回答的依據或注意事項...",
                    height=150,
                    key="new_judgment_basis"
                )

            st.divider()

            new_creator = st.text_input(
                "👤 建立人（必填）",
                placeholder="例如：皮皮、66",
                key="new_knowledge_creator"
            )

            if st.button(
                "➕ 新增這筆知識",
                type="primary",
                use_container_width=True,
                key="add_knowledge",
                disabled=not (
                    new_creator.strip()
                    and new_question_summary.strip()
                    and new_standard_reply.strip()
                )
            ):

                new_df = df.copy()

                # ----------------------------------------
                # 自動產生序號
                # ----------------------------------------

                if "序號" in new_df.columns:

                    numeric_seq = pd.to_numeric(
                        new_df["序號"],
                        errors="coerce"
                    )

                    if numeric_seq.notna().any():

                        new_sequence = int(
                            numeric_seq.max()
                        ) + 1

                    else:

                        new_sequence = 1

                else:

                    new_sequence = 1

                today = pd.Timestamp.now().strftime(
                    "%Y/%m/%d"
                )

                new_row = {}

                for column in new_df.columns:
                    new_row[column] = ""

                new_row["序號"] = new_sequence
                new_row["問題類型"] = new_problem_type
                new_row["商品編號"] = new_product_no
                new_row["問題摘要"] = new_question_summary
                new_row["標準回覆"] = new_standard_reply
                new_row["判斷依據"] = new_judgment_basis
                new_row["資料來源"] = new_source
                new_row["關鍵字"] = new_keywords
                new_row["建立人"] = new_creator
                new_row["建立日期"] = today
                new_row["更新人"] = new_creator
                new_row["更新日期"] = today

                new_df = pd.concat(
                    [
                        new_df,
                        pd.DataFrame([new_row])
                    ],
                    ignore_index=True
                )

                if save_knowledge_data(
                    new_df,
                    new_creator
                ):

                    st.success(
                        f"✅ 已新增知識 #{new_sequence}"
                    )

                    st.rerun()

            if "adding_announcement" in st.session_state:

                st.divider()

                with st.container(
                    border=True
                ):

                    st.markdown(
                        "### ➕ 新增公告"
                    )

                    st.caption(
                        "填寫完成後請輸入建立人姓名，再按「新增公告」。"
                    )

                    add_title = st.text_input(
                        "公告標題",
                        key="add_announcement_title"
                    )

                    add_content = st.text_area(
                        "公告內容",
                        height=180,
                        key="add_announcement_content"
                    )

                    date_col1, date_col2, date_col3 = st.columns(3)

                    with date_col1:

                        add_date = st.text_input(
                            "公告日期",
                            key="add_announcement_date"
                        )

                    with date_col2:

                        add_activity_date = st.text_input(
                            "活動日期",
                            key="add_announcement_activity_date"
                        )

                    with date_col3:

                        add_point_date = st.text_input(
                            "點數期限",
                            key="add_announcement_point_date"
                        )

                    add_editor_name = st.text_input(
                        "👤 建立人姓名（必填）",
                        placeholder="請輸入實際建立人的姓名，例如：皮皮、66",
                        key="add_editor_name_announcement"
                    )

                    add_col1, add_col2 = st.columns(2)

                    with add_col1:

                        save_new = st.button(
                            "💾 新增公告",
                            type="primary",
                            use_container_width=True,
                            disabled=not add_editor_name.strip(),
                            key="save_new_announcement"
                        )

                    with add_col2:

                        cancel_new = st.button(
                            "取消新增",
                            use_container_width=True,
                            key="cancel_new_announcement"
                        )

                    # --------------------------------------------------------
                    # 取消新增
                    # --------------------------------------------------------

                    if cancel_new:

                        del st.session_state[
                            "adding_announcement"
                        ]

                        st.rerun()

                    # --------------------------------------------------------
                    # 儲存新增公告
                    # --------------------------------------------------------

                    if save_new:

                        new_df = df.copy()

                        # 自動產生公告編號
                        if "公告編號" in new_df.columns:

                            numbers = pd.to_numeric(
                                new_df["公告編號"],
                                errors="coerce"
                            )

                            if numbers.notna().any():

                                new_number = int(
                                    numbers.max()
                                ) + 1

                            else:

                                new_number = 1

                        else:

                            new_number = 1

                        new_row = {
                            "公告編號": new_number,
                            "公告標題": add_title,
                            "公告內容": add_content,
                            "公告日期": add_date,
                            "活動日期": add_activity_date,
                            "點數期限": add_point_date,
                            "顯示狀態": "是"
                        }

                        new_df = pd.concat(
                            [
                                new_df,
                                pd.DataFrame(
                                    [new_row]
                                )
                            ],
                            ignore_index=True
                        )

                        sheet = find_sheet(
                            [
                                "內部公告",
                                "公告"
                            ]
                        )

                        if sheet:

                            success = save_sheet(
                                sheet,
                                new_df,
                                add_editor_name
                            )

                            if success:

                                st.success(
                                    "公告已新增！"
                                )

                                del st.session_state[
                                    "adding_announcement"
                                ]

                                st.rerun()

                            else:

                                st.error(
                                    "公告新增失敗，請稍後再試。"
                                )

                        else:

                            st.error(
                                "找不到公告資料工作表。"
                            )


            announcement_columns = (
                get_announcement_columns(
                    df
                )
            )

            for index, row in df.iterrows():

                render_announcement_admin(
                    row,
                    announcement_columns,
                    index
                )


            # ========================================================
            # 編輯公告
            # ========================================================

            if "editing_announcement" in st.session_state:

                edit_index = st.session_state[
                    "editing_announcement"
                ]

                edit_row = df.loc[
                    edit_index
                ]

                (
                    edit_title,
                    edit_content,
                    edit_date,
                    edit_activity_date,
                    edit_point_date
                ) = get_announcement_values(
                    edit_row,
                    announcement_columns
                )

                st.divider()

                with st.container(
                    border=True
                ):

                    st.markdown(
                        "### ✏️ 編輯公告"
                    )

                    st.caption(
                        "修改完成後請填寫修改人姓名，再按「儲存修改」。"
                    )

                    new_title = st.text_input(
                        "公告標題",
                        value=edit_title or "",
                        key="edit_announcement_title"
                    )

                    new_content = st.text_area(
                        "公告內容",
                        value=edit_content or "",
                        height=180,
                        key="edit_announcement_content"
                    )

                    date_col1, date_col2, date_col3 = st.columns(3)

                    with date_col1:

                        new_date = st.text_input(
                            "公告日期",
                            value=edit_date or "",
                            key="edit_announcement_date"
                        )

                    with date_col2:

                        new_activity_date = st.text_input(
                            "活動日期",
                            value=edit_activity_date or "",
                            key="edit_announcement_activity_date"
                        )

                    with date_col3:

                        new_point_date = st.text_input(
                            "點數期限",
                            value=edit_point_date or "",
                            key="edit_announcement_point_date"
                        )
                    new_status = st.selectbox(
                        "顯示狀態",
                        ["是", "否"],
                        index=0 if edit_row.get("顯示狀態", "是") == "是" else 1,
                        key="edit_announcement_status"
                    )
                    editor_name_announcement = st.text_input(
                        "👤 修改人姓名（必填）",
                        placeholder="請輸入實際修改人的姓名，例如：皮皮、66",
                        key="edit_editor_name_announcement"
                    )

                    col1, col2 = st.columns(2)

                    with col1:

                        save_edit = st.button(
                            "💾 儲存修改",
                            type="primary",
                            use_container_width=True,
                            disabled=not editor_name_announcement.strip(),
                            key="save_edit_announcement"
                        )

                    with col2:

                        cancel_edit = st.button(
                            "取消編輯",
                            use_container_width=True,
                            key="cancel_edit_announcement"
                        )

                    # --------------------------------------------------------
                    # 取消編輯
                    # --------------------------------------------------------

                    if cancel_edit:

                        del st.session_state[
                            "editing_announcement"
                        ]

                        st.rerun()
                    # --------------------------------------------------------
                    # 儲存修改：更新 Supabase
                    # --------------------------------------------------------

                    if save_edit:

                        title_column = announcement_columns.get(
                            "title"
                        )

                        content_column = announcement_columns.get(
                            "content"
                        )

                        date_column = announcement_columns.get(
                            "date"
                        )

                        activity_date_column = announcement_columns.get(
                            "activity_date"
                        )

                        point_date_column = announcement_columns.get(
                            "point_date"
                        )

                        # 取得目前正在編輯的公告編號
                        announcement_id = str(
                            df.at[edit_index, "公告編號"]
                        ).strip() if "公告編號" in df.columns else ""

                        if not announcement_id:

                            st.error(
                                "找不到公告編號，無法更新資料。"
                            )

                        else:

                            update_data = {
                                "title": new_title.strip(),
                                "content": new_content.strip(),
                                "announcement_date": new_date.strip(),
                                "event_date": new_activity_date.strip(),
                                "points_deadline": new_point_date.strip(),
                                "display_status": new_status,
                                "updated_by": editor_name_announcement.strip(),
                                "updated_at": datetime.now().isoformat()
                            }

                            try:

                                response = (
                                    supabase
                                    .table("announcements")
                                    .update(update_data)
                                    .eq(
                                        "announcement_id",
                                        announcement_id
                                    )
                                    .execute()
                                )

                                if response.data:

                                    load_announcements.clear()

                                    st.success(
                                        "公告已更新！"
                                    )

                                    del st.session_state[
                                        "editing_announcement"
                                    ]

                                    st.rerun()

                                else:

                                    st.error(
                                        "沒有更新任何資料，請確認公告編號是否正確。"
                                    )

                            except Exception as e:

                                st.error(
                                    f"公告更新失敗：{e}"
                                )

    # ========================================================
    # 公告管理
    # ========================================================
    with tab2:
        st.subheader("📢 公告管理")
        st.caption("集中管理公告內容、活動日期、點數期限及顯示狀態。")

        df = load_announcements()

        # --------------------------------------------------------
        # 新增公告按鈕：即使目前沒有公告，也可以新增
        # --------------------------------------------------------
        if st.button(
            "➕ 新增公告",
            type="primary",
            use_container_width=True,
            key="add_announcement"
        ):
            st.session_state["adding_announcement"] = True
            st.session_state.pop("editing_announcement", None)
            st.rerun()

        # --------------------------------------------------------
        # 新增公告表單
        # --------------------------------------------------------
        if st.session_state.get("adding_announcement", False):
            st.divider()

            with st.container(border=True):
                st.markdown("### ✨ 新增公告")
                st.caption("請填寫公告資訊，並輸入建立人姓名。")

                add_title = st.text_input(
                    "公告標題",
                    key="add_announcement_title"
                )

                add_content = st.text_area(
                    "公告內容",
                    height=180,
                    key="add_announcement_content"
                )

                date_col1, date_col2, date_col3 = st.columns(3)

                with date_col1:
                    add_date = st.text_input(
                        "公告日期",
                        key="add_announcement_date"
                    )

                with date_col2:
                    add_activity_date = st.text_input(
                        "活動日期",
                        key="add_announcement_activity_date"
                    )

                with date_col3:
                    add_point_date = st.text_input(
                        "點數期限",
                        key="add_announcement_point_date"
                    )

                add_editor_name = st.text_input(
                    "👤 建立人姓名（必填）",
                    placeholder="請輸入建立人的姓名",
                    key="add_editor_name_announcement"
                )

                save_col, cancel_col = st.columns(2)

                with save_col:
                    save_new = st.button(
                        "💾 儲存公告",
                        type="primary",
                        use_container_width=True,
                        key="save_new_announcement"
                    )

                with cancel_col:
                    cancel_new = st.button(
                        "取消新增",
                        use_container_width=True,
                        key="cancel_new_announcement"
                    )

                if cancel_new:
                    st.session_state.pop("adding_announcement", None)
                    st.rerun()

                if save_new:
                    if not add_editor_name.strip():
                        st.warning("請填寫建立人姓名。")

                    elif not add_title.strip():
                        st.warning("請填寫公告標題。")

                    elif not add_content.strip():
                        st.warning("請填寫公告內容。")

                    else:
                        try:
                            # 依目前公告編號產生下一個編號
                            numbers = pd.to_numeric(
                                df["公告編號"],
                                errors="coerce"
                            ) if "公告編號" in df.columns else pd.Series(
                                dtype="float64"
                            )

                            new_number = (
                                int(numbers.max()) + 1
                                if numbers.notna().any()
                                else 1
                            )

                            new_data = {
                                "announcement_id": new_number,
                                "title": add_title.strip(),
                                "content": add_content.strip(),
                                "announcement_date": add_date.strip(),
                                "event_date": add_activity_date.strip(),
                                "points_deadline": add_point_date.strip(),
                                "display_status": "是",
                                "created_by": add_editor_name.strip(),
                                "updated_by": add_editor_name.strip(),
                                "updated_at": datetime.now().isoformat()
                            }

                            response = (
                                supabase
                                .table("announcements")
                                .insert(new_data)
                                .execute()
                            )

                            if response.data:
                                history_saved = save_announcement_history(
                                    action="新增",
                                    announcement_id=new_number,
                                    title=add_title.strip(),
                                    editor_name=add_editor_name.strip()
                                )

                                load_announcements.clear()
                                st.session_state.pop(
                                    "adding_announcement", None
                                )

                                if history_saved:
                                    st.success("公告已新增，歷史紀錄已保存！")
                                else:
                                    st.warning(
                                        "公告已新增，但歷史紀錄寫入失敗，請確認 Excel 檔案是否可寫入。"
                                    )

                                st.rerun()

                            else:
                                st.error(
                                    "未收到新增結果，請重新確認公告資料。"
                                )

                        except Exception as e:
                            st.error(f"公告新增失敗：{e}")


        # --------------------------------------------------------
        # 公告列表
        # --------------------------------------------------------
        st.divider()
        st.markdown("### 📋 公告列表")

        # 新增成功後重新讀取資料；空資料時仍顯示新增表單入口
        df = load_announcements()

        if df.empty:
            st.info("目前尚無公告資料，請點選上方「新增公告」。")

        else:
            announcement_columns = get_announcement_columns(df)

            for index, row in df.iterrows():
                render_announcement_admin(
                    row,
                    announcement_columns,
                    index
                )

        # --------------------------------------------------------
        # 編輯公告
        # --------------------------------------------------------
        if (
            "editing_announcement" in st.session_state
            and not df.empty
        ):
            edit_index = st.session_state["editing_announcement"]

            if edit_index not in df.index:
                st.warning("找不到要編輯的公告，請重新選擇。")
                st.session_state.pop("editing_announcement", None)
                st.rerun()

            else:
                edit_row = df.loc[edit_index]

                (
                    edit_title,
                    edit_content,
                    edit_date,
                    edit_activity_date,
                    edit_point_date
                ) = get_announcement_values(
                    edit_row,
                    announcement_columns
                )

                current_status = str(
                    edit_row.get("顯示狀態", "是")
                ).strip()

                st.divider()

                with st.container(border=True):
                    st.markdown("### ✏️ 編輯公告")
                    st.caption("修改完成後請填寫修改人姓名，再儲存。")

                    new_title = st.text_input(
                        "公告標題",
                        value=str(edit_title or ""),
                        key="edit_announcement_title"
                    )

                    new_content = st.text_area(
                        "公告內容",
                        value=str(edit_content or ""),
                        height=180,
                        key="edit_announcement_content"
                    )

                    date_col1, date_col2, date_col3 = st.columns(3)

                    with date_col1:
                        new_date = st.text_input(
                            "公告日期",
                            value=str(edit_date or ""),
                            key="edit_announcement_date"
                        )

                    with date_col2:
                        new_activity_date = st.text_input(
                            "活動日期",
                            value=str(edit_activity_date or ""),
                            key="edit_announcement_activity_date"
                        )

                    with date_col3:
                        new_point_date = st.text_input(
                            "點數期限",
                            value=str(edit_point_date or ""),
                            key="edit_announcement_point_date"
                        )

                    new_status = st.selectbox(
                        "顯示狀態",
                        ["是", "否"],
                        index=0 if current_status == "是" else 1,
                        key="edit_announcement_status"
                    )

                    editor_name_announcement = st.text_input(
                        "👤 修改人姓名（必填）",
                        placeholder="請輸入修改人的姓名",
                        key="edit_editor_name_announcement"
                    )

                    edit_col1, edit_col2 = st.columns(2)

                    with edit_col1:
                        save_edit = st.button(
                            "💾 儲存修改",
                            type="primary",
                            use_container_width=True,
                            key="save_edit_announcement"
                        )

                    with edit_col2:
                        cancel_edit = st.button(
                            "取消編輯",
                            use_container_width=True,
                            key="cancel_edit_announcement"
                        )

                    if cancel_edit:
                        st.session_state.pop(
                            "editing_announcement", None
                        )
                        st.rerun()

                    if save_edit:
                        if not editor_name_announcement.strip():
                            st.warning("請填寫修改人姓名。")

                        elif not new_title.strip():
                            st.warning("請填寫公告標題。")

                        elif not new_content.strip():
                            st.warning("請填寫公告內容。")

                        else:
                            try:
                                announcement_id = edit_row.get("公告編號")

                                old_history_data = {
                                    "title": str(edit_title or ""),
                                    "content": str(edit_content or ""),
                                    "announcement_date": str(edit_date or ""),
                                    "event_date": str(edit_activity_date or ""),
                                    "points_deadline": str(edit_point_date or ""),
                                    "display_status": current_status
                                }
                                if pd.isna(announcement_id) or str(
                                    announcement_id
                                ).strip() == "":
                                    st.error(
                                        "找不到公告編號，無法更新資料。"
                                    )

                                else:
                                    update_data = {
                                        "title": new_title.strip(),
                                        "content": new_content.strip(),
                                        "announcement_date": new_date.strip(),
                                        "event_date": new_activity_date.strip(),
                                        "points_deadline": new_point_date.strip(),
                                        "display_status": new_status,
                                        "updated_by": editor_name_announcement.strip(),
                                        "updated_at": datetime.now().isoformat()
                                    }

                                    response = (
                                        supabase
                                        .table("announcements")
                                        .update(update_data)
                                        .eq(
                                            "announcement_id",
                                            announcement_id
                                        )
                                        .execute()
                                    )

                                    if response.data:
                                        new_history_data = {
                                            "title": new_title.strip(),
                                            "content": new_content.strip(),
                                            "announcement_date": new_date.strip(),
                                            "event_date": new_activity_date.strip(),
                                            "points_deadline": new_point_date.strip(),
                                            "display_status": new_status
                                        }

                                        history_saved = save_announcement_history(
                                            action="修改",
                                            announcement_id=announcement_id,
                                            title=new_title.strip(),
                                            editor_name=editor_name_announcement.strip(),
                                            old_data=old_history_data,
                                            new_data=new_history_data
                                        )
                                        load_announcements.clear()
                                        st.session_state.pop(
                                            "editing_announcement", None
                                        )

                                        if history_saved:
                                            st.success(
                                                "公告已更新，歷史紀錄已保存！"
                                            )
                                        else:
                                            st.warning(
                                                "公告已更新，但歷史紀錄寫入 Supabase 失敗，請檢查資料表權限及錯誤日誌。"
                                            )

                                        st.rerun()

                                    else:
                                        st.error(
                                            "沒有更新任何資料，請確認公告編號是否正確。"
                                        )

                            except Exception as e:
                                st.error(f"公告更新失敗：{e}")

        # --------------------------------------------------------
        # 公告修改歷程：讀取 Supabase
        # --------------------------------------------------------
        render_announcement_history()

    # ========================================================
    # 待審核問題
    # ========================================================

    with tab3:

        st.subheader(
            "📝 待審核問題"
        )

        st.caption(
            "客服小幫手找不到答案時，問題會先放在這裡。"
        )

        pending = load_pending_questions()

        if pending.empty:

            st.info(
                "目前沒有待審核問題。"
            )

        else:

            status_options = [
                "確認中",
                "已新增完成",
                "已有類似問題",
                "不需要建立知識"
            ]

            pending = pending.copy()

            pending["狀態"] = (
                pending["狀態"]
                .astype(str)
                .str.strip()
                .replace({
                    "待審核": "確認中",
                    "待處理": "確認中",
                    "": "確認中"
                })
            )

            show_all = st.checkbox(
                "📂 顯示全部問題",
                value=False,
                key="show_all_pending"
            )

            if show_all:

                display_pending = (
                    pending.copy()
                )

                st.caption(
                    f"目前顯示全部 {len(display_pending)} 筆問題"
                )

            else:

                display_pending = pending[
                    pending["狀態"]
                    == "確認中"
                ].copy()

                st.caption(
                    f"目前有 {len(display_pending)} 筆待處理問題"
                )

            if display_pending.empty:

                if show_all:

                    st.info(
                        "目前沒有待審核問題。"
                    )

                else:

                    st.success(
                        "🎉 目前沒有待處理問題！"
                    )

                    st.caption(
                        "如果想查看已處理的歷史資料，可以勾選「📂 顯示全部問題」。"
                    )

            else:

                edited_pending = st.data_editor(
                    display_pending,
                    use_container_width=True,
                    num_rows="dynamic",
                    height=550,
                    key="pending_editor",
                    column_config={

                        "提出時間": st.column_config.TextColumn(
                            "提出時間",
                            disabled=True
                        ),

                        "問題": st.column_config.TextColumn(
                            "問題"
                        ),

                        "建議分類": st.column_config.TextColumn(
                            "建議分類"
                        ),

                        "狀態": st.column_config.SelectboxColumn(
                            "狀態",
                            options=status_options
                        ),

                        "審核備註": st.column_config.TextColumn(
                            "審核備註"
                        )
                    }
                )

                editor_name_pending = st.text_input(
                    "👤 修改人姓名（必填）",
                    placeholder="請輸入實際修改人的姓名，例如：皮皮、66",
                    key="editor_name_pending",
                    help="儲存審核結果時必須填寫。"
                )

                if st.button(
                    "💾 儲存審核結果",
                    type="primary",
                    key="save_pending_questions",
                    disabled=not editor_name_pending.strip()
                ):

                    if show_all:

                        save_pending_df = (
                            edited_pending.copy()
                        )

                    else:

                        hidden_pending = pending[
                            pending["狀態"]
                            != "確認中"
                        ].copy()

                        save_pending_df = pd.concat(
                            [
                                hidden_pending,
                                edited_pending
                            ],
                            ignore_index=True
                        )

                    save_pending_df = (
                        save_pending_df
                        .reindex(
                            columns=PENDING_COLUMNS
                        )
                        .fillna("")
                    )

                    for col in PENDING_COLUMNS:

                        save_pending_df[col] = (
                            save_pending_df[col]
                            .astype(str)
                            .replace("nan", "")
                            .map(clean_text)
                        )

                    save_pending_df["狀態"] = (
                        save_pending_df["狀態"]
                        .replace({
                            "": "確認中",
                            "待審核": "確認中",
                            "待處理": "確認中"
                        })
                    )

                    if save_sheet(
                        "待審核問題",
                        save_pending_df,
                        editor_name_pending
                    ):

                        st.success(
                            "✅ 審核結果已儲存！"
                        )

                        st.rerun()

        render_change_history(
            "待審核問題",
            "pending_history"
        )

        st.divider()

        st.markdown(
            """
            #### 💡 建議工作方式

            ① 小幫手找不到答案

            ↓

            ② 按「提出這個問題給我審核」

            ↓

            ③ 到「待審核問題」確認

            ↓

            ④ 整理正確的問題與答案

            ↓

            ⑤ 加入正式知識庫

            ↓

            ⑥ 將狀態改成「已新增完成」

            ↓

            ⑦ 小幫手下次就可以找到
            """
        )

    # ========================================================
    # 客訴資料
    # ========================================================

    with tab4:

        st.subheader(
            "📊 客訴資料"
        )

        st.caption(
            "可以直接新增、修改或刪除客訴資料，完成後按下儲存。"
        )

        complaints = load_complaints()

        complaint_columns = [
            "序號",
            "日期",
            "月份",
            "購買通路",
            "商品編號",
            "商品名稱",
            "客訴類型",
            "客訴細項",
            "客訴內容",
            "處理方式",
            "處理結果",
            "備註"
        ]

        if complaints.empty:

            edit_df = pd.DataFrame(
                columns=complaint_columns
            )

        else:

            edit_df = complaints[
                complaint_columns
            ].copy()

        if "序號" in edit_df.columns:

            for i in range(
                len(edit_df)
            ):

                if not clean_text(
                    edit_df.iloc[i]["序號"]
                ):

                    edit_df.iloc[
                        i,
                        edit_df.columns.get_loc(
                            "序號"
                        )
                    ] = str(i + 1)

        edited_complaints = st.data_editor(
            edit_df,
            use_container_width=True,
            num_rows="dynamic",
            height=600,
            key="complaint_editor",
            column_config={

                "序號": st.column_config.TextColumn(
                    "序號",
                    help="儲存時會自動重新編號"
                ),

                "日期": st.column_config.TextColumn(
                    "日期",
                    help="例如：2026/09/26"
                ),

                "月份": st.column_config.TextColumn(
                    "月份",
                    help="例如：2026/09"
                )
            }
        )

        st.caption(
            "可以直接在表格內新增、修改或刪除資料。"
        )

        editor_name_complaints = st.text_input(
            "👤 修改人姓名（必填）",
            placeholder="請輸入實際修改人的姓名，例如：皮皮、66",
            key="editor_name_complaints",
            help="儲存客訴資料時必須填寫，姓名會記錄在修改歷程中。"
        )

        if st.button(
            "💾 儲存客訴資料",
            type="primary",
            key="save_complaints",
            disabled=not editor_name_complaints.strip()
        ):

            save_df = (
                edited_complaints.copy()
            )

            save_df = save_df.reindex(
                columns=complaint_columns
            )

            if not save_df.empty:

                save_df = save_df[
                    save_df.apply(
                        lambda row: any(
                            clean_text(value)
                            for value in row
                        ),
                        axis=1
                    )
                ].reset_index(
                    drop=True
                )

            if "序號" in save_df.columns:

                save_df["序號"] = [
                    str(i + 1)
                    for i in range(
                        len(save_df)
                    )
                ]

            if (
                "日期" in save_df.columns
                and "月份" in save_df.columns
            ):

                for i in range(
                    len(save_df)
                ):

                    date_value = clean_text(
                        save_df.at[
                            i,
                            "日期"
                        ]
                    )

                    month_value = clean_text(
                        save_df.at[
                            i,
                            "月份"
                        ]
                    )

                    if (
                        date_value
                        and not month_value
                    ):

                        parsed_date = pd.to_datetime(
                            date_value,
                            errors="coerce"
                        )

                        if not pd.isna(
                            parsed_date
                        ):

                            save_df.at[
                                i,
                                "月份"
                            ] = parsed_date.strftime(
                                "%Y/%m"
                            )

            if save_sheet(
                "客訴資料",
                save_df,
                editor_name_complaints
            ):

                st.success(
                    "客訴資料已儲存！"
                )

                st.rerun()

        render_change_history(
            "客訴資料",
            "complaints_history"
        )
    # ========================================================
    # 客訴資料
    # ========================================================
    with tab5:
        st.subheader("💻 系統索引")
        st.info("系統索引維護功能準備中。")

# ============================================================
# 浮動客服小幫手
# ============================================================

if "assistant_searched" not in st.session_state:

    st.session_state[
        "assistant_searched"
    ] = False

if "assistant_found" not in st.session_state:

    st.session_state[
        "assistant_found"
    ] = False

if "assistant_result" not in st.session_state:

    st.session_state[
        "assistant_result"
    ] = pd.DataFrame()

if "assistant_question_saved" not in st.session_state:

    st.session_state[
        "assistant_question_saved"
    ] = ""

if "assistant_submit_message" not in st.session_state:

    st.session_state[
        "assistant_submit_message"
    ] = ""


with st.popover("🐱"):

    st.subheader(
        "🐱 客服小幫手"
    )

    st.caption(
        "先從知識庫幫你找答案"
    )

    assistant_question = st.text_input(
        "顧客問題",
        placeholder="例如：床墊有味道怎麼辦？",
        key="floating_assistant_question"
    )

    assistant_search = st.button(
        "🔍 幫我找答案",
        use_container_width=True,
        type="primary",
        key="floating_assistant_search"
    )

    if assistant_search:

        question = clean_text(
            assistant_question
        )

        if not question:

            st.session_state[
                "assistant_searched"
            ] = False

            st.session_state[
                "assistant_found"
            ] = False

            st.session_state[
                "assistant_result"
            ] = pd.DataFrame()

            st.warning(
                "請先輸入顧客問題。"
            )

        else:

            found, result = (
                ask_customer_assistant(
                    question
                )
            )

            st.session_state[
                "assistant_searched"
            ] = True

            st.session_state[
                "assistant_found"
            ] = (
                found
                and not result.empty
            )

            st.session_state[
                "assistant_result"
            ] = result

            st.session_state[
                "assistant_question"
            ] = question

            st.session_state[
                "assistant_question_saved"
            ] = ""

            st.session_state[
                "assistant_submit_message"
            ] = ""

    if st.session_state.get(
        "assistant_searched",
        False
    ):

        result = st.session_state.get(
            "assistant_result",
            pd.DataFrame()
        )

        found = st.session_state.get(
            "assistant_found",
            False
        )

        question_to_submit = (
            st.session_state.get(
                "assistant_question",
                ""
            )
        )

        if found and not result.empty:

            best = result.iloc[0]

            title = clean_text(
                best.get(
                    "問題摘要",
                    "相關問題"
                )
            )

            answer = clean_text(
                best.get(
                    "標準回覆",
                    ""
                )
            )

            judgment = clean_text(
                best.get(
                    "判斷依據",
                    ""
                )
            )

            source = clean_text(
                best.get(
                    "資料來源",
                    ""
                )
            )

            score = float(
                best.get(
                    "_final_score",
                    0
                )
            )

            st.success(
                "🐾 找到相關資料"
            )

            st.markdown(
                "**問題**"
            )

            st.write(
                title
            )

            st.markdown(
                "**💬 建議回覆**"
            )

            st.write(
                answer
            )

            if judgment:

                st.markdown(
                    "**🔎 判斷依據**"
                )

                st.write(
                    judgment
                )

            if source:

                st.markdown(
                    "**📚 資料來源**"
                )

                st.write(
                    source
                )

            st.caption(
                f"相關度 {score:.1%}"
            )

            if len(result) > 1:

                st.divider()

                st.markdown(
                    "#### 🔎 其他相關資料"
                )

                for i in range(
                    1,
                    min(
                        len(result),
                        4
                    )
                ):

                    row = result.iloc[i]

                    other_title = clean_text(
                        row.get(
                            "問題摘要",
                            "相關問題"
                        )
                    )

                    other_score = float(
                        row.get(
                            "_final_score",
                            0
                        )
                    )

                    with st.container(
                        border=True
                    ):

                        st.write(
                            other_title
                        )

                        st.caption(
                            f"相關度 {other_score:.1%}"
                        )

        else:

            st.warning(
                "🐱 目前沒有找到足夠相關的資料"
            )

            st.write(
                "這個問題目前可能還沒有建立在知識庫裡。"
            )

            st.write(
                "可以先不要直接使用答案，"
                "建議送到「待審核問題」，"
                "確認後再加入正式知識庫。"
            )

            st.divider()

            st.markdown(
                "### 📝 要把這個問題送去審核嗎？"
            )

            st.caption(
                f"目前問題：{question_to_submit}"
            )

            saved_question = (
                st.session_state.get(
                    "assistant_question_saved",
                    ""
                )
            )

            if (
                saved_question
                and saved_question
                == question_to_submit
            ):

                st.success(
                    "🎉 已成功加入待審核問題！"
                )

                st.info(
                    "請到「🛠️ 資料維護 → 📝 待審核問題」查看。"
                )

            else:

                submit_pending = st.button(
                    "📝 提出這個問題給我審核",
                    type="primary",
                    use_container_width=True,
                    key="floating_submit_pending"
                )

                if submit_pending:

                    question_to_submit = clean_text(
                        question_to_submit
                    )

                    if not question_to_submit:

                        st.error(
                            "❌ 沒有取得要提交的問題，請重新搜尋一次。"
                        )

                    else:

                        success = add_pending_question(
                            question_to_submit
                        )

                        if success:

                            st.session_state[
                                "assistant_question_saved"
                            ] = question_to_submit

                            st.cache_data.clear()

                            st.success(
                                "🎉 已成功加入待審核問題！"
                            )

                            st.info(
                                "請到「🛠️ 資料維護 → 📝 待審核問題」查看。"
                            )

                            st.caption(
                                f"已提交問題：{question_to_submit}"
                            )

                        else:

                            st.error(
                                "❌ 提交失敗。"
                            )

                            st.warning(
                                "請查看上方顯示的錯誤訊息，確認 Excel 是否正在開啟。"
                            )
