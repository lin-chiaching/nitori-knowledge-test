import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import re
import html
from datetime import datetime
from pathlib import Path
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from openpyxl import load_workbook


# =========================================================
# 基本設定
# =========================================================

st.set_page_config(
    page_title="客服作業系統",
    page_icon="🐱",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# 路徑
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "DATA"
EXCEL_FILE = DATA_DIR / "知識庫實驗品.xlsx"


# =========================================================
# 日系貓咪風格 CSS
# =========================================================

st.markdown("""
<style>

    /* =========================
       整體
       ========================= */

    .stApp {
        background-color: #F7F4EE;
        color: #5F5A53;
    }

    .main {
        background-color: #F7F4EE;
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }


    /* =========================
       Sidebar
       ========================= */

    section[data-testid="stSidebar"] {
        background-color: #EEE8DD;
        border-right: 1px solid #DED4C6;
    }

    section[data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }

    .sidebar-cat {
        text-align: center;
        font-size: 46px;
        margin-bottom: 5px;
    }

    .sidebar-title {
        text-align: center;
        font-size: 21px;
        font-weight: 700;
        color: #5F5A53;
        margin-bottom: 2px;
    }

    .sidebar-subtitle {
        text-align: center;
        font-size: 10px;
        letter-spacing: 2px;
        color: #9A9186;
        margin-bottom: 25px;
    }

    .sidebar-footer {
        margin-top: 40px;
        padding: 15px;
        border-radius: 15px;
        background-color: #F7F1E8;
        color: #81786D;
        font-size: 13px;
        line-height: 1.7;
        text-align: center;
        border: 1px solid #E2D8CA;
    }


    /* =========================
       Hero
       ========================= */

    .hero-box {
        background: linear-gradient(
            135deg,
            #FFFDF9 0%,
            #F6EFE4 100%
        );
        border: 1px solid #E3D8CA;
        border-radius: 26px;
        padding: 42px 35px;
        margin-bottom: 28px;
        box-shadow: 0 8px 25px rgba(110, 95, 75, 0.06);
        text-align: center;
    }

    .hero-cat {
        font-size: 58px;
        line-height: 1;
        margin-bottom: 15px;
    }

    .hero-title {
        font-size: 36px;
        font-weight: 800;
        color: #5F5A53;
        margin-bottom: 5px;
    }

    .hero-subtitle {
        font-size: 12px;
        letter-spacing: 3px;
        color: #A1978B;
        margin-bottom: 18px;
    }

    .hero-message {
        font-size: 16px;
        color: #81786D;
    }


    /* =========================
       Section
       ========================= */

    .section-title {
        font-size: 24px;
        font-weight: 700;
        color: #625B52;
        margin-top: 20px;
        margin-bottom: 16px;
    }

    .section-subtitle {
        color: #9A9186;
        font-size: 13px;
        margin-bottom: 20px;
    }


    /* =========================
       卡片
       ========================= */

    .feature-card {
        background-color: #FFFDF9;
        border: 1px solid #E3D8CA;
        border-radius: 20px;
        padding: 25px;
        min-height: 165px;
        box-shadow: 0 5px 18px rgba(110, 95, 75, 0.05);
        margin-bottom: 10px;
    }

    .feature-icon {
        font-size: 34px;
        margin-bottom: 10px;
    }

    .feature-title {
        font-size: 18px;
        font-weight: 700;
        color: #625B52;
        margin-bottom: 8px;
    }

    .feature-text {
        font-size: 13px;
        line-height: 1.7;
        color: #8B8278;
    }


    /* =========================
       公告
       ========================= */

    .announcement-card {
        background-color: #FFFDF9;
        border: 1px solid #E3D8CA;
        border-left: 5px solid #C98568;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 14px;
        box-shadow: 0 4px 15px rgba(110, 95, 75, 0.04);
    }

    .announcement-title {
        font-size: 17px;
        font-weight: 700;
        color: #625B52;
        margin-bottom: 9px;
    }

    .announcement-content {
        font-size: 14px;
        line-height: 1.8;
        color: #716A61;
        white-space: pre-line;
    }

    .announcement-meta {
        font-size: 12px;
        color: #A39A90;
        margin-top: 12px;
    }


    /* =========================
       工具卡
       ========================= */

    .tool-card {
        background-color: #FFFDF9;
        border: 1px solid #E3D8CA;
        border-radius: 18px;
        padding: 20px;
        min-height: 145px;
        box-shadow: 0 4px 15px rgba(110, 95, 75, 0.04);
        margin-bottom: 8px;
    }

    .tool-icon {
        font-size: 30px;
        margin-bottom: 8px;
    }

    .tool-title {
        font-size: 16px;
        font-weight: 700;
        color: #625B52;
        margin-bottom: 7px;
    }

    .tool-description {
        font-size: 13px;
        color: #8B8278;
        line-height: 1.6;
    }


    /* =========================
       搜尋結果
       ========================= */

    .search-result {
        background-color: #FFFDF9;
        border: 1px solid #E3D8CA;
        border-radius: 18px;
        padding: 20px;
        margin-bottom: 15px;
        box-shadow: 0 4px 15px rgba(110, 95, 75, 0.04);
    }

    .result-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
    }

    .result-title {
        font-size: 17px;
        font-weight: 700;
        color: #625B52;
    }

    .score-badge {
        background-color: #E6EFE7;
        color: #66776B;
        border-radius: 20px;
        padding: 4px 10px;
        font-size: 12px;
        white-space: nowrap;
    }

    .result-info {
        color: #91887D;
        font-size: 12px;
        margin-bottom: 12px;
    }

    .result-answer {
        background-color: #F7F3EC;
        border-radius: 12px;
        padding: 13px 15px;
        font-size: 14px;
        line-height: 1.8;
        color: #6E675E;
        white-space: pre-line;
    }


    /* =========================
       小標籤
       ========================= */

    .tag {
        display: inline-block;
        background-color: #F1E8DC;
        color: #897B6D;
        padding: 4px 9px;
        border-radius: 12px;
        font-size: 11px;
        margin-right: 5px;
        margin-bottom: 5px;
    }


    /* =========================
       Streamlit 元件微調
       ========================= */

    div[data-testid="stButton"] > button {
        border-radius: 12px;
        border: 1px solid #D8CDBE;
        background-color: #FFFDF9;
        color: #665E55;
        transition: all 0.2s;
    }

    div[data-testid="stButton"] > button:hover {
        border-color: #C98568;
        color: #B56F53;
        background-color: #FBF3EC;
    }

    div[data-testid="stLinkButton"] > a {
        border-radius: 12px !important;
        border: 1px solid #D8CDBE !important;
        background-color: #FFFDF9 !important;
        color: #665E55 !important;
    }

    div[data-testid="stLinkButton"] > a:hover {
        border-color: #C98568 !important;
        color: #B56F53 !important;
        background-color: #FBF3EC !important;
    }

    input {
        border-radius: 12px !important;
    }

    textarea {
        border-radius: 12px !important;
    }

    div[data-baseweb="select"] > div {
        border-radius: 12px;
    }

</style>
""", unsafe_allow_html=True)


# =========================================================
# HTML / 文字清理工具
# =========================================================

def clean_html_text(value):
    """
    將 Excel 中可能存在的 HTML / Markdown / SVG 清掉。
    """

    if pd.isna(value):
        return ""

    text = str(value)

    # HTML entity
    text = html.unescape(text)

    # 移除 Markdown code fence
    text = re.sub(r"```(?:html|HTML|xml|XML)?", "", text)
    text = text.replace("```", "")

    # 移除完整 SVG
    text = re.sub(
        r"<svg[\s\S]*?</svg>",
        "",
        text,
        flags=re.IGNORECASE
    )

    # 移除單獨 svg
    text = re.sub(
        r"^\s*svg\s*$",
        "",
        text,
        flags=re.IGNORECASE | re.MULTILINE
    )

    # HTML 換行
    text = re.sub(
        r"<br\s*/?>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # 關閉區塊標籤換行
    text = re.sub(
        r"</(?:p|div|h[1-6]|li|tr)>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # 移除所有 HTML tag
    text = re.sub(r"<[^>]+>", "", text)

    # 清除常見殘留屬性
    text = re.sub(
        r'\b(?:id|class)\s*=\s*["\'][^"\']*["\']',
        "",
        text,
        flags=re.IGNORECASE
    )

    # 清除反斜線
    text = text.replace("\\", "")

    # 清理多餘空白
    text = re.sub(r"[ \t]+", " ", text)

    # 清理多餘換行
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)

    return text.strip()


def clean_announcement_title(value):
    """
    清理公告標題。
    """

    text = clean_html_text(value)

    # 移除 📌 後再統一由畫面加入
    text = text.replace("📌", "").strip()

    # 如果原本是【標題】格式，去掉外框
    if text.startswith("【") and text.endswith("】"):
        text = text[1:-1].strip()

    return text


def clean_announcement_content(value):
    """
    清理公告內容。
    """

    text = clean_html_text(value)

    return text


def safe_text(value):
    """
    將 NaN 等內容轉成空字串。
    """

    if pd.isna(value):
        return ""

    return str(value).strip()


# =========================================================
# Excel 讀取
# =========================================================

@st.cache_data
def load_excel_data():
    """
    讀取 Excel 所有工作表。
    """

    if not EXCEL_FILE.exists():
        return {}

    try:
        sheets = pd.read_excel(
            EXCEL_FILE,
            sheet_name=None,
            engine="openpyxl"
        )

        return sheets

    except Exception as e:
        st.error(f"Excel 讀取失敗：{e}")
        return {}


def get_sheet(sheet_name, default_columns=None):
    """
    取得指定工作表。
    """

    sheets = load_excel_data()

    if sheet_name in sheets:
        return sheets[sheet_name].copy()

    if default_columns:
        return pd.DataFrame(columns=default_columns)

    return pd.DataFrame()


# =========================================================
# 知識庫
# =========================================================

def load_knowledge():
    columns = [
        "序號",
        "大分類",
        "中分類",
        "小分類",
        "問題簡述",
        "解答",
        "判斷方式",
        "判斷使用的工具",
        "關鍵字",
        "建立人",
        "建立日期",
        "更新人",
        "更新日期"
    ]

    df = get_sheet("Sheet1", columns)

    if df.empty:
        return pd.DataFrame(columns=columns)

    for col in columns:
        if col not in df.columns:
            df[col] = ""

    df = df[columns].copy()

    for col in columns:
        df[col] = df[col].fillna("").astype(str)

    return df


def build_search_text(df):
    """
    建立搜尋文字。
    """

    search_columns = [
        "問題簡述",
        "解答",
        "判斷方式",
        "關鍵字",
        "大分類",
        "中分類",
        "小分類"
    ]

    existing = [
        col for col in search_columns
        if col in df.columns
    ]

    if not existing:
        return pd.Series([""] * len(df))

    return (
        df[existing]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
    )


@st.cache_resource
def build_search_engine(df):
    """
    建立 TF-IDF 搜尋引擎。
    """

    if df.empty:
        return None, None, None

    texts = build_search_text(df)

    vectorizer = TfidfVectorizer(
        analyzer="char",
        ngram_range=(1, 3),
        min_df=1,
        sublinear_tf=True
    )

    matrix = vectorizer.fit_transform(texts)

    return vectorizer, matrix, texts


def search_knowledge(df, query, top_n=10):
    """
    知識庫搜尋。
    """

    if df.empty:
        return pd.DataFrame()

    query = query.strip()

    if not query:
        return pd.DataFrame()

    vectorizer, matrix, texts = build_search_engine(df)

    if vectorizer is None:
        return pd.DataFrame()

    query_vector = vectorizer.transform([query])

    scores = cosine_similarity(
        query_vector,
        matrix
    ).flatten()

    result = df.copy()
    result["_score"] = scores

    result = result.sort_values(
        "_score",
        ascending=False
    ).head(top_n)

    return result


# =========================================================
# 公告
# =========================================================

def load_announcements():
    columns = [
        "公告標題",
        "公告內容",
        "公告日期",
        "活動日期",
        "點數"
    ]

    sheets = load_excel_data()

    # 嘗試找公告工作表
    possible_names = [
        "內部公告",
        "公告",
        "Announcements"
    ]

    df = None

    for name in possible_names:
        if name in sheets:
            df = sheets[name].copy()
            break

    if df is None:
        return pd.DataFrame(columns=columns)

    # 確保欄位存在
    for col in columns:
        if col not in df.columns:
            df[col] = ""

    # 清理
    if "公告標題" in df.columns:
        df["公告標題"] = df["公告標題"].apply(
            clean_announcement_title
        )

    if "公告內容" in df.columns:
        df["公告內容"] = df["公告內容"].apply(
            clean_announcement_content
        )

    for col in [
        "公告日期",
        "活動日期",
        "點數"
    ]:
        if col in df.columns:
            df[col] = df[col].apply(safe_text)

    return df[columns].copy()


# =========================================================
# 常用工具
# =========================================================

def load_tools():
    columns = [
        "工具名稱",
        "工具說明",
        "連結",
        "分類"
    ]

    sheets = load_excel_data()

    possible_names = [
        "常用工具",
        "工具",
        "Tools"
    ]

    df = None

    for name in possible_names:
        if name in sheets:
            df = sheets[name].copy()
            break

    if df is None:
        return pd.DataFrame(columns=columns)

    for col in columns:
        if col not in df.columns:
            df[col] = ""

    for col in columns:
        df[col] = df[col].fillna("").astype(str)

    return df[columns].copy()


def get_tool_icon(tool_name):
    """
    根據工具名稱自動判斷圖示。
    """

    name = str(tool_name)

    if "地圖" in name:
        return "🗺️"

    if "Power BI" in name or "BI" in name:
        return "📊"

    if "交接" in name:
        return "📋"

    if "信箱" in name or "Mail" in name:
        return "📧"

    if "Excel" in name:
        return "📗"

    if "Teams" in name:
        return "💬"

    if "SharePoint" in name:
        return "📁"

    return "🔗"


# =========================================================
# Excel 儲存
# =========================================================

def save_sheet(df, sheet_name):
    """
    將 DataFrame 寫回指定 Excel 工作表。
    """

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    try:
        if EXCEL_FILE.exists():
            wb = load_workbook(EXCEL_FILE)

            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                wb.remove(ws)

            wb.create_sheet(sheet_name)
            wb.save(EXCEL_FILE)

        else:
            with pd.ExcelWriter(
                EXCEL_FILE,
                engine="openpyxl"
            ) as writer:

                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

            return True

        # 重新寫入
        with pd.ExcelWriter(
            EXCEL_FILE,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="overlay"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

        load_excel_data.clear()
        build_search_engine.clear()

        return True

    except Exception as e:
        st.error(f"儲存失敗：{e}")
        return False


# =========================================================
# Sidebar
# =========================================================

with st.sidebar:

    st.markdown("""
    <div class="sidebar-cat">🐱</div>
    <div class="sidebar-title">客服作業系統</div>
    <div class="sidebar-subtitle">
        NITORI CUSTOMER SERVICE
    </div>
    """, unsafe_allow_html=True)

    page = st.radio(
        "功能選單",
        [
            "首頁",
            "知識庫搜尋",
            "門市地圖",
            "公告＆常用工具",
            "資料維護"
        ],
        label_visibility="collapsed"
    )

    st.markdown("""
    <div class="sidebar-footer">
        🌿 今天也一起把顧客問題解決掉吧！
        <br>
        小貓會陪你一起找資料。
    </div>
    """, unsafe_allow_html=True)


# =========================================================
# 首頁
# =========================================================

if page == "首頁":

    st.markdown("""
    <div class="hero-box">

        <div class="hero-cat">
            🐱
        </div>

        <div class="hero-title">
            客服作業系統
        </div>

        <div class="hero-subtitle">
            NITORI CUSTOMER SERVICE WORKSPACE
        </div>

        <div class="hero-message">
            🌿 今天也辛苦了，讓小貓陪你一起找資料
        </div>

    </div>
    """, unsafe_allow_html=True)

    st.markdown(
        '<div class="section-title">✨ 快速入口</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="feature-card">

            <div class="feature-icon">
                🔎
            </div>

            <div class="feature-title">
                知識庫搜尋
            </div>

            <div class="feature-text">
                輸入顧客問題，
                快速找到相關問題與客服回覆。
            </div>

        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "前往知識搜尋 →",
            key="home_search",
            use_container_width=True
        ):
            st.session_state["page_jump"] = "知識庫搜尋"
            st.rerun()

    with col2:
        st.markdown("""
        <div class="feature-card">

            <div class="feature-icon">
                🗺️
            </div>

            <div class="feature-title">
                門市地圖
            </div>

            <div class="feature-text">
                查詢門市位置、交通、
                停車及相關資訊。
            </div>

        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "查看門市地圖 →",
            key="home_map",
            use_container_width=True
        ):
            st.session_state["page_jump"] = "門市地圖"
            st.rerun()

    with col3:
        st.markdown("""
        <div class="feature-card">

            <div class="feature-icon">
                📢
            </div>

            <div class="feature-title">
                公告＆常用工具
            </div>

            <div class="feature-text">
                查看最新內部公告，
                也可以快速開啟工作工具。
            </div>

        </div>
        """, unsafe_allow_html=True)

        if st.button(
            "查看公告＆工具 →",
            key="home_tools",
            use_container_width=True
        ):
            st.session_state["page_jump"] = "公告＆常用工具"
            st.rerun()

    st.markdown(
        '<div class="section-title">📢 最新公告</div>',
        unsafe_allow_html=True
    )

    announcements = load_announcements()

    if announcements.empty:

        st.info("目前沒有公告。")

    else:

        # 最新公告排前面
        announcements = announcements.head(5)

        for _, row in announcements.iterrows():

            title = clean_announcement_title(
                row.get("公告標題", "")
            )

            content = clean_announcement_content(
                row.get("公告內容", "")
            )

            date = safe_text(
                row.get("公告日期", "")
            )

            if not title and not content:
                continue

            st.markdown(
                f"""
                <div class="announcement-card">

                    <div class="announcement-title">
                        📌 {html.escape(title)}
                    </div>

                    <div class="announcement-content">
                        {html.escape(content).replace(chr(10), "<br>")}
                    </div>

                    <div class="announcement-meta">
                        📅 公告日期：{html.escape(date)}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


# =========================================================
# 知識庫搜尋
# =========================================================

elif page == "知識庫搜尋":

    st.markdown("""
    <div class="hero-box">

        <div class="hero-cat">
            🐱
        </div>

        <div class="hero-title">
            小貓幫你找資料
        </div>

        <div class="hero-subtitle">
            KNOWLEDGE SEARCH
        </div>

        <div class="hero-message">
            輸入顧客問題，找到最相關的客服資訊
        </div>

    </div>
    """, unsafe_allow_html=True)

    df = load_knowledge()

    query = st.text_input(
        "搜尋問題",
        placeholder="例如：床墊重量、可以退貨嗎、門市有庫存嗎……",
        label_visibility="collapsed"
    )

    col1, col2 = st.columns([4, 1])

    with col1:
        st.caption(
            "💡 可以輸入顧客實際說法，不一定要完全符合資料庫中的問題。"
        )

    with col2:
        top_n = st.selectbox(
            "顯示筆數",
            [5, 10, 15, 20],
            index=1
        )

    if query.strip():

        results = search_knowledge(
            df,
            query,
            top_n=top_n
        )

        if results.empty:

            st.warning(
                "🐱 找不到相關資料，可以換個關鍵字試試看。"
            )

        else:

            st.markdown(
                f"""
                <div class="section-title">
                    🔎 搜尋結果
                </div>
                """,
                unsafe_allow_html=True
            )

            for _, row in results.iterrows():

                score = float(row["_score"])

                title = safe_text(
                    row.get("問題簡述", "")
                )

                answer = safe_text(
                    row.get("解答", "")
                )

                big = safe_text(
                    row.get("大分類", "")
                )

                middle = safe_text(
                    row.get("中分類", "")
                )

                small = safe_text(
                    row.get("小分類", "")
                )

                tags = ""

                for tag in [big, middle, small]:

                    if tag:
                        tags += (
                            f'<span class="tag">'
                            f'{html.escape(tag)}'
                            f'</span>'
                        )

                st.markdown(
                    f"""
                    <div class="search-result">

                        <div class="result-header">

                            <div class="result-title">
                                {html.escape(title)}
                            </div>

                            <div class="score-badge">
                                相關度 {score:.1%}
                            </div>

                        </div>

                        <div class="result-info">
                            {tags}
                        </div>

                        <div class="result-answer">
                            {html.escape(answer).replace(chr(10), "<br>")}
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True
                )

    else:

        st.markdown("""
        <div class="feature-card">

            <div class="feature-icon">
                🐾
            </div>

            <div class="feature-title">
                小貓正在等你輸入問題
            </div>

            <div class="feature-text">
                例如：
                「這張床架組裝需要多大空間？」
                <br>
                「這個商品有重量限制嗎？」
                <br>
                「可以門市試躺嗎？」
            </div>

        </div>
        """, unsafe_allow_html=True)


# =========================================================
# 門市地圖
# =========================================================

elif page == "門市地圖":

    st.markdown("""
    <div class="hero-box">

        <div class="hero-cat">
            🐱
        </div>

        <div class="hero-title">
            宜得利小地圖
        </div>

        <div class="hero-subtitle">
            STORE MAP
        </div>

        <div class="hero-message">
            🗺️ 查詢門市、交通、停車及周邊資訊
        </div>

    </div>
    """, unsafe_allow_html=True)

    # Power BI 嵌入網址
    power_bi_url = (
        "https://app.powerbi.com/view?"
        "r=eyJrIjoiYOUR_POWER_BI_LINK_HERE"
    )

    st.info(
        "如果你的 Power BI 嵌入網址已經有自己的網址，"
        "請把上面的 power_bi_url 換成你原本的網址。"
    )

    try:

        components.iframe(
            power_bi_url,
            height=750,
            scrolling=True
        )

    except Exception as e:

        st.error(
            f"Power BI 載入失敗：{e}"
        )


# =========================================================
# 公告＆常用工具
# =========================================================

elif page == "公告＆常用工具":

    st.markdown("""
    <div class="hero-box">

        <div class="hero-cat">
            🐱
        </div>

        <div class="hero-title">
            公告＆常用工具
        </div>

        <div class="hero-subtitle">
            ANNOUNCEMENTS & TOOLS
        </div>

        <div class="hero-message">
            📢 最新資訊與常用工作入口都放在這裡
        </div>

    </div>
    """, unsafe_allow_html=True)

    # -----------------------------------------------------
    # 公告
    # -----------------------------------------------------

    st.markdown(
        '<div class="section-title">📢 最新公告</div>',
        unsafe_allow_html=True
    )

    announcements = load_announcements()

    if announcements.empty:

        st.info("目前沒有公告。")

    else:

        for _, row in announcements.iterrows():

            title = clean_announcement_title(
                row.get("公告標題", "")
            )

            content = clean_announcement_content(
                row.get("公告內容", "")
            )

            date = safe_text(
                row.get("公告日期", "")
            )

            activity = safe_text(
                row.get("活動日期", "")
            )

            points = safe_text(
                row.get("點數", "")
            )

            if not title and not content:
                continue

            extra = ""

            if activity:
                extra += (
                    f"📅 活動期間："
                    f"{html.escape(activity)}"
                    f"<br>"
                )

            if points:
                extra += (
                    f"🎁 點數："
                    f"{html.escape(points)}"
                )

            st.markdown(
                f"""
                <div class="announcement-card">

                    <div class="announcement-title">
                        📌 {html.escape(title)}
                    </div>

                    <div class="announcement-content">
                        {html.escape(content).replace(chr(10), "<br>")}
                    </div>

                    <div class="announcement-meta">
                        📅 公告日期：{html.escape(date)}
                        <br>
                        {extra}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

    # -----------------------------------------------------
    # 工具
    # -----------------------------------------------------

    st.markdown(
        '<div class="section-title">🧰 常用工具</div>',
        unsafe_allow_html=True
    )

    tools_df = load_tools()

    if tools_df.empty:

        st.info("目前沒有常用工具。")

    else:

        # 如果有分類就依分類分組
        if "分類" in tools_df.columns:

            grouped = tools_df.groupby(
                tools_df["分類"].replace("", "其他")
            )

            for category, group in grouped:

                st.markdown(
                    f"""
                    <div class="section-subtitle">
                        {html.escape(str(category))}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                cols = st.columns(3)

                for index, (_, row) in enumerate(
                    group.iterrows()
                ):

                    col = cols[index % 3]

                    name = safe_text(
                        row.get("工具名稱", "")
                    )

                    description = safe_text(
                        row.get("工具說明", "")
                    )

                    link = safe_text(
                        row.get("連結", "")
                    )

                    icon = get_tool_icon(name)

                    with col:

                        st.markdown(
                            f"""
                            <div class="tool-card">

                                <div class="tool-icon">
                                    {icon}
                                </div>

                                <div class="tool-title">
                                    {html.escape(name)}
                                </div>

                                <div class="tool-description">
                                    {html.escape(description)}
                                </div>

                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                        if link:

                            st.link_button(
                                "開啟工具 →",
                                link,
                                use_container_width=True
                            )

        else:

            cols = st.columns(3)

            for index, (_, row) in enumerate(
                tools_df.iterrows()
            ):

                col = cols[index % 3]

                name = safe_text(
                    row.get("工具名稱", "")
                )

                description = safe_text(
                    row.get("工具說明", "")
                )

                link = safe_text(
                    row.get("連結", "")
                )

                icon = get_tool_icon(name)

                with col:

                    st.markdown(
                        f"""
                        <div class="tool-card">

                            <div class="tool-icon">
                                {icon}
                            </div>

                            <div class="tool-title">
                                {html.escape(name)}
                            </div>

                            <div class="tool-description">
                                {html.escape(description)}
                            </div>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    if link:

                        st.link_button(
                            "開啟工具 →",
                            link,
                            use_container_width=True
                        )


# =========================================================
# 資料維護
# =========================================================

elif page == "資料維護":

    st.markdown("""
    <div class="hero-box">

        <div class="hero-cat">
            🐱
        </div>

        <div class="hero-title">
            資料維護
        </div>

        <div class="hero-subtitle">
            DATA MANAGEMENT
        </div>

        <div class="hero-message">
            ✏️ 維護客服知識、公告及常用工具
        </div>

    </div>
    """, unsafe_allow_html=True)

    maintenance_type = st.radio(
        "選擇維護項目",
        [
            "知識庫",
            "內部公告",
            "常用工具"
        ],
        horizontal=True
    )

    # =====================================================
    # 知識庫維護
    # =====================================================

    if maintenance_type == "知識庫":

        df = load_knowledge()

        st.markdown(
            '<div class="section-title">📚 知識庫資料</div>',
            unsafe_allow_html=True
        )

        st.caption(
            f"目前共有 {len(df)} 筆資料"
        )

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            height=600,
            key="knowledge_editor"
        )

        if st.button(
            "💾 儲存知識庫",
            use_container_width=True
        ):

            if save_sheet(
                edited_df,
                "Sheet1"
            ):

                st.success(
                    "知識庫已成功儲存！"
                )

                st.rerun()

    # =====================================================
    # 公告維護
    # =====================================================

    elif maintenance_type == "內部公告":

        df = load_announcements()

        st.markdown(
            '<div class="section-title">📢 內部公告</div>',
            unsafe_allow_html=True
        )

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            height=550,
            key="announcement_editor"
        )

        if st.button(
            "💾 儲存公告",
            use_container_width=True
        ):

            if save_sheet(
                edited_df,
                "內部公告"
            ):

                st.success(
                    "公告已成功儲存！"
                )

                st.rerun()

    # =====================================================
    # 工具維護
    # =====================================================

    elif maintenance_type == "常用工具":

        df = load_tools()

        st.markdown(
            '<div class="section-title">🧰 常用工具</div>',
            unsafe_allow_html=True
        )

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            height=550,
            key="tools_editor"
        )

        if st.button(
            "💾 儲存常用工具",
            use_container_width=True
        ):

            if save_sheet(
                edited_df,
                "常用工具"
            ):

                st.success(
                    "常用工具已成功儲存！"
                )

                st.rerun()
