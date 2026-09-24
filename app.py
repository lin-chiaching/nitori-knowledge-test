import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from datetime import datetime
from openpyxl import load_workbook
import re
import html


# ============================================================
# 基本設定
# ============================================================

FILE_PATH = "DATA/知識庫實驗品.xlsx"

st.set_page_config(
    page_title="客服作業系統",
    page_icon="🐱",
    layout="wide"
)


# ============================================================
# 日系貓咪風格
# ============================================================

st.markdown(
    """
    <style>

    /* ========================================================
       全站
       ======================================================== */

    .stApp {
        background: #F7F4EE;
        color: #5F5A53;
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    /* 標題 */
    h1 {
        color: #514C45 !important;
        font-weight: 700 !important;
        letter-spacing: 0.03em;
    }

    h2 {
        color: #5B554D !important;
        font-weight: 650 !important;
    }

    h3 {
        color: #655F57 !important;
        font-weight: 650 !important;
    }

    h4 {
        color: #625C54 !important;
    }

    p {
        color: #706A62;
    }


    /* ========================================================
       側邊欄
       ======================================================== */

    [data-testid="stSidebar"] {
        background: #EEE8DD;
        border-right: 1px solid #DED4C5;
    }

    [data-testid="stSidebar"] > div:first-child {
        padding-top: 1.5rem;
    }

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #554F48 !important;
    }

    [data-testid="stSidebar"] .stRadio label {
        color: #625C54 !important;
        font-weight: 500;
    }

    [data-testid="stSidebar"] .stRadio label:hover {
        color: #B56F55 !important;
    }

    /* ========================================================
       按鈕
       ======================================================== */

    .stButton > button,
    .stLinkButton > a {
        border-radius: 12px !important;
        border: 1px solid #D8C9B8 !important;
        background: #FFFDF9 !important;
        color: #665F57 !important;
        font-weight: 600 !important;
        transition: all 0.2s ease;
    }

    .stButton > button:hover,
    .stLinkButton > a:hover {
        border-color: #C98568 !important;
        color: #B56F55 !important;
        background: #FFF9F3 !important;
        transform: translateY(-1px);
    }

    /* Primary */
    button[kind="primary"] {
        background: #C98568 !important;
        border-color: #C98568 !important;
        color: white !important;
    }

    button[kind="primary"]:hover {
        background: #B8735B !important;
        border-color: #B8735B !important;
        color: white !important;
    }


    /* ========================================================
       輸入框
       ======================================================== */

    div[data-baseweb="input"] > div,
    div[data-baseweb="textarea"] > div {
        background: #FFFDF9 !important;
        border: 1px solid #DCCFC0 !important;
        border-radius: 12px !important;
    }

    div[data-baseweb="input"] > div:focus-within,
    div[data-baseweb="textarea"] > div:focus-within {
        border-color: #C98568 !important;
        box-shadow: 0 0 0 2px rgba(201, 133, 104, 0.12) !important;
    }

    input,
    textarea {
        color: #5E5952 !important;
    }


    /* ========================================================
       首頁 Hero
       ======================================================== */

    .hero-card {
        background:
            linear-gradient(
                135deg,
                #FFFDF8 0%,
                #F5EEE4 100%
            );
        border: 1px solid #E1D5C5;
        border-radius: 24px;
        padding: 35px 40px;
        margin-bottom: 25px;
        box-shadow: 0 6px 20px rgba(99, 82, 64, 0.06);
    }

    .hero-cat {
        font-size: 76px;
        line-height: 1;
        margin-bottom: 8px;
    }

    .hero-title {
        font-size: 34px;
        font-weight: 750;
        color: #514C45;
        letter-spacing: 0.04em;
        margin-bottom: 8px;
    }

    .hero-subtitle {
        font-size: 15px;
        color: #8A8177;
        letter-spacing: 0.05em;
    }

    .hero-message {
        margin-top: 18px;
        display: inline-block;
        background: #F2E2D8;
        color: #8A5E4C;
        border-radius: 20px;
        padding: 8px 16px;
        font-size: 14px;
    }


    /* ========================================================
       快速入口
       ======================================================== */

    .quick-card {
        background: #FFFDF9;
        border: 1px solid #E1D8CB;
        border-radius: 18px;
        padding: 22px;
        min-height: 145px;
        box-shadow: 0 4px 14px rgba(90, 72, 52, 0.04);
        transition: all 0.2s ease;
    }

    .quick-card:hover {
        transform: translateY(-2px);
        border-color: #D3B8A6;
        box-shadow: 0 7px 18px rgba(90, 72, 52, 0.08);
    }

    .quick-icon {
        font-size: 30px;
        margin-bottom: 8px;
    }

    .quick-title {
        font-size: 17px;
        font-weight: 700;
        color: #5B554D;
        margin-bottom: 5px;
    }

    .quick-text {
        font-size: 13px;
        color: #8A837A;
        line-height: 1.7;
    }


    /* ========================================================
       公告卡片
       ======================================================== */

    .announcement-card {
        background: #FFFDF9;
        border: 1px solid #E3D9CC;
        border-radius: 18px;
        padding: 22px 24px;
        margin-bottom: 16px;
        box-shadow: 0 4px 15px rgba(90, 72, 52, 0.045);
        position: relative;
    }

    .announcement-card::before {
        content: "";
        position: absolute;
        left: 0;
        top: 18px;
        bottom: 18px;
        width: 4px;
        border-radius: 4px;
        background: #C98568;
    }

    .announcement-title {
        font-size: 18px;
        font-weight: 700;
        color: #5B554D;
        margin-bottom: 10px;
        padding-left: 5px;
    }

    .announcement-content {
        color: #716A61;
        line-height: 1.8;
        white-space: pre-line;
        padding-left: 5px;
    }

    .announcement-meta {
        color: #9A9187;
        font-size: 12px;
        margin-top: 12px;
        line-height: 1.8;
        padding-left: 5px;
    }


    /* ========================================================
       工具卡片
       ======================================================== */

    .tool-card {
        background: #FFFDF9;
        border: 1px solid #E2D8CB;
        border-radius: 18px;
        padding: 22px;
        min-height: 155px;
        margin-bottom: 8px;
        box-shadow: 0 4px 15px rgba(90, 72, 52, 0.045);
        transition: all 0.2s ease;
    }

    .tool-card:hover {
        transform: translateY(-3px);
        border-color: #CFAE9A;
        box-shadow: 0 8px 22px rgba(90, 72, 52, 0.09);
    }

    .tool-icon {
        font-size: 28px;
        margin-bottom: 8px;
    }

    .tool-title {
        font-size: 17px;
        font-weight: 700;
        color: #5B554D;
        margin-bottom: 7px;
    }

    .tool-description {
        font-size: 13px;
        line-height: 1.7;
        color: #8A837A;
        min-height: 44px;
    }

    .tool-category-title {
        color: #665F57;
        font-size: 20px;
        font-weight: 700;
        margin-top: 15px;
        margin-bottom: 15px;
        padding-left: 4px;
    }


    /* ========================================================
       搜尋結果
       ======================================================== */

    .search-header {
        background: #F2E7DB;
        border-radius: 18px;
        padding: 18px 22px;
        margin-bottom: 20px;
        border: 1px solid #E2D3C3;
    }

    .result-card {
        background: #FFFDF9;
        padding: 20px 22px;
        border-radius: 16px;
        border: 1px solid #E2D9CE;
        margin-bottom: 8px;
        box-shadow: 0 3px 12px rgba(90, 72, 52, 0.035);
    }

    .result-title {
        font-size: 17px;
        font-weight: 700;
        color: #5B554D;
    }

    .score-badge {
        display: inline-block;
        background: #E6EFE7;
        color: #667C69;
        padding: 4px 9px;
        border-radius: 10px;
        font-size: 11px;
        margin-top: 8px;
    }


    /* ========================================================
       分類標籤
       ======================================================== */

    .section-label {
        color: #8A5E4C;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 0.08em;
        margin-bottom: 5px;
    }


    /* ========================================================
       Streamlit Divider
       ======================================================== */

    hr {
        border-color: #E5DDD3 !important;
    }


    /* ========================================================
       Dataframe
       ======================================================== */

    [data-testid="stDataFrame"] {
        border: 1px solid #DED5CA;
        border-radius: 12px;
        overflow: hidden;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 公告內容清理
# ============================================================

def clean_announcement_content(text):

    if text is None:
        return ""

    try:
        if pd.isna(text):
            return ""
    except Exception:
        pass

    text = str(text)

    if not text.strip():
        return ""

    text = html.unescape(text)

    # Markdown code fence
    text = re.sub(
        r"```(?:html|HTML|text|TEXT)?",
        "",
        text
    )

    text = text.replace("```", "")

    # SVG
    text = re.sub(
        r"<svg\b[^>]*>.*?</svg\s*>",
        "",
        text,
        flags=re.IGNORECASE | re.DOTALL
    )

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

    text = re.sub(
        r"</p\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"</div\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"</h[1-6]\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"</li\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # 移除所有 HTML
    text = re.sub(
        r"<[^>]*>",
        "",
        text,
        flags=re.DOTALL
    )

    # 殘留屬性
    text = re.sub(
        r'\bid\s*=\s*"[^"]*"',
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r'\bclass\s*=\s*"[^"]*"',
        "",
        text,
        flags=re.IGNORECASE
    )

    # 跳脫符號
    text = text.replace("\\~", "~")
    text = text.replace("\\", "")

    lines = []

    for line in text.splitlines():

        line = line.strip()

        if not line:
            continue

        if line.lower() in [
            "svg",
            "div",
            "/div",
            "p",
            "/p"
        ]:
            continue

        line = re.sub(
            r"[ \t]+",
            " ",
            line
        )

        lines.append(line)

    text = "\n".join(lines)

    text = re.sub(
        r"\n{3,}",
        "\n\n",
        text
    )

    return text.strip()


# ============================================================
# 公告標題清理
# ============================================================

def clean_announcement_title(text):

    if text is None:
        return ""

    try:
        if pd.isna(text):
            return ""
    except Exception:
        pass

    text = str(text)

    text = html.unescape(text)

    text = re.sub(
        r"```(?:html|HTML|text|TEXT)?",
        "",
        text
    )

    text = text.replace("```", "")

    text = re.sub(
        r"<[^>]*>",
        "",
        text
    )

    text = text.replace("📌", "")

    text = text.strip()

    if text.startswith("【") and text.endswith("】"):
        text = text[1:-1].strip()

    return text


# ============================================================
# HTML 安全
# ============================================================

def safe_html_text(text):

    if text is None:
        return ""

    try:
        if pd.isna(text):
            return ""
    except Exception:
        pass

    return html.escape(str(text))


# ============================================================
# 讀取資料
# ============================================================

@st.cache_data
def load_data():

    df = pd.read_excel(
        FILE_PATH,
        sheet_name=0
    ).fillna("")

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


# ============================================================
# 搜尋詞
# ============================================================

@st.cache_data
def load_synonyms():

    try:

        df = pd.read_excel(
            FILE_PATH,
            sheet_name="搜尋詞對照"
        ).fillna("")

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        return df

    except Exception:

        return pd.DataFrame()


# ============================================================
# 公告
# ============================================================

@st.cache_data
def load_announcements():

    try:

        df = pd.read_excel(
            FILE_PATH,
            sheet_name="內部公告"
        ).fillna("")

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        required_columns = [
            "公告編號",
            "公告標題",
            "公告內容",
            "公告日期",
            "活動期限",
            "點數期限",
            "顯示狀態"
        ]

        for column in required_columns:

            if column not in df.columns:
                df[column] = ""

        df = df[required_columns]

        # 直接在載入時清理公告
        df["公告標題"] = df["公告標題"].apply(
            clean_announcement_title
        )

        df["公告內容"] = df["公告內容"].apply(
            clean_announcement_content
        )

        df["公告日期"] = df["公告日期"].apply(
            clean_announcement_content
        )

        df["活動期限"] = df["活動期限"].apply(
            clean_announcement_content
        )

        df["點數期限"] = df["點數期限"].apply(
            clean_announcement_content
        )

        return df

    except Exception:

        return pd.DataFrame()


# ============================================================
# 常用工具
# ============================================================

@st.cache_data
def load_tools():

    try:

        df = pd.read_excel(
            FILE_PATH,
            sheet_name="常用工具"
        ).fillna("")

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        required_columns = [
            "工具編號",
            "工具名稱",
            "工具說明",
            "工具網址",
            "工具分類",
            "顯示狀態",
            "顯示順序"
        ]

        for column in required_columns:

            if column not in df.columns:
                df[column] = ""

        return df[required_columns]

    except Exception:

        return pd.DataFrame()


# ============================================================
# 載入
# ============================================================

df = load_data()
synonym_df = load_synonyms()
announcement_df = load_announcements()
tool_df = load_tools()


# ============================================================
# 儲存 Excel
# ============================================================

def save_sheet(sheet_name, data):

    wb = load_workbook(FILE_PATH)

    ws = wb[sheet_name]

    if ws.max_row > 0:

        ws.delete_rows(
            1,
            ws.max_row
        )

    for col_num, column_name in enumerate(
        data.columns,
        1
    ):

        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )

    for row_num, row_data in enumerate(
        data.itertuples(index=False),
        2
    ):

        for col_num, value in enumerate(
            row_data,
            1
        ):

            ws.cell(
                row=row_num,
                column=col_num,
                value=value
            )

    wb.save(FILE_PATH)

    st.cache_data.clear()


def save_knowledge_data(data):

    save_sheet(
        "知識庫",
        data
    )


# ============================================================
# 搜尋詞擴充
# ============================================================

def expand_search_word(keyword):

    keyword = str(keyword).strip()

    if keyword == "":
        return keyword

    expanded_words = [keyword]

    if not synonym_df.empty:

        for _, row in synonym_df.iterrows():

            try:

                original = str(
                    row.iloc[0]
                ).strip()

                synonym = str(
                    row.iloc[1]
                ).strip()

                if original == keyword:
                    expanded_words.append(
                        synonym
                    )

                elif synonym == keyword:
                    expanded_words.append(
                        original
                    )

            except Exception:
                pass

    return " ".join(
        list(
            dict.fromkeys(
                expanded_words
            )
        )
    )


# ============================================================
# 側邊欄
# ============================================================

st.sidebar.markdown(
    """
    <div style="
        text-align:center;
        padding:10px 0 18px 0;
    ">
        <div style="font-size:52px;">🐱</div>
        <div style="
            font-size:20px;
            font-weight:700;
            color:#554F48;
        ">
            客服作業系統
        </div>
        <div style="
            font-size:11px;
            color:#91887E;
            letter-spacing:0.08em;
            margin-top:4px;
        ">
            NITORI CUSTOMER SERVICE
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.sidebar.divider()

page = st.sidebar.radio(
    "功能",
    [
        "🏠 首頁",
        "🔍 知識搜尋",
        "🗺️ 門市地圖",
        "📢 公告＆常用工具",
        "🛠️ 資料維護"
    ]
)

st.sidebar.divider()

st.sidebar.markdown(
    """
    <div style="
        text-align:center;
        color:#9A9187;
        font-size:12px;
        line-height:1.7;
        padding:5px;
    ">
        🐾 今天也一起<br>
        把顧客問題解決掉吧！
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# ① 首頁
# ============================================================

if page == "🏠 首頁":

    st.markdown(
        """
        <div class="hero-card">

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
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-label">QUICK ACCESS</div>',
        unsafe_allow_html=True
    )

    st.subheader("快速入口")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            """
            <div class="quick-card">
                <div class="quick-icon">🔍</div>
                <div class="quick-title">知識搜尋</div>
                <div class="quick-text">
                    快速查詢客服知識、商品問題與常見處理方式。
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            """
            <div class="quick-card">
                <div class="quick-icon">🗺️</div>
                <div class="quick-title">門市地圖</div>
                <div class="quick-text">
                    查詢門市地址、交通、停車及相關資訊。
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            """
            <div class="quick-card">
                <div class="quick-icon">📢</div>
                <div class="quick-title">公告＆工具</div>
                <div class="quick-text">
                    查看最新內部公告，以及日常工作常用工具。
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        '<div class="section-label">LATEST NEWS</div>',
        unsafe_allow_html=True
    )

    st.subheader("最新公告")

    if announcement_df.empty:

        st.info("目前沒有公告。")

    else:

        show_announcements = announcement_df[
            announcement_df["顯示狀態"]
            .astype(str)
            .str.strip()
            == "顯示"
        ]

        if show_announcements.empty:

            st.info("目前沒有公告。")

        else:

            show_announcements = show_announcements.tail(3)

            for _, row in show_announcements.iloc[::-1].iterrows():

                title = clean_announcement_title(
                    row.get("公告標題", "")
                )

                content = clean_announcement_content(
                    row.get("公告內容", "")
                )

                date = clean_announcement_content(
                    row.get("公告日期", "")
                )

                st.markdown(
                    f"""
                    <div class="announcement-card">

                        <div class="announcement-title">
                            📌 {safe_html_text(title)}
                        </div>

                        <div class="announcement-content">
                            {safe_html_text(content).replace(chr(10), "<br>")}
                        </div>

                        <div class="announcement-meta">
                            📅 公告日期：{safe_html_text(date)}
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True
                )


# ============================================================
# ② 知識搜尋
# ============================================================

elif page == "🔍 知識搜尋":

    st.markdown(
        '<div class="section-label">KNOWLEDGE SEARCH</div>',
        unsafe_allow_html=True
    )

    st.title("🔍 客服知識搜尋")

    st.markdown(
        """
        <div class="search-header">
            🐱 <b>小貓幫你找資料</b><br>
            <span style="font-size:13px;color:#8A8177;">
            輸入顧客問題、商品名稱或關鍵字即可開始搜尋
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    keyword = st.text_input(
        "搜尋",
        placeholder="例如：重量、床墊、退貨、配送、保固……",
        label_visibility="collapsed"
    )

    if keyword:

        search_text = expand_search_word(keyword)

        search_columns = []

        possible_columns = [
            "問題簡述",
            "解答",
            "判斷方式",
            "關鍵字",
            "大分類",
            "中分類",
            "小分類"
        ]

        for column in possible_columns:

            if column in df.columns:
                search_columns.append(column)

        if not search_columns:

            st.error(
                "知識庫沒有可以搜尋的欄位。"
            )

        else:

            search_data = (
                df[search_columns]
                .astype(str)
                .agg(" ".join, axis=1)
            )

            vectorizer = TfidfVectorizer(
                analyzer="char",
                ngram_range=(1, 3)
            )

            matrix = vectorizer.fit_transform(
                search_data
            )

            query_vector = vectorizer.transform(
                [search_text]
            )

            similarity = cosine_similarity(
                query_vector,
                matrix
            ).flatten()

            df_result = df.copy()

            df_result["_score"] = similarity

            df_result = df_result.sort_values(
                "_score",
                ascending=False
            )

            df_result = df_result[
                df_result["_score"] > 0
            ]

            st.caption(
                f"🔎 找到 {len(df_result)} 筆相關資料"
            )

            if len(df_result) == 0:

                st.warning(
                    "目前沒有找到相關資料，請換個關鍵字試試看。"
                )

            else:

                display_result = df_result.head(10)

                for _, row in display_result.iterrows():

                    score = row["_score"]

                    title = row.get(
                        "問題簡述",
                        "相關資料"
                    )

                    st.markdown(
                        f"""
                        <div class="result-card">

                            <div class="result-title">
                                📌 {safe_html_text(title)}
                            </div>

                            <div class="score-badge">
                                相關程度 {score:.1%}
                            </div>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    info_col1, info_col2, info_col3 = st.columns(3)

                    with info_col1:

                        if "大分類" in df.columns:

                            st.caption("大分類")
                            st.write(row["大分類"])

                    with info_col2:

                        if "中分類" in df.columns:

                            st.caption("中分類")
                            st.write(row["中分類"])

                    with info_col3:

                        if "小分類" in df.columns:

                            st.caption("小分類")
                            st.write(row["小分類"])

                    if "解答" in df.columns:

                        st.markdown("**📌 解答**")
                        st.write(row["解答"])

                    if "判斷方式" in df.columns:

                        if str(
                            row["判斷方式"]
                        ).strip():

                            st.markdown(
                                "**🔎 判斷方式**"
                            )

                            st.write(
                                row["判斷方式"]
                            )

                    if "判斷使用的工具" in df.columns:

                        if str(
                            row["判斷使用的工具"]
                        ).strip():

                            st.markdown(
                                "**🛠️ 判斷使用的工具**"
                            )

                            st.write(
                                row["判斷使用的工具"]
                            )

                    if "關鍵字" in df.columns:

                        if str(
                            row["關鍵字"]
                        ).strip():

                            st.caption(
                                "關鍵字：" +
                                str(row["關鍵字"])
                            )

                    st.divider()


# ============================================================
# ③ 門市地圖
# ============================================================

elif page == "🗺️ 門市地圖":

    st.markdown(
        '<div class="section-label">STORE MAP</div>',
        unsafe_allow_html=True
    )

    st.title("🗺️ 門市地圖")

    st.caption(
        "宜得利家居門市資訊查詢"
    )

    components.iframe(
        "https://app.powerbi.com/reportEmbed?reportId=851d09f1-f169-4049-9090-640f09cdc385&autoAuth=true&ctid=94d67acc-8c17-4d3b-8463-9829cf08ef55",
        height=800
    )


# ============================================================
# ④ 公告＆常用工具
# ============================================================

elif page == "📢 公告＆常用工具":

    st.markdown(
        '<div class="section-label">INFORMATION & TOOLS</div>',
        unsafe_allow_html=True
    )

    st.title("📢 公告＆常用工具")

    st.caption(
        "查看內部公告及常用系統、表單與文件"
    )

    st.divider()


    # ========================================================
    # 公告
    # ========================================================

    st.subheader("📢 內部公告")

    if announcement_df.empty:

        st.info("目前沒有公告資料。")

    else:

        show_announcements = announcement_df[
            announcement_df["顯示狀態"]
            .astype(str)
            .str.strip()
            == "顯示"
        ]

        if show_announcements.empty:

            st.info("目前沒有公告資料。")

        else:

            show_announcements = show_announcements.iloc[::-1]

            for _, row in show_announcements.iterrows():

                title = clean_announcement_title(
                    row.get("公告標題", "")
                )

                content = clean_announcement_content(
                    row.get("公告內容", "")
                )

                date = clean_announcement_content(
                    row.get("公告日期", "")
                )

                activity = clean_announcement_content(
                    row.get("活動期限", "")
                )

                point = clean_announcement_content(
                    row.get("點數期限", "")
                )

                meta = ""

                if date:
                    meta += f"📅 公告日期：{safe_html_text(date)}<br>"

                if activity:
                    meta += f"🗓️ 活動期限：{safe_html_text(activity)}<br>"

                if point:
                    meta += f"🪙 點數期限：{safe_html_text(point)}"

                st.markdown(
                    f"""
                    <div class="announcement-card">

                        <div class="announcement-title">
                            📌 {safe_html_text(title)}
                        </div>

                        <div class="announcement-content">
                            {safe_html_text(content).replace(chr(10), "<br>")}
                        </div>

                        <div class="announcement-meta">
                            {meta}
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True
                )


    # ========================================================
    # 常用工具
    # ========================================================

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        '<div class="section-label">QUICK TOOLS</div>',
        unsafe_allow_html=True
    )

    st.subheader("🔗 常用工具")

    if tool_df.empty:

        st.info("目前沒有常用工具。")

    else:

        show_tools = tool_df[
            tool_df["顯示狀態"]
            .astype(str)
            .str.strip()
            == "顯示"
        ]

        if show_tools.empty:

            st.info("目前沒有常用工具。")

        else:

            categories = show_tools[
                "工具分類"
            ].astype(str).unique()

            for category in categories:

                st.markdown(
                    f"""
                    <div class="tool-category-title">
                        {safe_html_text(category)}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                category_tools = show_tools[
                    show_tools["工具分類"]
                    .astype(str)
                    == category
                ]

                category_tools = category_tools.sort_values(
                    by="顯示順序"
                )

                tools_list = category_tools.to_dict(
                    "records"
                )

                for i in range(
                    0,
                    len(tools_list),
                    3
                ):

                    cols = st.columns(
                        3,
                        gap="medium"
                    )

                    for j, tool in enumerate(
                        tools_list[i:i + 3]
                    ):

                        with cols[j]:

                            tool_name = str(
                                tool.get(
                                    "工具名稱",
                                    ""
                                )
                            )

                            tool_description = str(
                                tool.get(
                                    "工具說明",
                                    ""
                                )
                            )

                            # 工具圖示
                            if "地圖" in tool_name:
                                icon = "🗺️"
                            elif "Power BI" in tool_name:
                                icon = "📊"
                            elif "交接" in tool_name:
                                icon = "📋"
                            elif "信箱" in tool_name:
                                icon = "📧"
                            elif "Excel" in tool_name:
                                icon = "📗"
                            elif "Teams" in tool_name:
                                icon = "💬"
                            else:
                                icon = "🔗"

                            st.markdown(
                                f"""
                                <div class="tool-card">

                                    <div class="tool-icon">
                                        {icon}
                                    </div>

                                    <div class="tool-title">
                                        {safe_html_text(tool_name)}
                                    </div>

                                    <div class="tool-description">
                                        {safe_html_text(tool_description)}
                                    </div>

                                </div>
                                """,
                                unsafe_allow_html=True
                            )

                            url = str(
                                tool.get(
                                    "工具網址",
                                    ""
                                )
                            ).strip()

                            tool_id = str(
                                tool.get(
                                    "工具編號",
                                    j
                                )
                            )

                            if (
                                url
                                and url != "待補"
                                and url.lower() != "nan"
                            ):

                                st.link_button(
                                    "開啟工具 →",
                                    url,
                                    use_container_width=True
                                )

                            else:

                                st.button(
                                    "網址待補",
                                    disabled=True,
                                    use_container_width=True,
                                    key=f"tool_{tool_id}"
                                )


# ============================================================
# ⑤ 資料維護
# ============================================================

elif page == "🛠️ 資料維護":

    st.markdown(
        '<div class="section-label">DATA MANAGEMENT</div>',
        unsafe_allow_html=True
    )

    st.title("🛠️ 資料維護")

    st.caption(
        "管理客服知識、內部公告及常用工具"
    )

    st.divider()

    maintenance_type = st.radio(
        "請選擇要維護的資料",
        [
            "📚 知識庫",
            "📢 內部公告",
            "🔗 常用工具"
        ],
        horizontal=True
    )


    # ========================================================
    # 知識庫
    # ========================================================

    if maintenance_type == "📚 知識庫":

        st.subheader("📋 目前知識庫資料")

        st.dataframe(
            df,
            use_container_width=True,
            height=400
        )

        st.divider()

        st.subheader("➕ 新增知識")

        new_data = {}

        for column in df.columns:

            if column == "序號":
                continue

            new_data[column] = st.text_input(
                column,
                key=f"new_{column}"
            )

        if st.button(
            "➕ 新增資料",
            type="primary"
        ):

            if "序號" in df.columns:

                try:

                    numbers = pd.to_numeric(
                        df["序號"],
                        errors="coerce"
                    )

                    max_number = numbers.max()

                    if pd.isna(max_number):
                        new_number = 1
                    else:
                        new_number = int(max_number) + 1

                except Exception:

                    new_number = len(df) + 1

                new_data["序號"] = new_number

            new_row = pd.DataFrame([new_data])

            new_row = new_row.reindex(
                columns=df.columns
            )

            df = pd.concat(
                [
                    df,
                    new_row
                ],
                ignore_index=True
            )

            save_knowledge_data(df)

            st.success("資料已新增！")

            st.rerun()

        st.divider()

        st.subheader("✏️ 修改知識")

        if len(df) > 0:

            selected_index = st.selectbox(
                "請選擇要修改的資料",
                options=df.index,
                format_func=lambda x:
                    f"{x + 1}｜" +
                    str(
                        df.loc[
                            x,
                            "問題簡述"
                        ]
                        if "問題簡述" in df.columns
                        else ""
                    )
            )

            selected_row = df.loc[selected_index]

            edit_data = {}

            for column in df.columns:

                if column == "序號":

                    st.text_input(
                        column,
                        value=str(
                            selected_row[column]
                        ),
                        disabled=True,
                        key=f"edit_knowledge_{column}"
                    )

                else:

                    edit_data[column] = st.text_area(
                        column,
                        value=str(
                            selected_row[column]
                        ),
                        key=f"edit_knowledge_{selected_index}_{column}"
                    )

            if st.button(
                "💾 儲存修改",
                type="primary"
            ):

                for column in edit_data:

                    df.loc[
                        selected_index,
                        column
                    ] = edit_data[column]

                if "更新日期" in df.columns:

                    df.loc[
                        selected_index,
                        "更新日期"
                    ] = datetime.now().strftime(
                        "%Y-%m-%d"
                    )

                if "更新人" in df.columns:

                    df.loc[
                        selected_index,
                        "更新人"
                    ] = "客服"

                save_knowledge_data(df)

                st.success("資料已更新！")

                st.rerun()


    # ========================================================
    # 公告維護
    # ========================================================

    elif maintenance_type == "📢 內部公告":

        st.subheader("📋 公告資料")

        if announcement_df.empty:

            st.info("目前沒有公告資料。")

        else:

            st.dataframe(
                announcement_df,
                use_container_width=True,
                height=300
            )

        st.divider()

        st.subheader("➕ 新增公告")

        announcement_title = st.text_input(
            "公告標題",
            key="announcement_title"
        )

        announcement_content = st.text_area(
            "公告內容",
            key="announcement_content"
        )

        announcement_date = st.text_input(
            "公告日期",
            value=datetime.now().strftime(
                "%Y-%m-%d"
            ),
            key="announcement_date"
        )

        activity_period = st.text_input(
            "活動期限",
            placeholder="例如：2026-09-24 ～ 2026-10-12",
            key="activity_period"
        )

        point_period = st.text_input(
            "點數期限",
            placeholder="例如：限期點數：2026/10/27～2026/12/13",
            key="point_period"
        )

        announcement_status = st.selectbox(
            "顯示狀態",
            ["顯示", "隱藏"],
            key="announcement_status"
        )

        if st.button(
            "➕ 新增公告",
            type="primary"
        ):

            if announcement_title.strip() == "":

                st.warning(
                    "請先輸入公告標題。"
                )

            else:

                if "公告編號" in announcement_df.columns:

                    numbers = pd.to_numeric(
                        announcement_df["公告編號"],
                        errors="coerce"
                    )

                    if numbers.notna().any():
                        new_id = int(numbers.max()) + 1
                    else:
                        new_id = 1

                else:

                    new_id = 1

                new_announcement = pd.DataFrame(
                    [
                        {
                            "公告編號": new_id,
                            "公告標題": clean_announcement_title(
                                announcement_title
                            ),
                            "公告內容": clean_announcement_content(
                                announcement_content
                            ),
                            "公告日期": clean_announcement_content(
                                announcement_date
                            ),
                            "活動期限": clean_announcement_content(
                                activity_period
                            ),
                            "點數期限": clean_announcement_content(
                                point_period
                            ),
                            "顯示狀態": announcement_status
                        }
                    ]
                )

                announcement_df = pd.concat(
                    [
                        announcement_df,
                        new_announcement
                    ],
                    ignore_index=True
                )

                save_sheet(
                    "內部公告",
                    announcement_df
                )

                st.success("公告已新增！")

                st.rerun()

        st.divider()

        st.subheader("✏️ 修改公告")

        if not announcement_df.empty:

            selected_announcement = st.selectbox(
                "選擇要修改的公告",
                options=announcement_df.index,
                format_func=lambda x:
                    f"{x + 1}｜" +
                    str(
                        announcement_df.loc[
                            x,
                            "公告標題"
                        ]
                    ),
                key="selected_announcement"
            )

            selected_row = announcement_df.loc[
                selected_announcement
            ]

            edit_title = st.text_input(
                "公告標題",
                value=clean_announcement_title(
                    selected_row.get(
                        "公告標題",
                        ""
                    )
                ),
                key="edit_announcement_title"
            )

            edit_content = st.text_area(
                "公告內容",
                value=clean_announcement_content(
                    selected_row.get(
                        "公告內容",
                        ""
                    )
                ),
                key="edit_announcement_content"
            )

            edit_date = st.text_input(
                "公告日期",
                value=clean_announcement_content(
                    selected_row.get(
                        "公告日期",
                        ""
                    )
                ),
                key="edit_announcement_date"
            )

            edit_activity = st.text_input(
                "活動期限",
                value=clean_announcement_content(
                    selected_row.get(
                        "活動期限",
                        ""
                    )
                ),
                key="edit_activity"
            )

            edit_point = st.text_input(
                "點數期限",
                value=clean_announcement_content(
                    selected_row.get(
                        "點數期限",
                        ""
                    )
                ),
                key="edit_point"
            )

            edit_status = st.selectbox(
                "顯示狀態",
                ["顯示", "隱藏"],
                index=(
                    0
                    if str(
                        selected_row.get(
                            "顯示狀態",
                            "顯示"
                        )
                    ).strip() == "顯示"
                    else 1
                ),
                key="edit_announcement_status"
            )

            if st.button(
                "💾 儲存公告修改",
                type="primary"
            ):

                announcement_df.loc[
                    selected_announcement,
                    "公告標題"
                ] = clean_announcement_title(
                    edit_title
                )

                announcement_df.loc[
                    selected_announcement,
                    "公告內容"
                ] = clean_announcement_content(
                    edit_content
                )

                announcement_df.loc[
                    selected_announcement,
                    "公告日期"
                ] = clean_announcement_content(
                    edit_date
                )

                announcement_df.loc[
                    selected_announcement,
                    "活動期限"
                ] = clean_announcement_content(
                    edit_activity
                )

                announcement_df.loc[
                    selected_announcement,
                    "點數期限"
                ] = clean_announcement_content(
                    edit_point
                )

                announcement_df.loc[
                    selected_announcement,
                    "顯示狀態"
                ] = edit_status

                save_sheet(
                    "內部公告",
                    announcement_df
                )

                st.success("公告已更新！")

                st.rerun()


    # ========================================================
    # 常用工具維護
    # ========================================================

    elif maintenance_type == "🔗 常用工具":

        st.subheader("📋 工具資料")

        if tool_df.empty:

            st.info("目前沒有工具資料。")

        else:

            st.dataframe(
                tool_df,
                use_container_width=True,
                height=300
            )

        st.divider()

        st.subheader("➕ 新增工具")

        tool_name = st.text_input(
            "工具名稱",
            key="tool_name"
        )

        tool_description = st.text_input(
            "工具說明",
            key="tool_description"
        )

        tool_url = st.text_input(
            "工具網址",
            placeholder="https://",
            key="tool_url"
        )

        tool_category = st.selectbox(
            "工具分類",
            [
                "🏢 內部系統",
                "🛒 網購平台",
                "📋 表單",
                "📚 文件／資料",
                "🔗 其他"
            ],
            key="tool_category"
        )

        tool_status = st.selectbox(
            "顯示狀態",
            ["顯示", "隱藏"],
            key="tool_status"
        )

        tool_order = st.number_input(
            "顯示順序",
            min_value=1,
            value=1,
            step=1,
            key="tool_order"
        )

        if st.button(
            "➕ 新增工具",
            type="primary"
        ):

            if tool_name.strip() == "":

                st.warning(
                    "請先輸入工具名稱。"
                )

            else:

                if "工具編號" in tool_df.columns:

                    numbers = pd.to_numeric(
                        tool_df["工具編號"],
                        errors="coerce"
                    )

                    if numbers.notna().any():
                        new_tool_id = int(
                            numbers.max()
                        ) + 1
                    else:
                        new_tool_id = 1

                else:

                    new_tool_id = 1

                new_tool = pd.DataFrame(
                    [
                        {
                            "工具編號": new_tool_id,
                            "工具名稱": tool_name,
                            "工具說明": tool_description,
                            "工具網址": tool_url,
                            "工具分類": tool_category,
                            "顯示狀態": tool_status,
                            "顯示順序": tool_order
                        }
                    ]
                )

                tool_df = pd.concat(
                    [
                        tool_df,
                        new_tool
                    ],
                    ignore_index=True
                )

                save_sheet(
                    "常用工具",
                    tool_df
                )

                st.success("工具已新增！")

                st.rerun()

        st.divider()

        st.subheader("✏️ 修改工具")

        if not tool_df.empty:

            selected_tool = st.selectbox(
                "選擇要修改的工具",
                options=tool_df.index,
                format_func=lambda x:
                    f"{x + 1}｜" +
                    str(
                        tool_df.loc[
                            x,
                            "工具名稱"
                        ]
                    ),
                key="selected_tool"
            )

            selected_row = tool_df.loc[
                selected_tool
            ]

            edit_tool_name = st.text_input(
                "工具名稱",
                value=str(
                    selected_row.get(
                        "工具名稱",
                        ""
                    )
                ),
                key="edit_tool_name"
            )

            edit_tool_description = st.text_input(
                "工具說明",
                value=str(
                    selected_row.get(
                        "工具說明",
                        ""
                    )
                ),
                key="edit_tool_description"
            )

            edit_tool_url = st.text_input(
                "工具網址",
                value=str(
                    selected_row.get(
                        "工具網址",
                        ""
                    )
                ),
                key="edit_tool_url"
            )

            category_options = [
                "🏢 內部系統",
                "🛒 網購平台",
                "📋 表單",
                "📚 文件／資料",
                "🔗 其他"
            ]

            current_category = str(
                selected_row.get(
                    "工具分類",
                    ""
                )
            )

            if current_category in category_options:
                category_index = category_options.index(
                    current_category
                )
            else:
                category_index = 0

            edit_tool_category = st.selectbox(
                "工具分類",
                category_options,
                index=category_index,
                key="edit_tool_category"
            )

            edit_tool_status = st.selectbox(
                "顯示狀態",
                ["顯示", "隱藏"],
                index=(
                    0
                    if str(
                        selected_row.get(
                            "顯示狀態",
                            "顯示"
                        )
                    ).strip() == "顯示"
                    else 1
                ),
                key="edit_tool_status"
            )

            current_order = pd.to_numeric(
                selected_row.get(
                    "顯示順序",
                    1
                ),
                errors="coerce"
            )

            if pd.isna(current_order):
                current_order = 1

            edit_tool_order = st.number_input(
                "顯示順序",
                min_value=1,
                value=int(current_order),
                step=1,
                key="edit_tool_order"
            )

            if st.button(
                "💾 儲存工具修改",
                type="primary"
            ):

                tool_df.loc[
                    selected_tool,
                    "工具名稱"
                ] = edit_tool_name

                tool_df.loc[
                    selected_tool,
                    "工具說明"
                ] = edit_tool_description

                tool_df.loc[
                    selected_tool,
                    "工具網址"
                ] = edit_tool_url

                tool_df.loc[
                    selected_tool,
                    "工具分類"
                ] = edit_tool_category

                tool_df.loc[
                    selected_tool,
                    "顯示狀態"
                ] = edit_tool_status

                tool_df.loc[
                    selected_tool,
                    "顯示順序"
                ] = edit_tool_order

                save_sheet(
                    "常用工具",
                    tool_df
                )

                st.success("工具已更新！")

                st.rerun()
