import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from pathlib import Path
import openpyxl
import re
import html


# =========================================================
# 基本設定
# =========================================================

st.set_page_config(
    page_title="客服作業系統",
    page_icon="🐱",
    layout="wide"
)


EXCEL_PATH = Path("DATA/知識庫實驗品.xlsx")

POWER_BI_URL = (
    "https://app.powerbi.com/reportEmbed?"
    "reportId=851d09f1-f169-4049-9090-640f09cdc385"
    "&autoAuth=true"
    "&ctid=94d67acc-8c17-4d3b-8463-9829cf08ef55"
)


# =========================================================
# CSS
# =========================================================

st.markdown(
    """
    <style>

    .stApp {
        background: #F7F4EE;
        color: #5F5A53;
    }

    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    h1, h2, h3 {
        color: #5F5A53;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #EEE8DD;
        border-right: 1px solid #DED4C5;
    }

    section[data-testid="stSidebar"] * {
        color: #5F5A53;
    }

    /* Button */
    .stButton > button {
        border-radius: 12px;
        border: 1px solid #D8C9B8;
        background: #FFFDF9;
        color: #5F5A53;
    }

    .stButton > button:hover {
        border-color: #BDA991;
        color: #5F5A53;
    }

    /* Hero */
    .hero {
        background: #FFFDF9;
        border: 1px solid #E2D8CC;
        border-radius: 24px;
        padding: 35px 30px 30px;
        text-align: center;
        box-shadow: 0 6px 18px rgba(100, 85, 65, 0.06);
        margin-bottom: 25px;
    }

    .hero-cat {
        font-size: 54px;
        margin-bottom: 8px;
    }

    .hero-title {
        font-size: 34px;
        font-weight: 700;
        color: #5F5A53;
        margin-bottom: 8px;
    }

    .hero-subtitle {
        font-size: 16px;
        color: #8A8177;
        margin-bottom: 18px;
    }

    .hero-message {
        display: inline-block;
        background: #F3EDE4;
        border-radius: 14px;
        padding: 10px 18px;
        color: #746B61;
        font-size: 14px;
    }

    /* Cards */
    .card {
        background: #FFFDF9;
        border: 1px solid #E2D8CC;
        border-radius: 18px;
        padding: 22px;
        margin-bottom: 18px;
        box-shadow: 0 4px 12px rgba(100, 85, 65, 0.04);
    }

    .card-title {
        font-size: 19px;
        font-weight: 700;
        color: #5F5A53;
        margin-bottom: 8px;
    }

    .card-text {
        color: #81776C;
        font-size: 14px;
        line-height: 1.7;
    }

    /* Tool card */
    .tool-card {
        background: #FFFDF9;
        border: 1px solid #E2D8CC;
        border-radius: 18px;
        padding: 20px;
        margin-bottom: 15px;
        box-shadow: 0 4px 12px rgba(100, 85, 65, 0.04);
    }

    .tool-card-title {
        font-size: 18px;
        font-weight: 700;
        color: #5F5A53;
        margin-bottom: 6px;
    }

    .tool-card-text {
        color: #81776C;
        font-size: 14px;
        line-height: 1.6;
    }

    /* Search result */
    .result-card {
        background: #FFFDF9;
        border: 1px solid #E2D8CC;
        border-radius: 18px;
        padding: 22px;
        margin-bottom: 16px;
        box-shadow: 0 4px 12px rgba(100, 85, 65, 0.04);
    }

    .result-title {
        font-size: 18px;
        font-weight: 700;
        color: #5F5A53;
        margin-bottom: 10px;
    }

    .result-meta {
        color: #8A8177;
        font-size: 13px;
        line-height: 1.7;
        margin-bottom: 12px;
    }

    .result-answer {
        background: #F7F2EA;
        border-radius: 12px;
        padding: 15px;
        color: #625B53;
        font-size: 14px;
        line-height: 1.8;
        white-space: pre-wrap;
    }

    .result-extra {
        margin-top: 12px;
        color: #81776C;
        font-size: 13px;
        line-height: 1.7;
    }

    .keyword-hit {
        display: inline-block;
        background: #EFE5D8;
        border-radius: 10px;
        padding: 3px 9px;
        margin-right: 5px;
        margin-bottom: 5px;
        font-size: 12px;
        color: #746B61;
    }

    /* Section */
    .section-title {
        font-size: 25px;
        font-weight: 700;
        color: #5F5A53;
        margin-bottom: 5px;
    }

    .section-subtitle {
        color: #8A8177;
        font-size: 14px;
        margin-bottom: 20px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 工具函式
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_search_text(value):
    """
    搜尋文字標準化：
    - 轉小寫
    - 移除空白
    - 移除部分標點
    """
    text = clean_text(value).lower()

    text = re.sub(r"\s+", "", text)

    text = text.replace(
        "，", ""
    ).replace(
        "。", ""
    ).replace(
        "、", ""
    ).replace(
        "！", ""
    ).replace(
        "!", ""
    ).replace(
        "？", ""
    ).replace(
        "?", ""
    ).replace(
        "：", ""
    ).replace(
        ":", ""
    ).replace(
        "；", ""
    ).replace(
        ";", ""
    ).replace(
        "（", ""
    ).replace(
        "）", ""
    ).replace(
        "(", ""
    ).replace(
        ")", ""
    )

    return text


def clean_announcement_content(value):
    text = clean_text(value)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    return html.unescape(text)


def clean_announcement_title(value):
    return clean_text(value)


def safe_html_text(value):
    return html.escape(clean_text(value)).replace("\n", "<br>")


# =========================================================
# Excel
# =========================================================

@st.cache_data(ttl=60)
def get_excel_sheets():
    if not EXCEL_PATH.exists():
        return []

    try:
        return openpyxl.load_workbook(
            EXCEL_PATH,
            read_only=True,
            data_only=True
        ).sheetnames
    except Exception:
        return []


def find_sheet(candidates, fallback_index=None):
    sheets = get_excel_sheets()

    for candidate in candidates:
        if candidate in sheets:
            return candidate

    if fallback_index is not None and fallback_index < len(sheets):
        return sheets[fallback_index]

    return None


@st.cache_data(ttl=60)
def load_data():
    sheet_name = find_sheet(
        ["知識庫", "問題集", "Sheet1"],
        fallback_index=0
    )

    if not sheet_name:
        return pd.DataFrame()

    try:
        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet_name,
            dtype=str
        )

        df = df.fillna("")

        df.columns = [
            str(col).strip()
            for col in df.columns
        ]

        return df

    except Exception as e:
        st.error(f"讀取知識庫失敗：{e}")
        return pd.DataFrame()


@st.cache_data(ttl=60)
def load_synonyms():
    sheet_name = find_sheet(
        ["搜尋詞對照", "關鍵字對照"]
    )

    if not sheet_name:
        return pd.DataFrame()

    try:
        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet_name,
            dtype=str
        )

        return df.fillna("")

    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=60)
def load_announcements():
    sheet_name = find_sheet(
        ["內部公告", "公告"]
    )

    if not sheet_name:
        return pd.DataFrame()

    try:
        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet_name,
            dtype=str
        )

        df = df.fillna("")

        if "內容" in df.columns:
            df["內容"] = df["內容"].apply(
                clean_announcement_content
            )

        if "標題" in df.columns:
            df["標題"] = df["標題"].apply(
                clean_announcement_title
            )

        return df

    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=60)
def load_tools():
    sheet_name = find_sheet(
        ["常用工具", "工具"]
    )

    if not sheet_name:
        return pd.DataFrame()

    try:
        df = pd.read_excel(
            EXCEL_PATH,
            sheet_name=sheet_name,
            dtype=str
        )

        return df.fillna("")

    except Exception:
        return pd.DataFrame()


# =========================================================
# 儲存 Excel
# =========================================================

def save_sheet(sheet_name, data):
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)

        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

        ws = wb.create_sheet(sheet_name)

        for col_idx, column in enumerate(data.columns, start=1):
            ws.cell(
                row=1,
                column=col_idx,
                value=str(column)
            )

        for row_idx, row in enumerate(
            data.itertuples(index=False),
            start=2
        ):
            for col_idx, value in enumerate(
                row,
                start=1
            ):
                if pd.isna(value):
                    value = ""
                ws.cell(
                    row=row_idx,
                    column=col_idx,
                    value=value
                )

        wb.save(EXCEL_PATH)

        get_excel_sheets.clear()
        load_data.clear()
        load_synonyms.clear()
        load_announcements.clear()
        load_tools.clear()

        return True, ""

    except Exception as e:
        return False, str(e)


def save_knowledge_data(data):
    sheet_name = find_sheet(
        ["知識庫", "問題集", "Sheet1"],
        fallback_index=0
    )

    if not sheet_name:
        sheet_name = "知識庫"

    return save_sheet(
        sheet_name,
        data
    )


# =========================================================
# 搜尋詞展開
# =========================================================

def expand_search_word(keyword):
    """
    從 Excel 的搜尋詞對照表取得同義詞。
    支援常見欄位：
    搜尋詞 / 原詞 / 關鍵字
    對照詞 / 同義詞 / 擴充詞
    """

    keyword = clean_text(keyword)

    if not keyword:
        return [keyword]

    words = [keyword]

    synonym_df = load_synonyms()

    if synonym_df.empty:
        return words

    columns = list(synonym_df.columns)

    source_col = None
    target_col = None

    source_candidates = [
        "搜尋詞",
        "原詞",
        "關鍵字",
        "詞語"
    ]

    target_candidates = [
        "對照詞",
        "同義詞",
        "擴充詞",
        "搜尋詞",
        "關鍵字"
    ]

    for col in source_candidates:
        if col in columns:
            source_col = col
            break

    for col in target_candidates:
        if col in columns:
            target_col = col
            break

    if source_col is None or target_col is None:
        return words

    keyword_normalized = normalize_search_text(keyword)

    for _, row in synonym_df.iterrows():
        source = normalize_search_text(
            row.get(source_col, "")
        )

        target = clean_text(
            row.get(target_col, "")
        )

        if not source or not target:
            continue

        if (
            keyword_normalized == source
            or source in keyword_normalized
            or keyword_normalized in source
        ):
            for item in re.split(
                r"[,，、/／;；\n]+",
                target
            ):
                item = item.strip()

                if item and item not in words:
                    words.append(item)

    return words


# =========================================================
# 知識搜尋
# =========================================================

def search_knowledge(keyword, top_n=10):

    df = load_data().copy()

    if df.empty:
        return pd.DataFrame()

    keyword = clean_text(keyword)

    if not keyword:
        return pd.DataFrame()

    # -----------------------------------------------------
    # 實際 Excel 欄位
    # -----------------------------------------------------

    fields = [
        "問題類型",
        "問題細項",
        "商品分類代碼",
        "商品名稱",
        "商品系列",
        "商品編號",
        "問題摘要",
        "標準回覆",
        "判斷方式",
        "判斷使用的工具",
        "關鍵字"
    ]

    fields = [
        field
        for field in fields
        if field in df.columns
    ]

    # -----------------------------------------------------
    # 欄位權重
    # -----------------------------------------------------

    weights = {
        "商品編號": 120,
        "關鍵字": 110,
        "問題摘要": 100,
        "商品名稱": 90,
        "商品系列": 80,
        "問題細項": 60,
        "問題類型": 40,
        "標準回覆": 25,
        "商品分類代碼": 20,
        "判斷方式": 15,
        "判斷使用的工具": 10
    }

    # -----------------------------------------------------
    # 搜尋詞
    # -----------------------------------------------------

    expanded_words = expand_search_word(keyword)

    normalized_words = [
        normalize_search_text(word)
        for word in expanded_words
        if normalize_search_text(word)
    ]

    query = normalize_search_text(keyword)

    # -----------------------------------------------------
    # 計算直接命中分數
    # -----------------------------------------------------

    direct_scores = []
    hit_details = []

    for _, row in df.iterrows():

        score = 0
        hits = []

        # 原始搜尋詞
        for field in fields:

            value = normalize_search_text(
                row.get(field, "")
            )

            if not value:
                continue

            weight = weights.get(field, 10)

            # 完整包含
            if query in value:

                score += weight

                hits.append(field)

                # 問題摘要 / 關鍵字 / 商品名稱
                # 再給一次小幅加分
                if field in [
                    "問題摘要",
                    "關鍵字",
                    "商品名稱"
                ]:
                    score += 20

            # 對照詞
            for expanded_word in normalized_words:

                if expanded_word == query:
                    continue

                if expanded_word in value:

                    score += int(weight * 0.65)

                    if field not in hits:
                        hits.append(field)

        # -------------------------------------------------
        # 文字拆分命中
        # 對比較長的搜尋詞增加準確度
        # -------------------------------------------------

        if len(query) >= 3:

            query_parts = [
                part
                for part in re.split(
                    r"[\s、，,]+",
                    query
                )
                if part
            ]

            for part in query_parts:

                if len(part) < 2:
                    continue

                for field in fields:

                    value = normalize_search_text(
                        row.get(field, "")
                    )

                    if part in value:

                        weight = weights.get(
                            field,
                            10
                        )

                        score += int(
                            weight * 0.15
                        )

        direct_scores.append(score)

        hit_details.append(
            "、".join(hits)
        )

    df["_direct_score"] = direct_scores
    df["_hit_fields"] = hit_details

    # -----------------------------------------------------
    # 短關鍵字
    #
    # 例如：
    # 重量
    # 尺寸
    # 配送
    # 退貨
    #
    # 不讓 TF-IDF 把完全沒直接命中的資料
    # 排到前面。
    # -----------------------------------------------------

    if len(query) <= 2:

        df = df[
            df["_direct_score"] > 0
        ].copy()

        if df.empty:
            return pd.DataFrame()

        df["_tfidf_score"] = 0.0

    else:

        # -------------------------------------------------
        # TF-IDF 作為輔助，不作主要判斷
        # -------------------------------------------------

        corpus = []

        for _, row in df.iterrows():

            text_parts = []

            for field in fields:

                value = normalize_search_text(
                    row.get(field, "")
                )

                if value:
                    text_parts.append(value)

            corpus.append(
                " ".join(text_parts)
            )

        try:

            vectorizer = TfidfVectorizer(
                analyzer="char",
                ngram_range=(1, 3),
                min_df=1,
                sublinear_tf=True
            )

            matrix = vectorizer.fit_transform(
                corpus + [query]
            )

            tfidf_scores = cosine_similarity(
                matrix[-1],
                matrix[:-1]
            ).ravel()

            df["_tfidf_score"] = tfidf_scores

        except Exception:

            df["_tfidf_score"] = 0.0

    # -----------------------------------------------------
    # 最終分數
    #
    # 直接命中 >> TF-IDF
    # -----------------------------------------------------

    df["_score"] = (
        df["_direct_score"] * 1.0
        + df["_tfidf_score"] * 20
    )

    # -----------------------------------------------------
    # 排序
    #
    # 先看直接命中
    # 再看總分
    # 再看 TF-IDF
    # -----------------------------------------------------

    df = df.sort_values(
        by=[
            "_direct_score",
            "_score",
            "_tfidf_score"
        ],
        ascending=False
    )

    df = df.head(top_n).copy()

    # -----------------------------------------------------
    # 分數顯示
    # -----------------------------------------------------

    max_score = df["_score"].max()

    if max_score and max_score > 0:

        df["_score"] = (
            df["_score"] / max_score * 100
        ).round(0)

    else:

        df["_score"] = 0

    # -----------------------------------------------------
    # 清除內部欄位
    # 保留 _score 給畫面使用
    # -----------------------------------------------------

    df = df.drop(
        columns=[
            "_search_text",
            "_direct_score",
            "_tfidf_score"
        ],
        errors="ignore"
    )

    return df


# =========================================================
# Sidebar
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div style="
            text-align:center;
            padding:10px 0 18px 0;
            font-size:38px;
        ">
        🐱
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div style="
            text-align:center;
            font-size:21px;
            font-weight:700;
            color:#5F5A53;
            margin-bottom:18px;
        ">
        客服作業系統
        </div>
        """,
        unsafe_allow_html=True
    )

    page = st.radio(
        "功能選單",
        [
            "🏠 首頁",
            "🔍 知識搜尋",
            "🗺️ 門市地圖",
            "📢 公告＆常用工具",
            "🛠️ 資料維護"
        ],
        label_visibility="collapsed"
    )


# =========================================================
# 首頁
# =========================================================

if page == "🏠 首頁":

    st.markdown(
        """
        <div class="hero">

            <div class="hero-cat">🐱</div>

            <div class="hero-title">
                客服作業系統
            </div>

            <div class="hero-subtitle">
                宜得利家居 Customer Service
            </div>

            <div class="hero-message">
                今天也一起把顧客問題快速找到吧！
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            """
            <div class="tool-card">

                <div class="tool-card-title">
                    🔍 知識搜尋
                </div>

                <div class="tool-card-text">
                    輸入顧客詢問的內容，
                    快速找到相關的知識庫資料與標準回覆。
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            """
            <div class="tool-card">

                <div class="tool-card-title">
                    🗺️ 門市地圖
                </div>

                <div class="tool-card-text">
                    查詢門市位置、交通、
                    停車及相關資訊。
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            """
            <div class="tool-card">

                <div class="tool-card-title">
                    📢 公告＆常用工具
                </div>

                <div class="tool-card-text">
                    查看內部公告，
                    並快速使用常見作業工具。
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("")

    df = load_data()

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            f"""
            <div class="card">

                <div class="card-title">
                    📚 知識庫
                </div>

                <div class="card-text">
                    目前共有
                    <b>{len(df):,}</b>
                    筆資料。
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:

        announcements = load_announcements()

        st.markdown(
            f"""
            <div class="card">

                <div class="card-title">
                    📢 內部公告
                </div>

                <div class="card-text">
                    目前共有
                    <b>{len(announcements):,}</b>
                    筆公告資料。
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


# =========================================================
# 知識搜尋
# =========================================================

elif page == "🔍 知識搜尋":

    st.markdown(
        '<div class="section-title">🔍 知識庫搜尋</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-subtitle">
            輸入顧客詢問的關鍵字，小貓會幫你找出相關資料。
        </div>
        """,
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(
        [5, 1]
    )

    with col1:

        keyword = st.text_input(
            "搜尋",
            placeholder="例如：重量、床包尺寸、退貨、配送、商品編號……",
            label_visibility="collapsed"
        )

    with col2:

        top_n = st.slider(
            "顯示筆數",
            min_value=5,
            max_value=30,
            value=10,
            step=5
        )

    if keyword.strip():

        results = search_knowledge(
            keyword,
            top_n
        )

        if results.empty:

            st.markdown(
                """
                <div class="card">

                    <div class="card-title">
                        🐱 找不到相關資料
                    </div>

                    <div class="card-text">
                        可以試試商品名稱、商品編號、
                        問題關鍵字或其他描述方式。
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

        else:

            st.caption(
                f"找到 {len(results)} 筆相關資料"
            )

            for _, row in results.iterrows():

                title = clean_text(
                    row.get("問題摘要", "")
                )

                if not title:
                    title = "未設定問題摘要"

                score = row.get(
                    "_score",
                    0
                )

                meta_parts = []

                for field in [
                    "問題類型",
                    "問題細項",
                    "商品名稱",
                    "商品系列",
                    "商品編號"
                ]:

                    value = clean_text(
                        row.get(field, "")
                    )

                    if value:

                        meta_parts.append(
                            f"{field}：{value}"
                        )

                meta_html = "<br>".join(
                    html.escape(x)
                    for x in meta_parts
                )

                answer = clean_text(
                    row.get(
                        "標準回覆",
                        ""
                    )
                )

                judgment = clean_text(
                    row.get(
                        "判斷方式",
                        ""
                    )
                )

                source = clean_text(
                    row.get(
                        "判斷使用的工具",
                        ""
                    )
                )

                keywords_text = clean_text(
                    row.get(
                        "關鍵字",
                        ""
                    )
                )

                keyword_hits = []

                search_targets = [
                    "問題摘要",
                    "關鍵字",
                    "商品名稱",
                    "商品系列",
                    "商品編號"
                ]

                normalized_keyword = normalize_search_text(
                    keyword
                )

                for field in search_targets:

                    value = clean_text(
                        row.get(
                            field,
                            ""
                        )
                    )

                    if (
                        value
                        and normalized_keyword
                        and normalized_keyword
                        in normalize_search_text(value)
                    ):

                        keyword_hits.append(
                            value
                        )

                hits_html = ""

                for hit in keyword_hits[:5]:

                    hits_html += (
                        '<span class="keyword-hit">'
                        + html.escape(hit)
                        + "</span>"
                    )

                extra_html = ""

                if judgment:

                    extra_html += (
                        "<div>"
                        "<b>【判斷方式】</b>"
                        "<br>"
                        + safe_html_text(judgment)
                        + "</div>"
                    )

                if source:

                    extra_html += (
                        "<div style='margin-top:8px;'>"
                        "<b>【判斷使用的工具】</b>"
                        "<br>"
                        + safe_html_text(source)
                        + "</div>"
                    )

                if keywords_text:

                    extra_html += (
                        "<div style='margin-top:8px;'>"
                        "<b>【關鍵字】</b>"
                        "<br>"
                        + safe_html_text(keywords_text)
                        + "</div>"
                    )

                st.markdown(
                    f"""
                    <div class="result-card">

                        <div style="
                            display:flex;
                            justify-content:space-between;
                            align-items:flex-start;
                            gap:15px;
                        ">

                            <div class="result-title">
                                {html.escape(title)}
                            </div>

                            <div style="
                                font-size:12px;
                                color:#9A8D7D;
                                white-space:nowrap;
                            ">
                                相關度 {score:.0f}%
                            </div>

                        </div>

                        <div class="result-meta">
                            {meta_html}
                        </div>

                        <div class="result-answer">
                            {safe_html_text(answer)}
                        </div>

                        <div class="result-extra">
                            {extra_html}
                        </div>

                        {
                            (
                                '<div style="margin-top:10px;">'
                                + hits_html
                                + "</div>"
                            )
                            if hits_html
                            else ""
                        }

                    </div>
                    """,
                    unsafe_allow_html=True
                )


# =========================================================
# 門市地圖
# =========================================================

elif page == "🗺️ 門市地圖":

    st.markdown(
        '<div class="section-title">🗺️ 宜得利小地圖</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-subtitle">
            查詢門市位置、交通、停車及相關資訊。
        </div>
        """,
        unsafe_allow_html=True
    )

    components.iframe(
        POWER_BI_URL,
        height=760,
        scrolling=True
    )


# =========================================================
# 公告＆常用工具
# =========================================================

elif page == "📢 公告＆常用工具":

    st.markdown(
        '<div class="section-title">📢 公告＆常用工具</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-subtitle">
            查看最新內部公告及常用作業工具。
        </div>
        """,
        unsafe_allow_html=True
    )

    tab1, tab2 = st.tabs(
        [
            "📢 內部公告",
            "🧰 常用工具"
        ]
    )

    with tab1:

        announcements = load_announcements()

        if announcements.empty:

            st.info("目前沒有公告資料。")

        else:

            for _, row in announcements.iterrows():

                title = clean_text(
                    row.get("標題", "")
                )

                content = clean_announcement_content(
                    row.get("內容", "")
                )

                st.markdown(
                    f"""
                    <div class="tool-card">

                        <div class="tool-card-title">
                            📢 {html.escape(title)}
                        </div>

                        <div class="tool-card-text">
                            {safe_html_text(content)}
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True
                )

    with tab2:

        tools = load_tools()

        if tools.empty:

            st.info("目前沒有常用工具資料。")

        else:

            for _, row in tools.iterrows():

                title = clean_text(
                    row.get(
                        "工具名稱",
                        row.get("名稱", "")
                    )
                )

                description = clean_text(
                    row.get(
                        "說明",
                        row.get("描述", "")
                    )
                )

                url = clean_text(
                    row.get(
                        "連結",
                        row.get("網址", "")
                    )
                )

                st.markdown(
                    f"""
                    <div class="tool-card">

                        <div class="tool-card-title">
                            🧰 {html.escape(title)}
                        </div>

                        <div class="tool-card-text">
                            {safe_html_text(description)}
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True
                )

                if url:

                    st.link_button(
                        "開啟工具",
                        url
                    )


# =========================================================
# 資料維護
# =========================================================

elif page == "🛠️ 資料維護":

    st.markdown(
        '<div class="section-title">🛠️ 資料維護</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-subtitle">
            維護知識庫、內部公告及常用工具資料。
        </div>
        """,
        unsafe_allow_html=True
    )

    tab1, tab2, tab3 = st.tabs(
        [
            "📚 知識庫",
            "📢 公告",
            "🧰 常用工具"
        ]
    )

    # -----------------------------------------------------
    # 知識庫
    # -----------------------------------------------------

    with tab1:

        df = load_data()

        if df.empty:

            st.warning(
                "目前無法讀取知識庫資料。"
            )

        else:

            edited_df = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                height=600,
                key="knowledge_editor"
            )

            if st.button(
                "💾 儲存知識庫",
                key="save_knowledge"
            ):

                success, message = save_knowledge_data(
                    edited_df
                )

                if success:

                    st.success(
                        "知識庫已成功儲存！"
                    )

                    st.rerun()

                else:

                    st.error(
                        f"儲存失敗：{message}"
                    )

    # -----------------------------------------------------
    # 公告
    # -----------------------------------------------------

    with tab2:

        announcements = load_announcements()

        if announcements.empty:

            st.info(
                "目前沒有公告資料。"
            )

        else:

            edited_announcements = st.data_editor(
                announcements,
                use_container_width=True,
                num_rows="dynamic",
                height=600,
                key="announcement_editor"
            )

            announcement_sheet = find_sheet(
                ["內部公告", "公告"]
            )

            if st.button(
                "💾 儲存公告",
                key="save_announcements"
            ):

                success, message = save_sheet(
                    announcement_sheet or "內部公告",
                    edited_announcements
                )

                if success:

                    st.success(
                        "公告已成功儲存！"
                    )

                    st.rerun()

                else:

                    st.error(
                        f"儲存失敗：{message}"
                    )

    # -----------------------------------------------------
    # 常用工具
    # -----------------------------------------------------

    with tab3:

        tools = load_tools()

        if tools.empty:

            st.info(
                "目前沒有常用工具資料。"
            )

        else:

            edited_tools = st.data_editor(
                tools,
                use_container_width=True,
                num_rows="dynamic",
                height=600,
                key="tools_editor"
            )

            tools_sheet = find_sheet(
                ["常用工具", "工具"]
            )

            if st.button(
                "💾 儲存常用工具",
                key="save_tools"
            ):

                success, message = save_sheet(
                    tools_sheet or "常用工具",
                    edited_tools
                )

                if success:

                    st.success(
                        "常用工具已成功儲存！"
                    )

                    st.rerun()

                else:

                    st.error(
                        f"儲存失敗：{message}"
                    )
