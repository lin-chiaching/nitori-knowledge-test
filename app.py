import streamlit as st
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from datetime import datetime
from openpyxl import load_workbook
import html


# =========================================================
# 基本設定
# =========================================================

FILE_PATH = "DATA/知識庫實驗品.xlsx"

st.set_page_config(
    page_title="客服知識庫",
    page_icon="🏠",
    layout="wide"
)


# =========================================================
# UI 樣式
# =========================================================

st.markdown(
    """
    <style>

    /* ==============================
       全站
       ============================== */

    .stApp {
        background-color: #F7F3ED;
    }

    .main .block-container {
        max-width: 1250px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    h1, h2, h3 {
        color: #4E4036;
    }

    p, label {
        color: #5C5149;
    }


    /* ==============================
       Sidebar
       ============================== */

    section[data-testid="stSidebar"] {
        background-color: #EFE7DC;
        border-right: 1px solid #D8C8B5;
    }

    section[data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }

    .sidebar-title {
        font-size: 23px;
        font-weight: 700;
        color: #4E4036;
        margin-bottom: 4px;
    }

    .sidebar-subtitle {
        font-size: 12px;
        color: #88786A;
        margin-bottom: 20px;
        line-height: 1.5;
    }

    .sidebar-count {
        background: #FFFDF9;
        border: 1px solid #DCCDBC;
        border-radius: 14px;
        padding: 14px 16px;
        margin-top: 18px;
        margin-bottom: 18px;
    }

    .sidebar-count-title {
        font-size: 12px;
        color: #958578;
        margin-bottom: 4px;
    }

    .sidebar-count-number {
        font-size: 25px;
        font-weight: 700;
        color: #C56F3A;
    }

    .sidebar-note {
        background: #F8E8D8;
        border-left: 4px solid #D9824B;
        border-radius: 10px;
        padding: 11px 13px;
        font-size: 12px;
        color: #735B49;
        line-height: 1.6;
        margin-top: 18px;
    }


    /* ==============================
       頁首
       ============================== */

    .page-header {
        background: #FFFDF9;
        border: 1px solid #E2D6C8;
        border-radius: 22px;
        padding: 28px 30px;
        margin-bottom: 22px;
        box-shadow: 0 4px 15px rgba(110, 88, 67, 0.05);
    }

    .page-title {
        font-size: 30px;
        font-weight: 700;
        color: #4E4036;
        margin-bottom: 6px;
    }

    .page-subtitle {
        font-size: 14px;
        color: #8B7C70;
        margin-bottom: 18px;
    }

    .page-decoration {
        text-align: right;
        font-size: 30px;
        letter-spacing: 8px;
        color: #CDB9A4;
    }


    /* ==============================
       搜尋區
       ============================== */

    .search-title {
        font-size: 17px;
        font-weight: 700;
        color: #55463C;
        margin-top: 5px;
        margin-bottom: 5px;
    }

    .search-subtitle {
        font-size: 13px;
        color: #958579;
        margin-bottom: 12px;
    }

    .hint-card {
        background: #FBF0E5;
        border: 1px solid #E9CFB9;
        border-radius: 15px;
        padding: 15px 17px;
        min-height: 112px;
    }

    .hint-title {
        color: #B76531;
        font-size: 14px;
        font-weight: 700;
        margin-bottom: 7px;
    }

    .hint-text {
        color: #79685B;
        font-size: 12px;
        line-height: 1.7;
    }

    .example-chip {
        display: inline-block;
        background: #FFFDF9;
        border: 1px solid #E3CDB8;
        color: #80644E;
        padding: 4px 8px;
        border-radius: 20px;
        margin: 3px 3px 0 0;
        font-size: 11px;
    }


    /* ==============================
       輸入框
       ============================== */

    div[data-baseweb="input"] {
        border-radius: 12px;
        border-color: #DCCDBD;
        background-color: #FFFDF9;
    }

    div[data-baseweb="input"]:focus-within {
        border-color: #D9824B;
        box-shadow: 0 0 0 1px #D9824B;
    }


    /* ==============================
       篩選條件
       ============================== */

    div[data-testid="stExpander"] {
        background-color: #FFFDF9;
        border: 1px solid #E2D6C8;
        border-radius: 14px;
    }


    /* ==============================
       搜尋結果
       ============================== */

    .result-header {
        margin-top: 18px;
        margin-bottom: 12px;
    }

    .result-title {
        font-size: 20px;
        font-weight: 700;
        color: #55463C;
    }

    .result-count {
        display: inline-block;
        background: #F4E3D4;
        color: #A85F30;
        border-radius: 20px;
        padding: 5px 12px;
        font-size: 12px;
        font-weight: 600;
        margin-left: 8px;
    }


    /* ==============================
       搜尋詞
       ============================== */

    .search-word-box {
        background: #F7EBDD;
        border-left: 4px solid #D9824B;
        border-radius: 10px;
        padding: 10px 14px;
        margin: 8px 0 16px 0;
        color: #765B48;
        font-size: 12px;
    }


    /* ==============================
       知識卡片
       ============================== */

    .knowledge-card {
        background: #FFFDF9;
        border: 1px solid #E1D5C7;
        border-radius: 18px;
        padding: 20px 22px 12px 22px;
        margin-top: 14px;
        margin-bottom: 0;
        box-shadow: 0 4px 14px rgba(110, 88, 67, 0.05);
    }

    .knowledge-title {
        color: #4F4138;
        font-size: 17px;
        font-weight: 700;
        margin-bottom: 10px;
        line-height: 1.5;
    }

    .tag {
        display: inline-block;
        padding: 4px 9px;
        margin: 0 5px 7px 0;
        border-radius: 20px;
        background: #F4E8DC;
        border: 1px solid #E4D0BC;
        color: #805D45;
        font-size: 11px;
    }

    .answer-title {
        color: #B96330;
        font-size: 13px;
        font-weight: 700;
        margin-top: 14px;
        margin-bottom: 7px;
    }

    .answer-box {
        background: #FBF5EE;
        border-left: 4px solid #D9824B;
        border-radius: 10px;
        padding: 13px 15px;
        color: #53483F;
        font-size: 14px;
        line-height: 1.8;
        white-space: pre-wrap;
    }


    /* ==============================
       詳細資料
       ============================== */

    .detail-box {
        background: #FAF7F2;
        border-radius: 10px;
        padding: 10px 14px;
        color: #6C5E54;
        font-size: 12px;
        line-height: 1.8;
        margin-top: 8px;
    }


    /* ==============================
       空白畫面
       ============================== */

    .empty-house {
        background: #FFFDF9;
        border: 1px solid #E2D6C8;
        border-radius: 22px;
        padding: 45px 30px;
        text-align: center;
        margin-top: 20px;
        box-shadow: 0 4px 15px rgba(110, 88, 67, 0.04);
    }

    .empty-icon {
        font-size: 52px;
        margin-bottom: 10px;
    }

    .empty-title {
        color: #59493E;
        font-size: 21px;
        font-weight: 700;
        margin-bottom: 8px;
    }

    .empty-text {
        color: #918176;
        font-size: 13px;
        line-height: 1.7;
        margin-bottom: 18px;
    }

    .empty-example {
        display: inline-block;
        background: #F8EEE5;
        border: 1px solid #E6D2BF;
        border-radius: 20px;
        padding: 7px 12px;
        margin: 4px;
        color: #80634E;
        font-size: 12px;
    }


    /* ==============================
       客服小角落
       ============================== */

    .service-corner {
        margin-top: 38px;
        background: #EFE5D8;
        border: 1px solid #DCC9B5;
        border-radius: 20px;
        padding: 20px 23px;
    }

    .service-corner-title {
        color: #5B493C;
        font-weight: 700;
        font-size: 15px;
        margin-bottom: 6px;
    }

    .service-corner-text {
        color: #806F61;
        font-size: 12px;
        line-height: 1.8;
    }


    /* ==============================
       維護頁
       ============================== */

    .maintenance-header {
        background: #FFFDF9;
        border: 1px solid #E2D6C8;
        border-radius: 20px;
        padding: 22px 25px;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(110, 88, 67, 0.05);
    }

    .maintenance-title {
        color: #514238;
        font-size: 25px;
        font-weight: 700;
    }

    .maintenance-subtitle {
        color: #8B7A6C;
        font-size: 13px;
        margin-top: 4px;
    }

    div[data-testid="stForm"] {
        background: #FFFDF9;
        border: 1px solid #E2D6C8;
        border-radius: 16px;
        padding: 18px;
    }


    /* ==============================
       Dataframe
       ============================== */

    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
        border: 1px solid #E0D4C6;
    }


    /* ==============================
       按鈕
       ============================== */

    button[kind="primary"] {
        background-color: #D9824B;
        border-color: #D9824B;
    }

    button[kind="primary"]:hover {
        background-color: #C66F38;
        border-color: #C66F38;
    }


    /* ==============================
       手機版
       ============================== */

    @media (max-width: 800px) {

        .main .block-container {
            padding: 1rem;
        }

        .page-title {
            font-size: 25px;
        }

        .page-decoration {
            font-size: 24px;
        }

        .knowledge-card {
            padding: 16px;
        }

    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# Excel 讀取
# =========================================================

def load_data():

    return pd.read_excel(
        FILE_PATH,
        sheet_name=0
    ).fillna("")


def load_synonyms():

    return pd.read_excel(
        FILE_PATH,
        sheet_name="搜尋詞對照"
    ).fillna("")


try:

    df = load_data()
    synonym_df = load_synonyms()

except Exception as e:

    st.error(f"讀取 Excel 時發生錯誤：{e}")
    st.stop()


# =========================================================
# Excel 安全寫入
# =========================================================

def save_data(new_df):

    wb = load_workbook(FILE_PATH)

    ws = wb.worksheets[0]

    ws.delete_rows(
        1,
        ws.max_row
    )

    for col_index, column_name in enumerate(
        new_df.columns,
        start=1
    ):

        ws.cell(
            row=1,
            column=col_index,
            value=column_name
        )

    for row_index, row in enumerate(
        new_df.itertuples(index=False),
        start=2
    ):

        for col_index, value in enumerate(
            row,
            start=1
        ):

            ws.cell(
                row=row_index,
                column=col_index,
                value=value
            )

    wb.save(FILE_PATH)


# =========================================================
# 搜尋詞辨識
# =========================================================

def find_standard_words(keyword):

    keyword = str(keyword).strip()

    matched = []

    for _, row in synonym_df.iterrows():

        standard_word = str(
            row.get(
                "標準搜尋詞",
                ""
            )
        ).strip()

        synonym_word = str(
            row.get(
                "同義詞／常見說法",
                ""
            )
        ).strip()

        if not standard_word:
            continue

        if not synonym_word:
            continue

        if synonym_word in keyword:

            matched.append(
                (
                    standard_word,
                    synonym_word
                )
            )

    result = []

    seen = set()

    for standard_word, synonym_word in matched:

        key = (
            standard_word,
            synonym_word
        )

        if key not in seen:

            seen.add(key)

            result.append(key)

    return result


# =========================================================
# 商品資訊辨識
# =========================================================

def find_product_words(keyword):

    words = []

    keyword = str(keyword).strip()

    if "商品系列" in df.columns:

        for value in df["商品系列"].astype(str).unique():

            value = value.strip()

            if value != "" and value in keyword:

                words.append(value)

    if "商品編號" in df.columns:

        for value in df["商品編號"].astype(str).unique():

            value = value.strip()

            if value != "" and value in keyword:

                words.append(value)

    if "商品名稱" in df.columns:

        for value in df["商品名稱"].astype(str).unique():

            value = value.strip()

            if value != "" and value in keyword:

                words.append(value)

    return list(
        dict.fromkeys(words)
    )


# =========================================================
# Sidebar
# =========================================================

st.sidebar.markdown(
    """
    <div class="sidebar-title">
        🏠 客服知識庫
    </div>

    <div class="sidebar-subtitle">
        NITORI Customer Service Knowledge
    </div>
    """,
    unsafe_allow_html=True
)


page = st.sidebar.radio(
    "功能",
    [
        "🔍 知識搜尋",
        "🛠️ 資料維護"
    ],
    label_visibility="collapsed"
)


st.sidebar.markdown(
    f"""
    <div class="sidebar-count">

        <div class="sidebar-count-title">
            📚 目前知識庫
        </div>

        <div class="sidebar-count-number">
            {len(df):,} 筆
        </div>

    </div>

    <div class="sidebar-note">
        💡 知識庫持續更新中<br>
        將最新確認的資訊整理後加入資料庫，
        讓大家都能快速找到一致的參考內容。
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 🔍 知識搜尋
# =========================================================

if page == "🔍 知識搜尋":

    # =====================================================
    # 頁首
    # =====================================================

    st.markdown(
        """
        <div class="page-header">

            <div class="page-title">
                🔍 客服知識搜尋
            </div>

            <div class="page-subtitle">
                快速找到顧客問題對應的標準資訊與回覆內容
            </div>

            <div class="page-decoration">
                🛋️　🪴　💡
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # =====================================================
    # 搜尋區
    # =====================================================

    st.markdown(
        """
        <div class="search-title">
            💬 顧客問題
        </div>

        <div class="search-subtitle">
            可以直接輸入顧客原本的說法，不需要先判斷分類。
        </div>
        """,
        unsafe_allow_html=True
    )


    search_col1, search_col2 = st.columns(
        [2.7, 1],
        gap="large"
    )


    with search_col1:

        keyword = st.text_input(
            "🔎 搜尋問題",
            placeholder="例如：這張椅子幾公斤？",
            label_visibility="collapsed"
        )


    with search_col2:

        st.markdown(
            """
            <div class="hint-card">

                <div class="hint-title">
                    💡 小提示
                </div>

                <div class="hint-text">
                    可以直接輸入顧客說法，例如：
                    <br>

                    <span class="example-chip">
                        這張椅子多重？
                    </span>

                    <span class="example-chip">
                        床墊可以退嗎？
                    </span>

                    <span class="example-chip">
                        有沒有展示品？
                    </span>
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


    # =====================================================
    # 篩選條件
    # =====================================================

    with st.expander(
        "⚙️ 篩選條件（不確定時可以不用設定）",
        expanded=False
    ):

        filter_col1, filter_col2, filter_col3 = st.columns(3)


        with filter_col1:

            problem_type_list = sorted(
                df["問題類型"]
                .astype(str)
                .replace("", pd.NA)
                .dropna()
                .unique()
                .tolist()
            )

            problem_type = st.selectbox(
                "問題類型",
                ["全部"] + problem_type_list
            )


        with filter_col2:

            problem_detail_list = sorted(
                df["問題細項"]
                .astype(str)
                .replace("", pd.NA)
                .dropna()
                .unique()
                .tolist()
            )

            problem_detail = st.selectbox(
                "問題細項",
                ["全部"] + problem_detail_list
            )


        with filter_col3:

            product_category_list = sorted(
                df["商品分類代碼"]
                .astype(str)
                .replace("", pd.NA)
                .dropna()
                .unique()
                .tolist()
            )

            product_category = st.selectbox(
                "商品分類代碼",
                ["全部"] + product_category_list
            )


        product_series_list = sorted(
            df["商品系列"]
            .astype(str)
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )

        product_series = st.selectbox(
            "商品系列",
            ["全部"] + product_series_list
        )


    # =====================================================
    # 搜尋
    # =====================================================

    if keyword:

        # -------------------------------------------------
        # 辨識搜尋詞
        # -------------------------------------------------

        matched_words = find_standard_words(
            keyword
        )

        standard_words = list(
            dict.fromkeys(
                [
                    x[0]
                    for x in matched_words
                ]
            )
        )


        # -------------------------------------------------
        # 辨識商品
        # -------------------------------------------------

        product_words = find_product_words(
            keyword
        )


        # -------------------------------------------------
        # 套用篩選
        # -------------------------------------------------

        filtered_df = df.copy()


        if problem_type != "全部":

            filtered_df = filtered_df[
                filtered_df["問題類型"].astype(str)
                == problem_type
            ]


        if problem_detail != "全部":

            filtered_df = filtered_df[
                filtered_df["問題細項"].astype(str)
                == problem_detail
            ]


        if product_category != "全部":

            filtered_df = filtered_df[
                filtered_df["商品分類代碼"].astype(str)
                == product_category
            ]


        if product_series != "全部":

            filtered_df = filtered_df[
                filtered_df["商品系列"].astype(str)
                == product_series
            ]


        # =================================================
        # 第一階段：標準搜尋詞
        # =================================================

        exact_result = pd.DataFrame()


        if standard_words:

            exact_result = filtered_df[
                filtered_df["問題摘要"]
                .astype(str)
                .apply(
                    lambda x:
                    any(
                        word in x
                        for word in standard_words
                    )
                )
            ].copy()


        # =================================================
        # 第二階段：關鍵字搜尋
        # =================================================

        keyword_result = pd.DataFrame()


        if standard_words:

            search_columns = [
                "問題摘要",
                "關鍵字",
                "問題細項",
                "商品名稱",
                "商品系列",
                "商品編號"
            ]

            search_columns = [
                col
                for col in search_columns
                if col in filtered_df.columns
            ]


            def keyword_match(row):

                text = " ".join(
                    str(row[col])
                    for col in search_columns
                )

                return any(
                    word in text
                    for word in standard_words
                )


            keyword_result = filtered_df[
                filtered_df.apply(
                    keyword_match,
                    axis=1
                )
            ].copy()


        # =================================================
        # 合併標準搜尋結果
        # =================================================

        if len(exact_result) > 0:

            result = exact_result.copy()

            result["搜尋方式"] = "標準問題詞"


        elif len(keyword_result) > 0:

            result = keyword_result.copy()

            result["搜尋方式"] = (
                "標準搜尋詞＋關鍵字"
            )


        else:

            # =============================================
            # 第三階段：相似度搜尋
            # =============================================

            search_columns = [
                "問題摘要",
                "關鍵字",
                "問題細項",
                "商品名稱",
                "商品系列",
                "商品編號"
            ]

            search_columns = [
                col
                for col in search_columns
                if col in filtered_df.columns
            ]


            if len(filtered_df) == 0:

                result = pd.DataFrame()


            else:

                filtered_df = filtered_df.copy()

                filtered_df["搜尋內容"] = (
                    filtered_df[search_columns]
                    .astype(str)
                    .agg(" ".join, axis=1)
                )

                expanded_keyword = (
                    keyword
                    + " "
                    + " ".join(standard_words)
                )


                try:

                    vectorizer = TfidfVectorizer(
                        analyzer="char",
                        ngram_range=(2, 4)
                    )

                    tfidf_matrix = (
                        vectorizer.fit_transform(
                            filtered_df["搜尋內容"]
                        )
                    )

                    query_vector = (
                        vectorizer.transform(
                            [expanded_keyword]
                        )
                    )

                    similarity = cosine_similarity(
                        query_vector,
                        tfidf_matrix
                    ).flatten()


                    filtered_df[
                        "相關程度"
                    ] = similarity


                    result = filtered_df[
                        filtered_df["相關程度"] > 0.05
                    ].sort_values(
                        "相關程度",
                        ascending=False
                    ).copy()


                    result["搜尋方式"] = (
                        "相似度搜尋"
                    )


                except Exception:

                    result = pd.DataFrame()


        # =================================================
        # 商品優先
        # =================================================

        if product_words and len(result) > 0:

            result["_商品符合"] = result.apply(

                lambda row:
                any(

                    word in str(
                        row.get(
                            "商品名稱",
                            ""
                        )
                    )

                    or word in str(
                        row.get(
                            "商品系列",
                            ""
                        )
                    )

                    or word in str(
                        row.get(
                            "商品編號",
                            ""
                        )
                    )

                    for word in product_words

                ),

                axis=1
            )

            result = result.sort_values(
                "_商品符合",
                ascending=False
            )


        # =================================================
        # 搜尋結果標題
        # =================================================

        st.markdown(
            f"""
            <div class="result-header">

                <span class="result-title">
                    📚 搜尋結果
                </span>

                <span class="result-count">
                    {len(result)} 筆
                </span>

            </div>
            """,
            unsafe_allow_html=True
        )


        # =================================================
        # 搜尋詞辨識
        # =================================================

        if matched_words:

            display_words = []

            for standard_word, synonym_word in matched_words:

                display_words.append(
                    f"{html.escape(str(synonym_word))}"
                    f" → "
                    f"{html.escape(str(standard_word))}"
                )


            st.markdown(
                """
                <div class="search-word-box">
                    <strong>🔎 搜尋詞辨識</strong>
                    &nbsp;&nbsp;
                """
                + "　".join(display_words)
                + """
                </div>
                """,
                unsafe_allow_html=True
            )


        # =================================================
        # 沒有結果
        # =================================================

        if len(result) == 0:

            st.markdown(
                """
                <div class="empty-house">

                    <div class="empty-icon">
                        🏠
                    </div>

                    <div class="empty-title">
                        這裡還沒有找到答案
                    </div>

                    <div class="empty-text">
                        可以換個顧客說法，
                        或先把篩選條件取消看看。
                    </div>

                    <span class="empty-example">
                        💬 這張椅子有多重？
                    </span>

                    <span class="empty-example">
                        💬 這個商品可以退嗎？
                    </span>

                    <span class="empty-example">
                        💬 門市有展示嗎？
                    </span>

                </div>
                """,
                unsafe_allow_html=True
            )


        # =================================================
        # 有結果
        # =================================================

        else:

            for _, row in result.head(20).iterrows():

                title = str(
                    row.get(
                        "問題摘要",
                        "未設定問題摘要"
                    )
                )

                title_html = html.escape(
                    title
                )


                # -----------------------------------------
                # 結果標籤
                # -----------------------------------------

                tags = []

                tag_fields = [
                    "問題類型",
                    "問題細項",
                    "商品系列"
                ]


                for field in tag_fields:

                    value = str(
                        row.get(
                            field,
                            ""
                        )
                    ).strip()


                    if value:

                        tags.append(
                            f'<span class="tag">'
                            f'{html.escape(value)}'
                            f'</span>'
                        )


                tags_html = "".join(
                    tags
                )


                # -----------------------------------------
                # 知識卡片標題
                # -----------------------------------------

                st.markdown(
                    f"""
                    <div class="knowledge-card">

                        <div class="knowledge-title">
                            {title_html}
                        </div>

                        <div>
                            {tags_html}
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True
                )


                # -----------------------------------------
                # 商品資訊
                # -----------------------------------------

                info_col1, info_col2, info_col3 = st.columns(3)


                with info_col1:

                    st.markdown(
                        f"""
                        <div class="detail-box">
                            <strong>商品名稱</strong><br>
                            {html.escape(
                                str(
                                    row.get(
                                        "商品名稱",
                                        ""
                                    )
                                )
                            )}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )


                with info_col2:

                    st.markdown(
                        f"""
                        <div class="detail-box">
                            <strong>商品系列</strong><br>
                            {html.escape(
                                str(
                                    row.get(
                                        "商品系列",
                                        ""
                                    )
                                )
                            )}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )


                with info_col3:

                    st.markdown(
                        f"""
                        <div class="detail-box">
                            <strong>商品編號</strong><br>
                            {html.escape(
                                str(
                                    row.get(
                                        "商品編號",
                                        ""
                                    )
                                )
                            )}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )


                # -----------------------------------------
                # 標準回覆
                # -----------------------------------------

                st.markdown(
                    """
                    <div class="answer-title">
                        📋 標準回覆
                    </div>
                    """,
                    unsafe_allow_html=True
                )


                answer_text = html.escape(
                    str(
                        row.get(
                            "標準回覆",
                            ""
                        )
                    )
                )


                st.markdown(
                    f"""
                    <div class="answer-box">
                        {answer_text}
                    </div>
                    """,
                    unsafe_allow_html=True
                )


                # -----------------------------------------
                # 詳細資訊
                # -----------------------------------------

                with st.expander(
                    "查看判斷依據與其他資訊"
                ):

                    st.write(
                        f"**問題類型：** "
                        f"{row.get('問題類型', '')}"
                    )

                    st.write(
                        f"**問題細項：** "
                        f"{row.get('問題細項', '')}"
                    )

                    st.write(
                        f"**商品分類代碼：** "
                        f"{row.get('商品分類代碼', '')}"
                    )

                    st.write(
                        f"**判斷依據：** "
                        f"{row.get('判斷依據', '')}"
                    )

                    st.write(
                        f"**資料來源：** "
                        f"{row.get('資料來源', '')}"
                    )

                    st.write(
                        f"**關鍵字：** "
                        f"{row.get('關鍵字', '')}"
                    )

                    st.write(
                        f"**搜尋方式：** "
                        f"{row.get('搜尋方式', '')}"
                    )

                    st.write(
                        f"**更新人：** "
                        f"{row.get('更新人', '')}"
                    )

                    st.write(
                        f"**更新日期：** "
                        f"{row.get('更新日期', '')}"
                    )


    # =====================================================
    # 尚未搜尋
    # =====================================================

    else:

        st.markdown(
            """
            <div class="empty-house">

                <div class="empty-icon">
                    🛋️
                </div>

                <div class="empty-title">
                    今天想查什麼？
                </div>

                <div class="empty-text">
                    把顧客的問題直接貼進來，
                    <br>
                    讓知識庫幫你找到可以參考的內容。
                </div>

                <span class="empty-example">
                    🔎 商品重量
                </span>

                <span class="empty-example">
                    🔎 門市展示
                </span>

                <span class="empty-example">
                    🔎 退換貨
                </span>

                <span class="empty-example">
                    🔎 配送問題
                </span>

            </div>
            """,
            unsafe_allow_html=True
        )


    # =====================================================
    # 客服小角落
    # =====================================================

    st.markdown(
        f"""
        <div class="service-corner">

            <div class="service-corner-title">
                🏠 客服小角落
            </div>

            <div class="service-corner-text">
                這裡整理的是客服日常可以查詢的知識內容。
                <br>
                如果發現資料需要修正或補充，可以到「資料維護」更新。
                <br>
                目前知識庫共有
                <strong>{len(df):,}</strong>
                筆資料。
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# 🛠️ 資料維護
# =========================================================

elif page == "🛠️ 資料維護":

    st.markdown(
        """
        <div class="maintenance-header">

            <div class="maintenance-title">
                🛠️ 知識庫資料維護
            </div>

            <div class="maintenance-subtitle">
                新增、查看及修改客服知識庫內容
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


    # =====================================================
    # 重新讀取最新 Excel
    # =====================================================

    df = load_data()


    # =====================================================
    # 分頁
    # =====================================================

    tab1, tab2, tab3 = st.tabs(
        [
            "📋 現有資料",
            "➕ 新增資料",
            "✏️ 修改資料"
        ]
    )


    # =====================================================
    # ① 現有資料
    # =====================================================

    with tab1:

        st.subheader(
            "📋 現有知識庫資料"
        )

        st.write(
            f"目前共有 **{len(df)}** 筆資料。"
        )


        search = st.text_input(
            "搜尋資料",
            placeholder="可輸入商品名稱、商品編號、問題摘要..."
        )


        display_df = df.copy()


        if search:

            search_columns = [
                "商品名稱",
                "商品系列",
                "商品編號",
                "問題摘要",
                "標準回覆",
                "關鍵字"
            ]


            search_columns = [
                col
                for col in search_columns
                if col in display_df.columns
            ]


            mask = display_df[
                search_columns
            ].astype(str).apply(
                lambda row:
                row.str.contains(
                    search,
                    case=False,
                    na=False
                ).any(),
                axis=1
            )


            display_df = display_df[
                mask
            ]


        st.write(
            f"符合條件：**{len(display_df)}** 筆"
        )


        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )


    # =====================================================
    # ② 新增資料
    # =====================================================

    with tab2:

        st.subheader(
            "➕ 新增知識庫資料"
        )

        st.info(
            "填寫完成後按下「新增資料」，資料會直接寫入 Excel。"
        )


        problem_type_list = sorted(
            df["問題類型"]
            .astype(str)
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )


        problem_detail_list = sorted(
            df["問題細項"]
            .astype(str)
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )


        category_list = sorted(
            df["商品分類代碼"]
            .astype(str)
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )


        with st.form("add_form"):

            col1, col2 = st.columns(2)


            with col1:

                add_problem_type = st.selectbox(
                    "問題類型",
                    problem_type_list
                )

                add_problem_detail = st.selectbox(
                    "問題細項",
                    problem_detail_list
                )

                add_category = st.selectbox(
                    "商品分類代碼",
                    category_list
                )

                add_product_name = st.text_input(
                    "商品名稱"
                )

                add_product_series = st.text_input(
                    "商品系列"
                )

                add_product_number = st.text_input(
                    "商品編號"
                )


            with col2:

                add_summary = st.text_input(
                    "問題摘要"
                )

                add_source = st.text_input(
                    "資料來源"
                )

                add_keyword = st.text_input(
                    "關鍵字"
                )

                add_creator = st.text_input(
                    "建立人"
                )


            add_answer = st.text_area(
                "標準回覆",
                height=180
            )


            add_basis = st.text_area(
                "判斷依據",
                height=150
            )


            submitted = st.form_submit_button(
                "➕ 新增資料"
            )


        # =================================================
        # 新增資料
        # =================================================

        if submitted:

            if not add_summary:

                st.warning(
                    "請至少填寫「問題摘要」。"
                )

            elif not add_answer:

                st.warning(
                    "請填寫「標準回覆」。"
                )

            else:

                numbers = pd.to_numeric(
                    df["序號"],
                    errors="coerce"
                )


                if numbers.notna().any():

                    new_id = int(
                        numbers.max()
                    ) + 1

                else:

                    new_id = 1


                today = datetime.now().strftime(
                    "%Y-%m-%d"
                )


                new_row = {

                    "序號": new_id,

                    "問題類型":
                        add_problem_type,

                    "問題細項":
                        add_problem_detail,

                    "商品分類代碼":
                        add_category,

                    "商品名稱":
                        add_product_name,

                    "商品系列":
                        add_product_series,

                    "商品編號":
                        add_product_number,

                    "問題摘要":
                        add_summary,

                    "標準回覆":
                        add_answer,

                    "判斷依據":
                        add_basis,

                    "資料來源":
                        add_source,

                    "關鍵字":
                        add_keyword,

                    "建立人":
                        add_creator,

                    "建立日期":
                        today,

                    "更新人":
                        add_creator,

                    "更新日期":
                        today
                }


                df = pd.concat(
                    [
                        df,
                        pd.DataFrame(
                            [new_row]
                        )
                    ],
                    ignore_index=True
                )


                try:

                    save_data(df)

                    st.success(
                        f"新增成功！序號：{new_id}"
                    )

                    st.rerun()


                except Exception as e:

                    st.error(
                        f"儲存失敗：{e}"
                    )


    # =====================================================
    # ③ 修改資料
    # =====================================================

    with tab3:

        st.subheader(
            "✏️ 修改知識庫資料"
        )

        st.write(
            "先選擇「序號」，找到資料後再進行修改。"
        )


        id_list = sorted(
            pd.to_numeric(
                df["序號"],
                errors="coerce"
            )
            .dropna()
            .astype(int)
            .tolist()
        )


        if id_list:

            selected_id = st.selectbox(
                "選擇要修改的序號",
                id_list
            )


            selected_rows = df[
                pd.to_numeric(
                    df["序號"],
                    errors="coerce"
                ) == selected_id
            ]


            if len(selected_rows) > 0:

                row = selected_rows.iloc[0]


                with st.form("edit_form"):

                    col1, col2 = st.columns(2)


                    with col1:

                        edit_problem_type = st.text_input(
                            "問題類型",
                            value=str(
                                row.get(
                                    "問題類型",
                                    ""
                                )
                            )
                        )


                        edit_problem_detail = st.text_input(
                            "問題細項",
                            value=str(
                                row.get(
                                    "問題細項",
                                    ""
                                )
                            )
                        )


                        edit_category = st.text_input(
                            "商品分類代碼",
                            value=str(
                                row.get(
                                    "商品分類代碼",
                                    ""
                                )
                            )
                        )


                        edit_product_name = st.text_input(
                            "商品名稱",
                            value=str(
                                row.get(
                                    "商品名稱",
                                    ""
                                )
                            )
                        )


                        edit_product_series = st.text_input(
                            "商品系列",
                            value=str(
                                row.get(
                                    "商品系列",
                                    ""
                                )
                            )
                        )


                        edit_product_number = st.text_input(
                            "商品編號",
                            value=str(
                                row.get(
                                    "商品編號",
                                    ""
                                )
                            )
                        )


                    with col2:

                        edit_summary = st.text_input(
                            "問題摘要",
                            value=str(
                                row.get(
                                    "問題摘要",
                                    ""
                                )
                            )
                        )


                        edit_source = st.text_input(
                            "資料來源",
                            value=str(
                                row.get(
                                    "資料來源",
                                    ""
                                )
                            )
                        )


                        edit_keyword = st.text_input(
                            "關鍵字",
                            value=str(
                                row.get(
                                    "關鍵字",
                                    ""
                                )
                            )
                        )


                        edit_creator = st.text_input(
                            "建立人",
                            value=str(
                                row.get(
                                    "建立人",
                                    ""
                                )
                            ),
                            disabled=True
                        )


                    edit_answer = st.text_area(
                        "標準回覆",
                        value=str(
                            row.get(
                                "標準回覆",
                                ""
                            )
                        ),
                        height=180
                    )


                    edit_basis = st.text_area(
                        "判斷依據",
                        value=str(
                            row.get(
                                "判斷依據",
                                ""
                            )
                        ),
                        height=150
                    )


                    edit_updater = st.text_input(
                        "更新人"
                    )


                    update_button = st.form_submit_button(
                        "💾 儲存修改"
                    )


                # =================================================
                # 儲存修改
                # =================================================

                if update_button:

                    if not edit_updater:

                        st.warning(
                            "請填寫「更新人」。"
                        )

                    else:

                        today = datetime.now().strftime(
                            "%Y-%m-%d"
                        )


                        index = selected_rows.index[0]


                        df.loc[
                            index,
                            "問題類型"
                        ] = edit_problem_type


                        df.loc[
                            index,
                            "問題細項"
                        ] = edit_problem_detail


                        df.loc[
                            index,
                            "商品分類代碼"
                        ] = edit_category


                        df.loc[
                            index,
                            "商品名稱"
                        ] = edit_product_name


                        df.loc[
                            index,
                            "商品系列"
                        ] = edit_product_series


                        df.loc[
                            index,
                            "商品編號"
                        ] = edit_product_number


                        df.loc[
                            index,
                            "問題摘要"
                        ] = edit_summary


                        df.loc[
                            index,
                            "標準回覆"
                        ] = edit_answer


                        df.loc[
                            index,
                            "判斷依據"
                        ] = edit_basis


                        df.loc[
                            index,
                            "資料來源"
                        ] = edit_source


                        df.loc[
                            index,
                            "關鍵字"
                        ] = edit_keyword


                        df.loc[
                            index,
                            "更新人"
                        ] = edit_updater


                        df.loc[
                            index,
                            "更新日期"
                        ] = today


                        try:

                            save_data(df)

                            st.success(
                                f"序號 {selected_id} 修改成功！"
                            )

                            st.rerun()


                        except Exception as e:

                            st.error(
                                f"儲存失敗：{e}"
                            )


        else:

            st.info(
                "目前沒有可修改的資料。"
            )
