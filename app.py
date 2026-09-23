import streamlit as st
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from datetime import datetime
from openpyxl import load_workbook


# =========================================================
# 基本設定
# =========================================================

FILE_PATH = "DATA/知識庫實驗品.xlsx"

st.set_page_config(
    page_title="客服知識庫",
    page_icon="🔍",
    layout="wide"
)


# =========================================================
# UI 外觀
# =========================================================

st.markdown(
    """
    <style>

    /* =========================
       整體頁面
       ========================= */

    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    /* =========================
       Sidebar
       ========================= */

    [data-testid="stSidebar"] {
        border-right: 1px solid #e5e7eb;
    }

    [data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }

    .sidebar-title {
        font-size: 1.25rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }

    .sidebar-subtitle {
        font-size: 0.85rem;
        color: #6b7280;
        margin-bottom: 1.2rem;
    }

    .sidebar-count {
        background: #f5f6f8;
        border-radius: 10px;
        padding: 12px 14px;
        margin-top: 15px;
        font-size: 0.9rem;
    }

    /* =========================
       主標題
       ========================= */

    .page-title {
        font-size: 2rem;
        font-weight: 750;
        margin-bottom: 0.25rem;
    }

    .page-subtitle {
        color: #6b7280;
        font-size: 0.95rem;
        margin-bottom: 1.5rem;
    }

    /* =========================
       搜尋區塊
       ========================= */

    .search-card {
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 22px 24px 18px 24px;
        background: #ffffff;
        margin-bottom: 20px;
    }

    .search-card-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 4px;
    }

    .search-card-description {
        font-size: 0.85rem;
        color: #6b7280;
        margin-bottom: 15px;
    }

    /* =========================
       結果區
       ========================= */

    .result-header {
        font-size: 1.25rem;
        font-weight: 700;
        margin-top: 8px;
        margin-bottom: 4px;
    }

    .result-count {
        color: #6b7280;
        font-size: 0.85rem;
        margin-bottom: 15px;
    }

    /* =========================
       知識卡片
       ========================= */

    .knowledge-card {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 14px;
        background: #ffffff;
    }

    .knowledge-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 12px;
        line-height: 1.5;
    }

    .knowledge-meta {
        font-size: 0.83rem;
        color: #6b7280;
        line-height: 1.8;
    }

    .answer-box {
        background: #f7f8fa;
        border-radius: 10px;
        padding: 15px 17px;
        margin-top: 15px;
        margin-bottom: 10px;
    }

    .answer-title {
        font-weight: 700;
        font-size: 0.9rem;
        margin-bottom: 8px;
    }

    .answer-text {
        font-size: 0.92rem;
        line-height: 1.7;
        white-space: pre-wrap;
    }

    /* =========================
       空白狀態
       ========================= */

    .empty-box {
        border: 1px dashed #d1d5db;
        border-radius: 12px;
        padding: 35px;
        text-align: center;
        color: #6b7280;
        margin-top: 15px;
    }

    .empty-icon {
        font-size: 2rem;
        margin-bottom: 8px;
    }

    .empty-title {
        font-weight: 700;
        color: #374151;
        margin-bottom: 5px;
    }

    /* =========================
       維護頁
       ========================= */

    .maintenance-header {
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 20px;
        background: #ffffff;
    }

    .maintenance-title {
        font-size: 1.15rem;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .maintenance-description {
        color: #6b7280;
        font-size: 0.88rem;
    }

    /* =========================
       表單區
       ========================= */

    .form-section-title {
        font-size: 1rem;
        font-weight: 700;
        margin-top: 10px;
        margin-bottom: 10px;
    }

    /* =========================
       隱藏部分 Streamlit 元素
       ========================= */

    div[data-testid="stDecoration"] {
        display: none;
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
    <div class="sidebar-title">🔍 客服知識庫</div>
    <div class="sidebar-subtitle">
        Customer Service Knowledge Base
    </div>
    """,
    unsafe_allow_html=True
)

st.sidebar.divider()

page = st.sidebar.radio(
    "功能",
    [
        "🔍 知識搜尋",
        "🛠️ 資料維護"
    ],
    label_visibility="collapsed"
)

st.sidebar.divider()

st.sidebar.markdown(
    f"""
    <div class="sidebar-count">
        📚 目前資料<br>
        <strong>{len(df):,}</strong> 筆
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 🔍 知識搜尋
# =========================================================

if page == "🔍 知識搜尋":

    st.markdown(
        """
        <div class="page-title">客服知識搜尋</div>
        <div class="page-subtitle">
            輸入顧客問題，快速找到可參考的標準資訊。
        </div>
        """,
        unsafe_allow_html=True
    )

    # =====================================================
    # 搜尋區塊
    # =====================================================

    st.markdown(
        """
        <div class="search-card">
            <div class="search-card-title">🔎 顧客問題</div>
            <div class="search-card-description">
                可直接輸入顧客原話，例如「這張椅子幾公斤？」、
                「可以退貨嗎？」或商品名稱。
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    keyword = st.text_input(
        "顧客問題",
        placeholder="例如：這張椅子幾公斤？",
        label_visibility="collapsed"
    )

    # =====================================================
    # 篩選條件
    # =====================================================

    with st.expander(
        "⚙️ 篩選條件",
        expanded=False
    ):

        col1, col2 = st.columns(2)

        with col1:

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

        with col2:

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

        col3, col4 = st.columns(2)

        with col3:

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

        with col4:

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
    # 開始搜尋
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
        # 顯示結果
        # =================================================

        st.divider()

        st.markdown(
            f"""
            <div class="result-header">
                搜尋結果
            </div>
            <div class="result-count">
                共找到 {len(result)} 筆相關資料
            </div>
            """,
            unsafe_allow_html=True
        )

        # 搜尋詞辨識結果

        if matched_words:

            display_words = []

            for standard_word, synonym_word in matched_words:

                display_words.append(
                    f"{synonym_word} → {standard_word}"
                )

            st.caption(
                "🔎 搜尋詞辨識："
                + "、".join(display_words)
            )

        # 沒有結果

        if len(result) == 0:

            st.markdown(
                """
                <div class="empty-box">
                    <div class="empty-icon">🔍</div>
                    <div class="empty-title">
                        目前沒有找到相關資料
                    </div>
                    <div>
                        可以嘗試更換顧客說法，
                        或取消部分篩選條件後再次搜尋。
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        # 有結果

        else:

            for _, row in result.head(20).iterrows():

                title = row.get(
                    "問題摘要",
                    "未設定問題摘要"
                )

                # =================================================
                # 知識卡片
                # =================================================

                st.markdown(
                    f"""
                    <div class="knowledge-card">
                        <div class="knowledge-title">
                            {title}
                        </div>

                        <div class="knowledge-meta">
                            <b>問題類型：</b>
                            {row.get('問題類型', '')}
                            &nbsp;&nbsp;|&nbsp;&nbsp;

                            <b>問題細項：</b>
                            {row.get('問題細項', '')}
                        </div>

                        <div class="knowledge-meta">
                            <b>商品名稱：</b>
                            {row.get('商品名稱', '')}
                            &nbsp;&nbsp;|&nbsp;&nbsp;

                            <b>商品系列：</b>
                            {row.get('商品系列', '')}
                        </div>

                        <div class="knowledge-meta">
                            <b>商品編號：</b>
                            {row.get('商品編號', '')}
                            &nbsp;&nbsp;|&nbsp;&nbsp;

                            <b>商品分類：</b>
                            {row.get('商品分類代碼', '')}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                # =================================================
                # 標準回覆
                # =================================================

                st.markdown(
                    f"""
                    <div class="answer-box">
                        <div class="answer-title">
                            📋 標準回覆
                        </div>

                        <div class="answer-text">
                            {str(row.get('標準回覆', '')).replace(chr(10), '<br>')}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                # =================================================
                # 其他資訊
                # =================================================

                with st.expander(
                    "查看判斷依據與其他資訊"
                ):

                    st.markdown(
                        "### 🧭 判斷依據"
                    )

                    st.write(
                        row.get(
                            "判斷依據",
                            ""
                        )
                    )

                    col1, col2 = st.columns(2)

                    with col1:

                        st.markdown(
                            "### 📚 資料來源"
                        )

                        st.write(
                            row.get(
                                "資料來源",
                                ""
                            )
                        )

                    with col2:

                        st.markdown(
                            "### 🏷️ 關鍵字"
                        )

                        st.write(
                            row.get(
                                "關鍵字",
                                ""
                            )
                        )

                    st.divider()

                    col3, col4 = st.columns(2)

                    with col3:

                        st.caption(
                            f"更新人：{row.get('更新人', '')}"
                        )

                    with col4:

                        st.caption(
                            f"更新日期：{row.get('更新日期', '')}"
                        )


    else:

        st.markdown(
            """
            <div class="empty-box">
                <div class="empty-icon">🔎</div>
                <div class="empty-title">
                    請輸入顧客問題
                </div>
                <div>
                    例如：這張椅子幾公斤？
                    可以退貨嗎？床墊可以試躺嗎？
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
        <div class="page-title">知識庫維護</div>
        <div class="page-subtitle">
            新增、查看及修改客服知識庫資料。
        </div>
        """,
        unsafe_allow_html=True
    )

    df = load_data()

    st.markdown(
        f"""
        <div class="maintenance-header">
            <div class="maintenance-title">
                🗂️ 知識庫資料管理
            </div>
            <div class="maintenance-description">
                目前共有 <b>{len(df):,}</b> 筆資料，
                可在下方進行查看、新增或修改。
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

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

        st.caption(
            f"符合條件：{len(display_df):,} 筆"
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

            st.markdown(
                '<div class="form-section-title">📌 基本資訊</div>',
                unsafe_allow_html=True
            )

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

            st.markdown(
                '<div class="form-section-title">📝 回覆內容</div>',
                unsafe_allow_html=True
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
                "➕ 新增資料",
                use_container_width=True
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

        st.caption(
            "先選擇序號，找到資料後再進行修改。"
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

                st.info(
                    f"目前正在編輯：序號 {selected_id}"
                )

                with st.form("edit_form"):

                    st.markdown(
                        '<div class="form-section-title">📌 基本資訊</div>',
                        unsafe_allow_html=True
                    )

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

                    st.markdown(
                        '<div class="form-section-title">📝 回覆內容</div>',
                        unsafe_allow_html=True
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
                        "💾 儲存修改",
                        use_container_width=True
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
