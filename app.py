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
    page_icon="🏠",
    layout="wide"
)


# =========================================================
# CSS：只負責外觀，不改搜尋邏輯
# =========================================================

st.markdown(
    """
    <style>

    /* ===== 整體背景 ===== */

    .stApp {
        background-color: #F7F3ED;
        color: #4E4036;
    }

    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }


    /* ===== 標題 ===== */

    h1, h2, h3 {
        color: #4E4036 !important;
    }

    h1 {
        font-weight: 700 !important;
    }

    h2, h3 {
        font-weight: 650 !important;
    }


    /* ===== Sidebar ===== */

    section[data-testid="stSidebar"] {
        background-color: #EFE7DC;
    }

    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: #4E4036 !important;
    }

    section[data-testid="stSidebar"] .stRadio label {
        color: #5E5148;
    }


    /* ===== Radio ===== */

    div[role="radiogroup"] label {
        font-weight: 500;
    }


    /* ===== Metric ===== */

    div[data-testid="stMetric"] {
        background-color: #FFFDF9;
        border: 1px solid #E2D6C8;
        border-radius: 12px;
        padding: 10px;
    }


    /* =====================================================
       搜尋框：加大、突出
       ===================================================== */

    div[data-baseweb="input"] {
        background-color: #FFFDF9;
        border: 2px solid #CDBBA7;
        border-radius: 14px;
        min-height: 52px;
        box-shadow: 0 3px 10px rgba(110, 88, 67, 0.08);
    }

    div[data-baseweb="input"]:focus-within {
        border: 2px solid #D9824B;
        box-shadow: 0 0 0 3px rgba(217, 130, 75, 0.12);
    }

    div[data-baseweb="input"] input {
        font-size: 17px !important;
        font-weight: 500 !important;
        color: #4E4036 !important;
    }

    div[data-baseweb="input"] input::placeholder {
        color: #A4978B !important;
        font-weight: 400 !important;
    }


    /* ===== Textarea ===== */

    div[data-baseweb="textarea"] {
        background-color: #FFFDF9;
        border-radius: 10px;
        border-color: #DCCDBD;
    }


    /* ===== Selectbox ===== */

    div[data-baseweb="select"] > div {
        background-color: #FFFDF9;
        border-radius: 10px;
        border-color: #DCCDBD;
    }


    /* ===== Expander ===== */

    div[data-testid="stExpander"] {
        background-color: #FFFDF9;
        border: 1px solid #E2D6C8;
        border-radius: 12px;
    }


    /* ===== 卡片 ===== */

    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFDF9;
        border-radius: 14px;
        border: 1px solid #E2D6C8;
    }


    /* ===== 提示 / Alert ===== */

    div[data-testid="stAlert"] {
        border-radius: 10px;
    }


    /* ===== 按鈕 ===== */

    .stButton > button {
        border-radius: 10px;
        border: 1px solid #DCCDBD;
        background-color: #FFFDF9;
        color: #4E4036;
    }

    .stButton > button:hover {
        border-color: #D9824B;
        color: #D9824B;
    }

    .stButton > button[kind="primary"] {
        background-color: #D9824B;
        border-color: #D9824B;
        color: white;
    }


    /* ===== Divider ===== */

    hr {
        border-color: #E2D6C8;
    }


    /* ===== Dataframe ===== */

    div[data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
    }


    /* ===== Tabs ===== */

    button[data-baseweb="tab"] {
        color: #6B5B4E;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color: #D9824B;
    }


    /* ===== 手機版 ===== */

    @media (max-width: 768px) {

        .main .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }

        div[data-baseweb="input"] input {
            font-size: 16px !important;
        }

    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# 讀取 Excel
# =========================================================

@st.cache_data
def load_data():
    return pd.read_excel(
        FILE_PATH,
        sheet_name=0
    ).fillna("")


@st.cache_data
def load_synonyms():
    return pd.read_excel(
        FILE_PATH,
        sheet_name="搜尋詞對照"
    ).fillna("")


df = load_data()
synonym_df = load_synonyms()


# =========================================================
# Excel 儲存
# =========================================================

def save_data(new_df):

    wb = load_workbook(FILE_PATH)

    ws = wb.worksheets[0]

    # 清除原本資料
    ws.delete_rows(
        1,
        ws.max_row
    )

    # 寫入欄位名稱
    for col_num, column_name in enumerate(
        new_df.columns,
        start=1
    ):
        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )

    # 寫入資料
    for row_num, row_data in enumerate(
        new_df.itertuples(index=False),
        start=2
    ):

        for col_num, value in enumerate(
            row_data,
            start=1
        ):
            ws.cell(
                row=row_num,
                column=col_num,
                value=value
            )

    wb.save(FILE_PATH)

    st.cache_data.clear()


# =========================================================
# 同義詞處理
# =========================================================

def find_standard_words(keyword):

    standard_words = []

    keyword = str(keyword).strip()

    if not keyword:
        return standard_words

    for _, row in synonym_df.iterrows():

        standard_word = str(
            row.get("標準搜尋詞", "")
        ).strip()

        synonyms = str(
            row.get("同義詞／常見說法", "")
        ).strip()

        if not standard_word:
            continue

        if standard_word in keyword:
            standard_words.append(
                standard_word
            )
            continue

        synonym_list = synonyms.replace(
            "、",
            ","
        ).split(",")

        for synonym in synonym_list:

            synonym = synonym.strip()

            if synonym and synonym in keyword:

                standard_words.append(
                    standard_word
                )

                break

    return list(
        dict.fromkeys(
            standard_words
        )
    )


# =========================================================
# 商品關鍵字處理
# =========================================================

def find_product_words(keyword):

    product_words = []

    keyword = str(keyword).strip()

    if not keyword:
        return product_words

    product_columns = [
        "商品系列",
        "商品編號",
        "商品名稱"
    ]

    for column in product_columns:

        if column not in df.columns:
            continue

        for value in df[column].astype(str):

            value = value.strip()

            if value and value in keyword:

                product_words.append(
                    value
                )

    return list(
        dict.fromkeys(
            product_words
        )
    )


# =========================================================
# Sidebar
# =========================================================

st.sidebar.title("🏠 客服知識庫")

st.sidebar.caption(
    "NITORI Customer Service Knowledge"
)

st.sidebar.divider()

page = st.sidebar.radio(
    "功能",
    [
        "🔍 知識搜尋",
        "🛠️ 資料維護"
    ]
)

st.sidebar.divider()

st.sidebar.metric(
    "📚 目前知識庫",
    f"{len(df):,} 筆"
)

st.sidebar.info(
    "💡 知識庫持續更新中..."
)


# =========================================================
# 知識搜尋
# =========================================================

if page == "🔍 知識搜尋":

    st.title("🔍 客服知識搜尋")

    st.caption(
        "輸入顧客的說法，快速找到對應的標準回覆與判斷依據。"
    )

    st.write("🛋️　🪴　💡")

    st.subheader("💬 顧客問題")

    st.caption(
        "可以直接輸入顧客原本的說法，不需要先判斷分類。"
    )

    # -----------------------------------------------------
    # 搜尋區
    # -----------------------------------------------------

    search_col1, search_col2 = st.columns(
        [2.7, 1],
        gap="large"
    )

    with search_col1:

        keyword = st.text_input(
            "🔎 搜尋問題",
            placeholder="例如：這張椅子幾公斤？",
            label_visibility="visible"
        )

    with search_col2:

        st.caption(
            "💡 小提示：直接輸入顧客的原話即可搜尋\n\n"
            "例如「椅子多重？」、「床墊可以退嗎？」"
        )


    # -----------------------------------------------------
    # 篩選條件
    # -----------------------------------------------------

    with st.expander(
        "⚙️ 進階篩選",
        expanded=False
    ):

        filter_col1, filter_col2 = st.columns(2)

        with filter_col1:

            if "問題類型" in df.columns:

                type_options = [
                    "全部"
                ] + sorted(
                    [
                        x
                        for x in df["問題類型"]
                        .astype(str)
                        .unique()
                        if x.strip()
                    ]
                )

                selected_type = st.selectbox(
                    "問題類型",
                    type_options
                )

            else:

                selected_type = "全部"


        with filter_col2:

            if "問題細項" in df.columns:

                detail_options = [
                    "全部"
                ] + sorted(
                    [
                        x
                        for x in df["問題細項"]
                        .astype(str)
                        .unique()
                        if x.strip()
                    ]
                )

                selected_detail = st.selectbox(
                    "問題細項",
                    detail_options
                )

            else:

                selected_detail = "全部"


        filter_col3, filter_col4 = st.columns(2)

        with filter_col3:

            if "商品分類代碼" in df.columns:

                category_options = [
                    "全部"
                ] + sorted(
                    [
                        x
                        for x in df["商品分類代碼"]
                        .astype(str)
                        .unique()
                        if x.strip()
                    ]
                )

                selected_category = st.selectbox(
                    "商品分類代碼",
                    category_options
                )

            else:

                selected_category = "全部"


        with filter_col4:

            if "商品系列" in df.columns:

                series_options = [
                    "全部"
                ] + sorted(
                    [
                        x
                        for x in df["商品系列"]
                        .astype(str)
                        .unique()
                        if x.strip()
                    ]
                )

                selected_series = st.selectbox(
                    "商品系列",
                    series_options
                )

            else:

                selected_series = "全部"


    # =====================================================
    # 搜尋
    # =====================================================

    if keyword.strip():

        standard_words = find_standard_words(
            keyword
        )

        product_words = find_product_words(
            keyword
        )

        result_df = df.copy()


        # -------------------------------------------------
        # 篩選
        # -------------------------------------------------

        if (
            selected_type != "全部"
            and "問題類型" in result_df.columns
        ):

            result_df = result_df[
                result_df["問題類型"]
                .astype(str)
                .eq(selected_type)
            ]


        if (
            selected_detail != "全部"
            and "問題細項" in result_df.columns
        ):

            result_df = result_df[
                result_df["問題細項"]
                .astype(str)
                .eq(selected_detail)
            ]


        if (
            selected_category != "全部"
            and "商品分類代碼" in result_df.columns
        ):

            result_df = result_df[
                result_df["商品分類代碼"]
                .astype(str)
                .eq(selected_category)
            ]


        if (
            selected_series != "全部"
            and "商品系列" in result_df.columns
        ):

            result_df = result_df[
                result_df["商品系列"]
                .astype(str)
                .eq(selected_series)
            ]


        # -------------------------------------------------
        # 搜尋階段 1
        # 標準搜尋詞 → 問題摘要
        # -------------------------------------------------

        exact_result = pd.DataFrame()

        if standard_words and "問題摘要" in result_df.columns:

            mask = result_df[
                "問題摘要"
            ].astype(str).apply(
                lambda x: any(
                    word in x
                    for word in standard_words
                )
            )

            exact_result = result_df[
                mask
            ]


        # -------------------------------------------------
        # 搜尋階段 2
        # 標準搜尋詞 → 多欄位
        # -------------------------------------------------

        keyword_result = pd.DataFrame()

        search_columns = [
            "問題摘要",
            "關鍵字",
            "問題細項",
            "商品名稱",
            "商品系列",
            "商品編號"
        ]

        search_columns = [
            column
            for column in search_columns
            if column in result_df.columns
        ]

        if standard_words and search_columns:

            def match_row(row):

                text = " ".join(
                    str(row[column])
                    for column in search_columns
                )

                return any(
                    word in text
                    for word in standard_words
                )

            mask = result_df.apply(
                match_row,
                axis=1
            )

            keyword_result = result_df[
                mask
            ]


        # -------------------------------------------------
        # TF-IDF 模糊搜尋
        # -------------------------------------------------

        fuzzy_result = pd.DataFrame()

        if (
            exact_result.empty
            and keyword_result.empty
            and len(result_df) > 0
        ):

            search_columns_for_tfidf = [
                column
                for column in [
                    "問題摘要",
                    "關鍵字",
                    "問題細項",
                    "商品名稱",
                    "商品系列",
                    "商品編號"
                ]
                if column in result_df.columns
            ]

            documents = result_df[
                search_columns_for_tfidf
            ].astype(str).agg(
                " ".join,
                axis=1
            ).tolist()

            vectorizer = TfidfVectorizer(
                analyzer="char",
                ngram_range=(2, 4)
            )

            try:

                matrix = vectorizer.fit_transform(
                    documents
                )

                query_vector = vectorizer.transform(
                    [keyword]
                )

                scores = cosine_similarity(
                    query_vector,
                    matrix
                )[0]

                result_df = result_df.copy()

                result_df["_score"] = scores

                fuzzy_result = result_df[
                    result_df["_score"] > 0.05
                ].sort_values(
                    "_score",
                    ascending=False
                )

            except ValueError:

                fuzzy_result = pd.DataFrame()


        # -------------------------------------------------
        # 合併結果
        # -------------------------------------------------

        if not exact_result.empty:

            final_result = exact_result.copy()

        elif not keyword_result.empty:

            final_result = keyword_result.copy()

        elif not fuzzy_result.empty:

            final_result = fuzzy_result.copy()

        else:

            final_result = pd.DataFrame()


        # -------------------------------------------------
        # 商品匹配優先
        # -------------------------------------------------

        if (
            product_words
            and not final_result.empty
        ):

            def product_match(row):

                text = " ".join(
                    str(row.get(column, ""))
                    for column in [
                        "商品系列",
                        "商品編號",
                        "商品名稱"
                    ]
                )

                return any(
                    word in text
                    for word in product_words
                )

            final_result = final_result.copy()

            final_result["_product_match"] = (
                final_result.apply(
                    product_match,
                    axis=1
                )
            )

            final_result = final_result.sort_values(
                "_product_match",
                ascending=False
            )


        # =================================================
        # 顯示結果
        # =================================================

        st.divider()

        if not final_result.empty:

            st.subheader(
                f"🔎 找到 {len(final_result)} 筆相關資料"
            )

            display_result = final_result.head(
                20
            )

            for _, row in display_result.iterrows():

                with st.container(
                    border=True
                ):

                    # -------------------------------------
                    # 標題
                    # -------------------------------------

                    title = str(
                        row.get(
                            "問題摘要",
                            "未提供問題摘要"
                        )
                    )

                    st.markdown(
                        f"### 💬 {title}"
                    )


                    # -------------------------------------
                    # 標籤
                    # -------------------------------------

                    tags = []

                    if str(
                        row.get(
                            "問題類型",
                            ""
                        )
                    ).strip():

                        tags.append(
                            f"📂 {row['問題類型']}"
                        )


                    if str(
                        row.get(
                            "問題細項",
                            ""
                        )
                    ).strip():

                        tags.append(
                            f"🏷️ {row['問題細項']}"
                        )


                    if str(
                        row.get(
                            "商品系列",
                            ""
                        )
                    ).strip():

                        tags.append(
                            f"🛋️ {row['商品系列']}"
                        )


                    if tags:

                        st.caption(
                            "　".join(tags)
                        )


                    # -------------------------------------
                    # 商品資訊
                    # -------------------------------------

                    info_col1, info_col2, info_col3 = st.columns(
                        3
                    )


                    with info_col1:

                        if "商品名稱" in row.index:

                            st.caption("商品名稱")

                            st.write(
                                str(
                                    row["商品名稱"]
                                )
                            )


                    with info_col2:

                        if "商品編號" in row.index:

                            st.caption("商品編號")

                            st.write(
                                str(
                                    row["商品編號"]
                                )
                            )


                    with info_col3:

                        if "商品分類代碼" in row.index:

                            st.caption("商品分類")

                            st.write(
                                str(
                                    row["商品分類代碼"]
                                )
                            )


                    # -------------------------------------
                    # 標準回覆
                    # -------------------------------------

                    if "標準回覆" in row.index:

                        st.success(
                            str(
                                row["標準回覆"]
                            )
                        )


                    # -------------------------------------
                    # 詳細資料
                    # -------------------------------------

                    with st.expander(
                        "查看判斷依據與資料來源"
                    ):

                        detail_col1, detail_col2 = st.columns(
                            2
                        )

                        with detail_col1:

                            st.markdown(
                                "**判斷依據**"
                            )

                            st.write(
                                str(
                                    row.get(
                                        "判斷依據",
                                        ""
                                    )
                                )
                            )


                        with detail_col2:

                            st.markdown(
                                "**資料來源**"
                            )

                            st.write(
                                str(
                                    row.get(
                                        "資料來源",
                                        ""
                                    )
                                )
                            )


                        if "關鍵字" in row.index:

                            st.markdown(
                                "**關鍵字**"
                            )

                            st.write(
                                str(
                                    row["關鍵字"]
                                )
                            )


                        if "建立人" in row.index:

                            st.caption(
                                f"建立人：{row.get('建立人', '')}　"
                                f"建立日期：{row.get('建立日期', '')}"
                            )


                        if "更新人" in row.index:

                            st.caption(
                                f"更新人：{row.get('更新人', '')}　"
                                f"更新日期：{row.get('更新日期', '')}"
                            )


        else:

            # -------------------------------------------------
            # 找不到結果
            # -------------------------------------------------

            st.info(
                "🔍 目前沒有找到相符的資料。\n\n"
                "可以試著換個說法，例如使用顧客實際會說的詞彙。"
            )


    else:

        # =================================================
        # 尚未搜尋
        # =================================================

        st.divider()

        st.subheader(
            "👋 歡迎使用客服知識庫"
        )

        st.write(
            "先輸入顧客的問題，系統會協助你找到相關的標準回覆。"
        )

        empty_col1, empty_col2, empty_col3 = st.columns(
            3
        )

        with empty_col1:

            st.info(
                "💬 **直接輸入問題**\n\n"
                "不用先判斷問題分類，直接輸入顧客說法即可。"
            )


        with empty_col2:

            st.info(
                "🔎 **搜尋相關知識**\n\n"
                "系統會從問題、關鍵字及商品資訊中尋找相關資料。"
            )


        with empty_col3:

            st.info(
                "📚 **查看標準回覆**\n\n"
                "找到資料後，可以直接查看標準回覆及判斷依據。"
            )


        st.divider()

        st.subheader(
            "🌿 客服小角落"
        )

        st.caption(
            "每天整理一點知識，讓下一次遇到相同問題時，可以更快找到答案。"
        )


# =========================================================
# 資料維護
# =========================================================

elif page == "🛠️ 資料維護":

    st.title("🛠️ 資料維護")

    st.caption(
        "可查看、建立及修改知識庫資料。"
    )


    # =====================================================
    # Tabs
    # =====================================================

    tab1, tab2, tab3 = st.tabs(
        [
            "📋 目前資料",
            "➕ 新增資料",
            "✏️ 修改資料"
        ]
    )


    # =====================================================
    # Tab 1：目前資料
    # =====================================================

    with tab1:

        st.subheader(
            "📋 知識庫目前資料"
        )

        data_search = st.text_input(
            "搜尋資料",
            placeholder="輸入關鍵字搜尋目前資料..."
        )


        display_df = df.copy()


        if data_search.strip():

            search_mask = display_df.astype(
                str
            ).apply(
                lambda column: column.str.contains(
                    data_search,
                    case=False,
                    na=False
                )
            ).any(
                axis=1
            )

            display_df = display_df[
                search_mask
            ]


        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )


    # =====================================================
    # Tab 2：新增資料
    # =====================================================

    with tab2:

        st.subheader(
            "➕ 新增知識"
        )

        new_id = (
            int(
                pd.to_numeric(
                    df["序號"],
                    errors="coerce"
                ).max()
            )
            + 1
            if "序號" in df.columns
            and not df.empty
            else 1
        )


        col1, col2 = st.columns(2)

        with col1:

            new_type = st.text_input(
                "問題類型"
            )

            new_detail = st.text_input(
                "問題細項"
            )

            new_category = st.text_input(
                "商品分類代碼"
            )

            new_product_name = st.text_input(
                "商品名稱"
            )

            new_series = st.text_input(
                "商品系列"
            )

            new_product_id = st.text_input(
                "商品編號"
            )


        with col2:

            new_summary = st.text_input(
                "問題摘要"
            )

            new_reply = st.text_area(
                "標準回覆",
                height=180
            )

            new_basis = st.text_area(
                "判斷依據",
                height=120
            )

            new_source = st.text_input(
                "資料來源"
            )

            new_keywords = st.text_input(
                "關鍵字"
            )

            new_creator = st.text_input(
                "建立人"
            )


        if st.button(
            "💾 儲存新增資料",
            type="primary"
        ):

            today = datetime.now().strftime(
                "%Y-%m-%d"
            )

            new_row = {
                "序號": new_id,
                "問題類型": new_type,
                "問題細項": new_detail,
                "商品分類代碼": new_category,
                "商品名稱": new_product_name,
                "商品系列": new_series,
                "商品編號": new_product_id,
                "問題摘要": new_summary,
                "標準回覆": new_reply,
                "判斷依據": new_basis,
                "資料來源": new_source,
                "關鍵字": new_keywords,
                "建立人": new_creator,
                "建立日期": today,
                "更新人": new_creator,
                "更新日期": today
            }


            new_df = pd.concat(
                [
                    df,
                    pd.DataFrame(
                        [new_row]
                    )
                ],
                ignore_index=True
            )


            save_data(
                new_df
            )

            st.success(
                "資料已新增完成！"
            )

            st.rerun()


    # =====================================================
    # Tab 3：修改資料
    # =====================================================

    with tab3:

        st.subheader(
            "✏️ 修改知識"
        )

        if "序號" in df.columns:

            id_list = sorted(
                pd.to_numeric(
                    df["序號"],
                    errors="coerce"
                ).dropna().astype(int).tolist()
            )

            selected_id = st.selectbox(
                "選擇要修改的序號",
                id_list
            )


            selected_row = df[
                pd.to_numeric(
                    df["序號"],
                    errors="coerce"
                ) == selected_id
            ].iloc[0]


            edit_col1, edit_col2 = st.columns(2)


            with edit_col1:

                edit_type = st.text_input(
                    "問題類型",
                    value=str(
                        selected_row.get(
                            "問題類型",
                            ""
                        )
                    )
                )

                edit_detail = st.text_input(
                    "問題細項",
                    value=str(
                        selected_row.get(
                            "問題細項",
                            ""
                        )
                    )
                )

                edit_category = st.text_input(
                    "商品分類代碼",
                    value=str(
                        selected_row.get(
                            "商品分類代碼",
                            ""
                        )
                    )
                )

                edit_product_name = st.text_input(
                    "商品名稱",
                    value=str(
                        selected_row.get(
                            "商品名稱",
                            ""
                        )
                    )
                )

                edit_series = st.text_input(
                    "商品系列",
                    value=str(
                        selected_row.get(
                            "商品系列",
                            ""
                        )
                    )
                )

                edit_product_id = st.text_input(
                    "商品編號",
                    value=str(
                        selected_row.get(
                            "商品編號",
                            ""
                        )
                    )
                )


            with edit_col2:

                edit_summary = st.text_input(
                    "問題摘要",
                    value=str(
                        selected_row.get(
                            "問題摘要",
                            ""
                        )
                    )
                )

                edit_reply = st.text_area(
                    "標準回覆",
                    value=str(
                        selected_row.get(
                            "標準回覆",
                            ""
                        )
                    ),
                    height=180
                )

                edit_basis = st.text_area(
                    "判斷依據",
                    value=str(
                        selected_row.get(
                            "判斷依據",
                            ""
                        )
                    ),
                    height=120
                )

                edit_source = st.text_input(
                    "資料來源",
                    value=str(
                        selected_row.get(
                            "資料來源",
                            ""
                        )
                    )
                )

                edit_keywords = st.text_input(
                    "關鍵字",
                    value=str(
                        selected_row.get(
                            "關鍵字",
                            ""
                        )
                    )
                )

                edit_updater = st.text_input(
                    "更新人"
                )


            if st.button(
                "💾 儲存修改",
                type="primary"
            ):

                today = datetime.now().strftime(
                    "%Y-%m-%d"
                )


                edit_index = df.index[
                    pd.to_numeric(
                        df["序號"],
                        errors="coerce"
                    ) == selected_id
                ][0]


                df.loc[
                    edit_index,
                    "問題類型"
                ] = edit_type

                df.loc[
                    edit_index,
                    "問題細項"
                ] = edit_detail

                df.loc[
                    edit_index,
                    "商品分類代碼"
                ] = edit_category

                df.loc[
                    edit_index,
                    "商品名稱"
                ] = edit_product_name

                df.loc[
                    edit_index,
                    "商品系列"
                ] = edit_series

                df.loc[
                    edit_index,
                    "商品編號"
                ] = edit_product_id

                df.loc[
                    edit_index,
                    "問題摘要"
                ] = edit_summary

                df.loc[
                    edit_index,
                    "標準回覆"
                ] = edit_reply

                df.loc[
                    edit_index,
                    "判斷依據"
                ] = edit_basis

                df.loc[
                    edit_index,
                    "資料來源"
                ] = edit_source

                df.loc[
                    edit_index,
                    "關鍵字"
                ] = edit_keywords

                df.loc[
                    edit_index,
                    "更新人"
                ] = edit_updater

                df.loc[
                    edit_index,
                    "更新日期"
                ] = today


                save_data(
                    df
                )

                st.success(
                    "資料已修改完成！"
                )

                st.rerun()
