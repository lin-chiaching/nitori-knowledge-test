import streamlit as st
import streamlit.components.v1 as components
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
# CSS
# =========================================================

st.markdown("""
<style>

    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    .result-card {
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #dddddd;
        margin-bottom: 12px;
    }

</style>
""", unsafe_allow_html=True)


# =========================================================
# 讀取資料
# =========================================================

@st.cache_data
def load_data():

    df = pd.read_excel(
        FILE_PATH,
        sheet_name=0
    ).fillna("")

    return df


@st.cache_data
def load_synonyms():

    try:

        df = pd.read_excel(
            FILE_PATH,
            sheet_name="搜尋詞對照"
        ).fillna("")

        return df

    except Exception:

        return pd.DataFrame()


df = load_data()
synonym_df = load_synonyms()


# =========================================================
# 儲存資料
# =========================================================

def save_data(df):

    wb = load_workbook(FILE_PATH)

    ws = wb[wb.sheetnames[0]]

    # 清除原本資料
    ws.delete_rows(
        1,
        ws.max_row
    )

    # 寫入欄位名稱
    for col_num, column_name in enumerate(df.columns, 1):

        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )

    # 寫入資料
    for row_num, row_data in enumerate(
        df.itertuples(index=False),
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


# =========================================================
# 搜尋詞處理
# =========================================================

def expand_search_word(keyword):

    keyword = str(keyword).strip()

    if keyword == "":
        return keyword

    expanded_words = [keyword]

    if not synonym_df.empty:

        for _, row in synonym_df.iterrows():

            try:

                original = str(row.iloc[0]).strip()
                synonym = str(row.iloc[1]).strip()

                if original == keyword:

                    expanded_words.append(synonym)

                elif synonym == keyword:

                    expanded_words.append(original)

            except Exception:

                pass

    return " ".join(
        list(dict.fromkeys(expanded_words))
    )


# =========================================================
# 商品名稱 / 商品編號相關
# =========================================================

def find_product_matches(keyword):

    keyword = str(keyword).strip()

    if keyword == "":
        return df.iloc[0:0]

    result = df[
        df.astype(str)
        .apply(
            lambda row: row.str.contains(
                keyword,
                case=False,
                regex=False
            ).any(),
            axis=1
        )
    ]

    return result


# =========================================================
# 側邊欄
# =========================================================

st.sidebar.title("🏠 客服作業系統")

st.sidebar.caption(
    "NITORI Customer Service"
)

st.sidebar.divider()


page = st.sidebar.radio(
    "功能",
    [
        "🔍 知識搜尋",
        "🗺️ 門市地圖",
        "🛠️ 資料維護"
    ]
)


# =========================================================
# ① 知識搜尋
# =========================================================

if page == "🔍 知識搜尋":

    st.title("🔍 客服知識搜尋")

    st.caption(
        "輸入顧客問題或關鍵字，快速找到相關客服資訊"
    )

    st.divider()


    # -----------------------------------------------------
    # 搜尋框
    # -----------------------------------------------------

    keyword = st.text_input(
        "請輸入搜尋內容",
        placeholder="例如：重量、床墊、退貨、配送、保固……"
    )


    if keyword:

        # -------------------------------------------------
        # 搜尋詞擴充
        # -------------------------------------------------

        search_text = expand_search_word(keyword)


        # -------------------------------------------------
        # 建立搜尋欄位
        # -------------------------------------------------

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


        # -------------------------------------------------
        # 組合搜尋文字
        # -------------------------------------------------

        search_data = (
            df[search_columns]
            .astype(str)
            .agg(" ".join, axis=1)
        )


        # -------------------------------------------------
        # TF-IDF
        # -------------------------------------------------

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


        # -------------------------------------------------
        # 排序
        # -------------------------------------------------

        df_result = df.copy()

        df_result["_score"] = similarity

        df_result = df_result.sort_values(
            "_score",
            ascending=False
        )


        # -------------------------------------------------
        # 只顯示有相關性的結果
        # -------------------------------------------------

        df_result = df_result[
            df_result["_score"] > 0
        ]


        st.write(
            f"找到 {len(df_result)} 筆相關資料"
        )


        # -------------------------------------------------
        # 顯示結果
        # -------------------------------------------------

        if len(df_result) == 0:

            st.warning(
                "目前沒有找到相關資料，請換個關鍵字試試看。"
            )

        else:

            display_result = df_result.head(10)


            for index, row in display_result.iterrows():

                score = row["_score"]


                # -----------------------------------------
                # 商品 / 問題標題
                # -----------------------------------------

                if "問題簡述" in df.columns:

                    title = row["問題簡述"]

                else:

                    title = "相關資料"


                st.markdown(
                    f"""
                    <div class="result-card">
                    <h4>{title}</h4>
                    </div>
                    """,
                    unsafe_allow_html=True
                )


                # -----------------------------------------
                # 分類
                # -----------------------------------------

                info_col1, info_col2, info_col3 = st.columns(3)


                with info_col1:

                    if "大分類" in df.columns:

                        st.caption("大分類")

                        st.write(
                            row["大分類"]
                        )


                with info_col2:

                    if "中分類" in df.columns:

                        st.caption("中分類")

                        st.write(
                            row["中分類"]
                        )


                with info_col3:

                    if "小分類" in df.columns:

                        st.caption("小分類")

                        st.write(
                            row["小分類"]
                        )


                # -----------------------------------------
                # 解答
                # -----------------------------------------

                if "解答" in df.columns:

                    st.markdown("**📌 解答**")

                    st.write(
                        row["解答"]
                    )


                # -----------------------------------------
                # 判斷方式
                # -----------------------------------------

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


                # -----------------------------------------
                # 判斷工具
                # -----------------------------------------

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


                # -----------------------------------------
                # 關鍵字
                # -----------------------------------------

                if "關鍵字" in df.columns:

                    if str(
                        row["關鍵字"]
                    ).strip():

                        st.caption(
                            "關鍵字："
                            + str(row["關鍵字"])
                        )


                # -----------------------------------------
                # 相似度
                # -----------------------------------------

                st.caption(
                    f"相關程度：{score:.1%}"
                )


                st.divider()


# =========================================================
# ② 門市地圖
# =========================================================

elif page == "🗺️ 門市地圖":

    st.title("🗺️ 門市地圖")

    st.caption(
        "宜得利家居門市資訊查詢"
    )

    st.divider()


    # -----------------------------------------------------
    # Power BI
    # -----------------------------------------------------

    components.iframe(
        "https://app.powerbi.com/groups/me/reports/851d09f1-f169-4049-9090-640f09cdc385/3b82b83a4737660969c9?experience=power-bi",
        height=800
    )


# =========================================================
# ③ 資料維護
# =========================================================

elif page == "🛠️ 資料維護":

    st.title("🛠️ 資料維護")

    st.caption(
        "新增、修改及管理客服知識庫資料"
    )

    st.divider()


    # -----------------------------------------------------
    # 顯示目前資料
    # -----------------------------------------------------

    st.subheader("📋 目前資料")

    st.dataframe(
        df,
        use_container_width=True,
        height=400
    )


    st.divider()


    # -----------------------------------------------------
    # 新增資料
    # -----------------------------------------------------

    st.subheader("➕ 新增資料")

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

        # 自動產生序號
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

                    new_number = int(
                        max_number
                    ) + 1

            except Exception:

                new_number = len(df) + 1


            new_data["序號"] = new_number


        # 建立新資料
        new_row = pd.DataFrame(
            [new_data]
        )


        # 調整欄位順序
        new_row = new_row.reindex(
            columns=df.columns
        )


        # 加入資料
        df = pd.concat(
            [df, new_row],
            ignore_index=True
        )


        # 儲存
        save_data(df)


        st.success(
            "資料已新增！"
        )

        st.rerun()


    st.divider()


    # -----------------------------------------------------
    # 修改資料
    # -----------------------------------------------------

    st.subheader("✏️ 修改資料")


    if len(df) > 0:

        selected_index = st.selectbox(
            "請選擇要修改的資料",
            options=df.index,
            format_func=lambda x:
                f"{x + 1}｜"
                + str(
                    df.loc[
                        x,
                        "問題簡述"
                    ]
                    if "問題簡述" in df.columns
                    else ""
                )
        )


        selected_row = df.loc[
            selected_index
        ]


        edit_data = {}


        for column in df.columns:

            if column == "序號":

                st.text_input(
                    column,
                    value=str(
                        selected_row[column]
                    ),
                    disabled=True,
                    key=f"edit_{column}"
                )

            else:

                edit_data[column] = st.text_area(
                    column,
                    value=str(
                        selected_row[column]
                    ),
                    key=f"edit_{selected_index}_{column}"
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


            # ---------------------------------------------
            # 更新日期
            # ---------------------------------------------

            if "更新日期" in df.columns:

                df.loc[
                    selected_index,
                    "更新日期"
                ] = datetime.now().strftime(
                    "%Y-%m-%d"
                )


            # ---------------------------------------------
            # 更新人
            # ---------------------------------------------

            if "更新人" in df.columns:

                df.loc[
                    selected_index,
                    "更新人"
                ] = "客服"


            # ---------------------------------------------
            # 儲存
            # ---------------------------------------------

            save_data(df)


            st.success(
                "資料已更新！"
            )

            st.rerun()
