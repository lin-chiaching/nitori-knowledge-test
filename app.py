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
# 頁面樣式
# ============================================================

st.markdown(
    """
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

    .announcement-card {
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #dddddd;
        margin-bottom: 15px;
        background-color: #fafafa;
    }

    .tool-card {
        padding: 18px;
        border-radius: 12px;
        border: 1px solid #dddddd;
        margin-bottom: 12px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 公告內容清理
# ============================================================

def clean_announcement_content(text):

    if pd.isna(text):
        return ""

    text = str(text)

    # --------------------------------------------------------
    # HTML entity 轉換
    # 例如 &nbsp; → 空白
    # --------------------------------------------------------

    text = html.unescape(text)

    # --------------------------------------------------------
    # 移除 SVG
    # --------------------------------------------------------

    text = re.sub(
        r"<svg.*?</svg>",
        "",
        text,
        flags=re.IGNORECASE | re.DOTALL
    )

    # 移除單獨出現的 svg 文字
    text = re.sub(
        r"\bsvg\b",
        "",
        text,
        flags=re.IGNORECASE
    )

    # --------------------------------------------------------
    # HTML 標籤處理
    # --------------------------------------------------------

    # h1～h6 開始標籤
    text = re.sub(
        r"<\s*h[1-6][^>]*>",
        "",
        text,
        flags=re.IGNORECASE
    )

    # h1～h6 結束標籤
    text = re.sub(
        r"</\s*h[1-6]\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # p 開始標籤
    text = re.sub(
        r"<\s*p[^>]*>",
        "",
        text,
        flags=re.IGNORECASE
    )

    # p 結束標籤
    text = re.sub(
        r"</\s*p\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # br → 換行
    text = re.sub(
        r"<\s*br\s*/?\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # div 開始標籤
    text = re.sub(
        r"<\s*div[^>]*>",
        "",
        text,
        flags=re.IGNORECASE
    )

    # div 結束標籤
    text = re.sub(
        r"</\s*div\s*>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    # --------------------------------------------------------
    # 移除剩餘 HTML 標籤
    # --------------------------------------------------------

    text = re.sub(
        r"<[^>]+>",
        "",
        text
    )

    # --------------------------------------------------------
    # 空白整理
    # --------------------------------------------------------

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    text = re.sub(
        r"\n[ \t]+",
        "\n",
        text
    )

    # 連續 3 行以上 → 最多保留 2 行
    text = re.sub(
        r"\n\s*\n\s*\n+",
        "\n\n",
        text
    )

    return text.strip()


# ============================================================
# HTML 顯示安全處理
# ============================================================

def safe_html_text(text):

    if pd.isna(text):
        return ""

    return html.escape(
        str(text)
    )


# ============================================================
# 讀取資料
# ============================================================

@st.cache_data
def load_data():

    df = pd.read_excel(
        FILE_PATH,
        sheet_name=0
    ).fillna("")

    # 清除欄位名稱前後空白
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


# ============================================================
# 讀取搜尋詞對照
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
# 讀取內部公告
# ============================================================

@st.cache_data
def load_announcements():

    try:

        df = pd.read_excel(
            FILE_PATH,
            sheet_name="內部公告"
        ).fillna("")

        # 清除欄位名稱前後空白
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

        # 如果 Excel 缺少欄位，自動建立
        for column in required_columns:

            if column not in df.columns:

                df[column] = ""

        # 固定欄位順序
        df = df[required_columns]

        return df

    except Exception:

        return pd.DataFrame()


# ============================================================
# 讀取常用工具
# ============================================================

@st.cache_data
def load_tools():

    try:

        df = pd.read_excel(
            FILE_PATH,
            sheet_name="常用工具"
        ).fillna("")

        # 清除欄位名稱前後空白
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

        # 如果 Excel 缺少欄位，自動建立
        for column in required_columns:

            if column not in df.columns:

                df[column] = ""

        # 固定欄位順序
        df = df[required_columns]

        return df

    except Exception:

        return pd.DataFrame()


# ============================================================
# 載入資料
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

    # 清除原本資料
    ws.delete_rows(
        1,
        ws.max_row
    )

    # 寫入標題
    for col_num, column_name in enumerate(
        data.columns,
        1
    ):

        ws.cell(
            row=1,
            column=col_num,
            value=column_name
        )

    # 寫入資料
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

    # 清除 Streamlit 快取
    st.cache_data.clear()


# ============================================================
# 儲存知識庫
# ============================================================

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

    expanded_words = [
        keyword
    ]

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

st.sidebar.title(
    "🐱 客服作業系統"
)

st.sidebar.caption(
    "NITORI Customer Service"
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


# ============================================================
# ① 首頁
# ============================================================

if page == "🏠 首頁":

    st.title(
        "🐱 歡迎來到客服作業系統"
    )

    st.subheader(
        "今天也辛苦了！"
    )

    st.divider()

    st.info(
        "這裡可以快速查詢客服知識、門市資訊、內部公告及常用工具。"
    )

    st.divider()

    st.subheader(
        "🚀 快速入口"
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            "### 🔍 知識搜尋"
        )

        st.write(
            "快速查詢客服知識與常見問題。"
        )

    with col2:

        st.markdown(
            "### 🗺️ 門市地圖"
        )

        st.write(
            "查詢門市位置、交通、停車等資訊。"
        )

    with col3:

        st.markdown(
            "### 📢 公告＆工具"
        )

        st.write(
            "查看最新內部公告及常用系統。"
        )

    st.divider()

    st.subheader(
        "📢 最新公告"
    )

    if announcement_df.empty:

        st.write(
            "目前沒有公告。"
        )

    else:

        # 只顯示「顯示」的公告
        show_announcements = announcement_df[
            announcement_df["顯示狀態"]
            .astype(str)
            .str.strip()
            == "顯示"
        ]

        if show_announcements.empty:

            st.write(
                "目前沒有公告。"
            )

        else:

            # 最新公告放前面
            show_announcements = show_announcements.tail(
                3
            )

            for _, row in show_announcements.iloc[::-1].iterrows():

                title = str(
                    row.get(
                        "公告標題",
                        ""
                    )
                )

                content = clean_announcement_content(
                    row.get(
                        "公告內容",
                        ""
                    )
                )

                date = str(
                    row.get(
                        "公告日期",
                        ""
                    )
                )

                safe_title = safe_html_text(
                    title
                )

                safe_content = safe_html_text(
                    content
                ).replace(
                    "\n",
                    "<br>"
                )

                safe_date = safe_html_text(
                    date
                )

                st.markdown(
                    f"""
                    <div class="announcement-card">

                    <h4>📌 {safe_title}</h4>

                    <p>{safe_content}</p>

                    <small>
                    公告日期：{safe_date}
                    </small>

                    </div>
                    """,
                    unsafe_allow_html=True
                )


# ============================================================
# ② 知識搜尋
# ============================================================

elif page == "🔍 知識搜尋":

    st.title(
        "🔍 客服知識搜尋"
    )

    st.caption(
        "輸入顧客問題或關鍵字，快速找到相關客服資訊"
    )

    st.divider()

    keyword = st.text_input(
        "請輸入搜尋內容",
        placeholder="例如：重量、床墊、退貨、配送、保固……"
    )

    if keyword:

        search_text = expand_search_word(
            keyword
        )

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

                search_columns.append(
                    column
                )

        if not search_columns:

            st.error(
                "知識庫沒有可以搜尋的欄位。"
            )

        else:

            search_data = (
                df[search_columns]
                .astype(str)
                .agg(
                    " ".join,
                    axis=1
                )
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

            st.write(
                f"找到 {len(df_result)} 筆相關資料"
            )

            if len(df_result) == 0:

                st.warning(
                    "目前沒有找到相關資料，請換個關鍵字試試看。"
                )

            else:

                display_result = df_result.head(
                    10
                )

                for _, row in display_result.iterrows():

                    score = row["_score"]

                    if "問題簡述" in df.columns:

                        title = row[
                            "問題簡述"
                        ]

                    else:

                        title = "相關資料"

                    st.markdown(
                        f"""
                        <div class="result-card">

                        <h4>{safe_html_text(title)}</h4>

                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    info_col1, info_col2, info_col3 = st.columns(
                        3
                    )

                    with info_col1:

                        if "大分類" in df.columns:

                            st.caption(
                                "大分類"
                            )

                            st.write(
                                row["大分類"]
                            )

                    with info_col2:

                        if "中分類" in df.columns:

                            st.caption(
                                "中分類"
                            )

                            st.write(
                                row["中分類"]
                            )

                    with info_col3:

                        if "小分類" in df.columns:

                            st.caption(
                                "小分類"
                            )

                            st.write(
                                row["小分類"]
                            )

                    if "解答" in df.columns:

                        st.markdown(
                            "**📌 解答**"
                        )

                        st.write(
                            row["解答"]
                        )

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
                                row[
                                    "判斷使用的工具"
                                ]
                            )

                    if "關鍵字" in df.columns:

                        if str(
                            row["關鍵字"]
                        ).strip():

                            st.caption(
                                "關鍵字："
                                + str(
                                    row["關鍵字"]
                                )
                            )

                    st.caption(
                        f"相關程度：{score:.1%}"
                    )

                    st.divider()


# ============================================================
# ③ 門市地圖
# ============================================================

elif page == "🗺️ 門市地圖":

    st.title(
        "🗺️ 門市地圖"
    )

    st.caption(
        "宜得利家居門市資訊查詢"
    )

    st.divider()

    components.iframe(
        "https://app.powerbi.com/reportEmbed?reportId=851d09f1-f169-4049-9090-640f09cdc385&autoAuth=true&ctid=94d67acc-8c17-4d3b-8463-9829cf08ef55",
        height=800
    )


# ============================================================
# ④ 公告＆常用工具
# ============================================================

elif page == "📢 公告＆常用工具":

    st.title(
        "📢 公告＆常用工具"
    )

    st.caption(
        "查看內部公告及常用系統、表單與文件"
    )

    st.divider()


    # ========================================================
    # 內部公告
    # ========================================================

    st.subheader(
        "📢 內部公告"
    )

    if announcement_df.empty:

        st.info(
            "目前沒有公告資料。"
        )

    else:

        show_announcements = announcement_df[
            announcement_df["顯示狀態"]
            .astype(str)
            .str.strip()
            == "顯示"
        ]

        if show_announcements.empty:

            st.info(
                "目前沒有公告資料。"
            )

        else:

            # 最新公告放前面
            show_announcements = show_announcements.iloc[::-1]

            for _, row in show_announcements.iterrows():

                title = str(
                    row.get(
                        "公告標題",
                        ""
                    )
                )

                content = clean_announcement_content(
                    row.get(
                        "公告內容",
                        ""
                    )
                )

                date = str(
                    row.get(
                        "公告日期",
                        ""
                    )
                )

                activity = str(
                    row.get(
                        "活動期限",
                        ""
                    )
                )

                point = str(
                    row.get(
                        "點數期限",
                        ""
                    )
                )

                # HTML 安全處理
                safe_title = safe_html_text(
                    title
                )

                content_html = safe_html_text(
                    content
                ).replace(
                    "\n",
                    "<br>"
                )

                safe_date = safe_html_text(
                    date
                )

                safe_activity = safe_html_text(
                    activity
                )

                safe_point = safe_html_text(
                    point
                )

                st.markdown(
                    f"""
                    <div class="announcement-card">

                    <h4>📌 {safe_title}</h4>

                    <p>{content_html}</p>

                    <p>
                    📅 公告日期：{safe_date}
                    </p>

                    <p>
                    🗓️ 活動期限：{safe_activity}
                    </p>

                    <p>
                    🪙 點數期限：{safe_point}
                    </p>

                    </div>
                    """,
                    unsafe_allow_html=True
                )


    st.divider()


    # ========================================================
    # 常用工具
    # ========================================================

    st.subheader(
        "🔗 常用工具"
    )

    if tool_df.empty:

        st.info(
            "目前沒有常用工具。"
        )

    else:

        show_tools = tool_df[
            tool_df["顯示狀態"]
            .astype(str)
            .str.strip()
            == "顯示"
        ]

        if show_tools.empty:

            st.info(
                "目前沒有常用工具。"
            )

        else:

            categories = show_tools[
                "工具分類"
            ].astype(str).unique()

            for category in categories:

                st.markdown(
                    f"### {safe_html_text(category)}",
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

                # 每三個工具一排
                for i in range(
                    0,
                    len(tools_list),
                    3
                ):

                    cols = st.columns(
                        3
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

                            st.markdown(
                                f"""
                                <div class="tool-card">

                                <h4>
                                🔗 {safe_html_text(tool_name)}
                                </h4>

                                <p>
                                {safe_html_text(tool_description)}
                                </p>

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
                                    "開啟",
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

    st.title(
        "🛠️ 資料維護"
    )

    st.caption(
        "管理客服知識、內部公告及常用工具"
    )

    st.divider()


    # ========================================================
    # 維護項目選擇
    # ========================================================

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
    # 知識庫維護
    # ========================================================

    if maintenance_type == "📚 知識庫":

        st.subheader(
            "📋 目前知識庫資料"
        )

        st.dataframe(
            df,
            use_container_width=True,
            height=400
        )

        st.divider()

        st.subheader(
            "➕ 新增知識"
        )

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

                        new_number = (
                            int(max_number) + 1
                        )

                except Exception:

                    new_number = len(df) + 1

                new_data["序號"] = new_number

            new_row = pd.DataFrame(
                [new_data]
            )

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

            save_knowledge_data(
                df
            )

            st.success(
                "資料已新增！"
            )

            st.rerun()

        st.divider()

        st.subheader(
            "✏️ 修改知識"
        )

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
                        if "問題簡述"
                        in df.columns
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
                            selected_row[
                                column
                            ]
                        ),
                        disabled=True,
                        key=f"edit_knowledge_{column}"
                    )

                else:

                    edit_data[column] = st.text_area(
                        column,
                        value=str(
                            selected_row[
                                column
                            ]
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

                save_knowledge_data(
                    df
                )

                st.success(
                    "資料已更新！"
                )

                st.rerun()


    # ========================================================
    # 內部公告維護
    # ========================================================

    elif maintenance_type == "📢 內部公告":

        st.subheader(
            "📋 公告資料"
        )

        if announcement_df.empty:

            st.info(
                "目前沒有公告資料。"
            )

        else:

            st.dataframe(
                announcement_df,
                use_container_width=True,
                height=300
            )

        st.divider()


        # ====================================================
        # 新增公告
        # ====================================================

        st.subheader(
            "➕ 新增公告"
        )

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
            [
                "顯示",
                "隱藏"
            ],
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

                        new_id = int(
                            numbers.max()
                        ) + 1

                    else:

                        new_id = 1

                else:

                    new_id = 1

                new_announcement = pd.DataFrame(
                    [
                        {
                            "公告編號": new_id,
                            "公告標題": announcement_title,
                            "公告內容": clean_announcement_content(
                                announcement_content
                            ),
                            "公告日期": announcement_date,
                            "活動期限": activity_period,
                            "點數期限": point_period,
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

                st.success(
                    "公告已新增！"
                )

                st.rerun()


        st.divider()


        # ====================================================
        # 修改公告
        # ====================================================

        st.subheader(
            "✏️ 修改公告"
        )

        if not announcement_df.empty:

            selected_announcement = st.selectbox(
                "選擇要修改的公告",

                options=announcement_df.index,

                format_func=lambda x:
                    f"{x + 1}｜"
                    + str(
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
                value=str(
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
                value=str(
                    selected_row.get(
                        "公告日期",
                        ""
                    )
                ),
                key="edit_announcement_date"
            )

            edit_activity = st.text_input(
                "活動期限",
                value=str(
                    selected_row.get(
                        "活動期限",
                        ""
                    )
                ),
                key="edit_activity"
            )

            edit_point = st.text_input(
                "點數期限",
                value=str(
                    selected_row.get(
                        "點數期限",
                        ""
                    )
                ),
                key="edit_point"
            )

            edit_status = st.selectbox(
                "顯示狀態",
                [
                    "顯示",
                    "隱藏"
                ],
                index=(
                    0
                    if str(
                        selected_row.get(
                            "顯示狀態",
                            "顯示"
                        )
                    ).strip()
                    == "顯示"
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
                ] = edit_title

                announcement_df.loc[
                    selected_announcement,
                    "公告內容"
                ] = clean_announcement_content(
                    edit_content
                )

                announcement_df.loc[
                    selected_announcement,
                    "公告日期"
                ] = edit_date

                announcement_df.loc[
                    selected_announcement,
                    "活動期限"
                ] = edit_activity

                announcement_df.loc[
                    selected_announcement,
                    "點數期限"
                ] = edit_point

                announcement_df.loc[
                    selected_announcement,
                    "顯示狀態"
                ] = edit_status

                save_sheet(
                    "內部公告",
                    announcement_df
                )

                st.success(
                    "公告已更新！"
                )

                st.rerun()


    # ========================================================
    # 常用工具維護
    # ========================================================

    elif maintenance_type == "🔗 常用工具":

        st.subheader(
            "📋 工具資料"
        )

        if tool_df.empty:

            st.info(
                "目前沒有工具資料。"
            )

        else:

            st.dataframe(
                tool_df,
                use_container_width=True,
                height=300
            )

        st.divider()


        # ====================================================
        # 新增工具
        # ====================================================

        st.subheader(
            "➕ 新增工具"
        )

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
            [
                "顯示",
                "隱藏"
            ],
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

                st.success(
                    "工具已新增！"
                )

                st.rerun()


        st.divider()


        # ====================================================
        # 修改工具
        # ====================================================

        st.subheader(
            "✏️ 修改工具"
        )

        if not tool_df.empty:

            selected_tool = st.selectbox(
                "選擇要修改的工具",

                options=tool_df.index,

                format_func=lambda x:
                    f"{x + 1}｜"
                    + str(
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
                [
                    "顯示",
                    "隱藏"
                ],
                index=(
                    0
                    if str(
                        selected_row.get(
                            "顯示狀態",
                            "顯示"
                        )
                    ).strip()
                    == "顯示"
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
                value=int(
                    current_order
                ),
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

                st.success(
                    "工具已更新！"
                )

                st.rerun()
