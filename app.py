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


# ============================================================
# 基本設定
# ============================================================

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


# ============================================================
# 日系貓咪 UI
# ============================================================

try:

    st.html("""
    <style>

    .stApp {
        background: #F7F4EE;
        color: #5F5A53;
    }

    [data-testid="stHeader"] {
        background: transparent !important;
    }

    .block-container {
        padding-top: 3.5rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    [data-testid="stTextInput"] input {
        height: 52px !important;
        font-size: 17px !important;
        padding: 0 18px !important;
        border: 2px solid #D8C9B8 !important;
        border-radius: 14px !important;
        background: #FFFDF9 !important;
    }

    [data-testid="stTextInput"] input:focus {
        border-color: #B9785E !important;
        box-shadow: 0 0 0 3px rgba(185,120,94,.12) !important;
    }

    [data-testid="stTextInput"] input::placeholder {
        color: #A9A096 !important;
    }

    h1, h2, h3, h4 {
        color: #514C45 !important;
    }

    p, label {
        color: #706A62;
    }

    [data-testid="stSidebar"] {
        background: #EEE8DD;
        border-right: 1px solid #DED4C5;
    }

    [data-testid="stSidebar"] > div:first-child {
        padding-top: 1.4rem;
    }

    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #554F48 !important;
    }

    .stButton > button {
        border-radius: 12px;
        border: 1px solid #D8C9B8;
        background: #FFFDF9;
        color: #5F5A53;
    }

    .stButton > button:hover {
        border-color: #B9785E;
        color: #8A5E4C;
    }

    input,
    textarea,
    [data-baseweb="select"] > div {
        border-radius: 10px !important;
    }

    /* =========================================================
       🐱 浮動客服小幫手
       ========================================================= */

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

        background: #FFFDF9 !important;
        border: 1px solid #D8C9B8 !important;

        box-shadow: 0 4px 15px rgba(90,70,50,.18) !important;

        font-size: 27px !important;
    }

    [data-testid="stPopover"] > button:hover {
        transform: scale(1.08);
        border-color: #B9785E !important;
    }

    [data-testid="stPopover"] > button:focus {
        outline: none !important;
        box-shadow: 0 4px 15px rgba(90,70,50,.18) !important;
    }

    </style>
    """)

except Exception:
    pass


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

    sheet = find_sheet(
        ["內部公告", "公告"]
    )

    if not sheet:
        return pd.DataFrame()

    df = pd.read_excel(
        EXCEL_PATH,
        sheet_name=sheet,
        dtype=str
    ).fillna("")

    for col in df.columns:

        if any(
            x in str(col)
            for x in ["內容", "說明", "公告"]
        ):

            df[col] = df[col].map(
                clean_announcement_content
            )

    for col in df.columns:

        if "標題" in str(col):

            df[col] = df[col].map(
                clean_announcement_title
            )

    return df


@st.cache_data(ttl=60)
def load_tools():

    sheet = find_sheet(
        ["常用工具", "工具"]
    )

    if not sheet:
        return pd.DataFrame()

    return pd.read_excel(
        EXCEL_PATH,
        sheet_name=sheet,
        dtype=str
    ).fillna("")


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

        # 確保欄位完整
        for col in PENDING_COLUMNS:

            if col not in df.columns:
                df[col] = ""

        # 只保留正確順序
        df = df[
            PENDING_COLUMNS
        ].copy()

        return df

    except Exception as e:

        st.error(
            f"讀取待審核問題失敗：{e}"
        )

        return pd.DataFrame(
            columns=PENDING_COLUMNS
        )


# ============================================================
# 儲存 Excel
# ============================================================

def save_sheet(sheet_name, data):

    if not EXCEL_PATH.exists():

        st.error(
            "找不到 Excel：DATA/知識庫實驗品.xlsx"
        )

        return False

    try:

        wb = openpyxl.load_workbook(
            EXCEL_PATH
        )

        # 如果工作表已存在，刪除後重新建立
        if sheet_name in wb.sheetnames:

            ws_old = wb[sheet_name]

            wb.remove(
                ws_old
            )

        ws = wb.create_sheet(
            sheet_name
        )

        # 寫入欄位名稱
        for c, col in enumerate(
            data.columns,
            1
        ):

            ws.cell(
                1,
                c,
                str(col)
            )

        # 寫入資料
        for r, row in enumerate(
            data.itertuples(index=False),
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

        # 清除 Streamlit 快取
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


def save_knowledge_data(data):

    sheet = find_sheet(
        ["知識庫", "問題集", "Sheet1"],
        0
    )

    if sheet:

        return save_sheet(
            sheet,
            data
        )

    return False


# ============================================================
# ⭐ 新增待審核問題
# ============================================================

def add_pending_question(question):

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

        # ----------------------------------------------------
        # 開啟 Excel
        # ----------------------------------------------------

        wb = openpyxl.load_workbook(
            EXCEL_PATH
        )

        sheet_name = "待審核問題"

        # ----------------------------------------------------
        # 如果沒有工作表 → 自動建立
        # ----------------------------------------------------

        if sheet_name not in wb.sheetnames:

            ws = wb.create_sheet(
                sheet_name
            )

            headers = PENDING_COLUMNS

            for c, header in enumerate(
                headers,
                1
            ):

                ws.cell(
                    1,
                    c,
                    header
                )

        else:

            ws = wb[sheet_name]

            # ------------------------------------------------
            # 確認第一列是否有標題
            # ------------------------------------------------

            existing_headers = [
                ws.cell(
                    1,
                    c
                ).value
                for c in range(
                    1,
                    max(
                        ws.max_column,
                        len(PENDING_COLUMNS)
                    ) + 1
                )
            ]

            existing_headers = [
                clean_text(x)
                for x in existing_headers
            ]

            # 如果標題不完整，補齊
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

        # ----------------------------------------------------
        # 檢查是否重複
        # ----------------------------------------------------

        existing_questions = []

        if ws.max_row >= 2:

            # 找「問題」欄位位置
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

            # 找不到就預設第 2 欄
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

        # ----------------------------------------------------
        # 新增一筆
        # ----------------------------------------------------

        new_row = [
            datetime.now().strftime(
                "%Y/%m/%d %H:%M"
            ),
            question,
            "",
            "待審核",
            ""
        ]

        ws.append(
            new_row
        )

        # ----------------------------------------------------
        # 儲存
        # ----------------------------------------------------

        wb.save(
            EXCEL_PATH
        )

        # 確認真的寫進去了
        wb_check = openpyxl.load_workbook(
            EXCEL_PATH,
            read_only=True,
            data_only=True
        )

        if sheet_name not in wb_check.sheetnames:

            wb_check.close()

            st.error(
                "❌ Excel 儲存後找不到「待審核問題」工作表。"
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
                and clean_text(row[1]) == question
            ):

                saved = True
                break

        wb_check.close()

        # 清除所有 cache
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

    df["_月份"] = df["_日期"].dt.strftime(
        "%Y/%m"
    )

    df["_月份"] = df["_月份"].fillna(
        df["月份"].apply(clean_text)
    )

    return df


# ============================================================
# 搜尋詞擴充
# ============================================================

def expand_search_word(keyword):

    keyword = clean_text(
        keyword
    )

    if not keyword:
        return ""

    synonym_df = load_synonyms()

    words = [
        keyword
    ]

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
                and keyword_normalized in left_normalized
            ):

                words.extend(
                    re.split(
                        r"[,，、/／\s]+",
                        right
                    )
                )

            elif (
                right_normalized
                and keyword_normalized in right_normalized
            ):

                words.append(
                    left
                )

    return " ".join(
        dict.fromkeys(
            [
                x
                for x in words
                if clean_text(x)
            ]
        )
    )


# ============================================================
# 搜尋 V2
# ============================================================

def split_search_terms(keyword):

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
        "嗎？",
        "嗎?",
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
        "一下子",
        "問題",
        "請",
        "幫我",
        "告訴我",
        "想知道",
        "請幫忙",
        "謝謝",
        "謝謝您",
        "謝謝妳"
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
        dict.fromkeys(
            terms
        )
    )


def build_search_terms(keyword):

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
        dict.fromkeys(
            [
                x
                for x in expanded_terms
                if x
            ]
        )
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

        if positions[-1] != max_start - 1:

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
        c
        for c in fields
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

    # --------------------------------------------------------
    # TF-IDF
    # --------------------------------------------------------

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
        corpus + [
            query_for_tfidf
        ]
    )

    scores = cosine_similarity(
        matrix[-1],
        matrix[:-1]
    ).ravel()

    result = df.copy()

    result["_score"] = scores

    # --------------------------------------------------------
    # 多關鍵字
    # --------------------------------------------------------

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

                # 完全命中
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

                # 模糊命中
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

            fuzzy_total += (
                best_term_score
            )

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
                coverage + multi_exact_bonus
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
                    len(meaningful_terms)
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

    # --------------------------------------------------------
    # 完整句子命中
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # 最終分數
    # --------------------------------------------------------

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
        result["_final_score"].clip(
            upper=1.0
        )
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
# 客服小幫手搜尋
# ============================================================

def ask_customer_assistant(question):

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
# 客訴統計
# ============================================================

# ============================================================
# 側邊欄
# ============================================================

with st.sidebar:

    st.markdown(
        "## 🐱"
    )

    st.markdown(
        "### 客服作業系統"
    )

    st.caption(
        "NITORI CUSTOMER SERVICE"
    )

    st.divider()

    page = st.radio(
        "功能選單",
        [
            "🏠 首頁",
            "🔍 知識搜尋",
            "📊 客訴資料統計",
            "🗺️ 門市地圖",
            "📢 公告＆常用工具",
            "🛠️ 資料維護"
        ],
        label_visibility="collapsed"
    )

    st.divider()

    st.caption(
        "🌿 讓小貓陪你一起找資料"
    )


# ============================================================
# 首頁
# ============================================================

if page == "🏠 首頁":

    st.title(
        "🐱 客服作業系統"
    )

    st.caption(
        "NITORI CUSTOMER SERVICE WORKSPACE"
    )

    st.info(
        "🌿 今天也辛苦了，讓小貓陪你一起找資料"
    )

    st.subheader(
        "QUICK ACCESS"
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        st.subheader(
            "🔍 知識庫搜尋"
        )

        st.write(
            "輸入顧客問題，快速找出相關作業資訊與罐頭回覆。"
        )

    with c2:

        st.subheader(
            "🗺️ 宜得利小地圖"
        )

        st.write(
            "查詢門市地址、交通、停車及門市相關資訊。"
        )

    with c3:

        st.subheader(
            "📢 公告＆常用工具"
        )

        st.write(
            "查看最新內部公告與工作中常用的連結工具。"
        )

    st.divider()

    st.subheader(
        "📢 LATEST ANNOUNCEMENTS"
    )

    ann = load_announcements()

    if ann.empty:

        st.info(
            "目前沒有公告資料。"
        )

    else:

        cols = list(
            ann.columns
        )

        title_col = next(
            (
                c
                for c in cols
                if "標題" in str(c)
            ),
            cols[0]
        )

        content_col = next(
            (
                c
                for c in cols
                if any(
                    x in str(c)
                    for x in [
                        "內容",
                        "說明"
                    ]
                )
            ),
            cols[1]
            if len(cols) > 1
            else cols[0]
        )

        date_col = next(
            (
                c
                for c in cols
                if "日期" in str(c)
            ),
            None
        )

        for _, row in ann.head(5).iterrows():

            title = clean_announcement_title(
                row[title_col]
            )

            content = clean_announcement_content(
                row[content_col]
            )

            date = (
                clean_text(
                    row[date_col]
                )
                if date_col
                else ""
            )

            with st.container(
                border=True
            ):

                st.subheader(
                    "📌 " + title
                )

                if date:

                    st.caption(
                        "📅 公告日期：" + date
                    )

                if content:

                    st.write(
                        content
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
                filtered[
                    "客訴類型"
                ].astype(str).str.contains(
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

            monthly_data = (
                monthly_data.sort_values(
                    "_月份"
                )
            )

            monthly_chart = (
                monthly_data.set_index(
                    "_月份"
                )
            )

            st.line_chart(
                monthly_chart[
                    "客訴件數"
                ],
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
                    filtered[
                        "客訴類型"
                    ].astype(str).str.strip()
                    != ""
                ]
                .groupby(
                    "客訴類型"
                )
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
                    filtered[
                        "客訴細項"
                    ].astype(str).str.strip()
                    != ""
                ]
                .groupby(
                    "客訴細項"
                )
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
                filtered[
                    "購買通路"
                ].astype(str).str.strip()
                != ""
            ]
            .groupby(
                "購買通路"
            )
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

        display_df = filtered[
            display_columns
        ].copy()

        st.dataframe(
            display_df,
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

    components.iframe(
        POWER_BI_URL,
        height=760,
        scrolling=True
    )


# ============================================================
# 公告＆常用工具
# ============================================================

elif page == "📢 公告＆常用工具":

    st.title(
        "📢 公告＆常用工具"
    )

    tab1, tab2 = st.tabs(
        [
            "📢 內部公告",
            "🔗 常用工具"
        ]
    )

    with tab1:

        ann = load_announcements()

        if ann.empty:

            st.info(
                "目前沒有公告資料。"
            )

        else:

            cols = list(
                ann.columns
            )

            title_col = next(
                (
                    c
                    for c in cols
                    if "標題" in str(c)
                ),
                cols[0]
            )

            content_col = next(
                (
                    c
                    for c in cols
                    if any(
                        x in str(c)
                        for x in [
                            "內容",
                            "說明"
                        ]
                    )
                ),
                cols[1]
                if len(cols) > 1
                else cols[0]
            )

            date_col = next(
                (
                    c
                    for c in cols
                    if "日期" in str(c)
                ),
                None
            )

            for _, row in ann.iterrows():

                title = clean_announcement_title(
                    row[title_col]
                )

                content = clean_announcement_content(
                    row[content_col]
                )

                date = (
                    clean_text(
                        row[date_col]
                    )
                    if date_col
                    else ""
                )

                with st.container(
                    border=True
                ):

                    st.subheader(
                        "📌 " + title
                    )

                    if date:

                        st.caption(
                            "📅 " + date
                        )

                    st.write(
                        content
                    )

    with tab2:

        tools = load_tools()

        if tools.empty:

            st.info(
                "目前沒有常用工具資料。"
            )

        else:

            cols = list(
                tools.columns
            )

            name_col = next(
                (
                    c
                    for c in cols
                    if any(
                        x in str(c)
                        for x in [
                            "名稱",
                            "工具",
                            "標題"
                        ]
                    )
                ),
                cols[0]
            )

            url_col = next(
                (
                    c
                    for c in cols
                    if any(
                        x in str(c)
                        for x in [
                            "網址",
                            "連結",
                            "URL",
                            "url"
                        ]
                    )
                ),
                None
            )

            desc_col = next(
                (
                    c
                    for c in cols
                    if any(
                        x in str(c)
                        for x in [
                            "說明",
                            "描述",
                            "備註"
                        ]
                    )
                ),
                None
            )

            grid = st.columns(3)

            for i, (_, row) in enumerate(
                tools.iterrows()
            ):

                name = clean_text(
                    row[name_col]
                )

                url = (
                    clean_text(
                        row[url_col]
                    )
                    if url_col
                    else ""
                )

                desc = (
                    clean_text(
                        row[desc_col]
                    )
                    if desc_col
                    else ""
                )

                with grid[
                    i % 3
                ]:

                    with st.container(
                        border=True
                    ):

                        st.subheader(
                            "🔗 " + name
                        )

                        if desc:

                            st.write(
                                desc
                            )

                        if url:

                            st.link_button(
                                "開啟工具",
                                url,
                                use_container_width=True
                            )


# ============================================================
# 資料維護
# ============================================================

elif page == "🛠️ 資料維護":

    st.title(
        "🛠️ 資料維護"
    )

    st.caption(
        "這裡可以直接編輯目前 Excel 中的資料，完成後按下儲存。"
    )

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "📚 知識庫",
            "📢 公告",
            "🔗 常用工具",
            "📝 待審核問題",
            "📊 客訴資料"
        ]
    )

    # ========================================================
    # 知識庫
    # ========================================================

    with tab1:

        df = load_data()

        if df.empty:

            st.warning(
                "找不到知識庫資料。"
            )

        else:

            edited = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                height=600,
                key="knowledge_editor"
            )

            if st.button(
                "💾 儲存知識庫",
                type="primary"
            ):

                if save_knowledge_data(
                    edited
                ):

                    st.success(
                        "知識庫已儲存！"
                    )

                    st.rerun()

    # ========================================================
    # 公告
    # ========================================================

    with tab2:

        df = load_announcements()

        if df.empty:

            st.warning(
                "找不到公告資料。"
            )

        else:

            edited = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                height=550,
                key="announcement_editor"
            )

            if st.button(
                "💾 儲存公告"
            ):

                sheet = find_sheet(
                    ["內部公告", "公告"]
                )

                if (
                    sheet
                    and save_sheet(
                        sheet,
                        edited
                    )
                ):

                    st.success(
                        "公告已儲存！"
                    )

                    st.rerun()

    # ========================================================
    # 常用工具
    # ========================================================

    with tab3:

        df = load_tools()

        if df.empty:

            st.warning(
                "找不到常用工具資料。"
            )

        else:

            edited = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                height=550,
                key="tools_editor"
            )

            if st.button(
                "💾 儲存常用工具"
            ):

                sheet = find_sheet(
                    ["常用工具", "工具"]
                )

                if (
                    sheet
                    and save_sheet(
                        sheet,
                        edited
                    )
                ):

                    st.success(
                        "常用工具已儲存！"
                    )

                    st.rerun()

    # ========================================================
    # 待審核問題
    # ========================================================

    with tab4:

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

            edited_pending = st.data_editor(
                pending,
                use_container_width=True,
                num_rows="dynamic",
                height=550,
                key="pending_editor"
            )

            if st.button(
                "💾 儲存審核結果",
                type="primary",
                key="save_pending_questions"
            ):

                if save_sheet(
                    "待審核問題",
                    edited_pending
                ):

                    st.success(
                        "待審核問題已更新！"
                    )

                    st.rerun()

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

            ⑥ 小幫手下次就可以找到
            """
        )

    # ========================================================
    # 客訴資料
    # ========================================================

    with tab5:

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
                    ] = str(
                        i + 1
                    )

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
                ),

                "購買通路": st.column_config.TextColumn(
                    "購買通路"
                ),

                "商品編號": st.column_config.TextColumn(
                    "商品編號"
                ),

                "商品名稱": st.column_config.TextColumn(
                    "商品名稱"
                ),

                "客訴類型": st.column_config.TextColumn(
                    "客訴類型"
                ),

                "客訴細項": st.column_config.TextColumn(
                    "客訴細項"
                ),

                "客訴內容": st.column_config.TextColumn(
                    "客訴內容"
                ),

                "處理方式": st.column_config.TextColumn(
                    "處理方式"
                ),

                "處理結果": st.column_config.TextColumn(
                    "處理結果"
                ),

                "備註": st.column_config.TextColumn(
                    "備註"
                )
            }
        )

        st.caption(
            "可以直接在表格內新增、修改或刪除資料。"
        )

        if st.button(
            "💾 儲存客訴資料",
            type="primary",
            key="save_complaints"
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
                            clean_text(
                                value
                            )
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

                        parsed_date = (
                            pd.to_datetime(
                                date_value,
                                errors="coerce"
                            )
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
                save_df
            ):

                st.success(
                    "客訴資料已儲存！"
                )

                st.rerun()


# ============================================================
# 🐱 浮動客服小幫手
# ============================================================

# ------------------------------------------------------------
# 初始化小幫手狀態
# ------------------------------------------------------------

if "assistant_searched" not in st.session_state:
    st.session_state["assistant_searched"] = False

if "assistant_found" not in st.session_state:
    st.session_state["assistant_found"] = False

if "assistant_result" not in st.session_state:
    st.session_state["assistant_result"] = pd.DataFrame()

if "assistant_question_saved" not in st.session_state:
    st.session_state["assistant_question_saved"] = ""

if "assistant_submit_message" not in st.session_state:
    st.session_state["assistant_submit_message"] = ""


# ============================================================
# 浮動小幫手
# ============================================================

with st.popover("🐱"):

    st.subheader(
        "🐱 客服小幫手"
    )

    st.caption(
        "先從知識庫幫你找答案"
    )

    # --------------------------------------------------------
    # 顧客問題輸入
    # --------------------------------------------------------

    assistant_question = st.text_input(
        "顧客問題",
        placeholder="例如：床墊有味道怎麼辦？",
        key="floating_assistant_question"
    )

    # --------------------------------------------------------
    # 搜尋按鈕
    # --------------------------------------------------------

    assistant_search = st.button(
        "🔍 幫我找答案",
        use_container_width=True,
        type="primary",
        key="floating_assistant_search"
    )

    # --------------------------------------------------------
    # 按下搜尋
    # --------------------------------------------------------

    if assistant_search:

        question = clean_text(
            assistant_question
        )

        if not question:

            st.session_state["assistant_searched"] = False
            st.session_state["assistant_found"] = False
            st.session_state["assistant_result"] = pd.DataFrame()

            st.warning(
                "請先輸入顧客問題。"
            )

        else:

            found, result = ask_customer_assistant(
                question
            )

            # ------------------------------------------------
            # 把搜尋結果記住
            # ------------------------------------------------

            st.session_state["assistant_searched"] = True

            st.session_state["assistant_found"] = (
                found and not result.empty
            )

            st.session_state["assistant_result"] = result

            st.session_state["assistant_question"] = question

            # 每次重新搜尋都清除上次提交狀態
            st.session_state["assistant_question_saved"] = ""

            st.session_state["assistant_submit_message"] = ""


    # ========================================================
    # 顯示搜尋結果
    # ========================================================

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

        question_to_submit = st.session_state.get(
            "assistant_question",
            ""
        )

        # ====================================================
        # 找到答案
        # ====================================================

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

            # ------------------------------------------------
            # 其他相關資料
            # ------------------------------------------------

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


        # ====================================================
        # 沒找到答案
        # ====================================================

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

            # ------------------------------------------------
            # 已經成功提交過
            # ------------------------------------------------

            saved_question = st.session_state.get(
                "assistant_question_saved",
                ""
            )

            if (
                saved_question
                and saved_question == question_to_submit
            ):

                st.success(
                    "🎉 已成功加入待審核問題！"
                )

                st.info(
                    "請到「🛠️ 資料維護 → 📝 待審核問題」查看。"
                )

            # ------------------------------------------------
            # 尚未提交
            # ------------------------------------------------

            else:

                submit_pending = st.button(
                    "📝 提出這個問題給我審核",
                    type="primary",
                    use_container_width=True,
                    key="floating_submit_pending"
                )

                if submit_pending:

                    # ----------------------------------------
                    # 先確認問題不是空的
                    # ----------------------------------------

                    question_to_submit = clean_text(
                        question_to_submit
                    )

                    if not question_to_submit:

                        st.error(
                            "❌ 沒有取得要提交的問題，請重新搜尋一次。"
                        )

                    else:

                        st.info(
                            "🔄 正在將問題加入待審核清單……"
                        )

                        # ------------------------------------
                        # 實際寫入 Excel
                        # ------------------------------------

                        success = add_pending_question(
                            question_to_submit
                        )

                        # ------------------------------------
                        # 寫入成功
                        # ------------------------------------

                        if success:

                            # 記住這個問題已經提交
                            st.session_state[
                                "assistant_question_saved"
                            ] = question_to_submit

                            # 清除 cache
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

                        # ------------------------------------
                        # 寫入失敗
                        # ------------------------------------

                        else:

                            st.error(
                                "❌ 提交失敗。"
                            )

                            st.warning(
                                "請查看上方顯示的錯誤訊息，確認 Excel 是否正在開啟。"
                            )
