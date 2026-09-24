import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from datetime import datetime
from pathlib import Path
import openpyxl
import re
import html

# ============================================================
# 基本設定
# ============================================================
st.set_page_config(page_title="客服作業系統", page_icon="🐱", layout="wide")

EXCEL_PATH = Path("DATA/知識庫實驗品.xlsx")
POWER_BI_URL = "https://app.powerbi.com/reportEmbed?reportId=851d09f1-f169-4049-9090-640f09cdc385&autoAuth=true&ctid=94d67acc-8c17-4d3b-8463-9829cf08ef55"

# ============================================================
# 日系貓咪 UI
# ============================================================
st.markdown("""
<style>
.stApp { background:#F7F4EE; color:#5F5A53; }
.block-container { padding-top:1.8rem; padding-bottom:3rem; max-width:1400px; }
h1,h2,h3,h4 { color:#514C45 !important; }
p, label { color:#706A62; }
[data-testid="stSidebar"] { background:#EEE8DD; border-right:1px solid #DED4C5; }
[data-testid="stSidebar"] > div:first-child { padding-top:1.4rem; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,[data-testid="stSidebar"] h3 { color:#554F48 !important; }
.stButton > button { border-radius:12px; border:1px solid #D8C9B8; background:#FFFDF9; color:#5F5A53; }
.stButton > button:hover { border-color:#B9785E; color:#8A5E4C; }
input, textarea, [data-baseweb="select"] > div { border-radius:10px !important; }
.hero { background:#FFFDF9; border:1px solid #E2D8CC; border-radius:24px; padding:35px 30px 30px; text-align:center; box-shadow:0 8px 25px rgba(90,70,50,.06); margin-bottom:28px; }
.hero-cat { font-size:54px; line-height:1; margin-bottom:12px; }
.hero-title { font-size:34px; font-weight:750; letter-spacing:.06em; color:#514C45; }
.hero-subtitle { margin-top:7px; font-size:12px; letter-spacing:.16em; color:#9B8F82; }
.hero-message { display:inline-block; margin-top:18px; padding:9px 17px; border-radius:999px; background:#F3E6D9; color:#8A5E4C; font-size:14px; }
.card { background:#FFFDF9; border:1px solid #E2D8CC; border-radius:17px; padding:19px; margin-bottom:15px; box-shadow:0 5px 18px rgba(90,70,50,.045); }
.card-title { font-weight:700; color:#62584F; font-size:17px; margin-bottom:8px; }
.card-meta { font-size:12px; color:#A19588; margin-bottom:8px; }
.card-text { color:#706A62; line-height:1.75; white-space:pre-wrap; }
.section-label { color:#8A5E4C; font-size:13px; font-weight:700; letter-spacing:.08em; margin:22px 0 9px; }
.tool-card { min-height:125px; background:#FFFDF9; border:1px solid #E2D8CC; border-radius:17px; padding:20px; }
.tool-icon { font-size:28px; margin-bottom:7px; }
.tool-name { font-size:17px; font-weight:700; color:#62584F; }
.tool-desc { margin-top:7px; color:#8B8279; line-height:1.6; font-size:13px; }
.result-card { background:#FFFDF9; border:1px solid #E1D7CB; border-radius:15px; padding:19px; margin:12px 0; }
.result-title { font-size:18px; font-weight:700; color:#5B554D; }
.result-meta { margin-top:5px; font-size:12px; color:#A19588; }
.result-answer { margin-top:12px; line-height:1.75; white-space:pre-wrap; color:#665F57; }
.score { color:#A56B50; font-size:12px; }
hr { border-color:#E5DDD3 !important; }
[data-testid="stDataFrame"] { border:1px solid #DED5CA; border-radius:12px; overflow:hidden; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 共用工具
# ============================================================
def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_announcement_content(text):
    """將公告中的 HTML / Markdown / SVG / 多餘標籤整理成純文字。"""
    if text is None or pd.isna(text):
        return ""
    text = html.unescape(str(text))
    text = re.sub(r"```(?:html|HTML|text|TEXT)?", "", text)
    text = text.replace("```", "")
    text = re.sub(r"<svg[\s\S]*?</svg>", "", text, flags=re.I)
    text = re.sub(r"\bsvg\b", "", text, flags=re.I)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"</(?:p|div|h[1-6]|li|tr)>", "\n", text, flags=re.I)
    text = re.sub(r"<li[^>]*>", "• ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\b(?:id|class)\s*=\s*['\"][^'\"]*['\"]", "", text, flags=re.I)
    text = text.replace("\\~", "~").replace("\\", "")
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def clean_announcement_title(text):
    text = clean_announcement_content(text)
    text = text.replace("📌", "").strip()
    if text.startswith("【") and text.endswith("】"):
        text = text[1:-1].strip()
    return text


def safe_html_text(value):
    return html.escape(clean_text(value)).replace("\n", "<br>")


# ============================================================
# Excel 載入
# ============================================================
@st.cache_data(ttl=60)
def get_excel_sheets():
    if not EXCEL_PATH.exists():
        return []
    return pd.ExcelFile(EXCEL_PATH).sheet_names


def find_sheet(candidates, fallback_index=None):
    sheets = get_excel_sheets()
    for name in candidates:
        if name in sheets:
            return name
    if fallback_index is not None and len(sheets) > fallback_index:
        return sheets[fallback_index]
    return None


@st.cache_data(ttl=60)
def load_data():
    sheet = find_sheet(["知識庫", "問題集", "Sheet1"], 0)
    if not sheet:
        return pd.DataFrame()
    return pd.read_excel(EXCEL_PATH, sheet_name=sheet).fillna("")


@st.cache_data(ttl=60)
def load_synonyms():
    sheet = find_sheet(["搜尋詞對照", "關鍵字對照"])
    if not sheet:
        return pd.DataFrame()
    return pd.read_excel(EXCEL_PATH, sheet_name=sheet).fillna("")


@st.cache_data(ttl=60)
def load_announcements():
    sheet = find_sheet(["內部公告", "公告"])
    if not sheet:
        return pd.DataFrame()
    df = pd.read_excel(EXCEL_PATH, sheet_name=sheet).fillna("")
    for col in df.columns:
        if any(x in str(col) for x in ["內容", "說明", "公告"]):
            df[col] = df[col].map(clean_announcement_content)
    for col in df.columns:
        if "標題" in str(col):
            df[col] = df[col].map(clean_announcement_title)
    return df


@st.cache_data(ttl=60)
def load_tools():
    sheet = find_sheet(["常用工具", "工具"])
    if not sheet:
        return pd.DataFrame()
    return pd.read_excel(EXCEL_PATH, sheet_name=sheet).fillna("")


# ============================================================
# 儲存
# ============================================================
def save_sheet(sheet_name, data):
    if not EXCEL_PATH.exists():
        st.error("找不到 Excel：DATA/知識庫實驗品.xlsx")
        return False
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            wb.remove(ws)
        ws = wb.create_sheet(sheet_name)
        for c, col in enumerate(data.columns, 1):
            ws.cell(1, c, col)
        for r, row in enumerate(data.itertuples(index=False), 2):
            for c, value in enumerate(row, 1):
                ws.cell(r, c, value)
        wb.save(EXCEL_PATH)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"儲存失敗：{e}")
        return False


def save_knowledge_data(data):
    sheet = find_sheet(["知識庫", "問題集", "Sheet1"], 0)
    if sheet:
        return save_sheet(sheet, data)
    return False


# ============================================================
# 搜尋
# ============================================================

def normalize_search_text(value):
    """統一搜尋文字格式，避免大小寫、空白、符號影響搜尋。"""
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()

    # 移除多餘空白
    text = re.sub(r"\s+", "", text)

    # 統一常見符號
    text = text.replace("台", "臺")
    text = text.replace("－", "-")
    text = text.replace("—", "-")
    text = text.replace("–", "-")

    return text


def expand_search_word(keyword):
    """搜尋同義詞／關鍵字對照。"""
    keyword = clean_text(keyword)

    if not keyword:
        return ""

    synonym_df = load_synonyms()

    words = [keyword]

    if not synonym_df.empty and len(synonym_df.columns) >= 2:
        a, b = synonym_df.columns[:2]

        for _, row in synonym_df.iterrows():
            left = clean_text(row[a])
            right = clean_text(row[b])

            if not left and not right:
                continue

            # 關鍵字命中左側
            if keyword in left:
                words.append(right)

            # 關鍵字命中右側
            elif keyword in right:
                words.append(left)

    # 拆開逗號、頓號、斜線等
    expanded = []

    for word in words:
        if not word:
            continue

        parts = re.split(r"[,，、/／\s]+", word)

        for part in parts:
            part = clean_text(part)

            if part:
                expanded.append(part)

    return " ".join(dict.fromkeys(expanded))


def search_knowledge(keyword, top_n=10):
    """
    知識庫搜尋：
    1. 先找真正包含關鍵字的資料
    2. 再用 TF-IDF 模糊搜尋補充
    3. 關鍵字直接命中的結果優先
    """

    df = load_data()

    if df.empty:
        return pd.DataFrame()

    keyword = clean_text(keyword)

    if not keyword:
        return df.head(top_n).copy()

    # --------------------------------------------------------
    # 可搜尋欄位
    # --------------------------------------------------------
    fields = [
        "問題簡述",
        "解答",
        "判斷方式",
        "判斷使用的工具",
        "關鍵字",
        "大分類",
        "中分類",
        "小分類"
    ]

    fields = [c for c in fields if c in df.columns]

    if not fields:
        return pd.DataFrame()

    # --------------------------------------------------------
    # 建立搜尋文字
    # --------------------------------------------------------
    search_df = df.copy()

    search_df["_search_text"] = (
        search_df[fields]
        .fillna("")
        .astype(str)
        .agg(" ".join, axis=1)
        .map(normalize_search_text)
    )

    # --------------------------------------------------------
    # 展開同義詞
    # --------------------------------------------------------
    expanded_query = expand_search_word(keyword)

    query_words = [
        normalize_search_text(x)
        for x in re.split(r"\s+", expanded_query)
        if normalize_search_text(x)
    ]

    # --------------------------------------------------------
    # 第一層：直接關鍵字命中
    # --------------------------------------------------------
    direct_scores = []

    normalized_keyword = normalize_search_text(keyword)

    for text in search_df["_search_text"]:

        score = 0

        # 完整關鍵字命中
        if normalized_keyword and normalized_keyword in text:
            score += 100

        # 同義詞命中
        for word in query_words:
            if word and word in text:
                score += 20

        direct_scores.append(score)

    search_df["_direct_score"] = direct_scores

    # --------------------------------------------------------
    # 第二層：TF-IDF 模糊搜尋
    # --------------------------------------------------------
    corpus = search_df["_search_text"].tolist()

    try:
        vectorizer = TfidfVectorizer(
            analyzer="char",
            ngram_range=(1, 3),
            min_df=1,
            sublinear_tf=True
        )

        matrix = vectorizer.fit_transform(
            corpus + [normalize_search_text(expanded_query)]
        )

        scores = cosine_similarity(
            matrix[-1],
            matrix[:-1]
        ).ravel()

    except Exception:
        scores = np.zeros(len(search_df))

    search_df["_tfidf_score"] = scores

    # --------------------------------------------------------
    # 第三層：綜合分數
    # --------------------------------------------------------
    search_df["_score"] = (
        search_df["_direct_score"]
        + search_df["_tfidf_score"] * 10
    )

    # --------------------------------------------------------
    # 排序
    # --------------------------------------------------------
    result = (
        search_df
        .sort_values(
            ["_direct_score", "_tfidf_score"],
            ascending=False
        )
        .head(top_n)
        .copy()
    )

    # --------------------------------------------------------
    # 清除搜尋用欄位
    # --------------------------------------------------------
    result = result.drop(
        columns=[
            "_search_text",
            "_direct_score",
            "_tfidf_score"
        ],
        errors="ignore"
    )

    # 保留原本的 _score 顯示相關度
    if "_score" in result.columns:
        # 原本的分數可能超過 1，因此轉成 0~1 顯示
        max_score = result["_score"].max()

        if max_score > 0:
            result["_score"] = result["_score"] / max_score
        else:
            result["_score"] = 0

    return result
    df = load_data()
    if df.empty:
        return pd.DataFrame()
    fields = ["問題簡述", "解答", "判斷方式", "關鍵字", "大分類", "中分類", "小分類"]
    fields = [c for c in fields if c in df.columns]
    if not fields:
        return pd.DataFrame()
    corpus = df[fields].fillna("").astype(str).agg(" ".join, axis=1)
    query = expand_search_word(keyword)
    if not query:
        return df.head(top_n).copy()
    vectorizer = TfidfVectorizer(analyzer="char", ngram_range=(1, 3), min_df=1)
    matrix = vectorizer.fit_transform(corpus.tolist() + [query])
    scores = cosine_similarity(matrix[-1], matrix[:-1]).ravel()
    result = df.copy()
    result["_score"] = scores
    result = result.sort_values("_score", ascending=False).head(top_n)
    return result


# ============================================================
# 側邊欄
# ============================================================
with st.sidebar:
    st.markdown("<div style='text-align:center;font-size:42px;'>🐱</div>", unsafe_allow_html=True)
    st.markdown("<h2 style='text-align:center;margin-top:0;'>客服作業系統</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;color:#95897D;'>NITORI CUSTOMER SERVICE</p>", unsafe_allow_html=True)
    st.divider()
    page = st.radio(
        "功能選單",
        ["🏠 首頁", "🔍 知識搜尋", "🗺️ 門市地圖", "📢 公告＆常用工具", "🛠️ 資料維護"],
        label_visibility="collapsed"
    )
    st.divider()
    st.caption("🌿 讓小貓陪你一起找資料")

# ============================================================
# 首頁
# ============================================================
if page == "🏠 首頁":
    st.markdown("""
    <div class="hero">
        <div class="hero-cat">🐱</div>
        <div class="hero-title">客服作業系統</div>
        <div class="hero-subtitle">NITORI CUSTOMER SERVICE WORKSPACE</div>
        <div class="hero-message">🌿 今天也辛苦了，讓小貓陪你一起找資料</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-label">QUICK ACCESS</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('<div class="tool-card"><div class="tool-icon">🔍</div><div class="tool-name">知識庫搜尋</div><div class="tool-desc">輸入顧客問題，快速找出相關作業資訊與罐頭回覆。</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="tool-card"><div class="tool-icon">🗺️</div><div class="tool-name">宜得利小地圖</div><div class="tool-desc">查詢門市地址、交通、停車及門市相關資訊。</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="tool-card"><div class="tool-icon">📢</div><div class="tool-name">公告＆常用工具</div><div class="tool-desc">查看最新內部公告與工作中常用的連結工具。</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label">LATEST ANNOUNCEMENTS</div>', unsafe_allow_html=True)
    ann = load_announcements()
    if ann.empty:
        st.info("目前沒有公告資料。")
    else:
        for _, row in ann.head(5).iterrows():
            cols = list(ann.columns)
            title_col = next((c for c in cols if "標題" in str(c)), cols[0])
            content_col = next((c for c in cols if any(x in str(c) for x in ["內容", "說明"])), cols[1] if len(cols) > 1 else cols[0])
            date_col = next((c for c in cols if "日期" in str(c)), None)
            title = clean_announcement_title(row[title_col])
            content = clean_announcement_content(row[content_col])
            date = clean_text(row[date_col]) if date_col else ""
            st.markdown(f"<div class='card'><div class='card-title'>📌 {safe_html_text(title)}</div><div class='card-text'>{safe_html_text(content)}</div>{('<div class=\"card-meta\">📅 公告日期：'+safe_html_text(date)+'</div>') if date else ''}</div>", unsafe_allow_html=True)

# ============================================================
# 知識搜尋
# ============================================================
elif page == "🔍 知識搜尋":
    st.markdown("# 🔍 知識庫搜尋")
    st.caption("輸入顧客詢問的關鍵字，小貓會幫你找出相關資料。")
    keyword = st.text_input("搜尋關鍵字", placeholder="例如：重量、床包尺寸、退貨、配送……")
    top_n = st.slider("顯示筆數", 5, 30, 10)
    if keyword.strip():
        result = search_knowledge(keyword, top_n)
        if result.empty:
            st.warning("目前沒有找到相關資料。")
        else:
            st.caption(f"共找到 {len(result)} 筆相關資料")
            for _, row in result.iterrows():
                title = clean_text(row.get("問題簡述", "未命名問題"))
                score = float(row.get("_score", 0))
                meta = "　｜　".join([clean_text(row.get(c, "")) for c in ["大分類", "中分類", "小分類"] if clean_text(row.get(c, ""))])
                answer = clean_text(row.get("解答", ""))
                judgment = clean_text(row.get("判斷方式", ""))
                tool = clean_text(row.get("判斷使用的工具", ""))
                extra = ""
                if judgment:
                    extra += f"\n\n【判斷方式】\n{judgment}"
                if tool:
                    extra += f"\n\n【判斷使用的工具】\n{tool}"
                st.markdown(f"<div class='result-card'><div class='result-title'>{safe_html_text(title)}</div><div class='result-meta'>{safe_html_text(meta)}　<span class='score'>相關度 {score:.1%}</span></div><div class='result-answer'>{safe_html_text(answer + extra)}</div></div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='card' style='text-align:center;padding:40px;'>🐾 輸入關鍵字開始搜尋</div>", unsafe_allow_html=True)

# ============================================================
# 門市地圖
# ============================================================
elif page == "🗺️ 門市地圖":
    st.markdown("# 🗺️ 宜得利小地圖")
    st.caption("可使用路名、捷運站、車站名稱、門市名稱或店號搜尋。")
    components.iframe(POWER_BI_URL, height=760, scrolling=True)

# ============================================================
# 公告＆工具
# ============================================================
elif page == "📢 公告＆常用工具":
    st.markdown("# 📢 公告＆常用工具")
    tab1, tab2 = st.tabs(["📢 內部公告", "🔗 常用工具"])
    with tab1:
        ann = load_announcements()
        if ann.empty:
            st.info("目前沒有公告資料。")
        else:
            for _, row in ann.iterrows():
                cols = list(ann.columns)
                title_col = next((c for c in cols if "標題" in str(c)), cols[0])
                content_col = next((c for c in cols if any(x in str(c) for x in ["內容", "說明"])), cols[1] if len(cols)>1 else cols[0])
                date_col = next((c for c in cols if "日期" in str(c)), None)
                title = clean_announcement_title(row[title_col])
                content = clean_announcement_content(row[content_col])
                date = clean_text(row[date_col]) if date_col else ""
                st.markdown(f"<div class='card'><div class='card-title'>📌 {safe_html_text(title)}</div><div class='card-text'>{safe_html_text(content)}</div>{('<div class=\"card-meta\">📅 '+safe_html_text(date)+'</div>') if date else ''}</div>", unsafe_allow_html=True)
    with tab2:
        tools = load_tools()
        if tools.empty:
            st.info("目前沒有常用工具資料。")
        else:
            cols = list(tools.columns)
            name_col = next((c for c in cols if any(x in str(c) for x in ["名稱", "工具", "標題"])), cols[0])
            url_col = next((c for c in cols if any(x in str(c) for x in ["網址", "連結", "URL", "url"])), None)
            desc_col = next((c for c in cols if any(x in str(c) for x in ["說明", "描述", "備註"])), None)
            grid = st.columns(3)
            for i, (_, row) in enumerate(tools.iterrows()):
                name = clean_text(row[name_col])
                url = clean_text(row[url_col]) if url_col else ""
                desc = clean_text(row[desc_col]) if desc_col else ""
                with grid[i % 3]:
                    st.markdown(f"<div class='tool-card'><div class='tool-icon'>🔗</div><div class='tool-name'>{safe_html_text(name)}</div><div class='tool-desc'>{safe_html_text(desc)}</div></div>", unsafe_allow_html=True)
                    if url:
                        st.link_button("開啟工具", url, use_container_width=True)

# ============================================================
# 資料維護
# ============================================================
elif page == "🛠️ 資料維護":
    st.markdown("# 🛠️ 資料維護")
    st.caption("這裡可以直接編輯目前 Excel 中的資料，完成後按下儲存。")
    tab1, tab2, tab3 = st.tabs(["📚 知識庫", "📢 公告", "🔗 常用工具"])

    with tab1:
        df = load_data()
        if df.empty:
            st.warning("找不到知識庫資料。")
        else:
            edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", height=600, key="knowledge_editor")
            if st.button("💾 儲存知識庫", type="primary"):
                if save_knowledge_data(edited):
                    st.success("知識庫已儲存！")
                    st.rerun()

    with tab2:
        df = load_announcements()
        if df.empty:
            st.warning("找不到公告資料。")
        else:
            edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", height=550, key="announcement_editor")
            if st.button("💾 儲存公告", type="primary"):
                sheet = find_sheet(["內部公告", "公告"])
                if sheet and save_sheet(sheet, edited):
                    st.success("公告已儲存！")
                    st.rerun()

    with tab3:
        df = load_tools()
        if df.empty:
            st.warning("找不到常用工具資料。")
        else:
            edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", height=550, key="tools_editor")
            if st.button("💾 儲存常用工具", type="primary"):
                sheet = find_sheet(["常用工具", "工具"])
                if sheet and save_sheet(sheet, edited):
                    st.success("常用工具已儲存！")
                    st.rerun()
