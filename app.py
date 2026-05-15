"""
DX支援事例 検索システム
意味検索（セマンティック検索）× タグ絞り込み
"""
from pathlib import Path

import chromadb
import streamlit as st
from sentence_transformers import SentenceTransformer

# ── 定数 ──────────────────────────────────────────────────
CHROMA_DIR      = Path("chroma_db")
COLLECTION_NAME = "dx_reports"
MODEL_NAME      = "paraphrase-multilingual-mpnet-base-v2"
NONE            = "（指定なし）"

# ══════════════════════════════════════════════════════════
# タグ定義
# ══════════════════════════════════════════════════════════
INDUSTRY_OPTIONS = [NONE,
    "建設業", "医療", "介護・福祉", "小売業", "卸売業",
    "飲食業・飲食店", "宿泊業", "情報通信業", "印刷業",
    "不動産業", "農業・林業・漁業", "物品賃貸業",
    "コンサルティング", "専門サービス業", "教育",
]
CHALLENGE_OPTIONS = [NONE,
    "会計・経理", "請求書・見積書", "集客・販促", "SNS活用",
    "在庫管理", "受発注管理", "棚卸・廃棄ロス",
    "手作業・アナログ", "属人化", "情報共有不足",
    "勤怠管理", "給与計算", "顧客管理", "予約管理",
    "セキュリティ", "人材不足", "ITリテラシー不足", "非効率",
]
SOLUTION_OPTIONS = [NONE,
    "kintone", "RPA", "生成AI・ChatGPT", "AI-OCR",
    "IT導入補助金", "DX補助金",
    "LINE公式アカウント",
    "電子カルテ", "予約システム",
    "弥生", "freee", "マネーフォワード", "PCA",
    "スマレジ", "Airpay",
    "ECサイト・ネットショップ", "ホームページ制作",
    "QRコード・バーコード", "IoT",
    "電子契約", "ペーパーレス",
    "DX戦略・計画",
]
MUNICIPALITY_OPTIONS = [NONE, "久留米市", "福岡県", "北九州市"]

# タグ表示名 → 実際に検索する語のリスト（複数表記に対応）
SEARCH_TERMS: dict[str, list[str]] = {
    "kintone":              ["kintone", "Kintone", "キントーン"],
    "生成AI・ChatGPT":      ["生成AI", "ChatGPT"],
    "AI-OCR":               ["AI-OCR", "OCR"],
    "会計・経理":           ["会計", "経理"],
    "請求書・見積書":       ["請求書", "見積書"],
    "集客・販促":           ["集客", "販促"],
    "SNS活用":              ["SNS", "Instagram", "Facebook"],
    "受発注管理":           ["受発注"],
    "棚卸・廃棄ロス":       ["棚卸", "廃棄ロス", "過剰在庫"],
    "手作業・アナログ":     ["手作業", "アナログ"],
    "情報共有不足":         ["情報共有"],
    "顧客管理":             ["顧客管理", "CRM"],
    "IT導入補助金":         ["IT導入補助金"],
    "DX補助金":             ["DX補助金", "久留米市DX補助金"],
    "LINE公式アカウント":   ["LINE"],
    "ECサイト・ネットショップ": ["ECサイト", "ネットショップ", "Shopify", "BASE"],
    "ホームページ制作":     ["ホームページ", "Webサイト"],
    "QRコード・バーコード": ["QRコード", "バーコード", "RFID"],
    "電子契約":             ["電子契約", "クラウドサイン"],
    "DX戦略・計画":         ["DX戦略", "DX推進"],
    "介護・福祉":           ["介護", "福祉"],
    "飲食業・飲食店":       ["飲食業", "飲食店", "居酒屋", "カフェ", "レストラン"],
    "農業・林業・漁業":     ["農業", "林業", "漁業"],
    "ITリテラシー不足":     ["ITリテラシー", "IT知識"],
    "専門サービス業":       ["専門サービス業", "専門・技術サービス業"],
}


def get_terms(tag: str) -> list[str]:
    return SEARCH_TERMS.get(tag, [tag])


def build_where_document(active_tags: list[str]):
    """タグリストから ChromaDB の where_document フィルタを構築"""
    conditions = []
    for tag in active_tags:
        terms = get_terms(tag)
        cond = ({"$contains": terms[0]}
                if len(terms) == 1
                else {"$or": [{"$contains": t} for t in terms]})
        conditions.append(cond)

    if not conditions:
        return None
    if len(conditions) == 1:
        return conditions[0]
    return {"$and": conditions}


# ── リソース読み込み（キャッシュ）──────────────────────────
@st.cache_resource(show_spinner="Embedding モデルを読み込み中（初回のみ）…")
def load_resources():
    model      = SentenceTransformer(MODEL_NAME)
    client     = chromadb.PersistentClient(path=str(CHROMA_DIR))
    collection = client.get_collection(COLLECTION_NAME)
    return model, collection


# ── 検索ロジック ───────────────────────────────────────────
def do_search(
    query: str,
    active_tags: list[str],
    n_results: int,
) -> tuple[dict | None, str | None]:

    model, collection = load_resources()
    where_doc = build_where_document(active_tags)

    # クエリが空の場合はタグをクエリ代わりに使う
    if query.strip():
        effective_query = query.strip()
    elif active_tags:
        effective_query = " ".join(
            get_terms(t)[0] for t in active_tags
        ) + " DX支援事例"
    else:
        return None, "検索クエリを入力するか、タグを選択してください。"

    embedding = model.encode(effective_query).tolist()

    try:
        results = collection.query(
            query_embeddings=[embedding],
            n_results=min(n_results, collection.count()),
            where_document=where_doc,
            include=["documents", "metadatas", "distances"],
        )
    except Exception as e:
        msg = str(e)
        if "Number of requested results" in msg:
            return None, "タグの絞り込みが厳しすぎます。条件を緩めてみてください。"
        return None, f"検索エラー: {msg}"

    if not results["documents"][0]:
        return None, "条件に一致する事例が見つかりませんでした。"

    return results, None


# ── 結果表示 ───────────────────────────────────────────────
def render_results(results: dict, query: str, active_tags: list[str]) -> None:
    docs      = results["documents"][0]
    metas     = results["metadatas"][0]
    distances = results["distances"][0]

    label = query.strip() or " / ".join(active_tags)
    st.success(f"「{label}」で **{len(docs)} 件**ヒット")

    TYPE_ICON = {"pdf": "🔴 PDF", "xlsx": "🟢 Excel", "xls": "🟡 XLS"}

    for rank, (doc, meta, dist) in enumerate(zip(docs, metas, distances), 1):
        similarity  = 1.0 - dist
        source      = meta.get("source", "不明")
        ftype       = meta.get("file_type", "")
        chunk_idx   = meta.get("chunk_index", 0) + 1
        total_ch    = meta.get("total_chunks", 1)
        type_icon   = TYPE_ICON.get(ftype, "📄")

        with st.expander(
            f"**#{rank}**  {Path(source).name}  ｜  {similarity:.1%}",
            expanded=(rank <= 3),
        ):
            col_path, col_meta = st.columns([3, 1])
            with col_path:
                st.caption(f"📁 `{source}`")
            with col_meta:
                st.caption(f"{type_icon}　チャンク {chunk_idx}/{total_ch}")

            st.progress(similarity)
            st.divider()

            # テキストプレビュー（400字）
            preview = doc[:400] + ("…" if len(doc) > 400 else "")
            st.text(preview)


PASSWORD = "dx2024"


def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    st.title("🔐 ログイン")
    pw = st.text_input("パスワード", type="password", key="pw_input")
    if st.button("ログイン", use_container_width=True):
        if pw == PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False


# ══════════════════════════════════════════════════════════
# Streamlit アプリ本体
# ══════════════════════════════════════════════════════════
def main() -> None:
    st.set_page_config(
        page_title="DX事例検索",
        page_icon="🔍",
        layout="wide",
    )

    if not check_password():
        st.stop()

    # ── サイドバー：タグフィルタ ─────────────────────────
    with st.sidebar:
        st.header("🏷️ タグで絞り込む")

        def clear_tags() -> None:
            for key in ("sb_industry", "sb_challenge", "sb_solution", "sb_municipality"):
                st.session_state[key] = NONE

        industry     = st.selectbox("業種",   INDUSTRY_OPTIONS,     key="sb_industry")
        challenge    = st.selectbox("課題",   CHALLENGE_OPTIONS,    key="sb_challenge")
        solution     = st.selectbox("解決策", SOLUTION_OPTIONS,     key="sb_solution")
        municipality = st.selectbox("自治体", MUNICIPALITY_OPTIONS, key="sb_municipality")

        active_tags = [
            t for t in [industry, challenge, solution, municipality]
            if t != NONE
        ]

        if active_tags:
            st.divider()
            st.markdown("**選択中のタグ**")
            for tag in active_tags:
                st.success(f"✓ {tag}")
            st.button("🗑️ タグをクリア", on_click=clear_tags, use_container_width=True)

        st.divider()
        n_results = st.slider("表示件数", min_value=3, max_value=20, value=10)

        st.divider()
        total = load_resources()[1].count()
        st.caption(f"インデックス件数: {total:,} チャンク")

    # ── メイン：検索フォーム ─────────────────────────────
    st.title("🔍 DX支援事例 検索システム")
    st.caption("意味検索（セマンティック検索）× タグ絞り込みで事例を探します")

    with st.form("search_form"):
        query = st.text_input(
            "検索クエリ（省略可：タグのみでも検索できます）",
            placeholder="例：在庫管理が課題の製造業の事例を教えて",
        )
        submitted = st.form_submit_button("🔍 検索する", use_container_width=True)

    # ── 検索実行・結果表示 ───────────────────────────────
    if submitted:
        with st.spinner("検索中…"):
            results, error = do_search(query, active_tags, n_results)

        if error:
            st.error(error)
        else:
            st.session_state["last_results"] = results
            st.session_state["last_query"]   = query
            st.session_state["last_tags"]    = active_tags

    # 結果はセッションに保持（フィルタ操作後も消えない）
    if "last_results" in st.session_state and st.session_state["last_results"]:
        render_results(
            st.session_state["last_results"],
            st.session_state.get("last_query", ""),
            st.session_state.get("last_tags", []),
        )
    elif not submitted:
        st.info(
            "**使い方**\n\n"
            "- 左サイドバーで **業種・課題・解決策・自治体** を選択\n"
            "- 検索クエリを入力（省略してタグのみでも検索可能）\n"
            "- 「検索する」ボタンを押す\n\n"
            "タグと意味検索を組み合わせて事例を絞り込めます。"
        )


if __name__ == "__main__":
    main()
