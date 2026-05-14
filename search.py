"""
DX支援報告書 RAG 検索システム

使い方:
  python search.py "製造業で在庫管理が課題の事例を探して"  # 直接検索
  python search.py                                          # 対話モード
  python search.py -n 3 "物流業のAI活用事例"              # 件数指定
"""
import sys
import io
import argparse
from pathlib import Path

# Windows コンソールの cp932 文字化け対策
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import chromadb
from sentence_transformers import SentenceTransformer

CHROMA_DIR      = Path("chroma_db")
COLLECTION_NAME = "dx_reports"
MODEL_NAME      = "paraphrase-multilingual-mpnet-base-v2"
DEFAULT_TOP_K   = 5
PREVIEW_CHARS   = 350  # 検索結果で表示するテキストの文字数


def load_resources():
    """モデルと ChromaDB コレクションを返す（再利用のため分離）"""
    if not CHROMA_DIR.exists():
        print("[エラー] インデックスが見つかりません。先に ingest.py を実行してください。")
        sys.exit(1)

    model = SentenceTransformer(MODEL_NAME)

    client = chromadb.PersistentClient(path=str(CHROMA_DIR))
    try:
        collection = client.get_collection(COLLECTION_NAME)
    except Exception:
        print("[エラー] コレクションが存在しません。先に ingest.py を実行してください。")
        sys.exit(1)

    if collection.count() == 0:
        print("[エラー] インデックスが空です。ingest.py を再実行してください。")
        sys.exit(1)

    return model, collection


def do_search(query: str, model, collection, n_results: int = DEFAULT_TOP_K) -> None:
    """クエリを受け取り、類似チャンクを表示する"""
    query_embedding = model.encode(query).tolist()

    results = collection.query(
        query_embeddings=[query_embedding],
        n_results=min(n_results, collection.count()),
        include=["documents", "metadatas", "distances"],
    )

    docs      = results["documents"][0]
    metadatas = results["metadatas"][0]
    distances = results["distances"][0]

    print(f"\n{'=' * 55}")
    print(f"  検索クエリ : {query}")
    print(f"  ヒット件数 : {len(docs)} 件")
    print(f"{'=' * 55}")

    if not docs:
        print("  該当する事例が見つかりませんでした。")
        return

    for rank, (doc, meta, dist) in enumerate(zip(docs, metadatas, distances), 1):
        similarity = 1.0 - dist  # cosine distance → similarity
        bar = "#" * int(similarity * 20)

        print(f"\n  【結果 {rank}】 類似度: {similarity:.1%}  {bar}")
        print(f"  ファイル : {meta['source']}")
        print(f"  チャンク : {meta['chunk_index'] + 1} / {meta['total_chunks']}")
        print(f"  {'─' * 48}")
        preview = doc[:PREVIEW_CHARS]
        if len(doc) > PREVIEW_CHARS:
            preview += "..."
        # インデント付きで表示
        for line in preview.split("\n"):
            print(f"  {line}")

    print(f"\n{'=' * 55}\n")


def interactive_mode(model, collection, n_results: int) -> None:
    print("=" * 55)
    print("  DX支援報告書 RAG 検索システム")
    print("  日本語で検索クエリを入力してください")
    print("  終了: 'q' または Ctrl+C")
    print("=" * 55)
    print("\n  検索例:")
    print("    製造業で在庫管理が課題の事例を探して")
    print("    小売業のデジタルマーケティング成功事例")
    print("    AIを活用した配送・物流の効率化事例")
    print("    コスト削減に成功した中堅企業のDX事例\n")

    while True:
        try:
            query = input("検索 > ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n終了します")
            break

        if query.lower() in ("q", "quit", "exit", "終了", ""):
            if query == "":
                continue
            print("終了します")
            break

        do_search(query, model, collection, n_results)


def main():
    parser = argparse.ArgumentParser(
        description="DX支援報告書 RAG 検索システム",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("query", nargs="*", help="検索クエリ（省略時は対話モード）")
    parser.add_argument("-n", "--num-results", type=int, default=DEFAULT_TOP_K,
                        help=f"取得件数（デフォルト: {DEFAULT_TOP_K}）")
    args = parser.parse_args()

    model, collection = load_resources()

    if args.query:
        do_search(" ".join(args.query), model, collection, args.num_results)
    else:
        interactive_mode(model, collection, args.num_results)


if __name__ == "__main__":
    main()
