import argparse
from rag_core import build_index

parser = argparse.ArgumentParser(description="建立漁業資料 RAG 索引")
parser.add_argument("--rebuild", action="store_true", help="刪除舊索引後重建")
args = parser.parse_args()
count, errors = build_index(args.rebuild)
print(f"完成：已建立 {count} 個檢索段落。")
if errors:
    print("\n以下檔案無法擷取：")
    print("\n".join(f"- {item}" for item in errors))
