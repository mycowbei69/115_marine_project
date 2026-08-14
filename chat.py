import argparse
from rag_core import MODELS, ask_ollama, retrieve

parser = argparse.ArgumentParser(description="使用本機 Ollama 漁業 RAG 問答")
parser.add_argument("question")
parser.add_argument("--model", choices=MODELS, default=MODELS[0])
parser.add_argument("--top-k", type=int, default=5)
args = parser.parse_args()
chunks = retrieve(args.question, args.top_k)
print(ask_ollama(args.question, args.model, chunks))
print("\n--- 檢索來源 ---")
for index, chunk in enumerate(chunks, 1):
    page = f" 第 {chunk.page} 頁" if chunk.page else ""
    print(f"[{index}] {chunk.source}{page}")
