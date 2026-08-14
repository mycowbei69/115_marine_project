"""Local document ingestion, Chinese FTS retrieval, and Ollama chat client."""
from __future__ import annotations

import os
import re
import sqlite3
import tempfile
import zipfile
from xml.etree import ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import fitz  # PyMuPDF
import requests
from docx import Document

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "my_data"
DB_PATH = ROOT / "rag_store.sqlite3"
OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://127.0.0.1:11434").rstrip("/")
MODELS = ("gemma2:9b", "deepseek-r1:8b", "qwen2.5-coder:7b", "deepseek-r1:7b")
SUPPORTED = {".pdf", ".doc", ".docx", ".odt", ".txt", ".md"}


@dataclass
class Chunk:
    source: str
    page: int | None
    position: int
    content: str


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def cjk_tokens(text: str) -> str:
    """FTS5's default tokenizer does not split CJK words; use 2-char grams."""
    normalized = re.sub(r"\s+", "", text.lower())
    grams = [normalized[i : i + 2] for i in range(len(normalized) - 1)]
    ascii_words = re.findall(r"[a-z0-9_]{2,}", text.lower())
    return " ".join(grams + ascii_words)


def split_text(text: str, limit: int = 700, overlap: int = 100) -> Iterable[str]:
    text = clean_text(text)
    while text:
        if len(text) <= limit:
            yield text
            return
        boundary = max(text.rfind(mark, 0, limit) for mark in "。！？；;\n")
        cut = boundary + 1 if boundary >= limit // 2 else limit
        yield text[:cut].strip()
        text = text[max(1, cut - overlap) :].lstrip()


def extract_pdf(path: Path) -> Iterable[tuple[int | None, str]]:
    pdf = fitz.open(path)
    try:
        for number, page in enumerate(pdf, start=1):
            yield number, page.get_text("text")
    finally:
        pdf.close()


def extract_doc(path: Path) -> Iterable[tuple[int | None, str]]:
    """Use installed Word through COM for legacy .doc; fail clearly if unavailable."""
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("需要 pywin32 與 Microsoft Word 才能讀取舊版 .doc") from exc
    word = win32com.client.DispatchEx("Word.Application")
    try:
        word.Visible = False
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "converted.txt"
            document = word.Documents.Open(str(path.resolve()), ReadOnly=True)
            try:
                document.SaveAs(str(output), FileFormat=2)  # wdFormatText
            finally:
                document.Close(False)
            yield None, output.read_text(encoding="utf-8", errors="ignore")
    finally:
        word.Quit()


def extract_file(path: Path) -> Iterable[tuple[int | None, str]]:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        yield from extract_pdf(path)
    elif suffix == ".docx":
        doc = Document(path)
        yield None, "\n".join(p.text for p in doc.paragraphs)
    elif suffix == ".odt":
        # ODT is a ZIP; parse content.xml without an additional dependency.
        with zipfile.ZipFile(path) as archive:
            xml = archive.read("content.xml")
        root = ET.fromstring(xml)
        yield None, "\n".join(part.strip() for part in root.itertext() if part.strip())
    elif suffix == ".doc":
        yield from extract_doc(path)
    else:
        yield None, path.read_text(encoding="utf-8", errors="ignore")


def connect() -> sqlite3.Connection:
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db


def setup_db(db: sqlite3.Connection, rebuild: bool) -> None:
    if rebuild:
        db.execute("DROP TABLE IF EXISTS chunks")
    db.execute("CREATE VIRTUAL TABLE IF NOT EXISTS chunks USING fts5(search_text, content UNINDEXED, source UNINDEXED, page UNINDEXED, position UNINDEXED)")
    db.commit()


def build_index(rebuild: bool = False) -> tuple[int, list[str]]:
    if not DATA_DIR.exists():
        raise FileNotFoundError(f"找不到資料目錄：{DATA_DIR}")
    db = connect()
    setup_db(db, rebuild)
    if not rebuild and db.execute("SELECT count(*) FROM chunks").fetchone()[0]:
        raise RuntimeError("索引已存在；如要重建請加上 --rebuild")
    count, errors = 0, []
    for path in sorted(p for p in DATA_DIR.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED):
        relative = str(path.relative_to(ROOT))
        try:
            position = 0
            for page, raw in extract_file(path):
                for text in split_text(raw):
                    tokens = cjk_tokens(text)
                    if tokens:
                        db.execute("INSERT INTO chunks(search_text, content, source, page, position) VALUES (?, ?, ?, ?, ?)", (tokens, text, relative, page, position))
                        count += 1
                        position += 1
        except Exception as exc:
            errors.append(f"{relative}: {exc}")
    db.commit()
    db.close()
    return count, errors


def retrieve(question: str, top_k: int = 5) -> list[Chunk]:
    query = cjk_tokens(question)
    if not query:
        return []
    db = connect()
    try:
        rows = db.execute("SELECT content, source, page, position FROM chunks WHERE chunks MATCH ? ORDER BY bm25(chunks) LIMIT ?", (" OR ".join(query.split()), top_k)).fetchall()
    finally:
        db.close()
    return [Chunk(row["source"], row["page"], row["position"], row["content"]) for row in rows]


def ask_ollama(question: str, model: str, chunks: list[Chunk]) -> str:
    if model not in MODELS:
        raise ValueError(f"不支援的模型：{model}")
    context = "\n\n".join(f"[來源 {i + 1}: {c.source}{f'，第 {c.page} 頁' if c.page else ''}]\n{c.content}" for i, c in enumerate(chunks))
    prompt = f"""你是臺灣漁業與養殖資料助理。請只依據提供的參考資料，以繁體中文回答問題。若資料不足，明確說「提供的資料不足以確認」，不要自行補充。回答中的事實後用 [來源編號] 標示出處。\n\n參考資料：\n{context or '（沒有找到相關資料）'}\n\n問題：{question}"""
    response = requests.post(f"{OLLAMA_HOST}/api/chat", json={"model": model, "stream": False, "messages": [{"role": "user", "content": prompt}]}, timeout=600)
    response.raise_for_status()
    return response.json()["message"]["content"]
