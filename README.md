# 台灣漁業 Hybrid-RAG 向量知識庫

本地端漁業知識問答系統，整合 **BGE-M3 語意向量**與 **BM25 關鍵字精準匹配**，完全在本機運行，不上傳任何資料。

---

## 系統架構

```
my_data/          ← 原始文件（PDF / DOCX / DOC / ODT）
    ↓
rag_core.py       ← 清洗 → 切塊 → 嵌入 → SQLite 向量庫
    ↓
rag_store.sqlite3 ← 向量索引（chunks + FTS5 全文搜尋）
    ↓
app.py            ← Streamlit 問答介面
```

### Hybrid-RAG 檢索流程

| 步驟 | 技術 | 說明 |
|------|------|------|
| 語意向量 | BGE-M3（`bge-m3:latest`） | 支援繁中 + 多語 + 8192 token 長文 |
| 關鍵字搜尋 | SQLite FTS5 + BM25 | Unigram + Bigram + Trigram，專業術語高召回 |
| 分數融合 | 動態 RRF 加權 | 含專有名詞時 BM25 權重提升至 58% |
| Metadata 過濾 | 後過濾 | 依魚種、文件類型、病害精準限縮 |

---

## 快速開始

### 1. 環境需求

- Python 3.11+
- [Ollama](https://ollama.com/)（本機運行 LLM 與向量模型）
- Windows：若要處理舊版 `.doc` 檔案，需安裝 Microsoft Word

### 2. 安裝

```powershell
# 建立虛擬環境
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 安裝依賴
pip install -r requirements.txt
```

### 3. 拉取向量模型

```powershell
ollama pull bge-m3
```

> BGE-M3 是 BAAI 開源的多語言向量模型，支援繁體中文、8192 token 上下文。

### 4. 建立向量索引

```powershell
# 首次建立
python build_index.py

# 強制重建（清除舊資料）
python build_index.py --rebuild
```

建立過程會顯示：
- 已處理的文件清單
- 嵌入進度百分比
- 無法讀取的檔案清單（及錯誤原因）

### 5. 啟動問答介面

```powershell
streamlit run app.py
```

瀏覽器開啟 `http://localhost:8501`

---

## 支援格式

| 格式 | 擷取方式 | 表格支援 | 圖片支援 |
|------|---------|---------|---------|
| `.pdf` | PyMuPDF | ✅ Markdown 化 | ✅ 存檔索引 |
| `.docx` | python-docx | ✅ | ❌ |
| `.doc` | Word COM（需 Microsoft Word）| ✅ | ❌ |
| `.odt` | ZIP + ElementTree XML | ✅ | ❌ |
| `.txt` / `.md` | 純文字讀取 | ❌ | ❌ |

> **`.doc` 說明**：舊版 Word 格式透過 `win32com` 呼叫 Microsoft Word 另存為 `.docx`，再用 python-docx 解析保留表格結構。若電腦未安裝 Word，建索引完成後終端機會列出該檔案的錯誤；請手動另存為 `.docx` 或 `.pdf` 後執行 `--rebuild`。

---

## 資料清洗規則

### 頁首/頁尾去除

自動過濾以下噪音（不影響語意內容）：
- 頁碼（`第 N 頁`、`N/M`、`— N —`）
- 出版機構標頭（農業部、水產試驗所等）
- 跨頁重複行（出現頻率 ≥ 60% 的行）
- 版權宣告、水平分隔線

### 術語校正清單

| 原文 | 校正為 |
|------|-------|
| 吳鍋魚、吳郭漁、吴郭鱼 | 吳郭魚 |
| 虱目漁、虱目渔 | 虱目魚 |
| 石班魚、石斑仔 | 石斑魚 |
| 金目鱸魚、鱸仔魚 | 金目鱸 |
| 七星鱸魚、海鱸、花鱸 | 七星鱸 |
| 烏仔魚 | 烏魚 |
| 泰國蝦 | 泰國蝦（*Macrobrachium rosenbergii*）|
| 草蝦 | 草蝦（*Penaeus monodon*）|
| 鰻苗、鰻仔 | 鰻魚苗 / 鰻魚 |

---

## 切塊策略

### 階層式切塊（Hierarchical Chunking）

依章節結構切分：`章 > 節 > 小節 > 段落`（最多 4 層）

```
第三章 水產養殖管理
  └─ 第四節 石斑魚
       └─ 二、疾病防治
            └─ （一）虹彩病毒  ← 此為一個切塊的 hierarchy
```

### 切塊參數

| 參數 | 數值 |
|------|------|
| 目標長度 | 420 字 |
| 最大長度 | 500 字 |
| Overlap | 70 字（跨頁銜接） |
| 斷點優先 | `。！？；;` → 換行 → 強制截斷 |

### Metadata（每切塊自動注入）

```json
{
  "fish_species": ["石斑魚"],
  "disease_name": ["虹彩病毒"],
  "doc_type": "病害圖鑑",
  "source": "my_data/技術手冊17-第三章水產養殖管理技術 第四節石斑魚.pdf，第 42 頁",
  "hierarchy": "第三章 水產養殖管理 > 第四節 石斑魚 > 二、疾病防治"
}
```

---

## 問答模型

可在介面中切換以下 Ollama 本機模型：

| 模型 | 下載指令 | 特色 |
|------|---------|------|
| `gemma2:9b` | `ollama pull gemma2:9b` | Google，繁中流暢 |
| `deepseek-r1:8b` | `ollama pull deepseek-r1:8b` | 推理能力強 |
| `qwen2.5-coder:7b` | `ollama pull qwen2.5-coder:7b` | 適合結構化回答 |
| `deepseek-r1:7b` | `ollama pull deepseek-r1:7b` | 輕量推理 |

---

## 環境變數

| 變數 | 預設值 | 說明 |
|------|--------|------|
| `OLLAMA_HOST` | `http://127.0.0.1:11434` | Ollama 服務位址 |
| `EMBED_MODEL` | `bge-m3:latest` | 向量嵌入模型 |

```powershell
# 使用遠端 Ollama 服務
$env:OLLAMA_HOST = "http://192.168.1.10:11434"
streamlit run app.py
```

---

## 命令列問答

```powershell
python chat.py "石斑魚常見養殖管理重點是什麼？" --model gemma2:9b
```

---

## 新增文件

1. 將 PDF / DOCX / DOC / ODT 放入 `my_data/`
2. 執行 `python build_index.py --rebuild` 重建索引

---

## 授權

本系統程式碼以 MIT 授權釋出；`my_data/` 中的文件版權歸各原始出版單位所有。
