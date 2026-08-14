# 漁業資料本機 RAG

此專案會讀取 `my_data/` 下的 PDF、DOC/DOCX、ODT、TXT 與 Markdown 文件，切成段落後建立 SQLite FTS5 檢索索引。問答完全在本機進行，並可在下列 Ollama 模型間切換：

- `gemma2:9b`
- `deepseek-r1:8b`
- `qwen2.5-coder:7b`
- `deepseek-r1:7b`

檢索採用中文 n-gram 斷詞，無須額外下載 embedding 模型；這讓您目前已有的四個模型可立即使用。每次回答都會列出資料來源與頁碼。

## 安裝與建索引

請在專案根目錄執行：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python build_index.py
```

建立索引後啟動網頁介面：

```powershell
streamlit run app.py
```

或從命令列提問：

```powershell
python chat.py "石斑魚常見的養殖管理重點是什麼？" --model gemma2:9b
```

## Ollama 設定

程式預設連到 `http://127.0.0.1:11434`。如果 Ollama 服務在其他主機，設定環境變數後再啟動：

```powershell
$env:OLLAMA_HOST = "http://192.168.1.10:11434"
```

可用 `GET $env:OLLAMA_HOST/api/tags` 確認服務與模型是否可用。此電腦目前未在 PATH 找到 `ollama` 指令；只要 Ollama 服務已啟動，程式仍可透過 API 使用模型。

## 支援格式與提醒

PDF、DOCX、ODT 可直接擷取。舊版 `.doc` 會自動透過 Windows 的 Microsoft Word 轉為文字（首次安裝已包含 `pywin32`）。若電腦沒有 Word，建索引完成後終端機會列出未擷取檔案；請先另存為 DOCX 或 PDF 後重建索引。

```powershell
python build_index.py --rebuild
```
