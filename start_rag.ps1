$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    python -m venv .venv
}

& .\.venv\Scripts\python.exe -m pip install -r requirements.txt

if (-not (Test-Path "rag_store.sqlite3")) {
    & .\.venv\Scripts\python.exe build_index.py
}

& .\.venv\Scripts\streamlit.exe run app.py
