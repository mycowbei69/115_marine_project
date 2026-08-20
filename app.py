"""漁業 Hybrid-RAG — Streamlit 前端介面。

功能：
- 側邊欄：模型選擇、段落數、魚種/文件類型/病害篩選、知識庫統計
- 主頁面：問題輸入、回答展示、來源段落詳情、查詢歷史
- 狀態指示：相關分數色彩標示（高/中/低）
"""
import streamlit as st
from rag_core import (
    DISEASES,
    DOC_TYPE_RULES,
    MODELS,
    SPECIES,
    ask_ollama,
    get_db_stats,
    retrieve,
)

# ── 頁面設定 ──────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="漁業知識庫 RAG",
    page_icon="🐟",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 樣式注入 ──────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    .score-high  { color: #22c55e; font-weight: bold; }
    .score-mid   { color: #f59e0b; font-weight: bold; }
    .score-low   { color: #ef4444; font-weight: bold; }
    .meta-tag    { background: #1e293b; border-radius: 6px; padding: 2px 8px;
                   font-size: 0.78rem; margin: 2px; display: inline-block; }
    .section-title { font-size: 0.75rem; color: #94a3b8; margin-bottom: 2px; }
    div[data-testid="stMetricValue"] { font-size: 1.6rem !important; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ── 工具函式 ──────────────────────────────────────────────────────────────────
def score_badge(score: float) -> str:
    if score >= 0.55:
        return f'<span class="score-high">● {score:.3f}</span>'
    if score >= 0.30:
        return f'<span class="score-mid">● {score:.3f}</span>'
    return f'<span class="score-low">● {score:.3f}</span>'


# ── 側邊欄 ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🐟 漁業知識庫")
    st.caption("BGE-M3 語意 ＋ BM25 混合檢索\n繁中漁業專業問答，完全本機運行")
    st.divider()

    # 模型與段落設定
    st.subheader("⚙️ 模型設定")
    model = st.selectbox("LLM 回答模型", MODELS, index=0)
    top_k = st.slider("檢索段落數 (top-k)", min_value=1, max_value=12, value=5)

    st.divider()

    # Metadata 篩選
    st.subheader("🔍 篩選條件")
    st.caption("設定後僅搜尋符合條件的段落")

    fish_opts = ["（不篩選）"] + list(SPECIES)
    fish_filter = st.selectbox("🐠 魚種", fish_opts)

    dtype_opts = ["（不篩選）"] + list(DOC_TYPE_RULES.keys()) + ["一般參考資料"]
    dtype_filter = st.selectbox("📄 文件類型", dtype_opts)

    disease_opts = ["（不篩選）"] + list(DISEASES)
    disease_filter = st.selectbox("🦠 病害類別", disease_opts)

    st.divider()

    # 知識庫統計
    st.subheader("📊 知識庫統計")
    try:
        stats = get_db_stats()
        if stats["total_chunks"] == 0:
            st.warning("⚠️ 知識庫尚未建立\n請執行 `python build_index.py --rebuild`")
        else:
            c1, c2 = st.columns(2)
            c1.metric("段落數", f"{stats['total_chunks']:,}")
            c2.metric("文件數", stats["total_sources"])

            if stats.get("doc_types"):
                st.markdown("**文件類型分布**")
                max_cnt = max(stats["doc_types"].values(), default=1)
                for dtype, cnt in sorted(stats["doc_types"].items(), key=lambda x: -x[1]):
                    st.progress(cnt / max_cnt, text=f"{dtype}（{cnt}）")

            if stats.get("species"):
                with st.expander("魚種分布"):
                    max_sp = max(stats["species"].values(), default=1)
                    for sp, cnt in sorted(stats["species"].items(), key=lambda x: -x[1]):
                        st.progress(cnt / max_sp, text=f"{sp}（{cnt}）")
    except Exception as e:
        st.error(f"無法讀取統計：{e}")

    st.divider()
    st.caption("向量模型：`bge-m3:latest`\n資料目錄：`my_data/`")


# ── 主頁面 ────────────────────────────────────────────────────────────────────
st.title("🐟 台灣漁業知識庫問答系統")
st.caption(
    "基於 **Hybrid-RAG**（BGE-M3 語意向量 + BM25 關鍵字精準匹配）的漁業專業問答。"
    "所有資料與模型皆在本機運行，不上傳任何資料。"
)

# ── 查詢歷史初始化 ────────────────────────────────────────────────────────────
if "history" not in st.session_state:
    st.session_state["history"] = []

# ── 快捷範例問題 ──────────────────────────────────────────────────────────────
st.markdown("**💡 範例問題：**")
example_questions = [
    "石斑魚養殖的水質管理重點？",
    "吳郭魚育種與品系選育方法？",
    "鰻魚寄生蟲病的防治方法？",
    "台灣核准使用的水產藥品有哪些？",
    "虹彩病毒的症狀與處理方式？",
]
eq_cols = st.columns(len(example_questions))
selected_example = ""
for col, eq in zip(eq_cols, example_questions):
    if col.button(eq, use_container_width=True, key=f"eq_{eq}"):
        selected_example = eq

# ── 問題輸入區 ────────────────────────────────────────────────────────────────
question_input = st.text_area(
    "🔎 輸入問題",
    value=selected_example,
    placeholder="例如：石斑魚虹彩病毒的症狀與防治方法？",
    height=100,
    label_visibility="collapsed",
)

btn_col1, btn_col2, btn_col3 = st.columns([2, 1, 5])
submit = btn_col1.button("🔍 提問", type="primary", use_container_width=True)
clear_hist = btn_col2.button("🗑️ 清歷史", use_container_width=True)

if clear_hist:
    st.session_state["history"] = []
    st.success("查詢歷史已清除")

# ── 篩選條件轉換 ──────────────────────────────────────────────────────────────
fish_arg = None if fish_filter.startswith("（") else fish_filter
dtype_arg = None if dtype_filter.startswith("（") else dtype_filter
disease_arg = None if disease_filter.startswith("（") else disease_filter

# ── 執行查詢 ──────────────────────────────────────────────────────────────────
if submit and question_input.strip():
    question = question_input.strip()

    # 顯示啟用的篩選條件
    active_filters = [
        f"🐠 {fish_arg}" if fish_arg else "",
        f"📄 {dtype_arg}" if dtype_arg else "",
        f"🦠 {disease_arg}" if disease_arg else "",
    ]
    active_str = "　".join(f for f in active_filters if f)
    if active_str:
        st.info(f"已啟用篩選：{active_str}")

    try:
        with st.spinner("🔍 搜尋相關文獻…"):
            chunks = retrieve(
                question,
                top_k=top_k,
                fish_species=fish_arg,
                doc_type=dtype_arg,
                disease=disease_arg,
            )

        if not chunks:
            st.warning(
                "⚠️ 未找到相關段落。\n"
                "建議：放寬篩選條件、修改問題關鍵字，或確認索引已建立（`python build_index.py --rebuild`）。"
            )
        else:
            with st.spinner(f"🤖 {model} 整合回答中…"):
                answer = ask_ollama(question, model, chunks)

            # ── 回答展示 ──────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### 💬 回答")
            st.markdown(answer)

            # 儲存歷史
            st.session_state["history"].append(
                {"q": question, "a": answer, "n": len(chunks), "model": model}
            )

            # ── 來源段落展示 ──────────────────────────────────────────────────
            st.markdown("---")
            with st.expander(f"📚 查看檢索到的 {len(chunks)} 個來源段落", expanded=False):
                for i, chunk in enumerate(chunks, 1):
                    page_str = f"，第 {chunk.page} 頁" if chunk.page else ""
                    st.markdown(
                        f"**{score_badge(chunk.score)} [{i}]** `{chunk.source}{page_str}`",
                        unsafe_allow_html=True,
                    )

                    # Metadata 標籤
                    tag_html = (
                        f'<span class="meta-tag">🐠 {", ".join(chunk.metadata["fish_species"])}</span>'
                        f'<span class="meta-tag">🦠 {", ".join(chunk.metadata["disease_name"])}</span>'
                        f'<span class="meta-tag">📄 {chunk.metadata["doc_type"]}</span>'
                    )
                    st.markdown(tag_html, unsafe_allow_html=True)

                    if chunk.metadata.get("hierarchy") and chunk.metadata["hierarchy"] != "未標記章節":
                        st.markdown(
                            f'<p class="section-title">章節：{chunk.metadata["hierarchy"]}</p>',
                            unsafe_allow_html=True,
                        )

                    st.text_area(
                        f"段落內容 [{i}]",
                        value=chunk.content,
                        height=180,
                        key=f"chunk_{i}_{chunk.id}",
                        label_visibility="collapsed",
                    )
                    st.divider()

    except Exception as exc:
        st.error(f"❌ 執行失敗：{exc}")
        st.info("請確認 Ollama 服務已啟動，且向量索引已建立。")

# ── 查詢歷史 ──────────────────────────────────────────────────────────────────
if st.session_state["history"]:
    st.markdown("---")
    with st.expander(f"📜 查詢歷史（共 {len(st.session_state['history'])} 筆）", expanded=False):
        for i, item in enumerate(reversed(st.session_state["history"]), 1):
            st.markdown(f"**Q{i} [{item['model']}]**　{item['q']}")
            st.markdown(f"> {item['a'][:300]}{'…' if len(item['a']) > 300 else ''}")
            st.caption(f"檢索 {item['n']} 個段落")
            st.divider()
