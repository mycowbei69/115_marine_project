import streamlit as st
from rag_core import MODELS, ask_ollama, retrieve

st.set_page_config(page_title="漁業資料 RAG", page_icon="🐟")
st.title("🐟 漁業資料本機 RAG")
st.caption("資料只在本機檢索；回答由本機 Ollama 模型產生。")
model = st.selectbox("回答模型", MODELS)
top_k = st.slider("參考段落數", 1, 10, 5)
question = st.text_area("問題", placeholder="例如：石斑魚養殖時水質管理的建議？")
if st.button("提問", type="primary") and question.strip():
    try:
        chunks = retrieve(question, top_k)
        with st.spinner("檢索資料並由模型回答中…"):
            answer = ask_ollama(question, model, chunks)
        st.markdown(answer)
        with st.expander(f"檢索到的 {len(chunks)} 個來源"):
            for i, chunk in enumerate(chunks, 1):
                page = f"，第 {chunk.page} 頁" if chunk.page else ""
                st.markdown(f"**[{i}] {chunk.source}{page}**")
                st.write(chunk.content)
    except Exception as exc:
        st.error(f"執行失敗：{exc}")
