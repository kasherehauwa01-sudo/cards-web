import streamlit as st
from pathlib import Path
from tempfile import TemporaryDirectory

from generate_cards import process_file_web


st.set_page_config(page_title="Генератор карточек", page_icon="🧾", layout="centered")

st.title("🧾 Генератор карточек со штрихкодами")

st.markdown(
    """
**Инструкция:**
1. Подготовьте файл формата **.xlsx**, где:
   - **первый столбец** содержит **ФИО**;
   - **второй столбец** — **числовой штрихкод**.
2. Загрузите подготовленный файл **.xlsx**.
3. Нажмите кнопку **«Сгенерировать»**.
4. **Готово.**
"""
)

st.divider()

uploaded = st.file_uploader("Загрузите Excel файл (.xlsx)", type=["xlsx"])
run = st.button("▶️ Сгенерировать", type="primary", disabled=(uploaded is None))

if run:
    if uploaded is None:
        st.warning("Сначала загрузите файл Excel (.xlsx).")
        st.stop()

    with st.spinner("Генерирую PDF..."):
        with TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            xlsx_path = tmp / uploaded.name
            xlsx_path.write_bytes(uploaded.getbuffer())

            pdf_path = process_file_web(xlsx_path)

            st.success("Готово!")
            st.download_button(
                label="⬇️ Скачать PDF",
                data=pdf_path.read_bytes(),
                file_name=pdf_path.name,
                mime="application/pdf",
            )
