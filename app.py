import streamlit as st
from pathlib import Path
from tempfile import TemporaryDirectory

from generate_cards import process_file_web


st.set_page_config(page_title="Генератор карточек", page_icon="🧾", layout="centered")

st.title("🧾 Генератор карточек со штрихкодами")
st.write("Загрузи Excel (.xlsx), нажми кнопку и скачай PDF. Настройки вшиты в код.")

xlsx_file = st.file_uploader("Excel файл (.xlsx)", type=["xlsx"])

run = st.button("▶️ Сгенерировать PDF", type="primary", disabled=(xlsx_file is None))

if run:
    if xlsx_file is None:
        st.warning("Сначала загрузи Excel (.xlsx)")
        st.stop()

    with st.spinner("Генерирую PDF..."):
        with TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            xlsx_path = tmp / xlsx_file.name
            xlsx_path.write_bytes(xlsx_file.getbuffer())

            pdf_path = process_file_web(xlsx_path)

            st.success("Готово!")
            st.download_button(
                label="⬇️ Скачать PDF",
                data=pdf_path.read_bytes(),
                file_name=pdf_path.name,
                mime="application/pdf",
            )