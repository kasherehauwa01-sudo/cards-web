import streamlit as st
from pathlib import Path
from tempfile import TemporaryDirectory

from generate_cards import process_file_web


st.set_page_config(page_title="Генератор карточек", page_icon="🧾", layout="centered")

st.title("🧾 Генератор карточек со штрихкодами")

st.markdown(
    """
**Инструкция:**
1. Подготовьте файл формата **.xls или .xlsx**, где:
   - **первый столбец** содержит **ФИО**;
   - **второй столбец** — **числовой штрихкод**.
2. Загрузите подготовленный файл **.xls / .xlsx**.
3. Нажмите кнопку **«Сгенерировать»**.
4. **Готово.**
"""
)

# Показываем пример таблицы (файл должен лежать рядом с app.py в репозитории)
example_img_path = Path(__file__).with_name("example_table.jpg")
if example_img_path.exists():
    st.image(str(example_img_path), caption="Пример заполнения таблицы", use_container_width=True)
else:
    st.info("Чтобы показать пример таблицы, добавьте файл `example_table.jpg` рядом с `app.py` в репозитории.")

st.divider()

uploaded = st.file_uploader("Загрузите Excel файл (.xls / .xlsx)", type=["xls", "xlsx"])
run = st.button("▶️ Сгенерировать", type="primary", disabled=(uploaded is None))

if run:
    if uploaded is None:
        st.warning("Сначала загрузите файл Excel.")
        st.stop()

    suffix = Path(uploaded.name).suffix.lower()

    # Твой текущий скрипт читает только .xlsx (openpyxl). .xls не поддерживается без конвертации.
    if suffix == ".xls":
        st.error("Формат .xls сейчас не поддерживается. Сохраните файл как .xlsx и загрузите снова.")
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
