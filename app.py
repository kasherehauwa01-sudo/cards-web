import streamlit as st
from pathlib import Path
from tempfile import TemporaryDirectory

from generate_cards import process_file_web

st.set_page_config(page_title="Генератор карточек", page_icon="🧾", layout="centered")

st.title("🧾 Генератор карточек со штрихкодами")
st.write("1) Загрузи config.json (один и тот же для всех файлов)\n2) Загрузи Excel (.xlsx)\n3) Нажми кнопку и скачай PDF")

st.subheader("Шаг 1. Конфигурация (одна на все)")
config_file = st.file_uploader("config.json", type=["json"])

st.subheader("Шаг 2. Excel")
xlsx_file = st.file_uploader("Excel файл (.xlsx)", type=["xlsx"])

run = st.button("▶️ Сгенерировать PDF", type="primary", disabled=(xlsx_file is None))

if run:
    if xlsx_file is None:
        st.warning("Сначала загрузи Excel (.xlsx)")
        st.stop()

    with st.spinner("Генерирую PDF..."):
        with TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            # сохраняем Excel
            xlsx_path = tmp / xlsx_file.name
            xlsx_path.write_bytes(xlsx_file.getbuffer())

            # сохраняем config.json рядом с Excel (если загружен)
            if config_file is not None:
                (tmp / "config.json").write_bytes(config_file.getbuffer())

            # запускаем обработку
            pdf_path = process_file_web(xlsx_path)

            st.success("Готово!")
            st.download_button(
                label="⬇️ Скачать PDF",
                data=pdf_path.read_bytes(),
                file_name=pdf_path.name,
                mime="application/pdf",
            )

            if config_file is None:
                st.info("config.json не загружен — использованы настройки по умолчанию.")
            else:
                st.info("Использован загруженный config.json.")
