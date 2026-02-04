import streamlit as st
from pathlib import Path
from tempfile import TemporaryDirectory

from generate_cards import process_file_web, read_excel_rows


st.set_page_config(page_title="Генератор карточек", page_icon="🧾", layout="centered")

st.title("🧾 Генератор карточек со штрихкодами")

st.markdown(
    """
**Инструкция:**
1. Подготовьте файл формата **.xlsx** или **.xls**.
2. Загрузите подготовленный файл **.xlsx** / **.xls**.
3. Отметьте нужные строки в таблице.
4. Нажмите кнопку **«Сгенерировать»**.
5. **Готово.**
"""
)

st.divider()

uploaded = st.file_uploader("Загрузите Excel файл (.xlsx / .xls)", type=["xlsx", "xls"])

entries = None
selection_state = {}
if uploaded is not None:
    upload_id = f"{uploaded.name}-{uploaded.size}"
    if st.session_state.get("uploaded_id") != upload_id:
        st.session_state["uploaded_id"] = upload_id
        st.session_state["row_selection"] = {}

    try:
        with TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            xlsx_path = tmp / uploaded.name
            xlsx_path.write_bytes(uploaded.getbuffer())
            entries = read_excel_rows(xlsx_path)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Не удалось прочитать Excel: {exc}")
        st.stop()

    selection_state = st.session_state.setdefault("row_selection", {})
    for row_idx, _, _ in entries:
        selection_state.setdefault(row_idx, False)

    st.subheader("Данные из файла")
    fio_query = st.text_input("Поиск по ФИО", key="fio_query")

    filtered_entries = entries
    if fio_query.strip():
        query = fio_query.strip().lower()
        filtered_entries = [
            entry
            for entry in entries
            # Ищем по первому слову в ФИО, чтобы фильтр срабатывал с первых символов.
            if entry[1].split() and entry[1].split()[0].lower().startswith(query)
        ]

    table_rows = [
        {
            "ФИО": fio,
            "Штрихкод": barcode,
            "Выбрать": selection_state.get(row_idx, False),
            "Строка": row_idx,
        }
        for row_idx, fio, barcode in filtered_entries
    ]

    edited_rows = st.data_editor(
        table_rows,
        use_container_width=True,
        hide_index=True,
        column_config={
            "ФИО": st.column_config.TextColumn(
                "ФИО",
                help="Отображается первое слово как фамилия и инициалы.",
                width="large",
            ),
            "Выбрать": st.column_config.CheckboxColumn(
                "Выбрать",
                help="Отметьте строки для генерации карточек.",
                default=False,
            )
        },
        disabled=["ФИО", "Штрихкод", "Строка"],
    )

    for row in edited_rows:
        selection_state[row["Строка"]] = row["Выбрать"]

    st.caption(f"Выбрано строк: {sum(selection_state.values())}")

run = st.button("▶️ Сгенерировать", type="primary", disabled=(uploaded is None))

if run:
    if uploaded is None:
        st.warning("Сначала загрузите файл Excel (.xlsx / .xls).")
        st.stop()

    if not entries:
        st.warning("В файле нет данных для обработки.")
        st.stop()

    selected_entries = [
        entry for entry in entries if selection_state.get(entry[0], False)
    ]
    if not selected_entries:
        st.warning("Отметьте хотя бы одну строку для генерации.")
        st.stop()

    with st.spinner("Генерирую PDF..."):
        with TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            xlsx_path = tmp / uploaded.name
            xlsx_path.write_bytes(uploaded.getbuffer())

            pdf_path = process_file_web(xlsx_path, entries=selected_entries)

            st.success("Готово!")
            st.download_button(
                label="⬇️ Скачать PDF",
                data=pdf_path.read_bytes(),
                file_name=pdf_path.name,
                mime="application/pdf",
            )
