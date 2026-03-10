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

    def sync_fio_query():
        # Обновляем фильтр по мере ввода, чтобы поиск срабатывал без Enter.
        st.session_state["fio_query"] = st.session_state.get("fio_query_input", "")

    def clear_fio_query():
        # Очищаем поле и фильтр по нажатию на иконку.
        st.session_state["fio_query_input"] = ""
        st.session_state["fio_query"] = ""

    fio_input_col, fio_clear_col = st.columns([1, 0.08])
    with fio_input_col:
        fio_query = st.text_input(
            "Поиск по ФИО",
            key="fio_query_input",
            on_change=sync_fio_query,
        )
    with fio_clear_col:
        st.button(
            "✖️",
            key="clear_fio_query",
            help="Очистить поиск по ФИО.",
            on_click=clear_fio_query,
        )
    if "fio_query" not in st.session_state:
        st.session_state["fio_query"] = fio_query
    fio_query = st.session_state.get("fio_query", fio_query)

    filtered_entries = entries
    if fio_query.strip():
        query = fio_query.strip().lower()
        filtered_entries = [
            entry
            for entry in entries
            # Ищем по первому слову в ФИО, чтобы фильтр срабатывал с первых символов.
            if entry[1].split() and entry[1].split()[0].lower().startswith(query)
        ]


    # Увеличиваем шрифт именно в колонке «ФИО» до 13px для лучшей читаемости.
    st.markdown(
        """
        <style>
        /* Колонка с индексом 1: 0 — чекбокс «Выбрать», 1 — «ФИО». */
        div[data-testid="stDataFrame"] div[role="gridcell"][data-col="1"] {
            font-size: 13px !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    table_rows = [
        {
            "Выбрать": selection_state.get(row_idx, False),
            "ФИО": fio,
            "Штрихкод": barcode,
        }
        for row_idx, fio, barcode in filtered_entries
    ]

    edited_rows = st.data_editor(
        table_rows,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Выбрать": st.column_config.CheckboxColumn(
                "Выбрать",
                help="Отметьте строки для генерации карточек.",
                default=False,
            ),
            "ФИО": st.column_config.TextColumn(
                "ФИО",
                help="Отображается первое слово как фамилия и инициалы.",
                width="large",
            ),
        },
        disabled=["ФИО", "Штрихкод"],
    )

    for row, entry in zip(edited_rows, filtered_entries):
        selection_state[entry[0]] = row.get("Выбрать", False)

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

            st.success("Готово! Файл начнёт скачиваться автоматически.")
            pdf_bytes = pdf_path.read_bytes()

            # Автоматически запускаем скачивание без отдельной кнопки.
            import base64

            b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
            auto_download = f"""
            <a id="auto-download" download="{pdf_path.name}"
               href="data:application/pdf;base64,{b64_pdf}"></a>
            <script>
              const link = document.getElementById("auto-download");
              if (link) {{
                link.click();
              }}
            </script>
            """
            st.components.v1.html(auto_download, height=0)
