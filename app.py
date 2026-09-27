"""Streamlit-фронтенд для взаимодействия с бэкендом преобразования координат."""

import streamlit as st
import pandas as pd
import requests
from io import BytesIO


BACKEND_URL = "http://127.0.0.1:8000"

st.set_page_config(
    page_title="Coordinate Converter",
    page_icon="🌐",
    layout="wide",
)

st.title("Преобразование координат")
st.markdown(
    "Загрузите Excel-файл с координатами в начальной системе. "
    "Сервис преобразует их по формулам перехода и вернёт отчёт."
)

# Проверка доступности бэкенда
with st.sidebar:
    st.header("Настройки")

    try:
        status = requests.get(BACKEND_URL, timeout=5)
        if status.status_code == 200:
            st.success("Бэкенд доступен")
        else:
            st.warning(f"Бэкенд вернул код {status.status_code}")
    except requests.RequestException as e:
        st.error(f"Бэкенд недоступен: {e}")

    st.divider()
    st.subheader("Параметры перехода")

    dx = st.number_input("dx", value=23.557, format="%.6f")
    dy = st.number_input("dy", value=-140.844, format="%.6f")
    dz = st.number_input("dz", value=-79.778, format="%.6f")
    wx = st.number_input("wx (угл. сек.)", value=-0.00230, format="%.6f")
    wy = st.number_input("wy (угл. сек.)", value=-0.34646, format="%.6f")
    wz = st.number_input("wz (угл. сек.)", value=-0.79421, format="%.6f")
    m = st.number_input("m", value=-0.00000228, format="%.10f")

    st.divider()
    st.subheader("Формат отчёта")
    output_format = st.radio(
        "Что скачать после обработки:",
        ["Markdown (.md)", "Word (.docx)"],
    )


uploaded_file = st.file_uploader("Выберите Excel-файл (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Предпросмотр
    try:
        df_preview = pd.read_excel(BytesIO(uploaded_file.getvalue()))
        st.subheader("Предварительный просмотр")
        st.dataframe(df_preview.head(10))

        col1, col2, col3 = st.columns(3)
        col1.metric("Строк", df_preview.shape[0])
        col2.metric("Столбцов", df_preview.shape[1])
        col3.metric("Пропусков", int(df_preview.isna().sum().sum()))
    except Exception as e:
        st.error(f"Не удалось прочитать файл: {e}")
        st.stop()

    st.divider()

    if st.button("Преобразовать", type="primary"):
        # Готовим запрос
        if output_format.startswith("Markdown"):
            endpoint = f"{BACKEND_URL}/process-excel/"
            mime = "text/markdown"
            ext = "md"
        else:
            endpoint = f"{BACKEND_URL}/process-excel-docx/"
            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ext = "docx"

        files = {
            "file": (uploaded_file.name, uploaded_file.getvalue(),
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        }
        data = {
            "dx": dx, "dy": dy, "dz": dz,
            "wx": wx, "wy": wy, "wz": wz,
            "m": m,
        }

        with st.spinner("Обрабатываю файл..."):
            try:
                response = requests.post(endpoint, files=files, data=data, timeout=60)
            except requests.RequestException as e:
                st.error(f"Ошибка соединения с бэкендом: {e}")
                st.stop()

        if response.status_code != 200:
            st.error(f"Бэкенд вернул ошибку {response.status_code}: {response.text}")
            st.stop()

        st.success("Отчёт готов")

        if ext == "md":
            # Показываем markdown прямо на странице
            report_text = response.content.decode("utf-8")
            st.subheader("Отчёт")
            st.markdown(report_text)

        st.download_button(
            label=f"Скачать отчёт (.{ext})",
            data=response.content,
            file_name=f"report.{ext}",
            mime=mime,
        )