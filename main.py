import os
import re
import tempfile
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Проверка ААС КНД", layout="wide")
st.title("Проверка выгрузок из мониторинга ААС КНД")

choice = st.radio(
    "Выберите тип проверки:",
    ["Перечень объектов", "Перечень объектов (по блокам)", "Проверка КНМ"],
)


# === Общие вспомогательные функции ===
def fix_datetime_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    df_copy = dataframe.copy()
    for column in df_copy.columns:
        if df_copy[column].apply(
                lambda x: isinstance(x, (datetime, pd.Timestamp, np.datetime64))
        ).any():
            df_copy[column] = df_copy[column].apply(
                lambda x: x.strftime("%d.%m.%Y")
                if (isinstance(x, (datetime, pd.Timestamp)) and not pd.isnull(x))
                else ("Не присвоена" if pd.isnull(x) else x)
            )
    return df_copy


def apply_style_to_sheet(sheet, wb, apply_auto_width: bool = False) -> None:
    # Заголовки
    for cell in sheet[1]:
        cell.font = Font(name="Calibri", size=11, bold=False)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = Border()

    # Данные
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row,
                               min_col=1, max_col=sheet.max_column):
        for cell in row:
            cell.font = Font(name="Calibri", size=11, bold=False)
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

    # Ширина столбцов из исходного файла (если есть)
    try:
        active_dims = wb.active.column_dimensions
    except Exception:
        active_dims = {}

    for col in sheet.columns:
        col_letter = col[0].column_letter
        if col_letter in active_dims:
            sheet.column_dimensions[col_letter].width = active_dims[col_letter].width

    # Автоширина, если надо
    if apply_auto_width:
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                except AttributeError:
                    continue
            sheet.column_dimensions[column].width = max_length + 2


def _get_series(df: pd.DataFrame, colname: str) -> pd.Series:
    """Безопасно вернуть Series по имени колонки.
    Если дубль — берём первую колонку; если нет колонки — пустая Series нужной длины."""
    if colname not in df.columns:
        return pd.Series([pd.NA] * len(df), index=df.index, dtype="object")
    obj = df[colname]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj


def _normalize_number_label(x) -> str:
    """' 9 ' / '9.0' / '09' -> '9'; иначе вернуть как есть."""
    s = str(x).strip()
    m = re.fullmatch(r"0*(\d+)(?:\.0+)?", s)
    return m.group(1) if m else s


def _normalize_numbered_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Нормализует имена колонок-номеров (убирает '.0', лидирующие нули) и
    выкидывает дубли."""
    df = df.copy()
    # Создаём Index — это важно, чтобы .duplicated() существовал и PyCharm не ругался
    new_cols = [_normalize_number_label(c) for c in df.columns]
    df.columns = pd.Index(new_cols)
    if df.columns.duplicated().any():
        df = df.loc[:, ~df.columns.duplicated()]
    return df


# Недопустимые символы в названиях листов Excel: : \ / ? * [ ]
_INVALID_SHEET_CHARS = re.compile(r"[:\\/*?\[\]]")


def _sanitize_sheet_title(raw: str, used: set) -> str:
    """Делает допустимое и уникальное имя листа (<=31 символ), убирает
    недопустимые знаки."""
    base = _INVALID_SHEET_CHARS.sub(" - ", str(raw)).strip() or "Лист"
    title = base[:31]
    if title in used:
        i = 2
        while True:
            suffix = f" ({i})"
            allowed = 31 - len(suffix)
            cand = (base[:allowed] if allowed > 0 else "") + suffix
            if cand not in used:
                title = cand
                break
            i += 1
    used.add(title)
    return title


# === Режим "Перечень объектов" (ваш исходный) ===
def run_object_check(uploaded_file) -> None:
    data_frame = pd.read_excel(uploaded_file, header=0, skiprows=[1])
    data_frame.columns = data_frame.columns.str.strip()
    st.success(f"Файл загружен: {uploaded_file.name}")

    subdivisions_column = (
        "Наименование соответствующего органа государственного пожарного надзора "
        "(ОНД и ПР - отдел надзорной деятельности и профилактической работы)"
    )
    if subdivisions_column not in data_frame.columns:
        st.error(f"В файле нет столбца: {subdivisions_column}")
        return

    data_frame[subdivisions_column] = data_frame[
        subdivisions_column
    ].str.replace(r"^ОНД и ПР по |^ОНД и ПР ", "", regex=True)

    all_subdivisions = (
        data_frame[subdivisions_column]
        .dropna()
        .astype(str)
        .map(lambda x: x.strip())
        .loc[lambda s: s != ""]
        .unique()
    )

    selected_subdivisions = st.multiselect(
        "Выберите подразделение для проверки (или оставьте все):",
        options=all_subdivisions,
        default=list(all_subdivisions),
    )

    date_column = "Дата присвоения категории риска"
    if date_column in data_frame.columns:
        data_frame[date_column] = data_frame[date_column].apply(
            lambda x: pd.to_datetime(x, format="%d.%m.%Y", errors="coerce")
            if not (isinstance(x, str) and x == "Не присвоена")
            else x
        )

    def make_error_conditions(df: pd.DataFrame) -> dict:
        return {
            "Категория риска не присвоена": df["Присвоенная категория риска"] == "Не присвоена",
            "Не пересчитана категория риска после обновления базовых показателей": (
                    (df[date_column]
                     .apply(lambda x: x if isinstance(x, (datetime, pd.Timestamp)) else pd.NaT)
                     < pd.to_datetime("2025-05-20"))
                    & ~(
                    (df["Функциональное назначение пожарных отсеков"]
                     == "Многоквартирный жилой дом высотой до 28 метров")
                    & (df["Присвоенная категория риска"] == "Умеренный риск")
            )
                    & ~(df["Присвоенная категория риска"] == "Не присвоена")
            ),
            "Отсутствуют собственники": df.get(
                "Количество собственников", pd.Series(1, index=df.index)
            )
                                        == 0,
            "Не внесены ИНН у КЛ": df.get("ИНН", pd.Series("")).notna()
                                   & df.get("ИНН", pd.Series("")).astype(str).str.contains("-"),
            "Функционалы содержат удаленные значения": df.get(
                "Функциональное назначение пожарных отсеков", pd.Series("")
            )
            .astype(str)
            .str.contains("DEL"),
            "Не определен тип, вид, подвид ЕРВК либо тип выбран «Результаты деятельности»": df.get(
                "Тип (ЕРВК)", pd.Series("")
            ).isin(["-", "Результаты деятельности"]),
        }

    error_conditions = make_error_conditions(data_frame)

    mkd_28m = data_frame[
        (data_frame.get("Функциональное назначение пожарных отсеков", pd.Series("")) ==
         "Многоквартирный жилой дом высотой до 28 метров")
        & (data_frame.get("Присвоенная категория риска", pd.Series("")) == "Умеренный риск")
        ]
    st.info(f"**Количество МКД до 28 м умеренного риска:** {mkd_28m.shape[0]}")

    if st.button("Провести проверку и скачать отчёт"):
        with st.spinner("Обрабатываем файл..."):
            filtered_data = data_frame[
                data_frame[subdivisions_column].isin(selected_subdivisions)
            ]

            summary_data = []
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
                    # Сводка по ошибкам
                    for error_name, condition in error_conditions.items():
                        error_data = filtered_data[condition]
                        error_count = error_data.shape[0]

                        grouped = (
                            error_data.groupby(subdivisions_column)
                            .size()
                            .reset_index(name="Количество ошибок")
                        )
                        grouped.columns = ["Подразделение", "Количество ошибок"]

                        grouped["Подразделение"] = grouped[
                            "Подразделение"
                        ].astype(str).str.replace("ОНД и ПР по ", "", regex=False)
                        grouped["Подразделение"] = grouped[
                            "Подразделение"
                        ].str.replace("ОНД и ПР", "", regex=False)
                        grouped = grouped.sort_values(
                            by="Количество ошибок", ascending=False
                        )
                        grouped = grouped[grouped["Количество ошибок"] > 0]

                        summary_data.append([error_name, "в ОНД и ПР по", error_count])
                        for _, row in grouped.iterrows():
                            summary_data.append(
                                [None, row["Подразделение"], row["Количество ошибок"]]
                            )

                    summary_df = pd.DataFrame(
                        summary_data, columns=["Тип ошибки", "Подразделение", "Количество ошибок"]
                    )
                    summary_df.to_excel(writer, sheet_name="Выявленные замечания", index=False)
                    ws_summary = writer.book["Выявленные замечания"]
                    apply_style_to_sheet(ws_summary, writer.book, apply_auto_width=True)

                    # Для каждого типа ошибки создаем лист с объектами
                    used_titles = {"Выявленные замечания"}
                    for error_name, condition in error_conditions.items():
                        error_data = filtered_data[condition]
                        if not error_data.empty:
                            error_data_out = fix_datetime_columns(error_data)

                            # Форматируем второй столбец как текст, чтобы не было научной нотации
                            if error_data_out.shape[1] > 1:
                                col_name_b = error_data_out.columns[1]
                                error_data_out[col_name_b] = error_data_out[col_name_b].astype(str)

                            sheet_title = _sanitize_sheet_title(error_name, used_titles)
                            error_data_out.to_excel(writer, sheet_name=sheet_title, index=False)
                            ws = writer.book[sheet_title]

                            # Вставляем строку с нумерацией (вторая строка)
                            cols_count = error_data_out.shape[1]
                            ws.insert_rows(2)
                            for i in range(cols_count):
                                ws.cell(row=2, column=i + 1).value = i + 1

                            # Форматирование ячеек
                            for row in ws.iter_rows():
                                for cell in row:
                                    cell.font = Font(name="Times New Roman", size=8)
                                    cell.alignment = Alignment(
                                        horizontal="center", vertical="center", wrap_text=True
                                    )
                                    cell.border = Border()

                            # Формат второго столбца как текст
                            if cols_count >= 2:
                                for row in range(1, ws.max_row + 1):
                                    ws.cell(row=row, column=2).number_format = "@"

                            ws.row_dimensions[1].height = 82.1
                            ws.row_dimensions[2].height = 10.2
                            for row in range(3, ws.max_row + 1):
                                ws.row_dimensions[row].height = 50.0

                            # Ширина столбцов - жёстко заданная
                            column_widths = {
                                1: 7.22, 2: 12.22, 3: 47.22, 4: 34.22, 5: 29.22,
                                6: 20.22, 7: 12.22, 8: 34.22, 9: 12.22, 10: 39.22,
                                11: 16.22, 12: 16.22, 13: 29.22, 14: 9.22, 15: 14.22,
                                16: 37.22, 17: 24.22, 18: 16.22, 19: 29.22, 20: 29.22,
                                21: 11.22, 22: 11.22, 23: 23.22,
                            }
                            for col_idx, width in column_widths.items():
                                col_letter = get_column_letter(col_idx)
                                ws.column_dimensions[col_letter].width = width

        st.success("Отчёт готов!")
        with open(tmp.name, "rb") as file:
            st.download_button("Скачать Excel-отчёт", file, "Проверка_перечня.xlsx")

        st.subheader("Общий отчёт по ошибкам:")
        st.dataframe(summary_df)


# === Режим "Проверка КНМ" (исправлен доступ к колонкам) ===
def check_erknm(data: pd.DataFrame) -> pd.Series:
    col9 = _get_series(data, "9")
    col28 = _get_series(data, "28")
    return (
            col9.isin(
                [
                    "Выездная проверка",
                    "Инспекционный визит",
                    "Рейдовый осмотр",
                    "Документарная проверка",
                ]
            )
            & col28.isnull()
    )


def check_predpisaniya(data: pd.DataFrame) -> pd.Series:
    col9 = _get_series(data, "9")
    col14 = _get_series(data, "14")
    col12 = _get_series(data, "12")
    col34 = pd.to_numeric(_get_series(data, "34"), errors="coerce").fillna(0)
    return (
            col9.isin(
                [
                    "Выездная проверка",
                    "Инспекционный визит",
                    "Рейдовый осмотр",
                    "Документарная проверка",
                ]
            )
            & (col14 == "Да")
            & (col12 == "Завершена")
            & (col34 == 0)
    )


def check_no_zaversheno(data: pd.DataFrame) -> pd.Series:
    s7 = (
        _get_series(data, "7")
        .astype(str)
        .str.replace("'", "", regex=False)
        .str.strip()
    )
    s7 = s7.mask(s7 == "-", pd.NA)
    s7 = pd.to_datetime(s7, dayfirst=True, errors="coerce")
    today = pd.to_datetime("today").normalize()
    col12 = _get_series(data, "12")
    status_filter = ~col12.isin(["Завершена", "Не согласовано"])
    date_filter = s7.isna() | (s7 < today)
    return status_filter & date_filter


def check_signed_acts(data: pd.DataFrame) -> pd.Series:
    col12 = _get_series(data, "12")
    col9 = _get_series(data, "9")
    col32 = pd.to_numeric(_get_series(data, "32"), errors="coerce").fillna(0)
    col33 = pd.to_numeric(_get_series(data, "33"), errors="coerce").fillna(0)
    return (
            col12.isin(["Завершена", "Подготовка результатов"])
            & col9.isin(
        [
            "Выездная проверка",
            "Инспекционный визит",
            "Рейдовый осмотр",
            "Документарная проверка",
        ]
    )
            & (col32 == 0)
            & (col33 == 0)
    )


def check_mesto_sostavleniya_akta(data: pd.DataFrame) -> pd.Series:
    col12 = _get_series(data, "12")
    col9 = _get_series(data, "9")
    col16 = _get_series(data, "16")
    return (
            (col12 == "Завершена")
            & col9.isin(
        [
            "Выездная проверка",
            "Инспекционный визит",
            "Рейдовый осмотр",
            "Документарная проверка",
        ]
    )
            & col16.isnull()
    )


def run_knm_check(uploaded_file) -> None:
    # читаем заголовок и данные с 3-уровневым MultiIndex
    original_header = pd.read_excel(uploaded_file, header=None, nrows=3, engine="openpyxl")
    df = pd.read_excel(uploaded_file, header=[0, 1, 2], engine="openpyxl")

    # используем 3-й уровень (номера), нормализуем и убираем дубли
    df.columns = df.columns.get_level_values(2)
    df = _normalize_numbered_columns(df)

    # столбец подразделений
    group_col = "18" if "18" in df.columns else df.columns[0]

    # список подразделений
    group_series = _get_series(df, group_col)
    subdivisions = (
        group_series.dropna()
        .astype(str)
        .str.strip()
        .str.replace(r"^ОНД и ПР по |^ОНД и ПР ", "", regex=True)
        .loc[lambda s: s != ""]
        .unique()
    )

    selected_subdivisions = st.multiselect(
        "Выберите подразделение для проверки (или оставьте все):",
        options=sorted(subdivisions),
        default=sorted(subdivisions),
    )

    if not selected_subdivisions:
        st.warning("Выберите хотя бы одно подразделение для запуска проверки.")
        return

    if st.button("Провести проверку и скачать отчёт"):
        with st.spinner("Проверка данных..."):
            # Фильтрация по выбранным подразделениям
            clean_group = (
                _get_series(df, group_col)
                .astype(str)
                .str.strip()
                .str.replace(r"^ОНД и ПР по |^ОНД и ПР ", "", regex=True)
            )
            df[group_col] = clean_group
            df = df[df[group_col].isin(selected_subdivisions)]

            # Проверки
            erk_nm_errors = check_erknm(df)
            predpisaniya_errors = check_predpisaniya(df)
            no_zaversheno_errors = check_no_zaversheno(df)
            signed_acts_errors = check_signed_acts(df)
            mesto_sostavleniya_akta_errors = check_mesto_sostavleniya_akta(df)

            error_types = [
                ("Не заполнены номера ЕРКНМ", erk_nm_errors),
                ("Не подписаны предписания", predpisaniya_errors),
                ("Не завершены КНМ", no_zaversheno_errors),
                ("Не подписаны акты", signed_acts_errors),
                ("Не заполнены места составления актов", mesto_sostavleniya_akta_errors),
            ]

            errors_data = []
            for error_name, error_series in error_types:
                error_rows = df[error_series].copy()
                error_rows["Тип ошибки"] = error_name
                error_counts = error_rows.groupby(group_col)["Тип ошибки"].count().reset_index()
                error_counts.columns = ["Подразделение", "Количество ошибок"]

                error_counts["Подразделение"] = error_counts["Подразделение"].replace(
                    {"ОНД и ПР по ": "", "ОНД и ПР ": ""}, regex=True
                )
                error_counts = error_counts.sort_values(by="Количество ошибок", ascending=False)

                error_type_row = pd.DataFrame(
                    [[error_name, "в ОНД и ПР по:", error_counts["Количество ошибок"].sum()]],
                    columns=["Тип ошибки", "Подразделение", "Количество ошибок"],
                )
                error_counts = pd.concat([error_type_row, error_counts], ignore_index=True)

                errors_data.append(error_counts)
                errors_data.append(pd.DataFrame([["", "", ""]], columns=["Тип ошибки", "Подразделение",
                                                                         "Количество ошибок"]))

            final_data = pd.concat(errors_data, ignore_index=True)

            wb = load_workbook(uploaded_file)
            directory = os.path.dirname(uploaded_file.name)
            filename = os.path.basename(uploaded_file.name)
            new_filename = f"Общий_отчет_ошибки_{filename}"
            new_file_path = os.path.join(directory, new_filename)

            with pd.ExcelWriter(new_file_path, engine="openpyxl") as writer:
                final_data.to_excel(writer, sheet_name="Выявленные замечания", index=False)
                sheet = writer.sheets["Выявленные замечания"]
                apply_style_to_sheet(sheet, wb, apply_auto_width=True)

                def add_sheet_with_header(sheet_name: str, error_series):
                    error_data = df[error_series].copy()

                    if "7" in error_data.columns:
                        error_data["7"] = pd.to_datetime(error_data["7"], dayfirst=True, errors="coerce").dt.strftime(
                            "%d.%m.%y")

                    ws = writer.book.create_sheet(title=sheet_name)

                    for row_idx in range(3):
                        for col_idx, value in enumerate(original_header.iloc[row_idx], start=1):
                            ws.cell(row=row_idx + 1, column=col_idx, value=value)

                    for r_idx, row in enumerate(error_data.itertuples(index=False), start=4):
                        for c_idx, val in enumerate(row, start=1):
                            ws.cell(row=r_idx, column=c_idx, value=val)

                    apply_style_to_sheet(ws, wb, apply_auto_width=False)

                add_sheet_with_header("Не заполнены номера ЕРКНМ", erk_nm_errors)
                add_sheet_with_header("Не подписаны предписания", predpisaniya_errors)
                add_sheet_with_header("Не завершены КНМ", no_zaversheno_errors)
                add_sheet_with_header("Не подписаны акты", signed_acts_errors)
                add_sheet_with_header("Не заполнены места составления актов", mesto_sostavleniya_akta_errors)

            st.success(f"Отчёт готов: {new_filename}")

            with open(new_file_path, "rb") as f:
                st.download_button("Скачать файл с ошибками КНМ", f, file_name=new_filename)


# === Режим "Перечень объектов (по блокам)" — по номерам столбцов, ускоренный ===
def run_object_check_by_blocks_numbered(uploaded_file) -> None:
    # 1-я строка — имена, 2-я — номера
    original_header = pd.read_excel(uploaded_file, header=None, nrows=2, engine="openpyxl")
    df = pd.read_excel(uploaded_file, header=[0, 1], engine="openpyxl")

    # оставляем в именах колонок именно номера из 2-й строки, нормализуем и убираем дубли
    df.columns = df.columns.get_level_values(1)
    df = _normalize_numbered_columns(df)

    st.success(f"Файл загружен: {uploaded_file.name}")

    # Опция для ускорения (по умолчанию только сводка)
    only_summary = st.checkbox("Только сводка (ускорить выгрузку)", value=True)

    # === Оставляем только строки с данными (начиная с третьей) ===
    df_data = df.iloc[2:].copy()

    # 5-й столбец — подразделение/«кого проверять»
    group_col = "5"
    group_series = _get_series(df_data, group_col).astype(str).str.strip().str.replace(
        r"^ОНД и ПР по |^ОНД и ПР ", "", regex=True
    )
    df_data[group_col] = group_series

    subdivisions = (
        group_series.dropna().astype(str).str.strip().loc[lambda s: s != ""].unique()
    )
    selected_subdivisions = st.multiselect(
        "Выберите подразделения для проверки (или оставьте все):",
        options=sorted(subdivisions),
        default=sorted(subdivisions),
    )
    if not selected_subdivisions:
        st.warning("Выберите хотя бы одно подразделение.")
        return

    filtered = df_data[df_data[group_col].isin(selected_subdivisions)].copy()

    # Приведение типов и серии для критериев
    cutoff_date = pd.Timestamp("2025-05-20")
    s7_txt = _get_series(filtered, "7").astype(str).str.strip()
    s8_txt = _get_series(filtered, "8").astype(str).str.strip()
    s9_dt = pd.to_datetime(
        _get_series(filtered, "9").astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA}),
        dayfirst=True,
        errors="coerce",
    )
    s10_txt = _get_series(filtered, "10").astype(str).str.strip()
    s11_txt = _get_series(filtered, "11").astype(str).str.strip()
    s12_txt = _get_series(filtered, "12").astype(str).str.strip()
    s13_txt = _get_series(filtered, "13").astype(str).str.strip().str.lower()
    s14_num = pd.to_numeric(_get_series(filtered, "14"), errors="coerce")
    s15_num = pd.to_numeric(_get_series(filtered, "15"), errors="coerce")
    s20_num = pd.to_numeric(_get_series(filtered, "20"), errors="coerce")
    s23_txt = _get_series(filtered, "23").astype(str).str.strip()

    # Новая проверка для пустых нумерованных пунктов в столбце 25
    def check_numbered_column_empty(series: pd.Series) -> pd.Series:
        mask = pd.Series(False, index=series.index)
        for idx, val in series.items():
            if pd.isna(val):
                continue
            lines = str(val).splitlines()
            for line in lines:
                # Если есть цифра + точка, но текст отсутствует
                if re.fullmatch(r"\s*\d+\.\s*", line):
                    mask.at[idx] = True
                    break
        return mask

    # Маски ошибок
    error_conditions = {
        "Категория риска не присвоена (столбец 7 пусто)": s7_txt.isin(["", "nan", "None"]),
        "Не выбран класс ФПО (столбец 8 пусто)": s8_txt.isin(["", "nan", "None"]),
        "Категория риска не пересчитана (столбец 9 дата < 20.05.2025 (с исключением МКД))": (
                (s9_dt.notna() & (s9_dt < cutoff_date))
                & ~(s23_txt == "1. Многоквартирный жилой дом высотой до 28 метров")
        ),
        "Не выбран вид собственности (столбец 10 пусто)": s10_txt.isin(["", "nan", "None"]),
        "Не выбрана многофункциональность здания (столбец 11 пусто)": s11_txt.isin(["", "nan", "None"]),
        "Не выбрано наличие границы с лесными участками (столбец 12 пусто)": s12_txt.isin(["", "nan", "None"]),
        "Не выбран расчет индекса индивидуализации (столбец 13 - Нет)": (s13_txt == "нет"),
        "Площадь не выбрана или выбрана неверно (столбец 14 пусто, 0, меньше 10, больше 999999)": (
                _get_series(filtered, "14").astype(str).str.strip().isin(["", "nan", "None", "0"])
                | s14_num.isna()
                | (s14_num < 10)
                | (s14_num > 999_999)
        ),
        "Этажность не выбрана или выбрана неверно (столбец 15 пусто, 0, больше 22)": (
                _get_series(filtered, "15").astype(str).str.strip().isin(["", "nan", "None"])
                | s15_num.fillna(0).eq(0)
                | (s15_num > 22)
        ),
        "Отсутствуют собственники (столбец 20 - 0)": s20_num.fillna(pd.NA).replace({pd.NA: -1}).eq(0),
        "Не внесены ИНН у КЛ (столбец 22 содержит -)": check_numbered_column_empty(_get_series(filtered, "22")),
        "Не выбрано функциональное назначение (столбец 23 не всё заполнено)": check_numbered_column_empty(_get_series(
            filtered, "23")),
        "Не выбрано наличие автоматической пожарной сигнализации (АПС) (столбец 24 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "24")),
        "Не выбрано дублирование сигнала системы АПС на пульт подразделений ПО (столбец 25 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "25")),
        "Не выбрано наличие СОУЭ (столбец 26 не всё заполнено)": check_numbered_column_empty(_get_series(filtered, "26")
                                                                                             ),
        "Не выбрано наличие ВПВ (столбец 27 не всё заполнено)": check_numbered_column_empty(_get_series(filtered, "27"))
        ,
        "Не выбрано наличие АУПТ (столбец 28 не всё заполнено)": check_numbered_column_empty(_get_series(filtered, "28")
                                                                                             ),
        "Не выбрано наличие противодымной вентиляции (столбец 29 не всё заполнено)": check_numbered_column_empty(
            _get_series(filtered, "29")),
        "Не выбрано наличие электроснабжения (столбец 45 не всё заполнено)": check_numbered_column_empty(_get_series(
            filtered, "45")),
        "Не выбрана степень огнестойкости (столбец 49 не всё заполнено)": check_numbered_column_empty(_get_series(
            filtered, "49")),
        "Не выбрано наличие на объекте ПО, обеспеченной ПТВ (столбец 50 не всё заполнено)": check_numbered_column_empty(
            _get_series(filtered, "50")),
        "Не выбрано наличие МГН (столбец 51 не всё заполнено)": check_numbered_column_empty(_get_series(filtered, "51"))
        ,
        "Не выставлен круглосуточный режим работы охранной организации (столбец 52 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "52")),
        "Не выбрана категория объекта по потенциальной РО (столбец 53 не всё заполнено)": check_numbered_column_empty(
            _get_series(filtered, "53")),
        "Не выбрано наличие людей в селитебной зоне (столбец 54 не всё заполнено)": check_numbered_column_empty(
            _get_series(filtered, "54")),
        "Не выбрано наличие круглосуточного пребывания или проживания МГН (столбец 55 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "55")),
        "Не выбрано наличие круглосуточного пребывания людей (столбец 56 не всё заполнено)": check_numbered_column_empty
        (_get_series(filtered, "56")),
        "Не выбрана категория по пожарной и взрывопожарной опасности (столбец 57 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "57")),
        "Не выбрано нахождение в рабочее время более 10 человек МГН (столбец 58 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "58")),
        "Не выбрано наличие открытых лестниц и (или) многосветных пространств (столбец 59 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "59")),
        "Не выбрана категория НУ по пожарной и взрывопожарной опасности (столбец 60 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "60")),
        "Не выбрано наличие постоянных рабочих мест (столбец 61 не всё заполнено)": check_numbered_column_empty(
            _get_series(filtered, "61")),
        "Не выбрано наличие привлеченной к охране организации (столбец 62 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "62")),
        "Не выбрана площадь (столбец 63 не всё заполнено)": check_numbered_column_empty(_get_series(filtered, "63")),
        "Не выбрана высотность (столбец 64 не всё заполнено)": check_numbered_column_empty(_get_series(filtered, "64")),
        "Не выбрано максимальное количество одновременно находящихся людей в здании (столбец 65 не всё заполнено)":
            check_numbered_column_empty(_get_series(filtered, "65")),
    }

    if st.button("Провести проверку (по номерам столбцов) и скачать отчёт"):
        with st.spinner("Формируем отчёт..."):
            summary_data = []
            used_titles = {"Выявленные замечания"}

            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
                    # Сводка
                    for err_name, mask in error_conditions.items():
                        mask = mask.fillna(False)
                        err_df = filtered[mask]

                        error_count = int(err_df.shape[0])

                        grouped = err_df.groupby(group_col).size().reset_index(name="Количество ошибок")
                        grouped.columns = ["Подразделение", "Количество ошибок"]

                        grouped["Подразделение"] = grouped["Подразделение"].astype(str).str.replace(
                            "ОНД и ПР по ", "", regex=False
                        )
                        grouped["Подразделение"] = grouped["Подразделение"].str.replace(
                            "ОНД и ПР", "", regex=False
                        )
                        if not grouped.empty:
                            grouped = grouped.sort_values(by="Количество ошибок", ascending=False)
                            grouped = grouped[grouped["Количество ошибок"] > 0]

                        summary_data.append([err_name, "в ОНД и ПР по", error_count])
                        for _, row in grouped.iterrows():
                            summary_data.append([None, row["Подразделение"], int(row["Количество ошибок"])])

                    summary_df = pd.DataFrame(
                        summary_data, columns=["Тип ошибки", "Подразделение", "Количество ошибок"]
                    )
                    summary_df.to_excel(writer, sheet_name="Выявленные замечания", index=False)
                    ws_summary = writer.sheets["Выявленные замечания"]
                    apply_style_to_sheet(ws_summary, writer.book, apply_auto_width=True)

                    # Детали
                    if not only_summary:
                        for err_name, mask in error_conditions.items():
                            mask = mask.fillna(False)
                            err_df = filtered[mask]
                            if err_df.empty:
                                continue

                            out = err_df.copy()
                            if "9" in out.columns:
                                out["9"] = pd.to_datetime(out["9"], dayfirst=True, errors="coerce").dt.strftime(
                                    "%d.%m.%Y")

                            title = _sanitize_sheet_title(err_name, used_titles)
                            ws = writer.book.create_sheet(title=title)

                            # Вставляем первые две строки из исходного файла
                            for row_idx in range(2):
                                for col_idx, value in enumerate(original_header.iloc[row_idx], start=1):
                                    ws.cell(row=row_idx + 1, column=col_idx, value=value)

                            # Вставляем данные с 3-й строки
                            for r_idx, row in enumerate(out.itertuples(index=False), start=3):
                                for c_idx, val in enumerate(row, start=1):
                                    cell = ws.cell(row=r_idx, column=c_idx)
                                    # Для второго столбца записываем строго как строку
                                    if c_idx == 2 and val is not None:
                                        cell.value = str(val)
                                    else:
                                        cell.value = val

                            # Форматирование ячеек
                            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                                for cell in row:
                                    cell.font = Font(name="Times New Roman", size=12)
                                    cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

                            # Формат второго столбца как текст
                            for row in range(3, ws.max_row + 1):
                                ws.cell(row=row, column=2).number_format = "@"

                            # Ширина столбцов
                            column_widths = {i: 29.22 for i in range(3, 66)}
                            column_widths.update({1: 7.22, 2: 17.22, 9: 34.22})
                            for col_idx, width in column_widths.items():
                                col_letter = get_column_letter(col_idx)
                                ws.column_dimensions[col_letter].width = width

                st.success("Отчёт по номерам столбцов готов!")
                with open(tmp.name, "rb") as file:
                    st.download_button(
                        "Скачать Excel-отчёт (по номерам)",
                        file,
                        "Проверка_перечня_по_номерам.xlsx",
                    )

        st.subheader("Общий отчёт по ошибкам (по номерам):")
        st.dataframe(summary_df)


# === Главная логика ===
uploaded_file = st.file_uploader("Загрузите Excel-файл", type=["xlsx"])
if uploaded_file:
    if choice == "Перечень объектов":
        run_object_check(uploaded_file)
    elif choice == "Перечень объектов (по блокам)":
        run_object_check_by_blocks_numbered(uploaded_file)
    else:
        run_knm_check(uploaded_file)
