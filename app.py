import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import json
import os

# --- Путь к файлу шаблона ---
TEMPLATE_FILE = "агрегация_шаблон.json"

st.set_page_config(page_title="Гибкая агрегация Excel", layout="wide")
st.title("Гибкая агрегация данных из Excel")

# --- Инициализация session_state ---
if 'group_keys' not in st.session_state:
    st.session_state.group_keys = []
if 'value_cols' not in st.session_state:
    st.session_state.value_cols = []
if 'agg_settings' not in st.session_state:
    st.session_state.agg_settings = {}
if 'result_df_full' not in st.session_state:
    st.session_state.result_df_full = None
if 'show_sum_products' not in st.session_state:
    st.session_state.show_sum_products = True
if 'show_sum_weights' not in st.session_state:
    st.session_state.show_sum_weights = True

# --- Функции для работы с шаблоном ---
def save_template():
    template = {
        "group_keys": st.session_state.group_keys,
        "value_cols": st.session_state.value_cols,
        "agg_settings": st.session_state.agg_settings,
        "show_sum_products": st.session_state.show_sum_products,
        "show_sum_weights": st.session_state.show_sum_weights
    }
    with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(template, f, ensure_ascii=False, indent=2)
    st.success(f"Шаблон сохранён в файл: {TEMPLATE_FILE}")

def load_template():
    if os.path.exists(TEMPLATE_FILE):
        with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
            template = json.load(f)
        st.session_state.group_keys = template.get("group_keys", [])
        st.session_state.value_cols = template.get("value_cols", [])
        st.session_state.agg_settings = template.get("agg_settings", {})
        st.session_state.show_sum_products = template.get("show_sum_products", True)
        st.session_state.show_sum_weights = template.get("show_sum_weights", True)
        st.success("Шаблон загружен!")
        st.rerun()  # обновляем страницу
    else:
        st.warning("Файл шаблона не найден.")

# --- Кнопки управления шаблоном ---
col_save, col_load = st.columns(2)
with col_save:
    if st.button("💾 Сохранить шаблон"):
        save_template()
with col_load:
    if st.button("📂 Загрузить шаблон"):
        load_template()

# --- Загрузка файла ---
uploaded_file = st.file_uploader("Загрузите Excel-файл (с заголовками в первой строке)", type=["xlsx", "xls"])

if uploaded_time := st.session_state.get("uploaded_file_time"):
    pass  # можно использовать для отслеживания смены файла

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        if df.empty:
            st.warning("Файл пуст.")
            st.stop()
        columns = list(df.columns)

        # --- 1. Выбор группировки ---
        st.subheader("1. Выберите иерархию группировки")
        group_keys = st.multiselect(
            "Ключи группировки (порядок важен!)",
            options=columns,
            default=st.session_state.group_keys,
            key="group_keys_input"
        )
        st.session_state.group_keys = group_keys

        # --- 2. Выбор столбцов для агрегации ---
        st.subheader("2. Настройте агрегацию для каждого столбца")
        value_cols = st.multiselect(
            "Столбцы для агрегации",
            options=columns,
            default=st.session_state.value_cols,
            key="value_cols_input"
        )
        st.session_state.value_cols = value_cols

        # --- Галочки для компонентов средневзвешенного ---
        st.subheader("3. Детализация средневзвешенного")
        col_check1, col_check2 = st.columns(2)
        with col_check1:
            show_sum_products = st.checkbox(
                "Показывать сумму произведений (показатель × вес)",
                value=st.session_state.show_sum_products,
                key="show_sum_products_checkbox"
            )
        with col_check2:
            show_sum_weights = st.checkbox(
                "Показывать сумму весов",
                value=st.session_state.show_sum_weights,
                key="show_sum_weights_checkbox"
            )
        st.session_state.show_sum_products = show_sum_products
        st.session_state.show_sum_weights = show_sum_weights

        if not group_keys:
            st.info("Выберите хотя бы один столбец для группировки.")
            st.stop()
        if not value_cols:
            st.info("Выберите хотя бы один столбец для агрегации.")
            st.stop()

        # --- Компактная настройка агрегации ---
        AGG_TYPES = ["Сумма", "Количество", "Среднее", "Медиана", "Средневзвешенное"]
        agg_settings = {}

        for col in value_cols:
            col_label, col_type, col_weight = st.columns([1.5, 1.2, 1.3])
            with col_label:
                st.markdown(f"**{col}**")
            with col_type:
                default_type = st.session_state.agg_settings.get(col, ("Сумма", None))[0]
                agg_type = st.selectbox(
                    "Тип",
                    options=AGG_TYPES,
                    index=AGG_TYPES.index(default_type) if default_type in AGG_TYPES else 0,
                    key=f"type_{col}",
                    label_visibility="collapsed"
                )
            weight_col = None
            with col_weight:
                if agg_type == "Средневзвешенное":
                    weight_options = columns
                    default_weight = st.session_state.agg_settings.get(col, (None, None))[1]
                    if default_weight in weight_options:
                        weight_col = st.selectbox(
                            "Вес",
                            options=weight_options,
                            index=weight_options.index(default_weight),
                            key=f"weight_{col}",
                            label_visibility="collapsed"
                        )
                    else:
                        weight_col = st.selectbox(
                            "Вес",
                            options=weight_options,
                            key=f"weight_{col}",
                            label_visibility="collapsed"
                        )
                    if weight_col == col:
                        st.warning("Вес ≠ значение")
                else:
                    st.empty()
            agg_settings[col] = (agg_type, weight_col)

        st.session_state.agg_settings = agg_settings

        # --- Кнопка расчёта ---
        if st.button("📊 Выполнить агрегацию"):
            df_clean = df.copy()
            grouped = df_clean.groupby(group_keys, dropna=False)
            result_df = grouped.size().reset_index().drop(columns=0)

            for col, (agg_type, weight_col) in agg_settings.items():
                if agg_type == "Сумма":
                    result_df[f"{col}_сумма"] = grouped[col].sum(numeric_only=True).values

                elif agg_type == "Количество":
                    result_df[f"{col}_количество"] = grouped[col].count().values

                elif agg_type == "Среднее":
                    result_df[f"{col}_среднее"] = grouped[col].mean(numeric_only=True).values

                elif agg_type == "Медиана":
                    result_df[f"{col}_медиана"] = grouped[col].median(numeric_only=True).values

                elif agg_type == "Средневзвешенное":
                    if weight_col is None:
                        st.error(f"Не указан вес для '{col}'.")
                        st.stop()

                    def weighted_mean(group):
                        vals = pd.to_numeric(group[col], errors='coerce')
                        weights = pd.to_numeric(group[weight_col], errors='coerce')
                        mask = vals.notna() & weights.notna()
                        if mask.sum() == 0:
                            return np.nan
                        return np.average(vals[mask], weights=weights[mask])
                    result_df[f"{col}_средневзвешенное_по_{weight_col}"] = grouped.apply(weighted_mean).values

                    if st.session_state.show_sum_products:
                        def sum_products(group):
                            vals = pd.to_numeric(group[col], errors='coerce')
                            weights = pd.to_numeric(group[weight_col], errors='coerce')
                            mask = vals.notna() & weights.notna()
                            return (vals[mask] * weights[mask]).sum()
                        result_df[f"{col}_взвеш_сумма_произведений"] = grouped.apply(sum_products).values

                    if st.session_state.show_sum_weights:
                        def sum_weights(group):
                            weights = pd.to_numeric(group[weight_col], errors='coerce')
                            return weights.sum()
                        result_df[f"{col}_взвеш_сумма_весов"] = grouped.apply(sum_weights).values

            st.session_state.result_df_full = result_df

        # --- Итоги ---
        if st.session_state.result_df_full is not None:
            base_df = st.session_state.result_df_full.copy()
            group_keys = st.session_state.group_keys

            numeric_cols = [
                col for col in base_df.columns 
                if col not in group_keys and pd.api.types.is_numeric_dtype(base_df[col])
            ]

            all_rows = [base_df]

            if len(group_keys) >= 1:
                first_key = group_keys[0]
                subtotal = base_df.groupby(first_key, dropna=False)[numeric_cols].sum().reset_index()
                for key in group_keys[1:]:
                    subtotal[key] = f"Итог по {first_key}"
                all_rows.append(subtotal)

            total_dict = {key: "Итог всего" for key in group_keys}
            total_dict.update({col: base_df[col].sum() for col in numeric_cols})
            total_row = pd.DataFrame([total_dict])
            all_rows.append(total_row)

            result_with_subtotals = pd.concat(all_rows, ignore_index=True)
            st.session_state.result_df_full = result_with_subtotals

        # --- Фильтрация и вывод ---
        if st.session_state.result_df_full is not None:
            st.subheader("4. Фильтрация по группам и ключам")
            filtered_df = st.session_state.result_df_full.copy()
            group_keys = st.session_state.group_keys

            if group_keys:
                filter_cols = st.columns(min(len(group_keys), 5))
                for idx, key in enumerate(group_keys):
                    with filter_cols[idx % len(filter_cols)]:
                        display_series = filtered_df[key].fillna("(пусто)").astype(str)
                        unique_vals = sorted(display_series.unique())
                        selected = st.multiselect(
                            f"Фильтр: {key}",
                            options=unique_vals,
                            default=st.session_state.get(f"filter_{key}", []),
                            key=f"filter_input_{key}"
                        )
                        st.session_state[f"filter_{key}"] = selected
                        if selected:
                            filtered_df = filtered_df[display_series.isin(selected)]

            st.subheader("Результат агрегации с итогами")
            st.dataframe(filtered_df, use_container_width=True)

            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Агрегация')
                return output.getvalue()

            excel_data = to_excel(filtered_df)
            st.download_button(
                label="📥 Скачать результат в Excel",
                data=excel_data,
                file_name="агрегация_результат.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ошибка: {e}")
else:
    st.info("Пожалуйста, загрузите Excel-файл.")