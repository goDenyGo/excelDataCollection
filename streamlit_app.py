import streamlit as st
import pandas as pd
import os

# Функція для завантаження Excel
def load_excel(file):
    return pd.ExcelFile(file)

# Функція для оновлення Excel-файлу
def save_excel(df, sheet_name):
    with pd.ExcelWriter("edited_file.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Функція для генерації унікального ID
def generate_unique_id(df, id_column):
    existing_ids = df[id_column].dropna().astype(str).tolist()
    new_id = 1
    while str(new_id) in existing_ids:
        new_id += 1
    return new_id

# Початок програми
st.title("Редагування файлу Excel")
uploaded_file = st.file_uploader("📂 Завантажте Excel-файл", type=["xls", "xlsx"])

if uploaded_file:
    xls = load_excel(uploaded_file)

    # Вибір листа
    sheet_name = st.selectbox("📑 Оберіть лист", xls.sheet_names, key="selected_sheet")

    if sheet_name:
        if "df" not in st.session_state or st.session_state["last_sheet"] != sheet_name:
            st.session_state["df"] = pd.read_excel(xls, sheet_name=sheet_name)
            st.session_state["last_sheet"] = sheet_name
            st.session_state["selected_index"] = None  # Очистка вибору

        df = st.session_state["df"]

        if not df.empty:
            st.subheader("🔍 Попередній перегляд таблиці")
            st.dataframe(df)  # **Тепер ця таблиця буде оновлюватися після редагування**

            # Вибір перших трьох колонок для пошуку
            if len(df.columns) >= 3:
                col1, col2, col3 = df.columns[:3]

                search1 = st.text_input(f"🔎 {col1}")
                search2 = st.text_input(f"🔎 {col2}")
                search3 = st.text_input(f"🔎 {col3}")

                if st.button("🔍 Пошук записів"):
                    found_rows = df[
                        (df[col1].astype(str).str.contains(search1, na=False, case=False)) &
                        (df[col2].astype(str).str.contains(search2, na=False, case=False)) &
                        (df[col3].astype(str).str.contains(search3, na=False, case=False))
                    ]

                    if not found_rows.empty:
                        st.write("✅ Знайдені записи:")
                        st.dataframe(found_rows)

                        # Формування списку для вибору
                        index_options = found_rows.index.tolist()
                        index_labels = [
                            f"{df.at[idx, col1]} | {df.at[idx, col2]} | {df.at[idx, col3]}" 
                            for idx in index_options
                        ]

                        selected_label = st.selectbox("✏️ Оберіть запис для редагування", index_labels)
                        selected_index = index_options[index_labels.index(selected_label)]
                        st.session_state["selected_index"] = selected_index

                    else:
                        st.warning("⚠️ Жодного запису не знайдено!")

            # Редагування знайденого запису
            if st.session_state.get("selected_index") is not None:
                selected_index = st.session_state["selected_index"]
                edited_values = {}

                for col in df.columns:
                    edited_values[col] = st.text_input(f"{col}", str(df.at[selected_index, col]))

                if st.button("💾 Зберегти зміни"):
                    for col in df.columns:
                        df.at[selected_index, col] = edited_values[col]

                    save_excel(df, sheet_name)
                    st.session_state["df"] = df  # **Оновлення збережених даних**
                    st.success("✅ Зміни збережені!")

                    st.experimental_rerun()  # **Перезапускає програму для оновлення таблиці**

            # Додавання нового запису
            st.subheader("➕ Додати новий запис")
            new_values = {}

            # Генерація нового ID, якщо колонка ID є
            id_column = None
            for col in [col1, col2, col3]:
                if "ID" in col.upper():
                    id_column = col
                    break

            for col in df.columns:
                if col == id_column:
                    new_values[col] = generate_unique_id(df, id_column)
                    st.text_input(f"{col} (автоматично)", str(new_values[col]), disabled=True)
                else:
                    new_values[col] = st.text_input(f"{col} (новий запис)", "")

            if st.button("✅ Додати запис"):
                new_row = pd.DataFrame([new_values])
                df = pd.concat([df, new_row], ignore_index=True)

                save_excel(df, sheet_name)
                st.session_state["df"] = df  # **Оновлення після додавання запису**
                sst.success(f"✅ Новий запис додано! (ID: {new_values.get('ID', 'Немає ID')})")


                st.rerun()  # **Правильний метод перезапуску в нових версіях Streamlit**
                
            # Кнопка завершення роботи
            if st.button("❌ Завершити роботу"):
                if st.confirm("Ви впевнені, що хочете завершити роботу?"):
                    if st.confirm("Зберегти зміни перед виходом?"):
                        save_excel(df, sheet_name)
                        st.success("✅ Дані збережено! Програму буде завершено.")
                    else:
                        st.warning("⚠️ Вихід без збереження!")

                    st.warning("🚪 Програма завершує роботу...")

                    st.session_state.clear()  # Очищення сесії перед виходом
                    os._exit(0)  # 🔴 Примусове завершення програми
            
            
            
          #  if st.button("❌ Завершити роботу"):
          #      save_changes = st.radio("💾 Зберегти зміни перед виходом?", ["Так", "Ні"])
          #
          #      if save_changes == "Так":
          #          save_excel(df, sheet_name)
          #          st.success("✅ Зміни збережені! Завантажте файл перед виходом:")
          #          st.download_button(
          #              label="⬇️ Завантажити оновлений файл",
          #              data=open("edited_file.xlsx", "rb").read(),
          #              file_name="edited_file.xlsx",
          #              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          #          )
          #
          #      st.warning("🚪 Ви вийшли з програми.")
          #      
          #      # **Завершення програми повністю**
          #      st.stop()

else:
    st.info("📂 Завантажте Excel-файл для початку роботи")