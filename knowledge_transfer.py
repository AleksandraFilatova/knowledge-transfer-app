# app.py
# ---------------------------
# Knowledge Transfer App (Google Sheets + локальный fallback)
# Исправлено:
# - запись в Google Sheets: update с A1-диапазоном, правильная конвертация колонки > 'Z'
# - чтение credentials из st.secrets (или из файла)
# - гарантия аркушей Lakes/Reports
# ---------------------------

import streamlit as st
from datetime import datetime
import os
import pandas as pd
import plotly.express as px
from PIL import Image
import base64
import requests
import time
import json

try:
    from openpyxl import load_workbook  # noqa: F401
except ImportError:
    pass

# ===== Google Sheets (новые) =====
try:
    import gspread
    from google.oauth2.service_account import Credentials
    from gspread.utils import rowcol_to_a1
    GOOGLE_SHEETS_AVAILABLE = True
    GS_IMPORT_ERROR = ""
except Exception as e:
    GOOGLE_SHEETS_AVAILABLE = False
    GS_IMPORT_ERROR = str(e)

# ==== CONFIG SECTION ====
# Локальная папка для резервных сохранений
LOCAL_DATA_DIR = os.path.join(os.path.expanduser("~"), "AppData", "Local", "StreamlitData")
os.makedirs(LOCAL_DATA_DIR, exist_ok=True)
EXCEL_FILE_PATH = os.path.join(LOCAL_DATA_DIR, "LakeHouse.xlsx")

# Google Sheets ID (замени на свой при необходимости)
GOOGLE_SHEETS_ID = "19Ge1PiHdeWt0mofW5YkxmectUchGcbclaHNim_XvmFM"
# Чтение из Google Sheets (CSV через gviz) — листы Lakes/Reports
GOOGLE_SHEETS_URL_LAKES = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}/gviz/tq?tqx=out:csv&sheet=Lakes"
GOOGLE_SHEETS_URL_REPORTS = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}/gviz/tq?tqx=out:csv&sheet=Reports"

# ----------------- Утилиты отображения -----------------
def display_image_from_path(image_path, caption=None, width=None):
    try:
        if image_path.startswith(('http://', 'https://')):
            st.image(image_path, caption=caption, width=width)
        elif os.path.exists(image_path):
            image = Image.open(image_path)
            st.image(image, caption=caption, width=width)
        else:
            st.warning(f"⚠️ Зображення не знайдено: {image_path}")
    except Exception as e:
        st.error(f"❌ Помилка при завантаженні зображення: {e}")

def display_image_from_base64(base64_string, caption=None, width=None):
    try:
        image_data = base64.b64decode(base64_string)
        st.image(image_data, caption=caption, width=width)
    except Exception as e:
        st.error(f"❌ Помилка при декодуванні зображення: {e}")

def process_text_with_images(text: str):
    if not text:
        return text
    import re
    image_pattern = r'\[IMAGE:(.*?)\]'
    matches = re.findall(image_pattern, text)
    if matches:
        parts = re.split(image_pattern, text)
        for i, part in enumerate(parts):
            if i % 2 == 0:
                if part.strip():
                    st.markdown(part)
            else:
                image_path = part.strip()
                if image_path.startswith('C:\\') and 'PL-notebook.png' in image_path:
                    github_url = "https://raw.githubusercontent.com/AleksandraFilatova/knowledge-transfer-app/main/Image/Sac-notebook.PNG"
                    display_image_from_path(github_url, width=600)
                elif 'github.com' in image_path and '/blob/' in image_path:
                    raw_url = image_path.replace('github.com', 'raw.githubusercontent.com').replace('/blob/', '/')
                    display_image_from_path(raw_url, width=600)
                else:
                    display_image_from_path(image_path, width=600)
    else:
        st.markdown(text)

# ----------------- Чтение Excel локально -----------------
@st.cache_data(ttl=300)
def load_lakes_and_reports(excel_path):
    try:
        xl = pd.ExcelFile(excel_path, engine='openpyxl')
        available_sheets = xl.sheet_names

        lakes_df = pd.read_excel(xl, 'Lakes', engine='openpyxl') if 'Lakes' in available_sheets else \
                   pd.read_excel(xl, available_sheets[0], engine='openpyxl')

        reports_df = pd.read_excel(xl, 'Reports', engine='openpyxl') if 'Reports' in available_sheets else \
                     pd.DataFrame()

        # названия (уникальные)
        lakes_names = list(lakes_df['LakeHouse'].dropna().unique()) if 'LakeHouse' in lakes_df.columns else list(lakes_df.iloc[:,0].dropna().unique())
        reports_names = list(reports_df.iloc[:,0].dropna().unique()) if not reports_df.empty else []
        return lakes_names, reports_names, lakes_df, reports_df

    except Exception as e:
        st.error(f"❌ Помилка при завантаженні файлу: {e}")
        st.warning("💡 Закрийте файл в Excel, дочекайтесь синхронізації OneDrive, оновіть сторінку.")
        return [], [], None, None

def create_default_excel_file(local_path):
    try:
        default_data = {
            'LakeHouse': [], 'Folder': [], 'Element': [], 'URL': [],
            'Загальна інформація про лейк': [], 'Внесення змін': []
        }
        df = pd.DataFrame(default_data)
        with pd.ExcelWriter(local_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Lakes', index=False)
            df.to_excel(writer, sheet_name='Reports', index=False)
        return True
    except Exception as e:
        st.error(f"❌ Помилка створення файлу: {e}")
        return False

def save_data_to_excel(df, filename, reports_table=None):
    try:
        st.info(f"💾 Резервне локальне збереження: {filename}")
        with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, sheet_name='Lakes', index=False)
            if reports_table is not None and not reports_table.empty:
                reports_table.to_excel(writer, sheet_name='Reports', index=False)
        st.success(f"✅ Локальний файл збережено: {os.path.abspath(filename)}")
        return True, filename
    except PermissionError as e:
        st.error(f"❌ Доступ до файлу: {e}")
        return False, None
    except Exception as e:
        st.error(f"❌ Помилка при локальному збереженні: {type(e).__name__}: {e}")
        return False, None

# ----------------- Аналитика (визуалки) -----------------
def analyze_lakes_data(lakes_df: pd.DataFrame):
    if lakes_df is None or lakes_df.empty:
        return {'total_lakes': 0, 'columns': [], 'missing_data': {}, 'unique_values': {}}
    analysis = {
        'total_lakes': len(lakes_df),
        'columns': list(lakes_df.columns),
        'missing_data': {c: lakes_df[c].isna().sum() for c in lakes_df.columns},
        'unique_values': {}
    }
    for c in lakes_df.columns:
        if lakes_df[c].dtype == 'object':
            analysis['unique_values'][c] = lakes_df[c].value_counts().to_dict()
    return analysis

def create_lakes_visualization(lakes_df):
    if lakes_df is None or lakes_df.empty:
        return None
    charts = {}
    if 'Status' in lakes_df.columns:
        status_counts = lakes_df['Status'].value_counts()
        charts['status_pie'] = px.pie(values=status_counts.values, names=status_counts.index,
                                      title="Розподіл лейків за статусом")
    if 'Update_Frequency' in lakes_df.columns:
        freq = lakes_df['Update_Frequency'].value_counts()
        charts['frequency_bar'] = px.bar(x=freq.index, y=freq.values, title="Частота оновлень лейків")
    if 'Workspace' in lakes_df.columns:
        charts['workspace_treemap'] = px.treemap(lakes_df, path=['Workspace'], title="Розподіл лейків по робочих просторах")
    return charts

def create_lake_details_card(lake_row: pd.Series):
    if lake_row is None or lake_row.empty:
        return "Немає даних про лейк"
    card_html = f"""
    <div style="
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px; border-radius: 10px; color: white; margin: 10px 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    ">
        <h3 style="margin: 0 0 15px 0; color: white;">🏞️ {lake_row.get('LakeHouse', 'Невідомий лейк')}</h3>
    """
    for col in lake_row.index:
        if pd.notna(lake_row[col]) and col not in ['LakeHouse', 'Folder', 'Element', 'Загальна інформація про лейк', 'Внесення змін']:
            display_col_name = {'Type': 'Тип', 'Опис': 'Опис', 'Оновлення': 'Оновлення', 'Особливості': 'Особливості'}.get(col, col)
            card_html += f'<p style="margin:5px 0;"><strong>{display_col_name}:</strong> {lake_row[col]}</p>'
    card_html += "</div>"
    return card_html

# ----------------- Чтение из Google Sheets (CSV) -----------------
def load_from_google_sheets():
    try:
        lakes_df = pd.read_csv(GOOGLE_SHEETS_URL_LAKES)
        try:
            reports_df = pd.read_csv(GOOGLE_SHEETS_URL_REPORTS)
        except Exception:
            reports_df = pd.DataFrame()
        lakes_names = list(lakes_df['LakeHouse'].dropna()) if 'LakeHouse' in lakes_df.columns else list(lakes_df.iloc[:,0].dropna())
        reports_names = list(reports_df.iloc[:,0].dropna()) if not reports_df.empty else []
        return lakes_names, reports_names, lakes_df, reports_df
    except Exception as e:
        st.error(f"❌ Помилка завантаження з Google Sheets (читання): {e}")
        return [], [], None, None

# ----------------- ЗАПИС в Google Sheets (исправленный) -----------------
def _get_gspread_client():
    """
    1) пробуем st.secrets['gcp_service_account'] (dict или JSON-строка)
    2) иначе файл service_account_credentials.json (рядом со скриптом или в домашней папке)
    """
    if not GOOGLE_SHEETS_AVAILABLE:
        raise RuntimeError(f"gspread/google-auth недоступны: {GS_IMPORT_ERROR}")

    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]

    # через st.secrets (рекомендовано для Streamlit Cloud)
    if "gcp_service_account" in st.secrets:
        sa_info = st.secrets["gcp_service_account"]
        if isinstance(sa_info, str):
            sa_info = json.loads(sa_info)
        creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
        return gspread.authorize(creds)

    # файл JSON
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, "service_account_credentials.json"),
        os.path.join(os.path.expanduser("~"), "service_account_credentials.json")
    ]
    for path in candidates:
        if os.path.exists(path):
            creds = Credentials.from_service_account_file(path, scopes=scopes)
            return gspread.authorize(creds)

    raise FileNotFoundError("Не найден ключ сервис-аккаунта: положи JSON в st.secrets['gcp_service_account'] "
                            "или файл service_account_credentials.json рядом со скриптом/в домашней папке.")

def _ensure_worksheet(sh, title, rows=1000, cols=50):
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)

def _update_sheet_with_dataframe(ws, df: pd.DataFrame):
    if df is None or df.empty:
        ws.clear()
        return
    # значения: заголовки + строки; приведение NaN к пустым строкам
    values = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    last_row = len(values)
    last_col = len(values[0]) if values else 1
    end_a1 = rowcol_to_a1(last_row, last_col)   # корректно и после 'Z'
    ws.clear()
    ws.update(f"A1:{end_a1}", values, value_input_option="RAW")

def save_to_google_sheets(df: pd.DataFrame, reports_table: pd.DataFrame | None = None) -> bool:
    try:
        gc = _get_gspread_client()
        sh = gc.open_by_key(GOOGLE_SHEETS_ID)

        # ВАЖНО: поделись таблицей с client_email сервис-аккаунта (Editor)!
        lakes_ws = _ensure_worksheet(sh, "Lakes", rows=max(1000, len(df)+10), cols=max(20, len(df.columns)+2))
        _update_sheet_with_dataframe(lakes_ws, df)

        if reports_table is not None and not reports_table.empty:
            reports_ws = _ensure_worksheet(sh, "Reports",
                                           rows=max(1000, len(reports_table)+10),
                                           cols=max(20, len(reports_table.columns)+2))
            _update_sheet_with_dataframe(reports_ws, reports_table)

        st.success("✅ Дані успішно збережено в Google Sheets!")
        return True

    except gspread.exceptions.APIError as api_err:
        st.error(f"❌ Google API error: {api_err}")
        st.info("🔎 Перевір: 1) сервіс-акаунт має доступ (Editor) до таблиці; 2) ID таблиці вірний; 3) назви листів 'Lakes'/'Reports'.")
        return False
    except FileNotFoundError as cred_err:
        st.error(f"❌ Креденшіали: {cred_err}")
        return False
    except Exception as e:
        st.error(f"❌ Несподівана помилка запису в Google Sheets: {e}")
        return False

# ==================== НАСТРОЙКИ СТОРІНКИ ====================
st.set_page_config(page_title="Knowledge Transfer App", page_icon="🧠", layout="wide", initial_sidebar_state="expanded")

# ==================== НАВІГАЦІЯ ====================
st.sidebar.title("🗂️ Навігація")
st.sidebar.markdown("### Оберіть розділ:")
section = st.sidebar.radio("", ["🏠 Головна", "💧 Оновлення LakeHouses", "📊 Оновлення PowerBI Report", "✏️ Редагування даних", "📞 Контакти та ресурси"])
st.sidebar.markdown("---")
st.sidebar.info(f"📅 Останнє оновлення:\n{datetime.now().strftime('%d.%m.%Y')}")

# Подсказка по кредам (если нет st.secrets и файла)
if not ("gcp_service_account" in st.secrets):
    CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "service_account_credentials.json")
    if not os.path.exists(CREDENTIALS_FILE):
        st.sidebar.markdown("---")
        st.sidebar.warning("⚠️ Google Sheets credentials не знайдено")
        uploaded_credentials = st.sidebar.file_uploader("Завантажте service_account_credentials.json", type=['json'], key='credentials_upload')
        if uploaded_credentials is not None:
            with open(CREDENTIALS_FILE, "wb") as f:
                f.write(uploaded_credentials.getbuffer())
            st.sidebar.success("✅ Credentials завантажено!")
            st.rerun()

# === Загрузка данных: сперва Google Sheets (CSV), затем локальный fallback ===
lakes, reports, lakes_table, reports_table = load_from_google_sheets()

if lakes_table is not None and not lakes_table.empty:
    st.sidebar.success(f"✅ Дані завантажено з Google Sheets ({len(lakes_table)} рядків)")
else:
    if os.path.exists(EXCEL_FILE_PATH):
        lakes, reports, lakes_table, reports_table = load_lakes_and_reports(EXCEL_FILE_PATH)
        st.sidebar.info(f"📂 Використовую локальний файл: `{os.path.abspath(EXCEL_FILE_PATH)}`")
    else:
        st.warning("⚠️ Файл LakeHouse.xlsx не знайдено. Завантажте Excel файл:")
        uploaded_file = st.file_uploader("Завантажте Excel файл", type=['xlsx', 'xls'])
        if uploaded_file is not None:
            with open(EXCEL_FILE_PATH, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success("✅ Файл завантажено! Оновлюємо дані...")
            st.cache_data.clear()
            lakes, reports, lakes_table, reports_table = load_lakes_and_reports(EXCEL_FILE_PATH)
            st.sidebar.info(f"📂 Локальний файл: `{os.path.abspath(EXCEL_FILE_PATH)}`")
        else:
            lakes, reports, lakes_table, reports_table = [], [], None, None
            st.info("👆 Завантажте Excel файл або підключіть Google Sheets у сайдбарі")

# ==================== ГОЛОВНА СТОРІНКА ====================
if section == "🏠 Головна":
    st.header("Вітаємо! 👋")
    st.markdown(f"""
    Ця база знань містить інформацію для підтримки та оновлення наших LakeHouses та Power BI Reports.
    """)
    col1, col2 = st.columns(2)
    with col1: st.metric("🏞️ Data Lakes", len(lakes) if lakes else 0)
    with col2: st.metric("📊 Power BI звіти", len(reports) if reports else 0)

# ==================== ОНОВЛЕННЯ DATA LAKES ====================
elif section == "💧 Оновлення LakeHouses":
    st.header("💧 Інструкції по оновленню LakeHouses")
    unique_lakes = []
    if lakes_table is not None and not lakes_table.empty:
        for col in ['LakeHouse', 'name', 'Name', 'назва', 'Назва', 'lake_name', 'Lake Name', 'Lakehouse']:
            if col in lakes_table.columns:
                unique_lakes = list(lakes_table[col].dropna().unique())
                break
        if not unique_lakes:
            unique_lakes = list(lakes_table.iloc[:,0].dropna().unique())

    lake_select_options = ["Всі лейки"] + unique_lakes + ["📊 Аналітика та візуалізація"]
    lake_name = st.selectbox("Оберіть Data Lake:", lake_select_options)

    if lake_name == "Всі лейки":
        st.info("👈 Оберіть конкретний лейк зі списку вище")
        if lakes_table is not None and not lakes_table.empty and 'LakeHouse' in lakes_table.columns:
            unique_lakes_vals = lakes_table['LakeHouse'].dropna().unique()
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("🏞️ Унікальних лейків", len(unique_lakes_vals))
            st.subheader("📋 Список всіх Data Lakes")
            if 'Загальна інформація про лейк' in lakes_table.columns:
                summary = lakes_table.groupby('LakeHouse').first().reset_index()
                st.dataframe(summary[['LakeHouse','Загальна інформація про лейк']], use_container_width=True, hide_index=True)
            else:
                st.dataframe(lakes_table[['LakeHouse']], use_container_width=True, hide_index=True)
        else:
            st.warning("Список лейків порожній або відсутня колонка 'LakeHouse'.")
    elif lake_name == "📊 Аналітика та візуалізація":
        st.subheader("📊 Аналітика та візуалізація лейків")
        if lakes_table is not None and not lakes_table.empty:
            analysis = analyze_lakes_data(lakes_table)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("🏞️ Всього лейків", analysis['total_lakes'])
            c2.metric("📊 Колонок даних", len(analysis['columns']))
            c3.metric("⚠️ Пропущених значень", sum(analysis['missing_data'].values()))
            c4.metric("📅 Останнє оновлення", datetime.now().strftime('%d.%m'))
            charts = create_lakes_visualization(lakes_table)
            if charts:
                for chart in charts.values():
                    st.plotly_chart(chart, use_container_width=True)
            st.subheader("🔍 Детальний аналіз")
            missing_df = pd.DataFrame(list(analysis['missing_data'].items()), columns=['Колонка','Пропущено'])
            missing_df = missing_df[missing_df['Пропущено'] > 0]
            if not missing_df.empty: st.dataframe(missing_df, use_container_width=True)
            else: st.success("✅ Пропущених даних немає!")
        else:
            st.warning("Немає даних для аналізу!")
    else:
        if lakes_table is not None and not lakes_table.empty:
            lake_data = None
            for col in ['LakeHouse', 'name', 'Name', 'назва', 'Назва', 'lake_name', 'Lake Name', 'Lakehouse']:
                if col in lakes_table.columns and lake_name in lakes_table[col].values:
                    lake_data = lakes_table[lakes_table[col] == lake_name]
                    break
            if (lake_data is None or lake_data.empty) and 'LakeHouse' in lakes_table.columns:
                uniq = lakes_table['LakeHouse'].dropna().unique()
                if len(uniq) == 1:
                    lake_data = lakes_table
                    st.info(f"💡 Використовуємо єдиний лейк: {uniq[0]}")
            if lake_data is not None and not lake_data.empty:
                st.success(f"🏞️ Вибрано лейк: **{lake_name}**")
                if 'Загальна інформація про лейк' in lake_data.columns and pd.notna(lake_data['Загальна інформація про лейк'].iloc[0]):
                    st.subheader("ℹ️ Загальна інформація про лейк")
                    st.info(lake_data['Загальна інформація про лейк'].iloc[0])
                if 'Folder' in lake_data.columns:
                    st.subheader("📁 Структура лейка")
                    unique_folders = lake_data['Folder'].dropna().unique()
                    if len(unique_folders) > 0:
                        st.write("**Доступні папки:**")
                        cols = st.columns(min(3, len(unique_folders)))
                        selected_folder = None
                        for i, folder in enumerate(unique_folders):
                            with cols[i % 3]:
                                if st.button(f"📂 {folder}", key=f"folder_{i}"):
                                    selected_folder = folder
                        if selected_folder:
                            st.success(f"📂 Вибрано папку: **{selected_folder}**")
                            folder_data = lake_data[lake_data['Folder'] == selected_folder]
                            st.subheader("🧩 Елементи папки")
                            display_columns = folder_data.columns[2:8]
                            if 'URL' in display_columns:
                                display_columns = [c for c in display_columns if c != 'URL']
                            if 'Element' in display_columns and 'URL' in folder_data.columns:
                                elements_df_display = folder_data[display_columns].copy()
                                url_dict = {idx: row.get('URL','').strip() for idx, row in folder_data.iterrows()
                                            if pd.notna(row.get('URL','')) and str(row.get('URL','')).strip()}
                                def create_link(row_data):
                                    element_name = row_data['Element']
                                    row_idx = row_data.name
                                    if row_idx in url_dict:
                                        url = url_dict[row_idx]
                                        return f'<a href="{url}" target="_blank" style="color:#1f77b4;text-decoration:underline;">{element_name}</a>'
                                    return element_name
                                elements_df_display['Element'] = elements_df_display.apply(create_link, axis=1)
                                st.markdown(elements_df_display.to_html(escape=False), unsafe_allow_html=True)
                                st.info(f"🔗 Активних посилань: {len(url_dict)}")
                            else:
                                st.dataframe(folder_data[display_columns], use_container_width=True, hide_index=True)
                            st.subheader("📝 Внесення змін")
                            changes_col = 'Внесення змін'
                            if changes_col in folder_data.columns and pd.notna(folder_data[changes_col].iloc[0]):
                                with st.expander("Показати деталі змін", expanded=True):
                                    process_text_with_images(folder_data[changes_col].iloc[0])
                            else:
                                st.info("Немає інформації про внесення змін для цієї папки.")
                        else:
                            st.info("👆 Натисніть на папку вище, щоб побачити її елементи")
                    else:
                        st.warning("⚠️ Папки не знайдено в даних")
                else:
                    st.warning("⚠️ Колонка 'Folder' не знайдена. Показую всі дані:")
                    st.dataframe(lake_data, use_container_width=True, hide_index=True)
            else:
                st.error(f"❌ Лейк '{lake_name}' не знайдено.")
        else:
            st.warning("⚠️ Дані лейків не завантажені.")

# ==================== РЕДАГУВАННЯ ДАНИХ ====================
elif section == "✏️ Редагування даних":
    st.header("✏️ Редагування даних")
    if lakes_table is not None and not lakes_table.empty:
        st.subheader("📊 Поточні дані")
        st.info("💡 Редагуйте дані прямо в таблиці. Зміни будуть записані у Google Sheets; якщо не вдасться — у локальний Excel (резерв).")

        edited_df = st.data_editor(
            lakes_table, use_container_width=True, num_rows="dynamic", key="data_editor"
        )

        if not edited_df.equals(lakes_table):
            # пробуем Google Sheets
            if save_to_google_sheets(edited_df, reports_table):
                st.cache_data.clear()
                time.sleep(1.2)
                st.rerun()
            else:
                # локальный резерв
                ok, saved = save_data_to_excel(edited_df, EXCEL_FILE_PATH, reports_table)
                if ok:
                    st.cache_data.clear()
                    time.sleep(1.2)
                    st.rerun()

        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 Оновити дані"):
                st.cache_data.clear()
                st.rerun()
        with col2:
            csv = (lakes_table if lakes_table is not None else pd.DataFrame()).to_csv(index=False)
            st.download_button("📥 Завантажити CSV", data=csv, file_name=f"lakes_data_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv")

        st.subheader("➕ Додати новий запис")
        with st.form("add_new_record"):
            c1, c2 = st.columns(2)
            with c1:
                new_lakehouse = st.text_input("LakeHouse *")
                new_folder = st.text_input("Folder *")
                new_element = st.text_input("Element *")
                new_url = st.text_input("URL")
            with c2:
                new_info = st.text_area("Загальна інформація про лейк")
                new_changes = st.text_area("Внесення змін")
            if st.form_submit_button("➕ Додати запис"):
                if new_lakehouse and new_folder and new_element:
                    new_row = {
                        'LakeHouse': new_lakehouse,
                        'Folder': new_folder,
                        'Element': new_element,
                        'URL': new_url or '',
                        'Загальна інформація про лейк': new_info or '',
                        'Внесення змін': new_changes or ''
                    }
                    new_df = pd.concat([lakes_table, pd.DataFrame([new_row])], ignore_index=True)

                    if save_to_google_sheets(new_df, reports_table):
                        st.cache_data.clear()
                        time.sleep(1.2)
                        st.rerun()
                    else:
                        st.warning("⚠️ Google Sheets недоступний. Зберігаю локально як резервну копію.")
                        ok, saved = save_data_to_excel(new_df, EXCEL_FILE_PATH, reports_table)
                        if ok:
                            st.cache_data.clear()
                            time.sleep(1.2)
                            st.rerun()
                else:
                    st.error("❌ Заповніть обов'язкові поля: LakeHouse, Folder, Element")
    else:
        st.warning("⚠️ Немає даних для редагування. Завантажте Excel або увімкніть Google Sheets.")

# ==================== КОНТАКТИ ТА РЕСУРСИ ====================
elif section == "📞 Контакти та ресурси":
    st.header("📞 Контакти та ресурси")
    st.subheader("👥 Наша команда")
    st.markdown("""
    ### 🏢 OurTeam

    **Zhovtiuk Svitlana**  
    Керівник групи  
    📧 s.zhovtiuk@darnytsia.ua

    **Filatova Oleksandra**  
    Менеджер з бізнес аналітики  
    📧 oleksandra.filatova@darnytsia.ua

    **Bohdanyk Oleksandr**  
    Менеджер з бізнес аналітики  
    📧 o.bohdanyk@darnytsia.ua

    **Taranenko Oleksandr**  
    Менеджер з бізнес аналітики  
    📧 o.taranenko@darnytsia.ua
    """)
    st.subheader("🔗 Корисні посилання")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        ### Внутрішні ресурси:
        - [SharePoint команди](https://darnytsia.sharepoint.com)
        - [Azure DevOps](https://dev.azure.com/darnitsa)
        - [Power BI Service](https://app.powerbi.com)
        """)
    with c2:
        st.markdown("""
        ### Зовнішні ресурси:
        - [Microsoft Learn](https://learn.microsoft.com)
        - [Power BI Community](https://community.powerbi.com)
        - [Streamlit Docs](https://docs.streamlit.io)
        """)

# ----------------- конец файла -----------------


