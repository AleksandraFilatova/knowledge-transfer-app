import streamlit as st
from datetime import datetime
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from PIL import Image
import base64
import requests
try:
    from openpyxl import load_workbook
except ImportError:
    openpyxl = None

# ==== CONFIG SECTION ====
# Путь к Excel с лейками и звітами. Використовуємо абсолютний шлях до папки з кодом
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(SCRIPT_DIR, "LakeHouse.xlsx")

# GitHub URL для файлу (raw формат)
GITHUB_RAW_URL = "https://raw.githubusercontent.com/AleksandraFilatova/knowledge-transfer-app/main/LakeHouse.xlsx"

# ======= Функція для читання інформації з Excel =========
@st.cache_data(ttl=300)
def load_lakes_and_reports(excel_path):
    """
    Завантажує дані з Excel файлу
    """
    try:
        xl = pd.ExcelFile(excel_path)
        available_sheets = xl.sheet_names
        
        lakes_df = None
        reports_df = None
        
        # Шукаємо лист з лейками (можливі варіанти назв)
        lake_sheet_names = ['Lakes']
        for sheet_name in lake_sheet_names:
            if sheet_name in available_sheets:
                lakes_df = pd.read_excel(xl, sheet_name)
                break
        
        # Якщо не знайшли спеціальний лист, спробуємо другий лист (якщо є)
        if lakes_df is None and len(available_sheets) > 1:
            lakes_df = pd.read_excel(xl, available_sheets[1])
        elif lakes_df is None and available_sheets:
            lakes_df = pd.read_excel(xl, available_sheets[0])
        
        # Шукаємо лист зі звітами
        report_sheet_names = ['Reports', 'reports', 'report', 'звіти', 'Power BI']
        for sheet_name in report_sheet_names:
            if sheet_name in available_sheets:
                reports_df = pd.read_excel(xl, sheet_name)
                break
        
        # Якщо не знайшли спеціальний лист, використаємо перший лист
        if reports_df is None and available_sheets:
            reports_df = pd.read_excel(xl, available_sheets[0])
        
        # Витягуємо назви
        lakes_names = []
        reports_names = []
        
        if lakes_df is not None and not lakes_df.empty:
            # Шукаємо колонку з назвами (можливі варіанти)
            name_columns = ['LakeHouse', 'name', 'Name', 'назва', 'Назва', 'lake_name', 'Lake Name', 'Lakehouse']
            name_col = None
            for col in name_columns:
                if col in lakes_df.columns:
                    name_col = col
                    break
            
            if name_col:
                lakes_names = list(lakes_df[name_col].dropna())
            else:
                # Якщо не знайшли колонку з назвами, використаємо першу колонку
                lakes_names = list(lakes_df.iloc[:, 0].dropna())
        
        if reports_df is not None and not reports_df.empty:
            name_columns = ['name', 'Name', 'назва', 'Назва', 'report_name', 'Report Name']
            name_col = None
            for col in name_columns:
                if col in reports_df.columns:
                    name_col = col
                    break
            
            if name_col:
                reports_names = list(reports_df[name_col].dropna())
            else:
                reports_names = list(reports_df.iloc[:, 0].dropna())
        
        return lakes_names, reports_names, lakes_df, reports_df
        
    except Exception as e:
        st.error(f"❌ Помилка при завантаженні файлу: {e}")
        return [], [], None, None

def analyze_lakes_data(lakes_df):
    """
    Аналізує дані лейків та повертає статистику
    """
    if lakes_df is None or lakes_df.empty:
        return {
            'total_lakes': 0,
            'columns': [],
            'missing_data': {},
            'unique_values': {}
        }
    
    analysis = {
        'total_lakes': len(lakes_df),
        'columns': list(lakes_df.columns),
        'missing_data': {},
        'unique_values': {}
    }
    
    # Аналіз пропущених даних
    for col in lakes_df.columns:
        missing_count = lakes_df[col].isna().sum()
        analysis['missing_data'][col] = missing_count
    
    # Аналіз унікальних значень для кожної колонки
    for col in lakes_df.columns:
        if lakes_df[col].dtype == 'object':  # Текстові колонки
            analysis['unique_values'][col] = lakes_df[col].value_counts().to_dict()
    
    return analysis

def display_image_from_path(image_path, caption=None, width=None):
    """
    Відображає зображення з файлового шляху або URL
    """
    try:
        # Перевіряємо, чи це URL
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
    """
    Відображає зображення з base64 рядка
    """
    try:
        # Декодуємо base64
        image_data = base64.b64decode(base64_string)
        st.image(image_data, caption=caption, width=width)
    except Exception as e:
        st.error(f"❌ Помилка при декодуванні зображення: {e}")

def process_text_with_images(text):
    """
    Обробляє текст та відображає зображення, якщо знайдені посилання на них
    """
    if not text:
        return text
    
    # Шукаємо посилання на зображення в тексті
    import re
    
    # Паттерн для пошуку посилань на зображення
    image_pattern = r'\[IMAGE:(.*?)\]'
    matches = re.findall(image_pattern, text)
    
    if matches:
        # Розділяємо текст на частини
        parts = re.split(image_pattern, text)
        
        for i, part in enumerate(parts):
            if i % 2 == 0:  # Текст
                if part.strip():
                    st.markdown(part)
            else:  # Шлях до зображення
                image_path = part.strip()
        # Перевіряємо, чи це локальний шлях, і замінюємо на GitHub URL
        if image_path.startswith('C:\\') and 'PL-notebook.png' in image_path:
            # Замінюємо локальний шлях на GitHub URL
            github_url = "https://raw.githubusercontent.com/AleksandraFilatova/knowledge-transfer-app/main/Image/Sac-notebook.PNG"
            display_image_from_path(github_url, width=600)
        elif 'github.com' in image_path and '/blob/' in image_path:
            # Автоматично виправляємо GitHub blob URL на raw URL
            raw_url = image_path.replace('github.com', 'raw.githubusercontent.com').replace('/blob/', '/')
            display_image_from_path(raw_url, width=600)
        else:
            display_image_from_path(image_path, width=600)
    else:
        # Якщо немає зображень, просто показуємо текст
        st.markdown(text)

def download_file_from_github(url, local_path):
    """
    Завантажує файл з GitHub
    """
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        with open(local_path, 'wb') as f:
            f.write(response.content)
        return True
    except Exception as e:
        return False

def save_data_to_excel(df, filename, lakes_table=None, reports_table=None):
    """
    Зберігає DataFrame в Excel файл з підтримкою множинних листів
    """
    try:
        # Відкриваємо існуючий файл, якщо він є
        if os.path.exists(filename):
            from openpyxl import load_workbook
            try:
                # Спробуємо зчитати існуючий файл
                existing_df = pd.ExcelFile(filename)
                
                # Якщо в файлі є лист "Reports", зберігаємо його
                if 'Reports' in existing_df.sheet_names and reports_table is not None:
                    # Зберігаємо з обома листами
                    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name='Lakes', index=False)
                        reports_table.to_excel(writer, sheet_name='Reports', index=False)
                else:
                    # Зберігаємо тільки оновлений лист
                    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name='Lakes', index=False)
                        if reports_table is not None:
                            reports_table.to_excel(writer, sheet_name='Reports', index=False)
            except:
                # Якщо не вдалося відкрити, просто перезапишемо
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Lakes', index=False)
                    if reports_table is not None:
                        reports_table.to_excel(writer, sheet_name='Reports', index=False)
        else:
            # Якщо файлу немає, створюємо новий
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Lakes', index=False)
                if reports_table is not None:
                    reports_table.to_excel(writer, sheet_name='Reports', index=False)
        
        return True, filename
    except Exception as e:
        st.error(f"❌ Помилка при збереженні: {e}")
        return False, None

def create_lakes_visualization(lakes_df):
    """
    Створює візуалізації для даних лейків
    """
    if lakes_df is None or lakes_df.empty:
        return None
    
    # Створюємо візуалізації
    charts = {}
    
    # 1. Кругова діаграма для статусу (якщо є колонка Status)
    if 'Status' in lakes_df.columns:
        status_counts = lakes_df['Status'].value_counts()
        fig_pie = px.pie(
            values=status_counts.values,
            names=status_counts.index,
            title="Розподіл лейків за статусом"
        )
        charts['status_pie'] = fig_pie
    
    # 2. Гістограма для частоти оновлень (якщо є колонка Update_Frequency)
    if 'Update_Frequency' in lakes_df.columns:
        fig_bar = px.bar(
            x=lakes_df['Update_Frequency'].value_counts().index,
            y=lakes_df['Update_Frequency'].value_counts().values,
            title="Частота оновлень лейків"
        )
        charts['frequency_bar'] = fig_bar
    
    # 3. Treemap для розподілу по робочих просторах (якщо є колонка Workspace)
    if 'Workspace' in lakes_df.columns:
        fig_treemap = px.treemap(
            lakes_df,
            path=['Workspace'],
            title="Розподіл лейків по робочих просторах"
        )
        charts['workspace_treemap'] = fig_treemap
    
    return charts

def create_lake_details_card(lake_row):
    """
    Створює детальну картку для конкретного лейка
    """
    if lake_row is None or lake_row.empty:
        return "Немає даних про лейк"
    
    # Створюємо HTML картку з інформацією
    card_html = """
    <div style="
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin: 10px 0;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    ">
        <h3 style="margin: 0 0 15px 0; color: white;">🏞️ {lake_name}</h3>
    """.format(lake_name=lake_row.get('LakeHouse', 'Невідомий лейк'))
    
    # Додаємо інформацію з усіх доступних колонок
    for col in lake_row.index:
        if pd.notna(lake_row[col]) and col not in ['LakeHouse', 'Folder', 'Element', 'Загальна інформація про лейк', 'Внесення змін']: # Виключаємо вже використані або спеціальні колонки
            value = lake_row[col]
            # Перекладаємо назви колонок для відображення
            display_col_name = {
                'Type': 'Тип',
                'Опис': 'Опис',
                'Оновлення': 'Оновлення',
                'Особливості': 'Особливості',
            }.get(col, col) # Якщо немає перекладу, використовуємо оригінальну назву
            card_html += f"<p style=\"margin: 5px 0;\"><strong>{display_col_name}:</strong> {value}</p>"
    
    card_html += "</div>"
    return card_html

# ==================== НАСТРОЙКИ СТОРІНКИ ====================
st.set_page_config(
    page_title="Knowledge Transfer App",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== НАВІГАЦІЯ ====================
st.sidebar.title("🗂️ Навігація")
st.sidebar.markdown("### Оберіть розділ:")

section = st.sidebar.radio(
    "",
    ["🏠 Головна", 
     "💧 Оновлення LakeHouses", 
     "📊 Оновлення PowerBI Report",
     "✏️ Редагування даних",
     "📞 Контакти та ресурси"]
)

st.sidebar.markdown("---")
st.sidebar.info(f"📅 Останнє оновлення:\n{datetime.now().strftime('%d.%m.%Y')}")


# === ДИНАМИЧЕСКИЙ ЗАПРОС таблицы Excel для Lakes & reports ===
# Перевіряємо, чи файл існує локально
if os.path.exists(EXCEL_FILE_PATH):
    lakes, reports, lakes_table, reports_table = load_lakes_and_reports(EXCEL_FILE_PATH)
    # Показуємо повідомлення, де зберігаються дані
    abs_path = os.path.abspath(EXCEL_FILE_PATH)
    st.sidebar.success(f"📂 Файл: `{abs_path}`")
else:
    # Якщо файл не знайдено локально, спробуємо завантажити з GitHub
    if download_file_from_github(GITHUB_RAW_URL, EXCEL_FILE_PATH):
        lakes, reports, lakes_table, reports_table = load_lakes_and_reports(EXCEL_FILE_PATH)
        abs_path = os.path.abspath(EXCEL_FILE_PATH)
        st.sidebar.success(f"✅ Файл завантажено з GitHub: `{abs_path}`")
    else:
        # Якщо не вдалося завантажити з GitHub, пропонуємо завантажити вручну
        st.warning("⚠️ Файл LakeHouse.xlsx не знайдено. Будь ласка, завантажте файл:")
        uploaded_file = st.file_uploader("Завантажте Excel файл", type=['xlsx', 'xls'])
        
        if uploaded_file is not None:
            # Зберігаємо завантажений файл
            with open(EXCEL_FILE_PATH, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.success("✅ Файл завантажено! Оновлюємо дані...")
            lakes, reports, lakes_table, reports_table = load_lakes_and_reports(EXCEL_FILE_PATH)
            abs_path = os.path.abspath(EXCEL_FILE_PATH)
            st.sidebar.info(f"📂 Локальний файл: `{abs_path}`")
        else:
            # Показуємо заглушку
            lakes, reports, lakes_table, reports_table = [], [], None, None
            st.info("👆 Завантажте Excel файл для початку роботи")

# ==================== ГОЛОВНА СТОРІНКА ====================
if section == "🏠 Головна":
    st.header("Вітаємо! 👋")
    st.markdown("""
    Ця база знань містить всю необхідну інформацію для підтримки та оновлення 
    наших LakeHouses та Power BI Reports.

    **⚡️ Тепер список лейків і звітів зчитується з таблиці Excel**  
    Можна легко коригувати склад без зміни коду!
    
    ### 📝 Як оновити дані:
    1. **Перейдіть в розділ "Редагування даних"** (в меню зліва)
    2. **Редагуйте дані прямо в таблиці** - зміни зберігаються автоматично
    3. **Додавайте нові записи** через форму
    4. **Зміни відображаються** миттєво для всіх користувачів
    
    **Excel файл:** `{}`  
    **Повний шлях:** `{}`  
    """.format(EXCEL_FILE_PATH, os.path.abspath(EXCEL_FILE_PATH)))
    col1, col2 = st.columns(2)
    with col1:
        st.metric("🏞️ Data Lakes", len(lakes) if lakes else 0)
    with col2:
        st.metric("📊 Power BI звіти", len(reports) if reports else 0)

# ==================== ОНОВЛЕННЯ DATA LAKES ====================
elif section == "💧 Оновлення LakeHouses":
    st.header("💧 Інструкції по оновленню LakeHouses")
    
    # Аналіз даних буде показано після вибору лейка
    
    # Отримуємо унікальні назви лейків (без дублювання)
    unique_lakes = []
    if lakes_table is not None and not lakes_table.empty:
        # Шукаємо колонку з назвами лейків
        name_columns = ['LakeHouse', 'name', 'Name', 'назва', 'Назва', 'lake_name', 'Lake Name', 'Lakehouse']
        name_col = None
        for col in name_columns:
            if col in lakes_table.columns:
                name_col = col
                break
        
        if name_col:
            unique_lakes = list(lakes_table[name_col].dropna().unique())
        else:
            unique_lakes = list(lakes_table.iloc[:, 0].dropna().unique())
    
    # Додаємо опції для вибору
    lake_select_options = ["Всі лейки"] + unique_lakes + ["📊 Аналітика та візуалізація"]
    
    lake_name = st.selectbox(
        "Оберіть Data Lake:",
        lake_select_options
    )
    
    if lake_name == "Всі лейки":
        st.info("👈 Оберіть конкретний лейк зі списку вище")
        
        # Показуємо тільки унікальні лейки з колонки LakeHouse
        if lakes_table is not None and not lakes_table.empty:
            # Отримуємо унікальні значення з колонки LakeHouse
            unique_lakes = lakes_table['LakeHouse'].dropna().unique()
            
            # Показуємо метрику тільки для кількості унікальних лейків
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("🏞️ Унікальних лейків", len(unique_lakes))
            with col2:
                st.metric("", "")  # Порожня картка
            with col3:
                st.metric("", "")  # Порожня картка
            with col4:
                st.metric("", "")  # Порожня картка
            
            st.subheader("📋 Список всіх Data Lakes")
            
            # Створюємо таблицю тільки з 2 колонками
            if 'LakeHouse' in lakes_table.columns and 'Загальна інформація про лейк' in lakes_table.columns:
                # Групуємо по LakeHouse та беремо перший запис для кожної групи
                summary_table = lakes_table.groupby('LakeHouse').first().reset_index()
                
                # Показуємо тільки потрібні колонки
                display_columns = ['LakeHouse', 'Загальна інформація про лейк']
                summary_display = summary_table[display_columns]
                
                st.dataframe(
                    summary_display, 
                    use_container_width=True,
                    hide_index=True
                )
                
                # Додаємо кнопку для експорту
                csv = summary_display.to_csv(index=False)
                st.download_button(
                    label="📥 Завантажити CSV",
                    data=csv,
                    file_name=f"lakes_summary_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            else:
                st.warning("⚠️ Не знайдено колонки 'LakeHouse' або 'Загальна інформація про лейк'")
        else:
            st.warning("Список лейків порожній у файлі Excel!")
    
    elif lake_name == "📊 Аналітика та візуалізація":
        st.subheader("📊 Аналітика та візуалізація лейків")
        
        if lakes_table is not None and not lakes_table.empty:
            # Показуємо аналіз даних
            analysis = analyze_lakes_data(lakes_table)
            
            # Показуємо загальну статистику
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("🏞️ Всього лейків", analysis['total_lakes'])
            with col2:
                st.metric("📊 Колонок даних", len(analysis['columns']))
            with col3:
                missing_total = sum(analysis['missing_data'].values())
                st.metric("⚠️ Пропущених значень", missing_total)
            with col4:
                st.metric("📅 Останнє оновлення", datetime.now().strftime('%d.%m'))
            
            # Створюємо візуалізації
            charts = create_lakes_visualization(lakes_table)
            
            if charts:
                st.subheader("📈 Візуалізації")
                
                # Показуємо доступні діаграми
                for chart_name, chart in charts.items():
                    if chart_name == 'status_pie':
                        st.plotly_chart(chart, use_container_width=True)
                    elif chart_name == 'frequency_bar':
                        st.plotly_chart(chart, use_container_width=True)
                    elif chart_name == 'workspace_treemap':
                        st.plotly_chart(chart, use_container_width=True)
            
            # Детальний аналіз
            st.subheader("🔍 Детальний аналіз")
            
            # Аналіз пропущених даних
            if any(analysis['missing_data'].values()):
                st.subheader("⚠️ Пропущені дані")
                missing_df = pd.DataFrame(list(analysis['missing_data'].items()), columns=['Колонка', 'Пропущено'])
                missing_df = missing_df[missing_df['Пропущено'] > 0]
                if not missing_df.empty:
                    st.dataframe(missing_df, use_container_width=True)
                else:
                    st.success("✅ Пропущених даних немає!")
            
            # Аналіз унікальних значень
            if analysis['unique_values']:
                st.subheader("📊 Унікальні значення")
                for col, values in analysis['unique_values'].items():
                    if values:
                        st.write(f"**{col}:**")
                        values_df = pd.DataFrame(list(values.items()), columns=['Значення', 'Кількість'])
                        st.dataframe(values_df, use_container_width=True)
        else:
            st.warning("Немає даних для аналізу!")
    
    else:
        # Детальна інструкція для конкретного lake
        if lakes_table is not None and not lakes_table.empty:
            # Фільтруємо дані для вибраного лейка
            lake_data = None
            
            # Шукаємо колонку з назвами лейків
            name_columns = ['LakeHouse', 'name', 'Name', 'назва', 'Назва', 'lake_name', 'Lake Name', 'Lakehouse']
            name_col = None
            for col in name_columns:
                if col in lakes_table.columns:
                    if lake_name in lakes_table[col].values:
                        name_col = col
                        lake_data = lakes_table[lakes_table[col] == lake_name]
                        break
            
            # Якщо не знайшли за назвою, але є тільки один унікальний лейк - використаємо всі дані
            if lake_data is None or lake_data.empty:
                # Перевіряємо, чи є тільки один унікальний лейк
                unique_lakes = []
                for col in name_columns:
                    if col in lakes_table.columns:
                        unique_lakes = list(lakes_table[col].dropna().unique())
                        if len(unique_lakes) == 1:
                            lake_data = lakes_table
                            st.info(f"💡 Використовуємо єдиний лейк: {unique_lakes[0]}")
                            break
            
            if lake_data is not None and not lake_data.empty:
                st.success(f"🏞️ Вибрано лейк: **{lake_name}**")
                
                # Показуємо аналіз даних для конкретного лейка
                if lakes_table is not None and not lakes_table.empty:
                    # Рахуємо унікальні лейки з колонки LakeHouse
                    unique_lakes_count = lakes_table['LakeHouse'].nunique()
                    
                    # Рахуємо унікальні елементи з колонки Element
                    unique_elements_count = lakes_table['Element'].nunique() if 'Element' in lakes_table.columns else 0
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("🏞️ Всього лейків", unique_lakes_count)
                    with col2:
                        st.metric("🧩 Кількість елементів", unique_elements_count)
                
                # Показуємо загальну інформацію про лейк з колонки "Загальна інформація про лейк"
                st.subheader("ℹ️ Загальна інформація про лейк")
                if 'Загальна інформація про лейк' in lake_data.columns and pd.notna(lake_data['Загальна інформація про лейк'].iloc[0]):
                    st.info(lake_data['Загальна інформація про лейк'].iloc[0])
                else:
                    st.info("Загальна інформація про лейк не надана")
                
                # Показуємо структуру лейка тільки якщо є колонка Folder
                if 'Folder' in lake_data.columns:
                    st.subheader("📁 Структура лейка")
                    
                    # Отримуємо унікальні папки
                    unique_folders = lake_data['Folder'].dropna().unique()
                    
                    if len(unique_folders) > 0:
                        st.write("**Доступні папки:**")
                        
                        # Створюємо кнопки для папок в колонках
                        cols = st.columns(min(3, len(unique_folders)))
                        selected_folder = None
                        
                        for i, folder in enumerate(unique_folders):
                            with cols[i % 3]:
                                if st.button(f"📂 {folder}", key=f"folder_{i}"):
                                    selected_folder = folder
                        
                        # Показуємо елементи вибраної папки
                        if selected_folder:
                            st.success(f"📂 Вибрано папку: **{selected_folder}**")
                            
                            # Фільтруємо дані по вибраній папці
                            folder_data = lake_data[lake_data['Folder'] == selected_folder]
                            
                            # Показуємо елементи папки (тільки стовпці з 3 по 8)
                            st.subheader("🧩 Елементи папки")
                            
                            # Вибираємо стовпці з 3 по 8 (індекси 2-7), але виключаємо URL
                            display_columns = folder_data.columns[2:8]
                            # Виключаємо колонку URL з відображення, якщо вона є
                            if 'URL' in display_columns:
                                display_columns = [col for col in display_columns if col != 'URL']
                            
                            if 'Element' in display_columns:
                                # Перевіряємо, чи є колонка URL
                                if 'URL' in folder_data.columns:
                                    # Створюємо копію для модифікації
                                    elements_df_display = folder_data[display_columns].copy()
                                    
                                    # Створюємо словник URL для кожного рядка (за індексом)
                                    url_dict = {}
                                    for idx, row in folder_data.iterrows():
                                        url_value = row.get('URL', '')
                                        if pd.notna(url_value) and url_value.strip():
                                            url_dict[idx] = url_value.strip()
                                    
                                    # Перетворюємо стовпець 'Element' на клікабельні посилання
                                    def create_link(row_data):
                                        element_name = row_data['Element']
                                        row_idx = row_data.name  # Отримуємо індекс рядка
                                        
                                        if row_idx in url_dict:
                                            url = url_dict[row_idx]
                                            return f'<a href="{url}" target="_blank" style="color: #1f77b4; text-decoration: underline;">{element_name}</a>'
                                        else:
                                            return element_name
                                    
                                    # Застосовуємо функцію до кожного рядка
                                    elements_df_display['Element'] = elements_df_display.apply(create_link, axis=1)
                                    
                                    # Показуємо таблицю з HTML посиланнями
                                    st.markdown(elements_df_display.to_html(escape=False), unsafe_allow_html=True)
                                    
                                    # Додаємо інформацію про кількість посилань
                                    active_links = len(url_dict)
                                    if active_links > 0:
                                        st.info(f"🔗 {active_links} з {len(folder_data)} елементів мають активні посилання")
                                else:
                                    # Якщо немає колонки URL, показуємо звичайну таблицю
                                    st.dataframe(folder_data[display_columns], use_container_width=True, hide_index=True)
                                    st.warning("⚠️ Колонка 'URL' не знайдена. Показуємо звичайну таблицю.")
                            else:
                                st.dataframe(folder_data[display_columns], use_container_width=True, hide_index=True)
                            
                            # Додаємо секцію "Внесення змін"
                            st.subheader("📝 Внесення змін")
                            changes_col = 'Внесення змін'
                            if changes_col in folder_data.columns and pd.notna(folder_data[changes_col].iloc[0]):
                                with st.expander("Показати деталі змін", expanded=True):
                                    changes_text = folder_data[changes_col].iloc[0]
                                    # Обробляємо текст з можливими зображеннями
                                    process_text_with_images(changes_text)
                            else:
                                st.info("Немає інформації про внесення змін для цієї папки.")
                        else:
                            st.info("👆 Натисніть на папку вище, щоб побачити її елементи")
                    else:
                        st.warning("⚠️ Папки не знайдено в даних")
                else:
                    # Якщо немає колонки Folder, показуємо всю таблицю
                    st.warning("⚠️ Колонка 'Folder' не знайдена. Показуємо всі дані:")
                    st.dataframe(lake_data, use_container_width=True, hide_index=True)
                
                # Секція "Внесення змін" тепер показується тільки після вибору папки
            else:
                st.error(f"❌ Лейк '{lake_name}' не знайдено в базі даних!")
                st.info("💡 Перевірте, чи правильно вказана назва лейка")
                
                # Показуємо доступні лейки для довідки
                if unique_lakes:
                    st.write("**Доступні лейки:**")
                    for lake in unique_lakes:
                        st.write(f"- {lake}")
        else:
            st.warning("⚠️ Дані лейків не завантажені. Перевірте файл Excel.")

# ==================== РЕДАГУВАННЯ ДАНИХ ====================
elif section == "✏️ Редагування даних":
    st.header("✏️ Редагування даних")
    
    if lakes_table is not None and not lakes_table.empty:
        st.subheader("📊 Поточні дані")
        st.info("💡 Редагуйте дані прямо в таблиці. Зміни зберігаються автоматично!")
        
        # Показуємо таблицю для редагування
        edited_df = st.data_editor(
            lakes_table,
            use_container_width=True,
            num_rows="dynamic",
            key="data_editor"
        )
        
        # Автоматичне збереження при змінах
        if not edited_df.equals(lakes_table):
            success, saved_file = save_data_to_excel(edited_df, EXCEL_FILE_PATH, 
                                                     lakes_table=None, reports_table=reports_table)
            if success:
                # Очищуємо кеш після збереження
                st.cache_data.clear()
                abs_path = os.path.abspath(saved_file)
                st.success(f"✅ Зміни збережено локально в: `{abs_path}`")
                st.info("💡 **Важливо:** Для синхронізації з іншими користувачами завантажте оновлений файл на GitHub вручну")
                st.rerun()
        
        # Кнопки для додаткових дій
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🔄 Оновити дані"):
                st.cache_data.clear()
                st.rerun()
        
        with col2:
            csv = edited_df.to_csv(index=False)
            st.download_button(
                label="📥 Завантажити CSV",
                data=csv,
                file_name=f"lakes_data_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
        # Додаємо новий рядок
        st.subheader("➕ Додати новий запис")
        
        with st.form("add_new_record"):
            col1, col2 = st.columns(2)
            
            with col1:
                new_lakehouse = st.text_input("LakeHouse *", help="Обов'язкове поле")
                new_folder = st.text_input("Folder *", help="Обов'язкове поле")
                new_element = st.text_input("Element *", help="Обов'язкове поле")
                new_url = st.text_input("URL")
            
            with col2:
                new_info = st.text_area("Загальна інформація про лейк")
                new_changes = st.text_area("Внесення змін")
            
            if st.form_submit_button("➕ Додати запис"):
                if new_lakehouse and new_folder and new_element:
                    new_row = {
                        'LakeHouse': new_lakehouse,
                        'Folder': new_folder,
                        'Element': new_element,
                        'URL': new_url if new_url else '',
                        'Загальна інформація про лейк': new_info if new_info else '',
                        'Внесення змін': new_changes if new_changes else ''
                    }
                    
                    # Додаємо новий рядок
                    new_df = pd.concat([lakes_table, pd.DataFrame([new_row])], ignore_index=True)
                    
                    success, saved_file = save_data_to_excel(new_df, EXCEL_FILE_PATH, 
                                                             lakes_table=None, reports_table=reports_table)
                    if success:
                        # Очищуємо кеш, щоб після перезапуску завантажити нові дані
                        st.cache_data.clear()
                        abs_path = os.path.abspath(saved_file)
                        st.success(f"✅ Новий запис додано та збережено в: `{abs_path}`")
                        st.rerun()
                else:
                    st.error("❌ Заповніть обов'язкові поля: LakeHouse, Folder, Element")
    else:
        st.warning("⚠️ Немає даних для редагування. Спочатку завантажте Excel файл.")

# ==================== КОНТАКТИ ТА РЕСУРСИ ====================
elif section == "📞 Контакти та ресурси":
    st.header("📞 Контакти та ресурси")
    st.subheader("👥 Наша команда")
    
    # Команда в одному блоці
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
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        ### Внутрішні ресурси:
        - [SharePoint команди](https://darnitsa.sharepoint.com)
        - [Azure DevOps](https://dev.azure.com/darnitsa)
        - [Power BI Service](https://app.powerbi.com)
        """)
    with col2:
        st.markdown("""
        ### Зовнішні ресурси:
        - [Microsoft Learn](https://learn.microsoft.com)
        - [Power BI Community](https://community.powerbi.com)
        - [Streamlit Docs](https://docs.streamlit.io)
        """)
    

# ==================== КОНТАКТИ ====================

