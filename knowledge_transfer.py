import streamlit as st
from datetime import datetime
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from PIL import Image
import base64

# ==== CONFIG SECTION ====
# Путь к Excel с лейками и звітами. Для Streamlit Cloud використовуємо відносний шлях
EXCEL_FILE_PATH = os.environ.get("KNOWLEDGE_TRANSFER_CONFIG_PATH", "LakeHouse.xlsx")  # Шлях до твого файлу

# ======= Функція для читання інформації з Excel =========
@st.cache_data(ttl=300)
def load_lakes_and_reports(excel_path):
    """
    Считывает список лейков и отчетов из excel-файла.
    Підтримує різні структури Excel файлів та автоматично визначає доступні листи.
    """
    if not os.path.exists(excel_path):
        st.warning(f"⚠️ Файл не знайдено: {excel_path}")
        return [], [], None, None
    
    try:
        xl = pd.ExcelFile(excel_path)
        available_sheets = xl.sheet_names
        # st.info(f"📋 Доступні листи в Excel: {', '.join(available_sheets)}")
        
        # Спробуємо знайти листи з лейками та звітами
        lakes_df = None
        reports_df = None
        
        # Шукаємо лист з лейками (можливі варіанти назв)
        lake_sheet_names = ['Lakes', 'lakes', 'lake', 'data_lakes', 'лейки', 'Data Lakes']
        for sheet_name in lake_sheet_names:
            if sheet_name in available_sheets:
                lakes_df = pd.read_excel(xl, sheet_name)
                # st.success(f"✅ Знайдено лист з лейками: '{sheet_name}'")
                break
        
        # Якщо не знайшли спеціальний лист, спробуємо другий лист (якщо є)
        if lakes_df is None and len(available_sheets) > 1:
            lakes_df = pd.read_excel(xl, available_sheets[1])
            # st.info(f"📋 Використовуємо другий лист: '{available_sheets[1]}'")
        elif lakes_df is None and available_sheets:
            lakes_df = pd.read_excel(xl, available_sheets[0])
            # st.info(f"📋 Використовуємо перший лист: '{available_sheets[0]}'")
        
        # Шукаємо лист зі звітами
        report_sheet_names = ['Reports', 'reports', 'report', 'звіти', 'Power BI']
        for sheet_name in report_sheet_names:
            if sheet_name in available_sheets:
                reports_df = pd.read_excel(xl, sheet_name)
                # st.success(f"✅ Знайдено лист зі звітами: '{sheet_name}'")
                break
        
        # Якщо не знайшли спеціальний лист, використаємо перший лист
        if reports_df is None and available_sheets:
            reports_df = pd.read_excel(xl, available_sheets[0])
            # st.info(f"📋 Використовуємо перший лист для звітів: '{available_sheets[0]}'")
        
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
        
    except Exception as ex:
        st.error(f"⚠️ Не вдалося зчитати файл {excel_path}: {ex}")
        st.error(f"Деталі помилки: {str(ex)}")
        return [], [], None, None

# ======= Функції для аналізу та візуалізації лейків =========
@st.cache_data(ttl=300)
def analyze_lakes_data(lakes_df):
    """
    Аналізує дані лейків та повертає статистику
    """
    if lakes_df is None or lakes_df.empty:
        return None
    
    analysis = {
        'total_lakes': len(lakes_df),
        'columns': list(lakes_df.columns),
        'data_types': lakes_df.dtypes.to_dict(),
        'missing_data': lakes_df.isnull().sum().to_dict(),
        'unique_values': {}
    }
    
    # Аналіз унікальних значень для кожної колонки
    for col in lakes_df.columns:
        if lakes_df[col].dtype == 'object':  # Текстові колонки
            analysis['unique_values'][col] = lakes_df[col].value_counts().to_dict()
    
    return analysis

def display_image_from_path(image_path, caption=None, width=None):
    """
    Відображає зображення з файлового шляху
    """
    try:
        if os.path.exists(image_path):
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
                display_image_from_path(image_path, width=600)
    else:
        # Якщо немає зображень, просто показуємо текст
        st.markdown(text)

def create_lakes_visualization(lakes_df):
    """
    Створює візуалізації для даних лейків
    """
    if lakes_df is None or lakes_df.empty:
        return None
    
    visualizations = {}
    
    # 1. Статус лейків (якщо є колонка status)
    if 'status' in lakes_df.columns:
        status_counts = lakes_df['status'].value_counts()
        fig_status = px.pie(
            values=status_counts.values, 
            names=status_counts.index,
            title="📊 Розподіл статусів лейків",
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        visualizations['status_pie'] = fig_status
    
    # 2. Частота оновлень (якщо є колонка update_freq)
    if 'update_freq' in lakes_df.columns:
        freq_counts = lakes_df['update_freq'].value_counts()
        fig_freq = px.bar(
            x=freq_counts.index, 
            y=freq_counts.values,
            title="⏰ Частота оновлень лейків",
            labels={'x': 'Частота оновлення', 'y': 'Кількість лейків'},
            color=freq_counts.values,
            color_continuous_scale='Blues'
        )
        fig_freq.update_layout(xaxis_tickangle=-45)
        visualizations['frequency_bar'] = fig_freq
    
    # 3. Workspace розподіл (якщо є колонка workspace)
    if 'workspace' in lakes_df.columns:
        workspace_counts = lakes_df['workspace'].value_counts()
        fig_workspace = px.treemap(
            names=workspace_counts.index,
            parents=[''] * len(workspace_counts),
            values=workspace_counts.values,
            title="🏢 Розподіл лейків по workspace"
        )
        visualizations['workspace_treemap'] = fig_workspace
    
    return visualizations

def create_lake_details_card(lake_row):
    """
    Створює детальну картку для конкретного лейка
    """
    if lake_row is None or lake_row.empty:
        return "Немає даних про лейк"
    
    # Визначаємо назву лейка
    lake_name = lake_row.get('name', lake_row.get('Name', lake_row.get('назва', 'Невідомий лейк')))
    
    # Створюємо HTML картку з інформацією
    card_html = f"""
    <div style="
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        margin: 10px 0;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    ">
        <h3 style="margin: 0 0 15px 0; color: white;">🏞️ {lake_name}</h3>
    """
    
    # Додаємо інформацію з усіх доступних колонок
    for col in lake_row.index:
        if pd.notna(lake_row[col]) and col.lower() not in ['name', 'назва']:
            value = lake_row[col]
            # Перекладаємо назви колонок на українську для кращого розуміння
            col_translations = {
                'workspace': 'Робочий простір',
                'update_freq': 'Частота оновлення',
                'last_update': 'Останнє оновлення',
                'status': 'Статус',
                'owner': 'Власник',
                'description': 'Опис',
                'size': 'Розмір',
                'location': 'Розташування',
                'components': 'Компоненти',
                'tables': 'Таблиці',
                'views': 'Представлення'
            }
            
            display_name = col_translations.get(col.lower(), col)
            card_html += f"""
            <div style="margin: 8px 0;">
                <strong>{display_name}:</strong> {value}
            </div>
            """
    
    card_html += "</div>"
    return card_html

# ДЕМО-ФАЙЛ для генерації якщо його не існує (тільки перший запуск):
def create_default_config_file(path):
    lakes_df = pd.DataFrame({
        "name": ["Sales_Lake", "Inventory_Lake", "HR_Lake", "Finance_Lake"],
        "workspace": ["Sales_Analytics", "Inventory_Analytics", "HR_Analytics", "Finance_Analytics"],
        "update_freq": ["Щодня 06:00", "Кожні 4 години", "Щодня 08:00", "Щотижня"],
        "last_update": ["08.10.2025 06:15", "08.10.2025 12:00", "08.10.2025 08:10", "07.10.2025"],
        "status": ["✅ OK", "✅ OK", "✅ OK", "⚠️ Затримка"]
    })
    reports_df = pd.DataFrame({
        "name": ["Sales Dashboard", "Inventory Report", "HR Analytics", "Financial Overview"],
        "workspace": ["Sales_Analytics", "Inventory_Analytics", "HR_Analytics", "Finance_Analytics"],
        "owner": ["Маркетинг", "Логістика", "HR", "Фінанси"],
        "update_freq": ["Щодня", "Щодня", "Щотижня", "Щомісяця"],
        "lake": ["Sales_Lake", "Inventory_Lake", "HR_Lake", "Finance_Lake"],
        "status": ["✅ OK", "✅ OK", "✅ OK", "⚠️ Потребує уваги"]
    })
    with pd.ExcelWriter(path) as writer:
        lakes_df.to_excel(writer, index=False, sheet_name="lakes")
        reports_df.to_excel(writer, index=False, sheet_name="reports")

if not os.path.exists(EXCEL_FILE_PATH):
    create_default_config_file(EXCEL_FILE_PATH)

# ==== STREAMLIT UI ====

st.set_page_config(
    page_title="База знань - Інструкції по роботі",
    page_icon="📚",
    layout="wide"
)

st.title("📚 База знань: Інструкції по оновленню звітів та лейків")
st.markdown("*Документація для команди Data Engineering*")
st.markdown("---")

st.sidebar.title("🗂️ Навігація")
st.sidebar.markdown("### Оберіть розділ:")

section = st.sidebar.radio(
    "",
    ["🏠 Головна", 
     "💧 Оновлення Data Lakes", 
     "📊 Оновлення Power BI звітів",
     "🔌 Підключення джерел",
     "🆘 Troubleshooting",
     "📞 Контакти та ресурси"]
)

st.sidebar.markdown("---")
st.sidebar.info(f"📅 Останнє оновлення:\n{datetime.now().strftime('%d.%m.%Y')}")

# Додаємо інформацію про доступ для колег
st.sidebar.markdown("---")
st.sidebar.markdown("🌐 **Доступ для колег:**")
st.sidebar.markdown("Для доступу з інших комп'ютерів:")
st.sidebar.markdown("1. Запустіть з командою:")
st.sidebar.code("streamlit run \"C:\\Users\\oleksandra.filatova\\OneDrive - PHARMACEUTICAL COMPANY DARNYTSIA\\Блокноти\\Streamlit\\knowledge_transfer.py\" --server.address 192.168.1.105")
st.sidebar.markdown("2. Дайте колегам посилання:")
st.sidebar.code("http://192.168.1.105:8501")

# === ДИНАМИЧЕСКИЙ ЗАПРОС таблицы Excel для Lakes & reports ===
# Перевіряємо, чи файл існує локально
if os.path.exists(EXCEL_FILE_PATH):
    lakes, reports, lakes_table, reports_table = load_lakes_and_reports(EXCEL_FILE_PATH)
else:
    # Якщо файл не знайдено, пропонуємо завантажити
    st.warning("⚠️ Файл LakeHouse.xlsx не знайдено. Будь ласка, завантажте файл:")
    uploaded_file = st.file_uploader("Завантажте Excel файл", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Зберігаємо завантажений файл
        with open("LakeHouse.xlsx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success("✅ Файл завантажено! Оновлюємо дані...")
        lakes, reports, lakes_table, reports_table = load_lakes_and_reports("LakeHouse.xlsx")
    else:
        # Показуємо заглушку
        lakes, reports, lakes_table, reports_table = [], [], None, None
        st.info("👆 Завантажте Excel файл для початку роботи")

# ==================== ГОЛОВНА СТОРІНКА ====================
if section == "🏠 Головна":
    st.header("Вітаємо! 👋")
    st.markdown("""
    Ця база знань містить всю необхідну інформацію для підтримки та оновлення 
    наших Data Lakes та Power BI звітів.

    **⚡️ Тепер список лейків і звітів зчитується з таблиці Excel**  
    Можна легко коригувати склад без зміни коду!
    
    **Excel файл:** `{}`  
    """.format(EXCEL_FILE_PATH))
    col1, col2 = st.columns(2)
    with col1:
        st.info("""
        **💧 Data Lakes**
        - Інструкції по оновленню
        - Графік оновлень
        - Список всіх лейків
        - Troubleshooting
        """)
        st.success("""
        **📊 Power BI Звіти**
        - Покрокові інструкції
        - Список звітів
        - Власники звітів
        - Часті помилки
        """)
    with col2:
        st.warning("""
        **🔌 Підключення джерел**
        - Connection strings
        - Облікові записи
        - Права доступу
        - API endpoints
        """)
        st.error("""
        **🆘 Що робити якщо...**
        - Звіт не оновлюється
        - Помилки підключення
        - Проблеми з даними
        - Екстрені контакти
        """)
    st.markdown("---")
    st.markdown("### 🚀 Швидкий старт")
    st.markdown("Оберіть розділ з меню зліва 👈")

# ==================== ОНОВЛЕННЯ DATA LAKES ====================
elif section == "💧 Оновлення Data Lakes":
    st.header("💧 Інструкції по оновленню Data Lakes")
    
    # Аналіз даних буде показано після вибору лейка
    
    # Отримуємо унікальні назви лейків (без дублювання)
    unique_lakes = []
    if lakes_table is not None and not lakes_table.empty:
        # Шукаємо колонку з назвами лейків
        name_columns = ['LakeHouse', 'name', 'Name', 'назва', 'Назва', 'lake_name', 'Lake Name']
        name_col = None
        for col in name_columns:
            if col in lakes_table.columns:
                name_col = col
                break
        
        if name_col:
            unique_lakes = lakes_table[name_col].dropna().unique().tolist()
        else:
            # Якщо не знайшли колонку з назвами, використаємо першу колонку
            unique_lakes = lakes_table.iloc[:, 0].dropna().unique().tolist()
    
    lake_select_options = ["Всі лейки", "📊 Аналітика та візуалізація"]
    if unique_lakes:
        lake_select_options += unique_lakes
    
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
            # Створюємо візуалізації
            visualizations = create_lakes_visualization(lakes_table)
            
            if visualizations:
                # Показуємо графіки в колонках
                if 'status_pie' in visualizations:
                    st.plotly_chart(visualizations['status_pie'], use_container_width=True)
                
                col1, col2 = st.columns(2)
                with col1:
                    if 'frequency_bar' in visualizations:
                        st.plotly_chart(visualizations['frequency_bar'], use_container_width=True)
                with col2:
                    if 'workspace_treemap' in visualizations:
                        st.plotly_chart(visualizations['workspace_treemap'], use_container_width=True)
                
                # Додаємо детальний аналіз
                st.subheader("🔍 Детальний аналіз")
                analysis = analyze_lakes_data(lakes_table)
                
                with st.expander("📈 Статистика по колонках"):
                    for col in analysis['columns']:
                        missing_count = analysis['missing_data'][col]
                        total_count = analysis['total_lakes']
                        completeness = ((total_count - missing_count) / total_count) * 100
                        
                        st.write(f"**{col}:** {completeness:.1f}% заповнено ({total_count - missing_count}/{total_count})")
                
                with st.expander("📊 Унікальні значення"):
                    for col, values in analysis['unique_values'].items():
                        st.write(f"**{col}:**")
                        for value, count in values.items():
                            st.write(f"  - {value}: {count}")
            else:
                st.info("Недостатньо даних для створення візуалізацій")
        else:
            st.warning("Немає даних для аналізу")
    
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
                        unique_lakes = lakes_table[col].dropna().unique().tolist()
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
                    
                    # Показуємо тільки унікальні папки
                    unique_folders = lake_data['Folder'].dropna().unique().tolist()
                    
                    if unique_folders:
                        st.write("**Доступні папки:**")
                        
                        # Створюємо клікабельні кнопки для папок
                        cols = st.columns(min(len(unique_folders), 3))  # Максимум 3 колонки
                        selected_folder = None
                        
                        for i, folder in enumerate(unique_folders):
                            col_idx = i % 3
                            with cols[col_idx]:
                                if st.button(f"📂 {folder}", key=f"folder_{folder}", use_container_width=True):
                                    selected_folder = folder
                        
                        # Показуємо деталі вибраної папки
                        if selected_folder:
                            st.success(f"📂 Вибрано папку: **{selected_folder}**")
                            
                            # Фільтруємо дані по вибраній папці
                            folder_data = lake_data[lake_data['Folder'] == selected_folder]
                            
                            # Показуємо елементи папки (тільки стовпці з 3 по 8)
                            st.subheader("🧩 Елементи папки")
                            
                            # Вибираємо стовпці з 3 по 8 (індекси 2-7), але виключаємо URL
                            display_columns = folder_data.columns[2:9]
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
                    st.subheader("📋 Доступні лейки:")
                    for lake in unique_lakes:
                        st.write(f"• {lake}")
                else:
                    st.warning("⚠️ Немає доступних лейків в базі даних")
        else:
            st.warning("⚠️ Немає даних про лейки")

# ==================== ОНОВЛЕННЯ POWER BI ====================
elif section == "📊 Оновлення Power BI звітів":
    st.header("📊 Інструкції по оновленню Power BI звітів")
    report_select_options = ["Всі звіти"]
    if reports:
        report_select_options += reports
    report_name = st.selectbox(
        "Оберіть звіт:",
        report_select_options
    )
    if report_name == "Всі звіти":
        st.info("👈 Оберіть конкретний звіт зі списку вище")
        st.subheader("📋 Список всіх звітів")
        if reports_table is not None and not reports_table.empty:
            st.dataframe(reports_table, use_container_width=True)
        else:
            st.warning("Список звітів порожній у файлі Excel!")
    else:
        # Підтягнути максимум detail по звіту з Excel
        if reports_table is not None and report_name in reports_table["name"].values:
            r_row = reports_table[reports_table["name"] == report_name].iloc[0]
            info_md = f"""
            **Назва звіту:** {r_row['name']}  
            **Workspace:** {r_row.get('workspace','')}  
            **Власник:** {r_row.get('owner','')}  
            **Частота оновлення:** {r_row.get('update_freq','')}  
            **Джерело даних (лейк):** {r_row.get('lake', '')}
            **Статус:** {r_row.get('status', '')}  
            """
        else:
            info_md = f"**Назва звіту:** {report_name}"
        st.success(f"Інструкція для: **{report_name}**")
        with st.expander("ℹ️ Інформація про звіт", expanded=True):
            st.markdown(info_md)
        st.subheader("📝 Як оновити звіт")
        with st.expander("Крок 1️⃣: Перевірка даних в Lakehouse", expanded=True):
            st.markdown("""
            Перевірте, що дані в пов'язаному Lakehouse актуальні, перед оновленням звіту.
            """)
            st.checkbox("✓ Дані в Lake актуальні", key=f"{report_name}_lake_data")
        with st.expander("Крок 2️⃣: Оновлення Dataset в Power BI Service"):
            st.markdown("""
            Оновіть dataset у Power BI — вручну або через plan/schedule.
            """)
            st.checkbox("✓ Dataset оновлено", key=f"{report_name}_pbirefresh")
        with st.expander("Крок 3️⃣: Перевірка звіту"):
            st.markdown("""
            Перевірте головні сторінки, фільтри, дати та візуали на коректність.
            """)
            st.checkbox("✓ Звіт працює коректно", key=f"{report_name}_ok")
        st.success("✅ Готово! Якщо всі чекбокси відмічені - звіт оновлено успішно")

# ==================== ПІДКЛЮЧЕННЯ ДЖЕРЕЛ ====================
elif section == "🔌 Підключення джерел":
    st.header("🔌 Підключення джерел даних")
    st.warning("⚠️ **ВАЖЛИВО:** Всі паролі зберігаються в Azure Key Vault. Ніколи не записуйте їх у відкритому вигляді!")
    source_type = st.selectbox(
        "Оберіть тип джерела:",
        ["Всі джерела", "SQL Server", "OData (1С)", "REST API", "SharePoint", "Excel файли"]
    )
    if source_type == "SQL Server":
        st.subheader("🗄️ SQL Server підключення")
        with st.expander("📍 Production SQL Server", expanded=True):
            st.markdown("""
            ### Connection String:
            ```
            Server=sql-prod-server.database.windows.net;
            Database=Production_DB;
            Authentication=Active Directory Integrated;
            ```
            ### Облікові дані:
            - **Username:** Зберігається в Key Vault (`sql-prod-username`)
            - **Password:** Зберігається в Key Vault (`sql-prod-password`)
            ### Як підключитися з Fabric:
            1. Data Factory → New Connection
            2. Оберіть "SQL Server"
            3. Введіть server name
            4. Authentication method: SQL Authentication
            5. Використайте credentials з Key Vault
            ### Таблиці:
            - `dbo.Sales` - дані продажів
            - `dbo.Customers` - клієнти
            - `dbo.Products` - продукти
            ### Відповідальний: Петров П.П.
            ### 📞 Контакт: petrov@company.com
            """)
    elif source_type == "OData (1С)":
        st.subheader("🔗 OData підключення до 1С")
        with st.expander("📍 1С Production OData", expanded=True):
            st.markdown("""
            ### Endpoint URL:
            ```
            https://1c-server.company.local/production/odata/standard.odata/
            ```
            ### Автентифікація:
            - **Тип:** Basic Authentication
            - **Username:** Зберігається в Key Vault (`1c-odata-username`)
            - **Password:** Зберігається в Key Vault (`1c-odata-password`)
            
            ### Як підключитися з Fabric:
            1. Створіть новий Data Source
            2. Оберіть "OData"
            3. Введіть URL endpoint
            4. Оберіть Basic Authentication
            5. Введіть credentials

            ### Доступні ендпоінти:
            - `Catalog_Номенклатура` - довідник номенклатури
            - `Document_РеалізаціяТоварівТаПослуг` - документи продажів
            - `InformationRegister_ЗалишкиТоварів` - залишки товарів
            
            ### ⚠️ Обмеження:
            - Максимум 1000 записів за запит (використовуйте $top і $skip)
            - Rate limit: 100 запитів на хвилину
            
            ### 💡 Приклад запиту:
            ```
            GET /Catalog_Номенклатура?$top=100&$select=Code,Description
            ```
            ### Відповідальний: Сидоров С.С.
            ### 📞 Контакт: sidorov@company.com
            """)
    else:
        st.info("Оберіть тип джерела зі списку вище, щоб побачити детальну інформацію")

# ==================== TROUBLESHOOTING ====================
elif section == "🆘 Troubleshooting":
    st.header("🆘 Вирішення проблем")
    st.markdown("Тут зібрані найчастіші проблеми та їх рішення")
    problem = st.selectbox(
        "Оберіть проблему:",
        [
            "Оберіть проблему...",
            "Pipeline падає з помилкою",
            "Дані не оновлюються",
            "Помилка підключення до джерела",
            "Power BI звіт показує старі дані",
            "Повільне оновлення",
            "Помилки автентифікації"
        ]
    )
    if problem == "Pipeline падає з помилкою":
        st.error("### ❌ Pipeline падає з помилкою")
        with st.expander("💡 Рішення 1: Перевірте логи", expanded=True):
            st.markdown("""
            ### Як подивитися логи:
            1. Відкрийте Fabric
            2. Знайдіть ваш pipeline
            3. Відкрийте історію запусків (Run history)
            4. Клікніть на проблемний запуск
            5. Перегляньте детальні логи
            
            ### Що шукати в логах:
            - 🔴 **"Timeout"** → джерело не відповідає, перевірте доступність
            - 🔴 **"Authentication failed"** → проблема з credentials
            - 🔴 **"Permission denied"** → немає прав доступу
            - 🔴 **"Schema mismatch"** → структура даних змінилася
            """)
        with st.expander("💡 Рішення 2: Перезапустіть pipeline"):
            st.markdown("""
            ### Кроки:
            1. Дочекайтеся завершення поточного запуску (навіть якщо він з помилкою)
            2. Натисніть "Run again"
            3. Якщо проблема повторюється - дивіться інші рішення
            """)
        with st.expander("💡 Рішення 3: Перевірте джерело даних"):
            st.markdown("""
            ### Як перевірити:
            1. Спробуйте підключитися до джерела вручну
            2. Виконайте простий запит
            3. Перевірте, чи доступний сервер
            
            ### Інструменти для перевірки:
            - SQL Server: SQL Server Management Studio
            - OData: браузер або Postman
            - API: Postman або curl
            """)
    elif problem == "Дані не оновлюються":
        st.error("### ⚠️ Дані не оновлюються")
        st.markdown("""
        ### Чеклист перевірки:
        """)
        check1 = st.checkbox("✓ Pipeline виконався успішно (без помилок)")
        check2 = st.checkbox("✓ В джерелі є нові дані")
        check3 = st.checkbox("✓ Dataset в Power BI оновлено після pipeline")
        check4 = st.checkbox("✓ Перевірив фільтри в звіті (можливо, відфільтровані нові дані)")
        check5 = st.checkbox("✓ Очистив кеш браузера")
        if all([check1, check2, check3, check4, check5]):
            st.success("Якщо всі пункти виконані, але дані все одно старі - зверніться до IT Support")
    elif problem == "Помилка підключення до джерела":
        st.error("### 🔌 Помилка підключення до джерела")
        st.markdown("""
        ### Можливі причини:
        1. **Неправильні credentials**
           - Перевірте Key Vault
           - Перевірте, чи не закінчився термін дії паролю
        2. **Джерело недоступне**
           - Перевірте, чи працює сервер
           - Можливо, проводяться технічні роботи
           - Перевірте firewall rules
        3. **Мережеві проблеми**
           - Перевірте VPN підключення
           - Перевірте, чи IP Fabric додано до whitelist
        4. **Закінчилися ліміти**
           - Можливо, перевищено ліміт запитів до API
           - Зачекайте 15-30 хвилин та спробуйте знову
        """)
    else:
        st.info("Оберіть проблему зі списку вище, щоб побачити рішення")
    st.markdown("---")
    st.warning("""
    ### 🆘 Якщо нічого не допомогло:
    1. **Зателефонуйте до IT Support:** +380 XX XXX-XX-XX
    2. **Напишіть в Teams:** канал #data-engineering-support
    3. **Email:** support@company.com
    ### ❗ Екстрені ситуації (звіти для керівництва не працюють):
    - Телефон технічного директора: +380 XX XXX-XX-XX
    - Telegram: @tech_director
    """)

# ==================== КОНТАКТИ ====================
elif section == "📞 Контакти та ресурси":
    st.header("📞 Контакти та ресурси")
    st.subheader("👥 Команда")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        ### Data Engineering Team

        **Іванов Іван Іванович**  
        Data Engineer (Lake/Pipeline)  
        📧 ivanov@company.com  
        📱 +380 XX XXX-XX-XX  
        💬 Teams: @ivanov

        ---

        **Петров Петро Петрович**  
        Data Engineer (Power BI)  
        📧 petrov@company.com  
        📱 +380 XX XXX-XX-XX  
        💬 Teams: @petrov
        """)
    with col2:
        st.markdown("""
        ### IT Support

        **Сидоров Сергій Сергійович**  
        System Administrator  
        📧 sidorov@company.com  
        📱 +380 XX XXX-XX-XX  
        💬 Teams: @sidorov

        ---

        **IT Support загальний**  
        📧 support@company.com  
        📱 +380 XX XXX-XX-XX (гаряча лінія)  
        🕐 Пн-Пт: 9:00-18:00
        """)
    st.markdown("---")
    st.subheader("🔗 Корисні посилання")
    st.markdown("""
    ### Робочі системи:
    - 🌐 [Microsoft Fabric Portal](https://fabric.microsoft.com)
    - 📊 [Power BI Service](https://app.powerbi.com)
    - 🔐 [Azure Key Vault](https://portal.azure.com)
    - 📂 [SharePoint - Документація](https://company.sharepoint.com/documentation)
    
    ### Документація:
    - 📚 [Microsoft Fabric Docs](https://learn.microsoft.com/fabric/)
    - 📚 [Power BI Docs](https://learn.microsoft.com/power-bi/)
    - 📚 [Внутрішня Wiki](https://wiki.company.local)
    
    ### Для навчання:
    - 🎓 [Microsoft Learn - Fabric](https://learn.microsoft.com/training/fabric/)
    - 🎓 [Power BI Training](https://learn.microsoft.com/training/powerplatform/power-bi)
    - 🎥 [Відео уроки (внутрішні)](https://company.sharepoint.com/videos)
    """)
    st.markdown("---")
    st.subheader("📝 Шаблони та скрипти")
    with st.expander("💾 Шаблон connection string для SQL"):
        st.code("""
Server=YOUR_SERVER.database.windows.net;
Database=YOUR_DATABASE;
Authentication=Active Directory Integrated;
        """, language="text")
    with st.expander("💾 Шаблон запиту до OData"):
        st.code("""
GET https://your-endpoint/EntityName?$top=100&$skip=0&$select=Field1,Field2&$filter=Date gt 2025-01-01
        """, language="text")
    with st.expander("💾 Скрипт перевірки даних в Lake"):
        st.code("""
SELECT 
    COUNT(*) as total_records,
    MAX(LoadDate) as last_load_date,
    MIN(LoadDate) as first_load_date,
    COUNT(DISTINCT CustomerID) as unique_customers
FROM Sales_Lake.FactSales
WHERE LoadDate >= DATEADD(day, -7, GETDATE())
        """, language="sql")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray;'>
    📚 База знань Data Engineering Team | Версія 1.0 | Жовтень 2025
</div>
""", unsafe_allow_html=True)
