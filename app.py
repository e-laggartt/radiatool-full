import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import tempfile
import io
import base64

st.set_page_config(
    page_title="RadiaTool Web v2.0",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Инициализация состояния сессии
if 'selected_items' not in st.session_state:
    st.session_state.selected_items = {}
if 'spec_data' not in st.session_state:
    st.session_state.spec_data = []
if 'matrix_data' not in st.session_state:
    st.session_state.matrix_data = []
if 'sheets' not in st.session_state:
    st.session_state.sheets = {}
if 'brackets_df' not in st.session_state:
    st.session_state.brackets_df = None

@st.cache_data
def load_data():
    """Загрузка всех данных"""
    try:
        # Загрузка основной матрицы
        matrix_path = "data/Матрица.xlsx"
        sheets = pd.read_excel(matrix_path, sheet_name=None, engine='openpyxl')
        
        # Загрузка кронштейнов
        brackets_path = "data/Кронштейны.xlsx"
        brackets_df = pd.read_excel(brackets_path, engine='openpyxl')
        brackets_df['Артикул'] = brackets_df['Артикул'].astype(str).str.strip()
        
        # Обработка данных матрицы
        for sheet_name, data in sheets.items():
            data['Артикул'] = data['Артикул'].astype(str).str.strip()
            data['Вес, кг'] = pd.to_numeric(data['Вес, кг'], errors='coerce').fillna(0)
            data['Объем, м3'] = pd.to_numeric(data['Объем, м3'], errors='coerce').fillna(0)
            data['Цена, руб'] = pd.to_numeric(data['Цена, руб'], errors='coerce').fillna(0)
            if 'Мощность, Вт' not in data.columns:
                data['Мощность, Вт'] = ''
        
        return sheets, brackets_df
        
    except Exception as e:
        st.error(f"Ошибка загрузки данных: {e}")
        return {}, None

def main():
    st.title("🔧 RadiaTool Web v2.0")
    st.markdown("---")
    
    # Загрузка данных
    if not st.session_state.sheets:
        st.session_state.sheets, st.session_state.brackets_df = load_data()
    
    # Боковая панель
    with st.sidebar:
        st.header("⚙️ Основные параметры")
        
        # Выбор подключения
        connection = st.selectbox(
            "Тип подключения",
            ["VK-правое", "VK-левое", "K-боковое"],
            key="connection"
        )
        
        # Динамический выбор типа радиатора
        if connection == "VK-левое":
            rad_types = ["10", "11", "30", "33"]
        else:
            rad_types = ["10", "11", "20", "21", "22", "30", "33"]
            
        rad_type = st.selectbox("Тип радиатора", rad_types, key="rad_type")
        
        # Тип крепления
        bracket_type = st.selectbox(
            "Тип крепления",
            ["Настенные кронштейны", "Напольные кронштейны", "Без кронштейнов"],
            key="bracket_type"
        )
        
        st.header("💰 Скидки")
        col1, col2 = st.columns(2)
        with col1:
            radiator_discount = st.number_input(
                "Радиаторы, %", 
                min_value=0.0, max_value=100.0, value=0.0, step=0.1,
                key="radiator_discount"
            )
        with col2:
            bracket_discount = st.number_input(
                "Кронштейны, %", 
                min_value=0.0, max_value=100.0, value=0.0, step=0.1,
                key="bracket_discount"
            )
        
        st.header("🔧 Действия")
        if st.button("🔄 Полный сброс", use_container_width=True):
            reset_all()
        
        if st.button("📊 Обновить матрицу", use_container_width=True):
            update_matrix_data(connection, rad_type)
    
    # Основной контент
    tab1, tab2, tab3 = st.tabs(["📊 Матрица радиаторов", "📋 Спецификация", "⚙️ Дополнительно"])
    
    with tab1:
        display_matrix_interface(connection, rad_type)
    
    with tab2:
        display_specification_interface(radiator_discount, bracket_discount, bracket_type)
    
    with tab3:
        display_additional_tools()

def update_matrix_data(connection, rad_type):
    """Обновление данных матрицы"""
    sheet_name = f"{connection} {rad_type}"
    if sheet_name in st.session_state.sheets:
        data = st.session_state.sheets[sheet_name]
        matrix_data = []
        
        for _, row in data.iterrows():
            try:
                name = str(row['Наименование'])
                name_parts = name.split('/')
                
                if len(name_parts) >= 3:
                    height_str = name_parts[-2].replace('мм', '').strip()
                    length_str = name_parts[-1].replace('мм', '').strip().split()[0]
                    
                    height = int(height_str) if height_str.isdigit() else 0
                    length = int(length_str) if length_str.isdigit() else 0
                    
                    matrix_data.append({
                        'articul': str(row['Артикул']).strip(),
                        'name': name,
                        'height': height,
                        'length': length,
                        'power': str(row.get('Мощность, Вт', '')),
                        'weight': float(row['Вес, кг']),
                        'volume': float(row['Объем, м3']),
                        'price': float(row.get('Цена, руб', 0))
                    })
            except Exception as e:
                continue
        
        st.session_state.matrix_data = matrix_data
        st.success("Матрица обновлена!")
    else:
        st.error(f"Лист '{sheet_name}' не найден")

def display_matrix_interface(connection, rad_type):
    """Отображение интерфейса матрицы"""
    
    if not st.session_state.matrix_data:
        update_matrix_data(connection, rad_type)
    
    st.header("Матрица радиаторов")
    
    # Фильтры
    col1, col2 = st.columns(2)
    with col1:
        min_height = st.number_input("Минимальная высота", value=300, step=100)
    with col2:
        max_height = st.number_input("Максимальная высота", value=900, step=100)
    
    # Фильтрация данных
    filtered_data = [item for item in st.session_state.matrix_data 
                    if min_height <= item['height'] <= max_height]
    
    if not filtered_data:
        st.warning("Нет данных для отображения")
        return
    
    # Создание матрицы
    heights = sorted(list(set(item['height'] for item in filtered_data)))
    lengths = sorted(list(set(item['length'] for item in filtered_data)))
    
    # Создаем сетку для ввода
    st.subheader("Введите количество радиаторов:")
    
    # Заголовок матрицы
    cols = st.columns(len(heights) + 1)
    with cols[0]:
        st.markdown("**Длина/Высота**")
    for i, height in enumerate(heights):
        with cols[i + 1]:
            st.markdown(f"**{height} мм**")
    
    # Строки матрицы
    for length in lengths:
        cols = st.columns(len(heights) + 1)
        
        with cols[0]:
            st.markdown(f"**{length} мм**")
        
        for i, height in enumerate(heights):
            with cols[i + 1]:
                item = next((x for x in filtered_data if x['length'] == length and x['height'] == height), None)
                if item:
                    current_qty = st.session_state.selected_items.get(item['articul'], 0)
                    new_qty = st.number_input(
                        "",
                        min_value=0,
                        value=current_qty,
                        key=f"matrix_{item['articul']}",
                        label_visibility="collapsed"
                    )
                    
                    if new_qty != current_qty:
                        st.session_state.selected_items[item['articul']] = new_qty
                        
                    # Подсказка при наведении
                    with st.expander("", expanded=False):
                        st.caption(f"Арт: {item['articul']}")
                        st.caption(f"Мощность: {item['power']} Вт")
                        st.caption(f"Вес: {item['weight']} кг")
                        st.caption(f"Объем: {item['volume']} м³")
                        st.caption(f"Цена: {item['price']} руб")
                else:
                    st.markdown("—")
    
    # Статистика
    total_selected = sum(st.session_state.selected_items.values())
    st.info(f"🎯 Выбрано радиаторов: {total_selected} шт")

def display_specification_interface(radiator_discount, bracket_discount, bracket_type):
    """Отображение интерфейса спецификации"""
    
    st.header("Спецификация")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("🔄 Рассчитать спецификацию", use_container_width=True):
            calculate_specification(radiator_discount, bracket_discount, bracket_type)
    with col2:
        if st.button("💾 Экспорт в Excel", use_container_width=True):
            export_to_excel(radiator_discount, bracket_discount, bracket_type)
    with col3:
        if st.button("📋 Копировать артикулы", use_container_width=True):
            copy_articuls()
    
    # Отображение спецификации
    if st.session_state.spec_data:
        display_specification_table()
        
        # Итоговая информация
        display_totals()
    else:
        st.info("Рассчитайте спецификацию для отображения данных")

def calculate_specification(radiator_discount, bracket_discount, bracket_type):
    """Расчет полной спецификации"""
    
    try:
        spec_data = []
        
        # Обрабатываем радиаторы
        for articul, qty in st.session_state.selected_items.items():
            if qty <= 0:
                continue
                
            # Ищем товар в данных
            for sheet_name, data in st.session_state.sheets.items():
                product = data[data['Артикул'].str.strip() == articul]
                if not product.empty:
                    product = product.iloc[0]
                    price = float(product.get('Цена, руб', 0))
                    discounted_price = round(price * (1 - radiator_discount / 100), 2)
                    total = round(discounted_price * qty, 2)
                    
                    spec_data.append({
                        "№": len(spec_data) + 1,
                        "Артикул": articul,
                        "Наименование": product['Наименование'],
                        "Мощность, Вт": product.get('Мощность, Вт', ''),
                        "Цена, руб (с НДС)": price,
                        "Скидка, %": radiator_discount,
                        "Цена со скидкой, руб (с НДС)": discounted_price,
                        "Кол-во": qty,
                        "Сумма, руб (с НДС)": total,
                        "Тип": "Радиатор"
                    })
                    break
        
        # Добавляем кронштейны
        if bracket_type != "Без кронштейнов" and st.session_state.brackets_df is not None:
            brackets = calculate_brackets(bracket_type, bracket_discount)
            spec_data.extend(brackets)
        
        st.session_state.spec_data = spec_data
        st.success(f"Спецификация рассчитана: {len(spec_data)} позиций")
        
    except Exception as e:
        st.error(f"Ошибка расчета: {e}")

def calculate_brackets(bracket_type, bracket_discount):
    """Расчет кронштейнов"""
    brackets = []
    
    if not st.session_state.selected_items:
        return brackets
    
    # Упрощенный расчет кронштейнов
    bracket_counts = {}
    
    for articul, qty in st.session_state.selected_items.items():
        if qty <= 0:
            continue
            
        # Находим параметры радиатора
        for sheet_name, data in st.session_state.sheets.items():
            product = data[data['Артикул'].str.strip() == articul]
            if not product.empty:
                product = product.iloc[0]
                name = product['Наименование']
                
                try:
                    name_parts = name.split('/')
                    if len(name_parts) >= 3:
                        height = int(name_parts[-2].replace('мм', '').strip())
                        length = int(name_parts[-1].replace('мм', '').strip().split()[0])
                        rad_type = sheet_name.split()[-1]
                        
                        # Расчет кронштейнов по типу
                        if bracket_type == "Настенные кронштейны":
                            if rad_type in ["10", "11"]:
                                bracket_counts["К9.2L"] = bracket_counts.get("К9.2L", 0) + 2 * qty
                                bracket_counts["К9.2R"] = bracket_counts.get("К9.2R", 0) + 2 * qty
                                if 1700 <= length <= 2000:
                                    bracket_counts["К9.3-40"] = bracket_counts.get("К9.3-40", 0) + 1 * qty
                            elif rad_type in ["20", "21", "22", "30", "33"]:
                                art_map = {300: "К15.4300", 400: "К15.4400", 500: "К15.4500", 
                                          600: "К15.4600", 900: "К15.4900"}
                                if height in art_map:
                                    art = art_map[height]
                                    if 400 <= length <= 1600:
                                        bracket_counts[art] = bracket_counts.get(art, 0) + 2 * qty
                                    elif 1700 <= length <= 2000:
                                        bracket_counts[art] = bracket_counts.get(art, 0) + 3 * qty
                        
                        elif bracket_type == "Напольные кронштейны":
                            # Аналогичная логика для напольных кронштейнов
                            if rad_type in ["10", "11"]:
                                if 300 <= height <= 400:
                                    main_art = "КНС450"
                                elif 500 <= height <= 600:
                                    main_art = "КНС470"
                                elif height == 900:
                                    main_art = "КНС4100"
                                else:
                                    main_art = None
                                
                                if main_art:
                                    bracket_counts[main_art] = bracket_counts.get(main_art, 0) + 2 * qty
                                    if 1700 <= length <= 2000:
                                        bracket_counts["КНС430"] = bracket_counts.get("КНС430", 0) + 1 * qty
                            
                            # ... остальные типы кронштейнов
                                
                except Exception as e:
                    continue
                break
    
    # Формируем записи кронштейнов
    for bracket_art, total_qty in bracket_counts.items():
        if total_qty > 0:
            bracket_info = st.session_state.brackets_df[
                st.session_state.brackets_df['Артикул'] == bracket_art
            ]
            if not bracket_info.empty:
                bracket_info = bracket_info.iloc[0]
                price = float(bracket_info.get('Цена, руб', 0))
                discounted_price = round(price * (1 - bracket_discount / 100), 2)
                total = round(discounted_price * total_qty, 2)
                
                brackets.append({
                    "№": len(brackets) + 1,
                    "Артикул": bracket_art,
                    "Наименование": bracket_info['Наименование'],
                    "Мощность, Вт": '',
                    "Цена, руб (с НДС)": price,
                    "Скидка, %": bracket_discount,
                    "Цена со скидкой, руб (с НДС)": discounted_price,
                    "Кол-во": total_qty,
                    "Сумма, руб (с НДС)": total,
                    "Тип": "Кронштейн"
                })
    
    return brackets

def display_specification_table():
    """Отображение таблицы спецификации"""
    
    df = pd.DataFrame(st.session_state.spec_data)
    
    # Форматируем числовые колонки
    display_df = df.copy()
    numeric_cols = ['Цена, руб (с НДС)', 'Цена со скидкой, руб (с НДС)', 'Сумма, руб (с НДС)']
    for col in numeric_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].map(lambda x: f"{x:,.2f}" if pd.notna(x) else "")
    
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "№": st.column_config.NumberColumn(width="small"),
            "Артикул": st.column_config.TextColumn(width="medium"),
            "Наименование": st.column_config.TextColumn(width="large"),
            "Мощность, Вт": st.column_config.TextColumn(width="small"),
            "Цена, руб (с НДС)": st.column_config.TextColumn(width="medium"),
            "Скидка, %": st.column_config.NumberColumn(width="small"),
            "Цена со скидкой, руб (с НДС)": st.column_config.TextColumn(width="medium"),
            "Кол-во": st.column_config.NumberColumn(width="small"),
            "Сумма, руб (с НДС)": st.column_config.TextColumn(width="medium"),
        }
    )

def display_totals():
    """Отображение итоговой информации"""
    
    if not st.session_state.spec_data:
        return
    
    spec_df = pd.DataFrame(st.session_state.spec_data)
    
    # Расчет итогов
    total_sum = spec_df['Сумма, руб (с НДС)'].sum()
    total_qty_radiators = spec_df[spec_df['Тип'] == 'Радиатор']['Кол-во'].sum()
    total_qty_brackets = spec_df[spec_df['Тип'] == 'Кронштейн']['Кол-во'].sum()
    
    # Расчет веса и объема
    total_weight = 0
    total_volume = 0
    for item in st.session_state.spec_data:
        if item['Тип'] == 'Радиатор':
            articul = item['Артикул']
            qty = item['Кол-во']
            
            for sheet_name, data in st.session_state.sheets.items():
                product = data[data['Артикул'].str.strip() == articul]
                if not product.empty:
                    total_weight += float(product.iloc[0]['Вес, кг']) * qty
                    total_volume += float(product.iloc[0]['Объем, м3']) * qty
                    break
    
    # Отображение в колонках
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Общая сумма", f"{total_sum:,.2f} руб")
    
    with col2:
        st.metric("Радиаторы / Кронштейны", f"{total_qty_radiators} / {total_qty_brackets}")
    
    with col3:
        st.metric("Общий вес", f"{total_weight:.1f} кг")
    
    with col4:
        st.metric("Общий объем", f"{total_volume:.3f} м³")

def export_to_excel(radiator_discount, bracket_discount, bracket_type):
    """Экспорт в Excel"""
    
    if not st.session_state.spec_data:
        st.warning("Нет данных для экспорта")
        return
    
    try:
        # Создаем Excel файл
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Основная спецификация
            spec_df = pd.DataFrame(st.session_state.spec_data)
            spec_df.to_excel(writer, sheet_name='Спецификация', index=False)
            
            # Настройка ширины колонок
            worksheet = writer.sheets['Спецификация']
            column_widths = {'A': 8, 'B': 15, 'C': 60, 'D': 12, 'E': 15, 
                           'F': 10, 'G': 20, 'H': 10, 'I': 15}
            for col, width in column_widths.items():
                worksheet.column_dimensions[col].width = width
        
        excel_data = output.getvalue()
        
        # Кнопка скачивания
        st.download_button(
            label="📥 Скачать Excel файл",
            data=excel_data,
            file_name=f"Расчет_стоимости_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Ошибка экспорта: {e}")

def copy_articuls():
    """Копирование артикулов в буфер"""
    
    if not st.session_state.spec_data:
        st.warning("Нет данных для копирования")
        return
    
    articuls = [item['Артикул'] for item in st.session_state.spec_data if item['Тип'] == 'Радиатор']
    articuls_text = '\n'.join(articuls)
    
    # В Streamlit нет прямого доступа к буферу, поэтому предлагаем скачать текстовый файл
    st.download_button(
        label="📋 Скачать артикулы (TXT)",
        data=articuls_text,
        file_name="артикулы.txt",
        mime="text/plain",
        use_container_width=True
    )

def display_additional_tools():
    """Дополнительные инструменты"""
    
    st.header("Дополнительные функции")
    
    tab1, tab2, tab3 = st.tabs(["Импорт данных", "Информация", "Настройки"])
    
    with tab1:
        st.subheader("Импорт из Excel/CSV")
        
        uploaded_file = st.file_uploader("Загрузите файл спецификации", 
                                       type=['xlsx', 'xls', 'csv'])
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, sep=';')
                else:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                
                st.success(f"Файл загружен: {len(df)} строк")
                
                # Показываем предпросмотр
                st.subheader("Предпросмотр данных")
                st.dataframe(df.head(10), use_container_width=True)
                
                # Автоматическое определение столбцов
                art_col = None
                qty_col = None
                
                for col in df.columns:
                    col_lower = str(col).lower()
                    if any(x in col_lower for x in ['артикул', 'art', 'код']):
                        art_col = col
                    elif any(x in col_lower for x in ['кол-во', 'количество', 'qty']):
                        qty_col = col
                
                if art_col and qty_col:
                    st.info(f"Автоопределение: Артикул - {art_col}, Количество - {qty_col}")
                    
                    if st.button("Импортировать данные"):
                        import_data(df, art_col, qty_col)
                else:
                    st.warning("Не удалось автоматически определить столбцы")
                
            except Exception as e:
                st.error(f"Ошибка загрузки файла: {e}")
    
    with tab2:
        st.subheader("Справка и информация")
        
        st.markdown("""
        ### 📖 Инструкция по использованию
        
        1. **Выбор параметров** - в боковой панели выберите тип подключения, радиатора и крепления
        2. **Заполнение матрицы** - введите количество радиаторов в соответствующих ячейках
        3. **Расчет спецификации** - нажмите кнопку для расчета полной спецификации
        4. **Экспорт** - скачайте результаты в Excel
        
        ### 🔧 Основные возможности
        
        - Подбор радиаторов по типоразмерам
        - Автоматический расчет кронштейнов
        - Учет скидок на радиаторы и кронштейны
        - Экспорт в Excel и копирование артикулов
        - Импорт данных из существующих спецификаций
        """)
    
    with tab3:
        st.subheader("Настройки")
        
        st.number_input("Размер матрицы (строк)", value=20, key="matrix_size")
        st.checkbox("Показывать подсказки", value=True, key="show_tooltips")
        st.checkbox("Авторасчет при изменении", value=True, key="auto_calculate")

def import_data(df, art_col, qty_col):
    """Импорт данных из файла"""
    
    try:
        imported_count = 0
        
        for _, row in df.iterrows():
            art = str(row[art_col]).strip()
            qty = int(float(row[qty_col])) if pd.notna(row[qty_col]) else 0
            
            if qty > 0 and art:
                # Ищем артикул в данных
                for sheet_name, data in st.session_state.sheets.items():
                    if art in data['Артикул'].astype(str).str.strip().values:
                        st.session_state.selected_items[art] = (
                            st.session_state.selected_items.get(art, 0) + qty
                        )
                        imported_count += 1
                        break
        
        st.success(f"Импортировано {imported_count} позиций")
        st.rerun()
        
    except Exception as e:
        st.error(f"Ошибка импорта: {e}")

def reset_all():
    """Полный сброс всех данных"""
    st.session_state.selected_items = {}
    st.session_state.spec_data = []
    st.success("Все данные сброшены!")
    st.rerun()

if __name__ == "__main__":
    main()