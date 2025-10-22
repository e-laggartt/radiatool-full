import streamlit as st
import pandas as pd
import io
import tempfile
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import base64

# Настройка страницы
st.set_page_config(
    page_title="RadiaTool Web v1.9",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

class RadiatorWebApp:
    def __init__(self):
        self.sheets = {}
        self.brackets_df = pd.DataFrame()
        self.entry_values = {}
        self.initialize_session_state()
        self.load_data()
    
    def initialize_session_state(self):
        """Инициализация состояния сессии"""
        defaults = {
            'connection_var': 'VK-правое',
            'radiator_type_var': '10',
            'bracket_var': 'Настенные кронштейны',
            'radiator_discount': 0,
            'bracket_discount': 0,
            'show_tooltips': False,
            'spec_data': None
        }
        
        for key, value in defaults.items():
            if key not in st.session_state:
                st.session_state[key] = value
    
    def load_data(self):
        """Загрузка данных из встроенных Excel файлов"""
        try:
            # Загрузка матрицы радиаторов
            matrix_path = "data/Матрица.xlsx"
            if os.path.exists(matrix_path):
                self.sheets = pd.read_excel(matrix_path, sheet_name=None, engine='openpyxl')
                
                # Обработка кронштейнов
                if "Кронштейны" in self.sheets:
                    self.brackets_df = self.sheets["Кронштейны"].copy()
                    self.brackets_df['Артикул'] = self.brackets_df['Артикул'].astype(str).str.strip()
                    del self.sheets["Кронштейны"]
                
                # Обработка остальных листов
                for sheet_name, data in self.sheets.items():
                    data['Артикул'] = data['Артикул'].astype(str).str.strip()
                    data['Вес, кг'] = pd.to_numeric(data['Вес, кг'], errors='coerce').fillna(0)
                    data['Объем, м3'] = pd.to_numeric(data['Объем, м3'], errors='coerce').fillna(0)
                    
        except Exception as e:
            st.error(f"Ошибка загрузки данных: {str(e)}")
    
    def create_interface(self):
        """Создание веб-интерфейса"""
        st.title("🔧 RadiaTool Web v1.9")
        st.markdown("---")
        
        # Боковая панель - меню
        with st.sidebar:
            st.header("Меню")
            
            if st.button("🔄 Создать спецификацию METEOR"):
                self.generate_spec("excel")
            
            if st.button("📊 Создать файл METEOR CSV"):
                self.generate_spec("csv")
            
            st.markdown("---")
            st.header("Загрузка данных")
            
            uploaded_file = st.file_uploader("Загрузить спецификацию", 
                                           type=['xlsx', 'xls', 'csv'])
            if uploaded_file:
                self.handle_file_upload(uploaded_file)
            
            st.markdown("---")
            st.header("Информация")
            
            if st.button("📖 Инструкция"):
                self.show_instruction()
            
            if st.button("📄 Лицензионное соглашение"):
                self.show_license()
        
        # Основная область - матрица радиаторов
        self.create_matrix_interface()
        
        # Область спецификации
        self.create_spec_preview()
    
    def create_matrix_interface(self):
        """Создание интерфейса матрицы радиаторов"""
        st.header("Матрица радиаторов")
        
        # Выбор параметров
        col1, col2, col3 = st.columns(3)
        
        with col1:
            connection = st.selectbox(
                "Вид подключения",
                ["VK-правое", "VK-левое", "K-боковое"],
                index=0,
                key="connection_var"
            )
        
        with col2:
            # Доступные типы радиаторов в зависимости от подключения
            if st.session_state.connection_var == "VK-левое":
                types = ["10", "11", "30", "33"]
            else:
                types = ["10", "11", "20", "21", "22", "30", "33"]
            
            radiator_type = st.selectbox(
                "Тип радиатора",
                types,
                index=0,
                key="radiator_type_var"
            )
        
        with col3:
            bracket_type = st.selectbox(
                "Тип крепления",
                ["Настенные кронштейны", "Напольные кронштейны", "Без кронштейнов"],
                index=0,
                key="bracket_var"
            )
        
        # Скидки
        col4, col5 = st.columns(2)
        with col4:
            st.session_state.radiator_discount = st.number_input(
                "Скидка на радиаторы, %",
                min_value=0.0,
                max_value=100.0,
                value=0.0,
                step=0.5
            )
        
        with col5:
            st.session_state.bracket_discount = st.number_input(
                "Скидка на кронштейны, %",
                min_value=0.0,
                max_value=100.0,
                value=0.0,
                step=0.5
            )
        
        # Матрица радиаторов
        self.display_radiator_matrix()
    
    def display_radiator_matrix(self):
        """Отображение матрицы радиаторов"""
        sheet_name = f"{st.session_state.connection_var} {st.session_state.radiator_type_var}"
        
        if sheet_name not in self.sheets:
            st.error(f"Лист '{sheet_name}' не найден")
            return
        
        data = self.sheets[sheet_name]
        lengths = list(range(400, 2100, 100))
        heights = [300, 400, 500, 600, 900]
        
        # Создание матрицы
        st.subheader("Матрица выбора радиаторов")
        st.markdown("**Высота радиаторов, мм →**")
        
        # Заголовки столбцов (высоты)
        cols = st.columns(len(heights) + 1)
        with cols[0]:
            st.markdown("**Длина ↓**")
        for j, h in enumerate(heights):
            with cols[j + 1]:
                st.markdown(f"**{h}**")
        
        # Строки матрицы
        for i, length in enumerate(lengths):
            cols = st.columns(len(heights) + 1)
            
            with cols[0]:
                st.markdown(f"**{length}**")
            
            for j, height in enumerate(heights):
                with cols[j + 1]:
                    self.create_matrix_cell(sheet_name, data, length, height)
    
    def create_matrix_cell(self, sheet_name, data, length, height):
        """Создание ячейки матрицы"""
        pattern = f"/{height}/{length}"
        match = data[data['Наименование'].str.contains(pattern, na=False)]
        
        if not match.empty:
            product = match.iloc[0]
            art = str(product['Артикул']).strip()
            
            # Получение текущего значения
            current_value = self.entry_values.get((sheet_name, art), "")
            
            # Поле ввода
            new_value = st.text_input(
                "",
                value=current_value,
                key=f"{sheet_name}_{art}",
                label_visibility="collapsed",
                placeholder="0"
            )
            
            # Сохранение значения
            if new_value != current_value:
                if new_value.strip():
                    self.entry_values[(sheet_name, art)] = new_value
                else:
                    self.entry_values.pop((sheet_name, art), None)
            
            # Подсказка при наведении
            if st.session_state.show_tooltips and st.session_state.get(f"hover_{sheet_name}_{art}"):
                power = product.get('Мощность, Вт', '')
                weight = product.get('Вес, кг', '')
                volume = product.get('Объем, м3', '')
                
                st.caption(f"Арт: {art}")
                st.caption(f"Мощность: {power} Вт")
                st.caption(f"Вес: {weight} кг")
    
    def create_spec_preview(self):
        """Создание предпросмотра спецификации"""
        st.header("Предпросмотр спецификации")
        
        if st.button("🔄 Обновить спецификацию"):
            spec_data = self.prepare_spec_data()
            if spec_data is not None:
                st.session_state.spec_data = spec_data
        
        if st.session_state.spec_data is not None:
            self.display_spec_table(st.session_state.spec_data)
            
            # Кнопки экспорта
            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 Экспорт в Excel"):
                    self.download_excel(st.session_state.spec_data)
            with col2:
                if st.button("📄 Экспорт в CSV"):
                    self.download_csv(st.session_state.spec_data)
    
    def display_spec_table(self, spec_data):
        """Отображение таблицы спецификации"""
        # Форматирование данных для отображения
        display_data = spec_data.copy()
        display_data['Цена, руб (с НДС)'] = display_data['Цена, руб (с НДС)'].round(2)
        display_data['Цена со скидкой, руб (с НДС)'] = display_data['Цена со скидкой, руб (с НДС)'].round(2)
        display_data['Сумма, руб (с НДС)'] = display_data['Сумма, руб (с НДС)'].round(2)
        
        st.dataframe(
            display_data,
            use_container_width=True,
            hide_index=True
        )
        
        # Итоги
        total_sum = spec_data["Сумма, руб (с НДС)"].sum()
        total_qty_radiators = sum(spec_data.query("Наименование.str.contains('Радиатор')")["Кол-во"])
        total_qty_brackets = sum(spec_data.query("Наименование.str.contains('Кронштейн')")["Кол-во"])
        
        st.markdown(f"**Итого:** Количество: {total_qty_radiators} / {total_qty_brackets} | Сумма: {total_sum:.2f} руб")
    
    def prepare_spec_data(self):
        """Подготовка данных для спецификации (аналогично десктопной версии)"""
        try:
            spec_data = []
            radiator_data = []
            bracket_data = []
            brackets_temp = {}

            # Обработка радиаторов
            for (sheet_name, art), value in self.entry_values.items():
                if value and sheet_name in self.sheets:
                    qty_radiator = self.parse_quantity(value)
                    mask = self.sheets[sheet_name]['Артикул'] == art
                    product = self.sheets[sheet_name].loc[mask]
                    
                    if product.empty:
                        continue
                    
                    product = product.iloc[0]
                    price = float(product['Цена, руб'])
                    discount = st.session_state.radiator_discount
                    discounted_price = round(price * (1 - discount / 100), 2)
                    total = round(discounted_price * qty_radiator, 2)
                    
                    radiator_data.append({
                        "№": len(radiator_data) + 1,
                        "Артикул": str(product['Артикул']).strip(),
                        "Наименование": str(product['Наименование']),
                        "Мощность, Вт": float(product.get('Мощность, Вт', 0)),
                        "Цена, руб (с НДС)": float(price),
                        "Скидка, %": float(discount),
                        "Цена со скидкой, руб (с НДС)": float(discounted_price),
                        "Кол-во": int(qty_radiator),
                        "Сумма, руб (с НДС)": float(total)
                    })

                    # Обработка кронштейнов
                    if st.session_state.bracket_var != "Без кронштейнов":
                        # Упрощенный расчет кронштейнов
                        brackets = self.calculate_brackets_simple(
                            str(product['Наименование']),
                            qty_radiator
                        )
                        
                        for art_bracket, qty_bracket in brackets:
                            mask_bracket = self.brackets_df['Артикул'] == art_bracket
                            bracket_info = self.brackets_df.loc[mask_bracket]
                            
                            if bracket_info.empty:
                                continue
                                
                            key = art_bracket.strip()
                            if key not in brackets_temp:
                                price_bracket = float(bracket_info.iloc[0]['Цена, руб'])
                                discount_bracket = st.session_state.bracket_discount
                                discounted_price_bracket = round(price_bracket * (1 - discount_bracket / 100), 2)
                                
                                brackets_temp[key] = {
                                    "Артикул": art_bracket,
                                    "Наименование": str(bracket_info.iloc[0]['Наименование']),
                                    "Цена, руб (с НДС)": float(price_bracket),
                                    "Скидка, %": float(discount_bracket),
                                    "Цена со скидкой, руб (с НДС)": float(discounted_price_bracket),
                                    "Кол-во": 0,
                                    "Сумма, руб (с НДС)": 0.0
                                }
                            
                            brackets_temp[key]["Кол-во"] += int(qty_bracket)
                            brackets_temp[key]["Сумма, руб (с НДС)"] += round(
                                brackets_temp[key]["Цена со скидкой, руб (с НДС)"] * qty_bracket, 2
                            )

            # Формирование данных кронштейнов
            for b in brackets_temp.values():
                bracket_data.append({
                    "№": len(radiator_data) + len(bracket_data) + 1,
                    "Артикул": str(b["Артикул"]),
                    "Наименование": str(b["Наименование"]),
                    "Мощность, Вт": 0.0,
                    "Цена, руб (с НДС)": float(b["Цена, руб (с НДС)"]),
                    "Скидка, %": float(b["Скидка, %"]),
                    "Цена со скидкой, руб (с НДС)": float(b["Цена со скидкой, руб (с НДС)"]),
                    "Кол-во": int(b["Кол-во"]),
                    "Сумма, руб (с НДС)": float(b["Сумма, руб (с НДС)"])
                })

            # Объединение данных
            combined_data = radiator_data + bracket_data
            
            if not combined_data:
                st.warning("Нет данных для формирования спецификации")
                return None

            # Создание DataFrame
            df = pd.DataFrame(
                combined_data,
                columns=[
                    "№", "Артикул", "Наименование", "Мощность, Вт",
                    "Цена, руб (с НДС)", "Скидка, %",
                    "Цена со скидкой, руб (с НДС)", "Кол-во",
                    "Сумма, руб (с НДС)"
                ]
            )
            
            return df
            
        except Exception as e:
            st.error(f"Ошибка подготовки спецификации: {str(e)}")
            return None
    
    def calculate_brackets_simple(self, radiator_name, qty_radiator):
        """Упрощенный расчет кронштейнов"""
        brackets = []
        
        if "тип 10" in radiator_name or "тип 11" in radiator_name:
            brackets.append(("К9.2L", 2 * qty_radiator))
            brackets.append(("К9.2R", 2 * qty_radiator))
        elif "тип 20" in radiator_name or "тип 21" in radiator_name or "тип 22" in radiator_name:
            brackets.append(("К15.4500", 2 * qty_radiator))
        elif "тип 30" in radiator_name or "тип 33" in radiator_name:
            brackets.append(("К15.4500", 3 * qty_radiator))
        
        return brackets
    
    def parse_quantity(self, value):
        """Парсинг количества (аналогично десктопной версии)"""
        try:
            if not value:
                return 0
            
            if isinstance(value, (int, float)):
                return int(round(float(value)))
            
            value = str(value).strip()
            
            # Удаление лишних '+'
            while value.startswith('+'):
                value = value[1:]
            while value.endswith('+'):
                value = value[:-1]
            
            if not value:
                return 0
            
            # Суммирование частей
            parts = value.split('+')
            total = 0
            for part in parts:
                part = part.strip()
                if part:
                    total += int(round(float(part)))
                    
            return total
        except:
            return 0
    
    def download_excel(self, spec_data):
        """Скачивание Excel файла"""
        try:
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                spec_data.to_excel(writer, sheet_name='Спецификация', index=False)
                
                # Форматирование
                workbook = writer.book
                worksheet = writer.sheets['Спецификация']
                
                # Заголовки
                for cell in worksheet[1]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
            
            output.seek(0)
            
            st.download_button(
                label="📥 Скачать Excel файл",
                data=output,
                file_name="Спецификация_радиаторов.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"Ошибка создания Excel: {str(e)}")
    
    def download_csv(self, spec_data):
        """Скачивание CSV файла"""
        try:
            csv_data = spec_data[['Артикул', 'Кол-во']].to_csv(index=False, sep=';')
            
            st.download_button(
                label="📥 Скачать CSV файл",
                data=csv_data,
                file_name="спецификация.csv",
                mime="text/csv"
            )
            
        except Exception as e:
            st.error(f"Ошибка создания CSV: {str(e)}")
    
    def handle_file_upload(self, uploaded_file):
        """Обработка загруженного файла"""
        try:
            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            elif uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep=';')
            else:
                st.error("Неподдерживаемый формат файла")
                return
            
            # Автоопределение столбцов
            art_col = None
            qty_col = None
            
            for col in df.columns:
                col_lower = str(col).lower()
                if 'артикул' in col_lower or 'art' in col_lower:
                    art_col = col
                elif 'кол-во' in col_lower or 'количество' in col_lower:
                    qty_col = col
            
            if art_col is None:
                art_col = df.columns[0]
            if qty_col is None and len(df.columns) > 1:
                qty_col = df.columns[1]
            
            if qty_col is None:
                st.error("Не найден столбец с количеством")
                return
            
            # Обработка данных
            loaded_count = 0
            for _, row in df.iterrows():
                art = str(row[art_col]).strip()
                qty = self.parse_quantity(row[qty_col])
                
                if qty > 0 and art:
                    # Поиск артикула в данных
                    for sheet_name, sheet_data in self.sheets.items():
                        if art in sheet_data['Артикул'].astype(str).str.strip().values:
                            self.entry_values[(sheet_name, art)] = str(qty)
                            loaded_count += 1
                            break
            
            st.success(f"Загружено {loaded_count} позиций")
            
        except Exception as e:
            st.error(f"Ошибка обработки файла: {str(e)}")
    
    def generate_spec(self, file_type):
        """Генерация спецификации"""
        spec_data = self.prepare_spec_data()
        if spec_data is not None:
            st.session_state.spec_data = spec_data
            st.success("Спецификация сгенерирована!")
    
    def show_instruction(self):
        """Показать инструкцию"""
        st.info("""
        **ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ**
        
        1. Выберите параметры радиатора в верхней части
        2. Заполните матрицу количествами
        3. Нажмите "Обновить спецификацию"
        4. Скачайте результат в нужном формате
        """)
    
    def show_license(self):
        """Показать лицензионное соглашение"""
        st.info("""
        **ЛИЦЕНЗИОННОЕ СОГЛАШЕНИЕ**
        
        Программное обеспечение предназначено для формирования спецификаций 
        на радиаторы METEOR. Все права защищены.
        """)

def main():
    app = RadiatorWebApp()
    app.create_interface()

if __name__ == "__main__":
    main()