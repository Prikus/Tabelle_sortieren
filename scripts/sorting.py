import pandas as pd
import toml
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class ExcelSorter:
    def __init__(self, file_path_save):
        current_date = datetime.now().strftime("%d-%m-%Y")

        if file_path_save == "":
            self.file_path_save = os.path.join(os.path.expanduser("~"), "Desktop", f"filtered_{current_date}.xlsx")
        else:
            # Добавляем расширение .xlsx если его нет
            if not file_path_save.endswith('.xlsx'):
                file_path_save += f"/filtered_{current_date}.xlsx"
            self.file_path_save = file_path_save
            
        self.config_path = "config.toml"
    
    def sortExcel(self, lieferung_value, aufschlag_value1, aufschlag_value2, file_path, is_checked):
        # Загрузка конфигурации
        config = toml.load(self.config_path)

        if lieferung_value == "":
            lieferung_value = config['default']['lieferung_value']
        if aufschlag_value1 == "":
            aufschlag_value1 = config['default']['aufschlag_value1']
        if aufschlag_value2 == "":
            aufschlag_value2 = config['default']['aufschlag_value2']

        # Преобразуем значения в числовой тип
        lieferung_value = float(lieferung_value)
        aufschlag_value1 = float(aufschlag_value1)
        aufschlag_value2 = float(aufschlag_value2)
        
        # Чтение значения первой ячейки и проверка
        df = pd.read_excel(file_path)

        # Безопасное извлечение значения первой ячейки
        first_cell_value = df.iat[0, 0]  # Используем .iat для доступа к первой ячейке
        print(f"Значение первой ячейки: '{first_cell_value}'")
        
        # Преобразуем значение в строку для корректного сравнения
        first_cell_value = str(first_cell_value).strip() if first_cell_value is not None else None

        if  first_cell_value == "Brands":
            print("Условие выполнено: первая ячейка пуста")
            self.methodSortZwei(aufschlag_value2, file_path, is_checked)
        else:
            print(f"Условие не выполнено: первая ячейка содержит '{first_cell_value}'")
            self.methodSortEins(lieferung_value, aufschlag_value1, aufschlag_value2, file_path, is_checked)



    def methodSortEins(self, lieferung_value, aufschlag_value1, aufschlag_value2, file_path, is_checked):
        # Загрузка данных из Excel-файла
        df = pd.read_excel(file_path)
        
        config = toml.load(self.config_path)
        bezeichnung_filter = config['dataSet']['data']

        # Оставляем только нужные столбцы
        columns_to_keep = ["Artikelnummer", "Bezeichnung", "Preis", "Zustand"]
        filtered_df = df[columns_to_keep].copy()

        # Фильтрация по первому слову в столбце Bezeichnung
        filtered_df.loc[:, 'First_Word'] = filtered_df['Bezeichnung'].str.split().str[0]
        filtered_df = filtered_df[filtered_df['First_Word'].isin(bezeichnung_filter)]

        # Удаляем временный столбец 'First_Word'
        filtered_df.drop(columns=['First_Word'], inplace=True)

        # Если is_checked равно True, добавляем столбец 'Kaufpreis'
        if is_checked:
            # Сохраняем оригинальные цены перед преобразованием
            filtered_df['Kaufpreis'] = filtered_df['Preis'].copy()

        # Преобразуем столбец 'Preis' в числовой формат
        filtered_df['Preis'] = pd.to_numeric(
            filtered_df['Preis'].astype(str).str.replace('€', '').str.replace(' ', ''),
            errors='coerce'
        )

        # Создаем новый столбец для форматированных цен
        filtered_df['Formatted_Preis'] = filtered_df['Preis'].copy()

        if 'Preis' in filtered_df.columns:
            for index, row in filtered_df.iterrows():
                price = row['Preis']
                
                if pd.notna(price):
                    if price < 500:
                        price = price + aufschlag_value1 + lieferung_value
                        price = price * 1.19
                    else:
                        price = price + (price * aufschlag_value2 / 100) + lieferung_value
                        price = price * 1.19

                    # Записываем форматированную цену в отдельный столбец
                    filtered_df.at[index, 'Formatted_Preis'] = f"{price:.2f} €"

        # Заменяем столбец 'Preis' на форматированный
        filtered_df['Preis'] = filtered_df['Formatted_Preis']
        filtered_df.drop(columns=['Formatted_Preis'], inplace=True)

        try:
            # Сохраняем результат через ExcelWriter
            with pd.ExcelWriter(self.file_path_save, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False)
            
            # Подгоняем размеры столбцов
            self.autofit_columns(self.file_path_save)
            
            print(f"Фильтрация завершена. Новый файл сохранён по пути: {self.file_path_save}")
        except Exception as e:
            print(f"Ошибка при сохранении файла: {str(e)}")
            # Попробуем сохранить на рабочий стол, если возникла ошибка
            backup_path = os.path.join(os.path.expanduser("~"), "Desktop", "filtered_example_backup.xlsx")
            filtered_df.to_excel(backup_path, index=False)
            self.autofit_columns(backup_path)
            print(f"Файл сохранен в резервное расположение: {backup_path}")


    def methodSortZwei(self, aufschlag_value2, file_path, is_checked):
        # Читаем файл
        df = pd.read_excel(file_path)
        
        # Создаем новый DataFrame с реальными заголовками из второй строки
        headers = df.iloc[0]
        filtered_df = df.iloc[1:].copy()
        filtered_df.columns = headers
        filtered_df.reset_index(drop=True, inplace=True)
        
        # Если is_checked равно True, добавляем столбец 'Kaufpreis'
        if is_checked:
            # Сохраняем оригинальные цены перед преобразованием
            filtered_df['Kaufpreis'] = filtered_df['Selling_Price'].copy()



        filtered_df['Selling_Online_Price'] = pd.to_numeric(
            filtered_df['Selling_Online_Price'].astype(str).str.replace('€', '').str.replace(' ', ''),
            errors='coerce'
        )
        filtered_df['Selling_Price'] = pd.to_numeric(
            filtered_df['Selling_Price'].astype(str).str.replace('€', '').str.replace(' ', ''),
            errors='coerce'
        )
        
        # Создаем новый столбец для форматированных цен
        filtered_df['Formatted_Preis'] = filtered_df['Selling_Price'].copy()

        if 'Selling_Price' in filtered_df.columns:
            for index, row in filtered_df.iterrows():
                price = row['Selling_Price']
                
                if pd.notna(price):
                    price = (price + (price * aufschlag_value2 / 100))* 1.19

                    # Убедитесь, что столбец существует и имеет тип данных object (строка)
                    filtered_df['Formatted_Preis'] = filtered_df['Formatted_Preis'].astype(str)

                    # Запись форматированного значения в столбец
                    filtered_df.at[index, 'Formatted_Preis'] = f"{price:.2f} €"


        # Заменяем столбец 'Preis' на форматированный
        filtered_df['Selling_Price'] = filtered_df['Formatted_Preis']
        filtered_df.drop(columns=['Formatted_Preis'], inplace=True)

           # Добавляем символ евро к столбцу Selling_Online_Price
        filtered_df['Selling_Online_Price'] = filtered_df['Selling_Online_Price'].astype(str) + ' €'
        filtered_df.reset_index(drop=True, inplace=True)

        try:
            # Сохраняем результат через ExcelWriter
            with pd.ExcelWriter(self.file_path_save, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, header=True)
            
            # Подгоняем размеры столбцов
            self.autofit_columns(self.file_path_save)
            
            print(f"Фильтрация завершена. Новый файл сохранён по пути: {self.file_path_save}")
        except Exception as e:
            print(f"Ошибка при сохранении файла: {str(e)}")
            # Попробуем сохранить на рабочий стол, если возникла ошибка
            backup_path = os.path.join(os.path.expanduser("~"), "Desktop", "filtered_example_backup.xlsx")
            filtered_df.to_excel(backup_path, index=False, header=True)
            self.autofit_columns(backup_path)
            print(f"Файл сохранен в резервное расположение: {backup_path}")

    def autofit_columns(self, filename):
        """Автоматическая подгонка ширины столбцов"""
        workbook = load_workbook(filename)
        worksheet = workbook.active

        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        workbook.save(filename)


    def read_first_column_value(self, file_path):
        """
        Читает значение первой строки первого столбца Excel файла.
        Если первая строка пустая, возвращает значение второй строки.
        """
        try:
            # Читаем Excel файл
            df = pd.read_excel(file_path)
            
            # Получаем первый столбец
            first_column = df.iloc[:, 0]
            
            # Проверяем первую строку
            if pd.notna(first_column.iloc[0]):
                return first_column.iloc[0]
            
            # Если первая строка пустая, возвращаем вторую строку
            if len(first_column) > 1 and pd.notna(first_column.iloc[1]):
                return first_column.iloc[1]
        
        except Exception as e:
            print(f"Ошибка при чтении файла: {e}")