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


    def sortExcel(self, lieferung_value, aufschlag_value1, aufschlag_value2, file_path, is_checked):
        # Загрузка данных из Excel-файла
        df = pd.read_excel(file_path)

        # Загрузка конфигурации
        config = toml.load(self.config_path)
        bezeichnung_filter = config['dataSet']['data']

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