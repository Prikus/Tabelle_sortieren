import pandas as pd
import toml
import os
import requests
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from concurrent.futures import ThreadPoolExecutor

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
    
    def sortExcel(self, lieferung_value, aufschlag_value1, aufschlag_value2, file_path, checkbox_kaufpreis, checkbox_steuer):
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
            self.methodSortZwei(aufschlag_value1, file_path, checkbox_kaufpreis, checkbox_steuer)
        else:
            print(f"Условие не выполнено: первая ячейка содержит '{first_cell_value}'")
            self.methodSortEins(lieferung_value, aufschlag_value1, aufschlag_value2, file_path, checkbox_kaufpreis, checkbox_steuer)

    def calculate(
                self,
                preis,
                lieferung_value, 
                aufschlag_value1,
                aufschlag_value2, 
                checkbox_steuer
                ):

            STEUER = 1.19
            if preis < 500:
                preis = preis + aufschlag_value1 + lieferung_value
            else:
                preis = preis + (preis * aufschlag_value2 / 100) + lieferung_value

            if checkbox_steuer:
                preis *= STEUER

            return round(preis, 2)

    def methodSortEins(self, lieferung_value, aufschlag_value1, aufschlag_value2, file_path, checkbox_kaufpreis, checkbox_steuer, export_to_csv=True):
        df = pd.read_excel(file_path)
        config = toml.load(self.config_path)
        bezeichnung_filter = config['dataSet']['data']

        # Сохраняем значение первой строки из колонки "Artikelnummer" во временную переменную


        df = df[["Artikelnummer", "Bezeichnung", "Preis", "Zustand"]].copy()
        df['Artikelnummer_raw'] = df['Artikelnummer'].astype(str)
        df['Artikelnummer'] = df['Artikelnummer'].astype(str) + df['Zustand'].astype(str).str.upper()
        df['First_Word'] = df['Bezeichnung'].str.split().str[0]
        df = df[df['First_Word'].isin(bezeichnung_filter)].drop(columns=['First_Word'])
        


        df['bilder'] = 'https://telefoneria.com/wp-content/uploads/products/bilder/' + df['Artikelnummer_raw'] + '.webp'
        
        def beschreibung_abrufen(artikelnummer_raw):
            url = f"https://telefoneria.com/wp-content/uploads/products/beschreibung/{artikelnummer_raw}.html"
            try:
                resp = requests.get(url, timeout=10)  # Можно поменять timeout обратно на 10
                resp.encoding = 'utf-8'
                if resp.status_code == 200:
                    print(f"\033[92mBeschreibung geladen für {artikelnummer_raw}\033[0m")
                    return resp.text.replace('\n', '').replace('\r', '')
                else:
                    print(f"\033[91mFehler beim Laden der Datei {url}: {resp.status_code}\033[0m")
            except Exception as e:
                print(f"\033[91mFehler beim Zugriff auf {url}: {e}\033[0m")
            return " "  # Bei Fehler oder fehlender Datei immer ein Leerzeichen

        def kurze_beschreibung_abrufen(artikelnummer_raw):
            url = f"https://telefoneria.com/wp-content/uploads/products/kurzebeschreibung/{artikelnummer_raw}.html"
            try:
                resp = requests.get(url, timeout=10)  # Можно поменять timeout обратно на 10
                resp.encoding = 'utf-8'
                if resp.status_code == 200:
                    print(f"\033[92mBeschreibung geladen für {artikelnummer_raw}\033[0m")
                    return resp.text.replace('\n', '').replace('\r', '')
                else:
                    print(f"\033[91mFehler beim Laden der Datei {url}: {resp.status_code}\033[0m")
            except Exception as e:
                print(f"\033[91mFehler beim Zugriff auf {url}: {e}\033[0m")
            return " "  # Bei Fehler oder fehlender Datei immer ein Leerzeichen
        
        def fetch_all_beschreibungen_parallel(funktion, artikelnummer_list, max_threads=60):
            with ThreadPoolExecutor(max_workers=max_threads) as executor:
                return list(executor.map(funktion, artikelnummer_list))
            
        df['Beschreibung'] = fetch_all_beschreibungen_parallel(beschreibung_abrufen, df['Artikelnummer_raw'].tolist())
        df['Kurzbeschreibung'] = fetch_all_beschreibungen_parallel(kurze_beschreibung_abrufen, df['Artikelnummer_raw'].tolist())


        df['Preis'] = pd.to_numeric(df['Preis'].astype(str).str.replace('€', '').str.replace(' ', ''), errors='coerce')

        if checkbox_kaufpreis:
            df['Kaufpreis'] = df['Preis']

        def apply_calc(price):
            if pd.notna(price):
                return self.calculate(price, lieferung_value, aufschlag_value1, aufschlag_value2, checkbox_steuer)
            return price

        df['Preis'] = df['Preis'].apply(apply_calc)
        df['Preis'] = df['Preis'].apply(lambda x: f"{x:.2f} €" if pd.notna(x) else x)

        def get_kategorie_from_bezeichnung(bezeichnung: str) -> str:
            name = bezeichnung.lower()

            if ("ipad" in name or "tab" in name) and "yealink" not in name:
                if "apple" in name:
                    return "Tablet > Apple"
                elif "samsung" in name:
                    return "Tablet > Samsung"
                return "Tablet"


            elif "iphone" in name or ("galaxy" in name and "watch" not in name and "tab" not in name):
                if "apple" in name:
                    return "Handy > Apple"
                elif "samsung" in name:
                    return "Handy > Samsung"
                return "Handy"

            elif "macbook" in name or "chromebook" in name or "notebook" in name or "book" in name:
                if "apple" in name:
                    return "Notebook > Apple"
                elif "samsung" in name:
                    return "Notebook > Samsung"
                return "Notebook"

            elif "watch" in name:
                if "apple" in name:
                    return "Smartwatch > Apple"
                elif "samsung" in name:
                    return "Smartwatch > Samsung"
                return "Smartwatch"

            elif "airpods" in name or "buds" in name or "headset" in name:
                if "apple" in name:
                    return "Kopfhörer > Apple"
                elif "samsung" in name:
                    return "Kopfhörer > Samsung"
                elif "yealink" in name:
                    return "Kopfhörer > Yealink"
                return "Kopfhörer"

            return "Sonstiges"

        df['Kategorien'] = df['Bezeichnung'].apply(get_kategorie_from_bezeichnung)
        df = df.rename(columns={'Bezeichnung': 'Name'})

        # Удаляем временный столбец 
        df.drop(columns=['Artikelnummer_raw'], inplace=True)

        try:
            if export_to_csv:
                # Сохраняем в CSV с запятыми
                csv_path = self.file_path_save.replace('.xlsx', '.csv')
                df.to_csv(csv_path, index=False, sep=',')
                print(f"Файл CSV сохранён: {csv_path}")
            else:
                # Сохраняем в Excel
                with pd.ExcelWriter(self.file_path_save, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                self.autofit_columns(self.file_path_save)
                print(f"Файл Excel сохранён: {self.file_path_save}")

        except Exception as e:
            print(f"Ошибка при сохранении: {e}")
            backup_path = os.path.join(os.path.expanduser("~"), "Desktop",
                                    f"filtered_example_backup.{ 'csv' if export_to_csv else 'xlsx' }")
            if export_to_csv:
                df.to_csv(backup_path, index=False, sep=',')
            else:
                df.to_excel(backup_path, index=False)
                self.autofit_columns(backup_path)
            print(f"Резервный файл сохранён: {backup_path}")



    def methodSortZwei(self, aufschlag_value1, file_path, checkbox_kaufpreis, checkbox_steuer):
        # Читаем файл
        df = pd.read_excel(file_path)
        
        # Создаем новый DataFrame с реальными заголовками из второй строки
        headers = df.iloc[0]
        filtered_df = df.iloc[1:].copy()
        filtered_df.columns = headers
        filtered_df.reset_index(drop=True, inplace=True)
        
        # Если checkbox_kaufpreis равно True, добавляем столбец 'Kaufpreis'
        if checkbox_kaufpreis:
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
                    formatted_price = self.calculate(
                        price,
                        0,
                        aufschlag_value1,
                        0,
                        checkbox_steuer
                    )
                    
                    filtered_df.at[index, 'Formatted_Preis'] = f"{formatted_price:.2f} €"


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
