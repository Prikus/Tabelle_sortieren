import PyInstaller.__main__
import os

assets_path = os.path.abspath("assets")
gui_path = os.path.abspath("gui")
scripts_path = os.path.abspath("scripts")

PyInstaller.__main__.run([
    'main.py',  # Главный файл приложения
    '--name=Tabelle_Sortieren',  # Имя выходного файла
    '--onefile',  # Собрать в один файл
    '--windowed',  # Без консольного окна
    f'--add-data={assets_path}{os.pathsep}assets',  # Добавляем папку assets
    f'--add-data={gui_path}{os.pathsep}gui',  # Добавляем папку gui
    f'--add-data={scripts_path}{os.pathsep}scripts',  # Добавляем папку scripts
    '--icon=assets/ico.ico',  # Используем вашу иконку
    '--clean',  # Очищаем кэш перед сборкой
])