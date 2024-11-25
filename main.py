import sys
import os
from PyQt5.QtWidgets import QApplication
from gui.window import MainWindow

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())    

def resource_path(relative_path):
    """Получение абсолютного пути к ресурсу"""
    try:
        # PyInstaller создает временную папку и хранит путь в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)    

if __name__ == "__main__":
    main()
