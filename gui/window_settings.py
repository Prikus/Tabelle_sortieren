from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QCheckBox, 
    QLabel, QLineEdit, QFileDialog, QDialog
)
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QPainter, QColor, QBrush, QPen, QIcon

class SettingsWindow(QDialog):
    def __init__(self):
        super().__init__()

        self.init_window()
        self.setup_ui()


    def init_window(self):
        """Настройки окна."""
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setGeometry(80, 100, 200, 200)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background-color: transparent;")
        self.setWindowIcon(QIcon("assets/ico.ico"))    

    def setup_ui(self):
        central_widget = QWidget(self)
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(10, 10, 10, 10)
        central_layout.setSpacing(10)
    
         # Верхняя панель
        top_panel = self.create_top_panel()
        central_layout.addWidget(top_panel)

        self.setLayout(central_layout)
    
    def create_top_panel(self):
        top_panel = QWidget(self)
        top_layout = QHBoxLayout(top_panel)
        top_layout.setContentsMargins(10, 0, 0, 0)

        # Заголовок
        title_label = QLabel("Einstellungen")
        title_label.setStyleSheet("""
            QLabel {
                color: #FF9840;
                font-size: 14px;
                font-weight: bold;
            }
        """)
        top_layout.addWidget(title_label)

        # Растягивающийся элемент
        top_layout.addStretch()

        # Кнопки
        buttons_widget = self.create_buttons()
        top_layout.addWidget(buttons_widget)

        return top_panel
    
    def create_buttons(self):
        """Создает кнопки закрытия и сворачивания."""
        buttons_widget = QWidget()
        buttons_layout = QHBoxLayout(buttons_widget)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(5)

        # Кнопка закрытия
        close_button = QPushButton(self)
        close_button.setFixedSize(30, 30)
        close_button.setIcon(QIcon("assets/close_icon.svg"))
        close_button.setIconSize(close_button.size())
        close_button.clicked.connect(self.close)
        close_button.setStyleSheet(self.button_style())
        buttons_layout.addWidget(close_button)


        return buttons_widget
    
    def input_style(self):
        """Возвращает стиль для полей ввода."""
        return """
            QLineEdit {
                padding: 5px;
                border: 1px solid #FF9840;
                border-radius: 5px;
                background-color: transparent;
                color: #FF9840;
                font-size: 14px;
            }
            QLineEdit:focus {
                border-color: #FF6600;
            }
        """

    def button_style(self):
        """Возвращает стиль для кнопок."""
        return """
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
                border-radius: 15px;
            }
        """

    def checkbox_style(self):
        """Возвращает стиль для чекбоксов."""
        return """
            QCheckBox {
                spacing: 8px;
                color: #FF9840;
                font-size: 14px;
                font-weight: bold;
            }
            
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 1.5px solid #FF9840;
                border-radius: 4px;
                background-color: transparent;
            }
            
            QCheckBox::indicator:unchecked:hover {
                border-color: #FF6600;
            }
            
            QCheckBox::indicator:checked {
                background-color: #FF9840;
                border-color: #FF9840;
                image: url(assets/check_white.svg);  /* Путь к белой галочке */
            }
            
            QCheckBox::indicator:checked:hover {
                background-color: #FF6600;
                border-color: #FF6600;
            }
        """
    
    def paintEvent(self, event):
        """Отрисовка окна с закругленными углами и оранжевой пунктирной рамкой."""
        # Основной закрашенный прямоугольник с закругленными углами
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QBrush(QColor("#303030")))
        painter.setPen(QPen(Qt.NoPen))

        rounded_rect = self.rect().adjusted(0, 0, -1, -1)
        painter.drawRoundedRect(rounded_rect, 10, 10)

        # Добавляем пунктирную оранжевую рамку
        border_painter = QPainter(self)
        border_painter.setRenderHint(QPainter.Antialiasing)
        dash_pen = QPen(QColor("#FF9840"))
        dash_pen.setStyle(Qt.DashLine)
        dash_pen.setWidth(2)
        border_painter.setPen(dash_pen)

        border_rect = self.rect().adjusted(1, 1, -2, -2)
        border_painter.drawRoundedRect(border_rect, 10, 10)

        painter.end()
        border_painter.end()
