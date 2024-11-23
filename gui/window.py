import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QLabel, QLineEdit
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QPainter, QColor, QBrush, QPen, QIcon, QFont

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Tabelle_sortieren")
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setGeometry(100, 100, 600, 400)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background-color: transparent;")

        self.setWindowIcon(QIcon("assets/ico.ico"))

        # Центральный виджет
        central_widget = QWidget(self)
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(10, 10, 10, 10)

        # Верхняя панель с заголовком и кнопками
        top_panel = QWidget(self)
        top_layout = QHBoxLayout(top_panel)
        top_layout.setContentsMargins(10, 0, 0, 0)  # Добавляем отступ слева для текста
        
        # Добавляем текст
        title_label = QLabel("Tabelle_sortieren")
        title_label.setStyleSheet("""
            QLabel {
                color: #FF9840;  /* Оранжевый цвет */
                font-size: 14px;
                font-weight: bold;
            }
        """)
        top_layout.addWidget(title_label)

        # Добавляем растягивающийся элемент между текстом и кнопками
        top_layout.addStretch()

        # Создаем виджет для группировки кнопок
        buttons_widget = QWidget()
        buttons_layout = QHBoxLayout(buttons_widget)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.setSpacing(5)

        # Кнопка сворачивания
        hide_button = QPushButton(self)
        hide_button.setFixedSize(30, 30)
        hide_icon = QIcon("assets/hide_icon.svg")
        hide_button.setIcon(hide_icon)
        hide_button.setIconSize(hide_button.size())
        hide_button.clicked.connect(self.showMinimized)
        hide_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
                border-radius: 15px;
            }
        """)

        # Кнопка закрытия
        close_button = QPushButton(self)
        close_button.setFixedSize(30, 30)
        close_icon = QIcon("assets/close_icon.svg")
        close_button.setIcon(close_icon)
        close_button.setIconSize(close_button.size())
        close_button.clicked.connect(self.close)
        close_button.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                padding: 0;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 0.1);
                border-radius: 15px;
            }
        """)

        # Добавляем кнопки в layout кнопок
        buttons_layout.addWidget(hide_button)
        buttons_layout.addWidget(close_button)

        # Добавляем виджет с кнопками в верхний layout
        top_layout.addWidget(buttons_widget)

        central_layout.addWidget(top_panel)
        central_layout.addStretch()

        self.setCentralWidget(central_widget)

          # Добавляем текст "Lieferung" и поле ввода с символом евро
        left_panel = QWidget(self)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(10, 10, 10, 10)

        # Текст "Lieferung"
        lieferung_label = QLabel("Lieferung")
        lieferung_label.setStyleSheet("""
            QLabel {
                color: #FF9840;
                font-size: 14px;
                font-weight: bold;
            }
        """)
        left_layout.addWidget(lieferung_label)

        # Поле для ввода
        lieferung_input = QLineEdit(self)
        lieferung_input.setPlaceholderText("Введите сумму...")
        lieferung_input.setStyleSheet("""
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
        """)

        # Добавляем поле ввода и евро-символ
        lieferung_input.setAlignment(Qt.AlignRight)
        lieferung_input.textChanged.connect(lambda: lieferung_input.setText(lieferung_input.text() + " €"))

        left_layout.addWidget(lieferung_input)

        central_layout.addWidget(left_panel)

        self.setCentralWidget(central_widget)

        # Перемещение окна
        self.is_dragging = False
        self.offset = QPoint(0, 0)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QBrush(QColor("#303030")))
        painter.setPen(QPen(Qt.NoPen))

        rounded_rect = self.rect().adjusted(0, 0, -1, -1)
        painter.drawRoundedRect(rounded_rect, 10, 10)

        painter.end()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.is_dragging = True
            self.offset = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.is_dragging and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.offset)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.is_dragging = False
