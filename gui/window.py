from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QHBoxLayout, QCheckBox, 
    QLabel, QLineEdit, QFileDialog, QProgressBar
)
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QPainter, QColor, QBrush, QPen, QIcon
from scripts.sorting import ExcelSorter

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_window()
        self.setup_ui()
        self.setup_drag()

        self.lieferung_input
        self.aufschlag_input1
        self.aufschlag_input2
        self.path_input
        self.checkbox_kaufpreis = False
        self.checkbox_steuer = True

    def init_window(self):
        """Настройки окна."""
        self.setWindowTitle("Tabelle_sortieren")
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setGeometry(80, 100, 400, 400)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background-color: transparent;")
        self.setWindowIcon(QIcon("assets/ico.ico"))

    def setup_ui(self):
        """Создание пользовательского интерфейса."""        
        central_widget = QWidget(self)
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(10, 10, 10, 10)
        central_layout.setSpacing(10)

        # Верхняя панель
        top_panel = self.create_top_panel()
        central_layout.addWidget(top_panel)

        # Центральный контейнер с левой панелью и правой областью
        center_layout = QHBoxLayout()
        left_panel = self.create_left_panel()
        right_drag_area = self.create_right_drag_area()

        # Разделим пространство между левой панелью и правой областью
        center_layout.addWidget(left_panel)
        center_layout.addWidget(right_drag_area)

        central_layout.addLayout(center_layout)

        # Растягивающийся элемент для закрепления нижней панели внизу
        central_layout.addStretch()

        # Добавляем прогресс-бар перед нижней панелью
        progress_bar = self.create_progress_bar()
        self.progress_bar = progress_bar
        central_layout.addWidget(progress_bar)      

        # Нижняя панель
        bottom_panel = self.create_bottom_panel()
        central_layout.addWidget(bottom_panel)

        # Проверка инициализации полей для отладки
        print("lieferung_input:", self.lieferung_input)
        print("aufschlag_input1:", self.aufschlag_input1)
        print("aufschlag_input2:", self.aufschlag_input2)
        print("path_input:", self.path_input)

        self.setCentralWidget(central_widget)

    def create_top_panel(self):
        """Создает верхнюю панель с заголовком и кнопками."""
        top_panel = QWidget(self)
        top_layout = QHBoxLayout(top_panel)
        top_layout.setContentsMargins(10, 0, 0, 0)

        # Заголовок
        title_label = QLabel("Tabelle_sortieren")
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

        #Кнопка настроек todo
        settings_button = QPushButton(self)
        settings_button.setFixedSize(30, 30)
        settings_button.setIcon(QIcon("assets/settings_icon.svg"))
        settings_button.setIconSize(settings_button.size())
        settings_button.clicked.connect(self.settings)
        settings_button.setStyleSheet(self.button_style())
        buttons_layout.addWidget(settings_button)

        # Кнопка сворачивания
        hide_button = QPushButton(self)
        hide_button.setFixedSize(30, 30)
        hide_button.setIcon(QIcon("assets/hide_icon.svg"))
        hide_button.setIconSize(hide_button.size())
        hide_button.clicked.connect(self.showMinimized)
        hide_button.setStyleSheet(self.button_style())
        buttons_layout.addWidget(hide_button)

        # Кнопка закрытия
        close_button = QPushButton(self)
        close_button.setFixedSize(30, 30)
        close_button.setIcon(QIcon("assets/close_icon.svg"))
        close_button.setIconSize(close_button.size())
        close_button.clicked.connect(self.close)
        close_button.setStyleSheet(self.button_style())
        buttons_layout.addWidget(close_button)


        return buttons_widget


    def create_progress_bar(self):
        """Создает и возвращает стилизованный прогресс-бар."""
        progress_bar = QProgressBar(self)
        progress_bar.setMinimum(0)
        progress_bar.setMaximum(100)
        progress_bar.setValue(0)
        progress_bar.setFixedHeight(25)
        progress_bar.setContentsMargins(60, 0, 60, 0)
        progress_bar.setTextVisible(False)
        progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #FF9840;
                border-radius: 6px;
                background: #303030;
                color: #ffffff;
                font-size: 13px;
            }
            QProgressBar::chunk {
                background-color: #FF9840;
                border-radius: 4px;
            }
        """)
        return progress_bar

    def create_left_panel(self):
        """Создает левую панель с текстами и полями ввода."""
        left_panel = QWidget(self)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(10, 10, 10, 10)

        # Текст "Lieferung "
        lieferung_label = QLabel("Lieferung")
        lieferung_label.setStyleSheet("""
            QLabel {
                color: #FF9840;
                font-size: 14px;
                font-weight: bold;
            }
        """)
        left_layout.addWidget(lieferung_label)

        # Поле для ввода Lieferung
        lieferung_input = QLineEdit(self)
        lieferung_input.setPlaceholderText(".....€")
        lieferung_input.setAlignment(Qt.AlignLeft)
        lieferung_input.setFixedWidth(150)
        lieferung_input.setStyleSheet(self.input_style())
        left_layout.addWidget(lieferung_input, alignment=Qt.AlignLeft)
        self.lieferung_input = lieferung_input


        # Текст "Aufschlag"
        aufschlag_label = QLabel("Aufschlag ")
        aufschlag_label.setStyleSheet("""
            QLabel {
                color: #FF9840;
                font-size: 14px;
                font-weight: bold;
            }
        """)
        left_layout.addWidget(aufschlag_label)

        # Поля для ввода Aufschlag <500 и >500
        aufschlag_input1 = QLineEdit(self)
        aufschlag_input1.setPlaceholderText(".....€")
        aufschlag_input1.setAlignment(Qt.AlignLeft)
        aufschlag_input1.setFixedWidth(150)
        aufschlag_input1.setStyleSheet(self.input_style())
        left_layout.addWidget(aufschlag_input1, alignment=Qt.AlignLeft)
        self.aufschlag_input1 = aufschlag_input1    


        aufschlag_input2 = QLineEdit(self)
        aufschlag_input2.setPlaceholderText(".....%")
        aufschlag_input2.setAlignment(Qt.AlignLeft)
        aufschlag_input2.setFixedWidth(150)
        aufschlag_input2.setStyleSheet(self.input_style())
        left_layout.addWidget(aufschlag_input2, alignment=Qt.AlignLeft)
        self.aufschlag_input2 = aufschlag_input2


        Startpreis_checkbox = QCheckBox("Kaufpreis")
        Startpreis_checkbox.setStyleSheet(self.checkbox_style())
        left_layout.addWidget(Startpreis_checkbox)

        Steuer_checkbox = QCheckBox("Steuer")
        Steuer_checkbox.setStyleSheet(self.checkbox_style())
        Steuer_checkbox.setChecked(True)  # Устанавливаем галочку
        left_layout.addWidget(Steuer_checkbox)

        # Подключение сигнала
        Startpreis_checkbox.stateChanged.connect(self.on_startpreis_changed)    
        Steuer_checkbox.stateChanged.connect(self.on_steuer_changed)
        return left_panel

    def create_bottom_panel(self):
        """Создает нижнюю панель с кнопкой и выбором пути."""
        bottom_panel = QWidget(self)
        bottom_layout = QHBoxLayout(bottom_panel)
        bottom_layout.setContentsMargins(10, 10, 10, 10)
        bottom_layout.setSpacing(10)

        # Поле выбора пути
        path_input = QLineEdit(self)
        path_input.setPlaceholderText("Speicherpfad auswählen...")
        path_input.setAlignment(Qt.AlignLeft)
        path_input.setFixedWidth(300)
        path_input.setStyleSheet(self.input_style())
        self.path_input = path_input
        bottom_layout.addWidget(path_input, alignment=Qt.AlignLeft)

        # Кнопка выбора пути
        browse_button = QPushButton()
        browse_button.setIcon(QIcon("assets/path_icon.svg"))
        browse_button.setFixedSize(30, 30)
        browse_button.setStyleSheet(self.button_style())
        browse_button.clicked.connect(lambda: self.select_path(path_input))
        bottom_layout.addWidget(browse_button)

        # Растягивающийся элемент, чтобы кнопка Starten была прижата к правому краю
        bottom_layout.addStretch()

        # Кнопка Starten
        starten_button = QPushButton("Starten")
        starten_button.setFixedSize(100, 30)
        starten_button.setStyleSheet("""
            QPushButton {
                background-color: #FF9840;
                color: white;
                border-radius: 5px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FF6600;
            }
        """)
        starten_button.clicked.connect(self.on_starten_click)
        bottom_layout.addWidget(starten_button, alignment=Qt.AlignRight)

        return bottom_panel

    def create_right_drag_area(self):
        """Создает область для перетаскивания файлов справа с пунктирной обводкой и индикатором в центре."""
        drag_area = QWidget(self)
        drag_area.setStyleSheet("""
            QWidget {
                border: 2px dashed #FF9840;
                border-radius: 10px;
                background-color: rgba(0, 0, 0, 0.1);
                min-height: 100px;
                min-width: 100px;
            }
        """)
        drag_area.setAcceptDrops(True)
        drag_area.dragEnterEvent = self.dragEnterEvent
        drag_area.dropEvent = self.dropEvent

        # Создаем layout для центрирования
        layout = QVBoxLayout(drag_area)
        layout.setContentsMargins(0, 0, 0, 0)  # Убираем отступы layout
        
        # Индикатор состояния
        self.status_icon = QLabel()
        self.status_icon.setAlignment(Qt.AlignCenter)
        self.status_icon.setPixmap(QIcon("assets/empty_icon.svg").pixmap(48, 48))
        self.status_icon.setStyleSheet("""
            background-color: transparent;
            border: none;  /* Убираем обводку у иконки */
        """)
        
        # Добавляем индикатор в layout
        layout.addWidget(self.status_icon, alignment=Qt.AlignCenter)
        
        return drag_area

    def on_starten_click(self):
        """Обработка нажатия кнопки Starten."""
        try:
            # Проверяем, что поля существуют и значения читаются
            lieferung_value = self.lieferung_input.text()
            aufschlag_value1 = self.aufschlag_input1.text()
            aufschlag_value2 = self.aufschlag_input2.text()
            path_save = self.path_input.text()
            path_file = self.dragged_file_path
            checkbox_kaufpreis = self.checkbox_kaufpreis
            checkbox_steuer = self.checkbox_steuer
            print("Получены значения:")
            print(f"Lieferung: {lieferung_value}")
            print(f"Aufschlag <500: {aufschlag_value1}")
            print(f"Aufschlag >500: {aufschlag_value2}")
            print(f"Путь сохранение: {path_save}")
            print(f"Путь файла: {path_file}")
            print(f"Галочка1: {checkbox_kaufpreis}")
            print(f"Галочка2: {checkbox_steuer}")

            # Передаем данные в сортировщик
            sortier = ExcelSorter(path_save)
            sortier.sortExcel(
                lieferung_value, 
                aufschlag_value1, 
                aufschlag_value2, 
                path_file, 
                checkbox_kaufpreis, 
                checkbox_steuer
                )

            print("Сортировка выполнена успешно!")

        except ValueError as ve:
            print(f"Ошибка ввода: {ve}")
        except Exception as e:
            print(f"Произошла ошибка при сортировке: {e}")



    def select_path(self, path_input):
        """Открывает диалог выбора пути сохранения."""
        file_dialog = QFileDialog(self)
        file_path = file_dialog.getExistingDirectory(self, "Выберите папку")
        if file_path:
            path_input.setText(file_path)

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

    def setup_drag(self):
        """Настройка перемещения окна."""
        self.is_dragging = False
        self.offset = QPoint(0, 0)

    def paintEvent(self, event):
        """Отрисовка окна с закругленными углами."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QBrush(QColor("#303030")))
        painter.setPen(QPen(Qt.NoPen))

        rounded_rect = self.rect().adjusted(0, 0, -1, -1)
        painter.drawRoundedRect(rounded_rect, 10, 10)

        painter.end()

    def mousePressEvent(self, event):
        """Обработка нажатия мыши для перемещения окна."""
        if event.button() == Qt.LeftButton:
            self.is_dragging = True
            self.offset = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        """Обработка перемещения мыши для перемещения окна."""
        if self.is_dragging and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.offset)
            event.accept()

    def mouseReleaseEvent(self, event):
        """Обработка отпускания мыши."""
        self.is_dragging = False

    def dragEnterEvent(self, event):
        """Обработка перетаскивания файла."""
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def on_startpreis_changed(self, state):
        """Обработчик изменения состояния чекбокса."""
        self.checkbox_kaufpreis = state == Qt.Checked

    def on_steuer_changed(self, state):
        """Обработчик изменения состояния чекбокса."""
        self.checkbox_steuer = state == Qt.Checked

    def dropEvent(self, event):
        """Обработка отпускания файла."""
        if event.mimeData().hasUrls():
            file_path = event.mimeData().urls()[0].toLocalFile()
            # Проверяем расширение файла
            if not file_path.endswith(('.xls', '.xlsx')):
                print(f"Неподдерживаемый файл: {file_path}")
                self.status_icon.setPixmap(QIcon("assets/error_icon.svg").pixmap(48, 48))
                return
            
            # Файл успешно перетащен
            self.dragged_file_path = file_path
            self.status_icon.setPixmap(QIcon("assets/check_icon.svg").pixmap(48, 48))
            print(f"Файл загружен: {file_path}")

    def settings(self):
        from gui.window_settings import SettingsWindow
        settings_window = SettingsWindow()
        settings_window.exec_()
