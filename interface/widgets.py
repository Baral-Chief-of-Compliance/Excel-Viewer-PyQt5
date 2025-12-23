from PyQt5.QtWidgets import QLabel, QPushButton, QVBoxLayout,\
QHBoxLayout, QWidget, QFileDialog, QMessageBox


class ExcelLoadDialog(QWidget):
    def __init__(self):
        super(ExcelLoadDialog, self).__init__()
        self.initUi()
        self.excel_file = None

    def initUi(self):
        """Инициализируем UI"""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # Кнопка для загрузки
        self.select_excel_button = QPushButton("Выбрать Excel файл")
        self.select_excel_button.clicked.connect(self.select_excel_file)
        self.select_excel_button.setMinimumHeight(40)

        layout.addWidget(self.select_excel_button)

        #Qtlabel для отображения информации о файле
        self.excel_info_layout = QHBoxLayout()

        self.excel_label = QLabel("Файл не выбран")
        self.excel_label.setWordWrap(True)

        #кнопка для открепления файла
        self.clear_button = QPushButton("x")
        self.clear_button.setFixedSize(25, 25)
        self.clear_button.setToolTip("Очистить")
        self.clear_button.clicked.connect(self.clear_excel_file)
        self.clear_button.setEnabled(False)

        self.excel_info_layout.addWidget(self.excel_label, 1)
        self.excel_info_layout.addWidget(self.clear_button)

        layout.addLayout(self.excel_info_layout)

        self.setLayout(layout)


    def select_excel_file(self):
        """Выбрать excel файл """
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите Excel файл",
                "",
                "Excel Files (*.xlsx *.xls);;All Files (*)"
            )

            if file_path:
                if not (file_path.endswith('.xlsx') or file_path.endswith('.xls')):
                    QMessageBox.warning(self, "Предупреждение", "Выберите файл с расширением .xlsx или .xls")
                    return
                
                #логика обработки файла
        except Exception as ex:
            QMessageBox.critical(self, "Ошибка", "Не удалось открыть файл: {}".format(ex))                



    def clear_excel_file(self):
        """Открепить файл"""
        pass