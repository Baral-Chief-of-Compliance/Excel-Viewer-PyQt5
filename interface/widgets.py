from PyQt5.QtWidgets import QLabel, QPushButton, QVBoxLayout,\
QHBoxLayout, QWidget, QFileDialog, QMessageBox, QTableWidget,\
QTableWidgetItem
from PyQt5.QtCore import pyqtSignal
from PyQt5 import QtWidgets
from openpyxl import Workbook

from utils.excel_converters import ExcelConverter


class ExcelLoadDialog(QWidget):

    excel_loaded = pyqtSignal(object)
    excel_info_loaded = pyqtSignal(dict)
    




    def __init__(self):
        super(ExcelLoadDialog, self).__init__()
        self.select_excel_button : QPushButton = None
        self.excel_label : QLabel = None
        self.clear_button : QPushButton = None 

        self.excel_info_layout : QHBoxLayout = QHBoxLayout()
        self.excel_file = None
        self.workbook = None
        
        self.init_ui()



    def init_select_excel_btn(self):
        """Настроить кнопки для выбора excel файла"""
        self.select_excel_button = QPushButton("Выбрать Excel файл")
        self.select_excel_button.clicked.connect(self.select_excel_file)
        self.select_excel_button.setMinimumHeight(40)

    def init_excel_label(self):
        """Настроить отображение информации о excel файле"""
        self.excel_label = QLabel("Файл не выбран")
        self.excel_label.setWordWrap(True)

    def set_excel_label(self, filepath : str):
        """Сброс текст на кнопке с excel файлом"""
        self.excel_label.setText(filepath)


    def init_clear_button(self):
        """Настроить кнопку открепления файла"""
        self.clear_button = QPushButton("x")
        self.clear_button.setFixedSize(25, 25)
        self.clear_button.setToolTip("Очистить")
        self.clear_button.clicked.connect(self.clear_excel_file)
        self.clear_button.setEnabled(False)

    def enable_claer_button(self):
        """Сделать активным кнопку открепления"""
        self.clear_button.setEnabled(True)

    def disable_claer_button(self):
        """Сдеалать неактивным кнопку открепления"""
        self.clear_button.setEnabled(False)

        

    def init_ui(self):
        """Инициализируем UI"""
        layout = QVBoxLayout()
        layout.setSpacing(10)

        # Кнопка для загрузки
        self.init_select_excel_btn()

        layout.addWidget(self.select_excel_button)

        #Qtlabel для отображения информации о файле
        self.init_excel_label()

        #кнопка для открепления файла
        self.init_clear_button()

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
                ec = ExcelConverter(filepath=file_path)
                
                self.workbook = ec.get_work_book()
                self.excel_info = ec.data_analysis()

                self.set_excel_label(file_path)
                self.enable_claer_button()


                self.excel_loaded.emit(self.workbook)
                self.excel_info_loaded.emit(self.excel_info)


        except Exception as ex:
            print(ex)
            QMessageBox.critical(self, "Ошибка", "Не удалось открыть файл: {}".format(ex))


    def clear_excel_file(self):
        """Открепить файл"""
        self.set_excel_label("Файл не выбран")
        self.disable_claer_button()
        self.workbook = None

        self.excel_loaded.emit(None)


class ExcelTableWidget(QTableWidget):

    wb : Workbook = None


    def __init__(self, wb : Workbook):
        super(ExcelTableWidget, self).__init__()
        self.wb = wb
        self.init_table()
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


    def init_table(self):
        """Инициализация таблицы"""
        if self.wb is not None:
            cols = self.wb.worksheets[0].iter_cols(
                min_col=1,
                max_col=self.wb.worksheets[0].max_column
            )

            rows = self.wb.worksheets[0].iter_rows(
                min_row=1,
                max_row=self.wb.worksheets[0].max_row

            )

            self.setColumnCount(self.wb.worksheets[0].max_column)

            self.setRowCount(self.wb.worksheets[0].max_row)


            for row in rows:
                for cell in row:
                    item = QTableWidgetItem(cell.value)
                    self.setItem(
                        cell.row,
                        cell.column,
                        item
                        
                    )

            self.resizeColumnsToContents()


class ExcelInfoWidget(QWidget):
    info : dict = None


    def __init__(self, info: dict):
        super(ExcelInfoWidget, self).__init__()
        self.title = QLabel("Информация о таблице")
        self.layout : QHBoxLayout = QHBoxLayout()
        self.info = info
        self.init_ui()


    def init_titel(self):
        """Настроить заголовок для блока c информацией о таблице"""
        self.title.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #2c3e50;
                margin: 10px 0;
                padding: 5px;
                border-bottom: 2px solid #3498db;
            }
        """)

    def init_ui(self):
        self.layout.setSpacing(10)

        self.init_titel()

        self.layout.addWidget(self.title)

        self.setLayout(self.layout)