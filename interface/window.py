from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget

from .widgets import ExcelLoadDialog, ExcelTableWidget, ExcelInfoWidget


class MainWindow(QMainWindow):
        def __init__(
                        self,
                        title : str ='Приложение',
                        witdh : int = 300,
                        height : int = 200,
                        start_x : int = 200,
                        start_y : int = 200,
                ):
                super(MainWindow, self).__init__()
                self.resize(witdh, height)
                self.move(start_x, start_y)
                self.setWindowTitle(title)

                self.table_widget = None
                self.table_info_widget = None

                self.initUi()

        def initUi(self):
                """Инициализируем UI"""
                centeral_widget = QWidget()
                self.setCentralWidget(centeral_widget)

                layout = QVBoxLayout()

                centeral_widget.setLayout(layout)

                self.excel_load_dialog = ExcelLoadDialog()

                self.excel_load_dialog.excel_loaded.connect(self.handle_excel_loaded)
                self.excel_load_dialog.excel_info_loaded.connect(self.handle_info_about_excel)

                layout.addWidget(self.excel_load_dialog)

                layout.addStretch(1)


        def handle_excel_loaded(self, workbook):
                """Обработка загрузки excel фаайла"""
                if self.table_widget is not None:
                        self.table_widget.deleteLater()
                        self.table_widget = None

                if workbook is not None:
                        self.table_widget = ExcelTableWidget(wb=workbook)

                        central_layout = self.centralWidget().layout()
            
                        item = central_layout.takeAt(central_layout.count() - 1)
            
                        # Добавляем таблицу
                        central_layout.addWidget(self.table_widget)
            
                        central_layout.addStretch(0)

                        self.table_widget.setMinimumHeight(int(self.height() * 0.7))


        def handle_info_about_excel(self, info):
                """Обработка загрузки информации об excel таблице"""
                if self.table_info_widget is not None:
                        self.table_info_widget.deleteLater()
                        self.table_info_widget = None

                if info is not None:
                        self.table_info_widget = ExcelInfoWidget(info=info)


                        central_layout = self.centralWidget().layout()

                        central_layout.addWidget(self.table_info_widget)

                        central_layout.addStretch(0)

                        self.table_info_widget.setMinimumHeight(int(self.height() * 0.2))
