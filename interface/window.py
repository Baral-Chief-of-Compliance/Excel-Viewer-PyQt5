from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget

from .widgets import ExcelLoadDialog, ExcelTableWidget


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

                self.initUi()

        def initUi(self):
                """Инициализируем UI"""
                centeral_widget = QWidget()
                self.setCentralWidget(centeral_widget)

                layout = QVBoxLayout()

                centeral_widget.setLayout(layout)

                self.excel_load_dialog = ExcelLoadDialog()

                self.excel_load_dialog.excel_loaded.connect(self.handle_excel_loaded)

                layout.addWidget(self.excel_load_dialog)

                layout.addStretch(1)


        def handle_excel_loaded(self, workbook):
                """Обработка загрузки excel фаайла"""
                if self.table_widget is not None:
                        self.table_widget.deleteLater()
                        self.table_widget = None

                if workbook is not None:
                        self.table_widget = ExcelTableWidget(wb=workbook)

                        # Получаем layout из central_widget
                        central_layout = self.centralWidget().layout()
            
                        # Удаляем stretch перед добавлением таблицы
                        item = central_layout.takeAt(central_layout.count() - 1)
            
                        # Добавляем таблицу
                        central_layout.addWidget(self.table_widget)
            
                        # Добавляем stretch обратно
                        central_layout.addStretch(0)
            
                         # Устанавливаем размер таблицы (80% от доступного пространства)
                        self.table_widget.setMinimumHeight(int(self.height() * 0.7))
        
