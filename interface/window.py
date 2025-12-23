from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QWidget

from .widgets import ExcelLoadDialog


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
                self.initUi()

        def initUi(self):
                """Инициализируем UI"""
                centeral_widget = QWidget()
                self.setCentralWidget(centeral_widget)

                layout = QVBoxLayout()

                centeral_widget.setLayout(layout)

                self.excel_load_dialog = ExcelLoadDialog()

                layout.addWidget(self.excel_load_dialog)
        
