from PyQt5.QtWidgets import QPushButton, QWidget, QLabel

from .dialogs import show_file_dialog


def init_open_files_btn(widget : QWidget, label : QLabel) -> QPushButton:
    file_btn = QPushButton('Выбрать файл', widget)
    file_btn.clicked.connect(lambda : show_file_dialog(widget, label))