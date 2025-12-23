import sys

from PyQt5.QtWidgets import QApplication


def make_app() -> QApplication:
    """Создает QApplication"""
    app = QApplication(sys.argv)
    return app