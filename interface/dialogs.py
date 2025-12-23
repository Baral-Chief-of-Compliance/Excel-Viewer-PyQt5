from PyQt5.QtWidgets import QFileDialog, QLabel, QWidget


def show_file_dialog(widget : QWidget, label : QLabel):
    """Показать диалоговое окно с выбором файла"""
    file_name = QFileDialog.getOpenFileName(
        widget,
        'Выберите Excel файл',
        '',
        'Excel Files (*.xls *.xlsx);;All Files (*)'
    )

    if file_name:
            label.setText(f'Выбран файл: {file_name}')
            # Здесь можно добавить обработку выбранного файла
            print(f'Выбран файл: {file_name}')