from datetime import datetime

from PyQt5.QtWidgets import QLabel, QVBoxLayout,\
QHBoxLayout, QWidget ,QTableWidget, QTableWidgetItem
from PyQt5 import QtWidgets

from interface.styles import QLABEL_INFO_EXCEL_STYLE, QLABEL_BRIED_INFO_TABLE


class ExcelInfoWidget(QWidget):
    info : dict = None
    KEYS = [
        'brief',
        'headers',
        'cols',
    ]

    def __init__(self, info: dict):
        super(ExcelInfoWidget, self).__init__()
        
        self.init_status = True

        for k in self.KEYS:
            if k not in info:
                self.init_status = False

        if self.init_status:
            self.info = info
            self.title = QLabel("Информация о таблице")
            self.layout : QVBoxLayout = QVBoxLayout()
            self.excel_cols_info = QTableWidget()
            self.brief_info : QVBoxLayout = QVBoxLayout()
            self.init_ui()


    def init_titel(self):
        """Настроить заголовок для блока c информацией о таблице"""
        self.title.setStyleSheet(QLABEL_INFO_EXCEL_STYLE)

    
    def get_label_for_brief(self, lable : str) -> QLabel:
        """Получить QLabel для краткой информации о таблице"""
        q_label = QLabel(lable)
        q_label.setStyleSheet(QLABEL_BRIED_INFO_TABLE)
        return q_label


    def init_brief_info_table(self):
        """Настроить краткую информацию о таблице (колонки, строки, список имен)"""

        info_filename = QHBoxLayout()
        info_filename.addWidget(self.get_label_for_brief('Наименование файла'))
        info_filename.addWidget(self.get_label_for_brief(str(self.info['brief']['filename'])))
        self.brief_info.addLayout(info_filename)

        info_sheetname = QHBoxLayout()
        info_sheetname.addWidget(self.get_label_for_brief('Наименование листа'))
        info_sheetname.addWidget(self.get_label_for_brief(str(self.info['brief']['sheetname'])))
        self.brief_info.addLayout(info_sheetname)

        info_row = QHBoxLayout()
        info_row.addWidget(self.get_label_for_brief('Количество строк'))
        info_row.addWidget(self.get_label_for_brief(str(self.info['brief']['rows_num'])))
        self.brief_info.addLayout(info_row)

        info_col = QHBoxLayout()
        info_col.addWidget(self.get_label_for_brief('Количество столбцов'))
        info_col.addWidget(self.get_label_for_brief(str(self.info['brief']['cols_num'])))
        self.brief_info.addLayout(info_col)


    
    def init_cols_info_table(self):
        """Настроить таблицу с информацией о каждой колонке"""
        self.excel_cols_info.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        self.excel_cols_info.setColumnCount(6)
        self.excel_cols_info.setRowCount(len(self.info['cols']) + 1)

        #Иницилизируем heade таблицы
        self.excel_cols_info.setItem(0,0,QTableWidgetItem('Столбец'))
        self.excel_cols_info.setItem(0,1,QTableWidgetItem('Тип'))
        self.excel_cols_info.setItem(0,2,QTableWidgetItem('Пустых'))
        self.excel_cols_info.setItem(0,3,QTableWidgetItem('Min'))
        self.excel_cols_info.setItem(0,4,QTableWidgetItem('Max'))
        self.excel_cols_info.setItem(0,5,QTableWidgetItem('Mean'))


        for index_row, ic in enumerate(self.info['cols']):
            for index_col, k in enumerate(ic.keys()):
                self.excel_cols_info.setItem(
                    index_row + 1,
                    index_col,
                    QTableWidgetItem(ic[k])
                )

        self.excel_cols_info.resizeColumnsToContents()


    def init_ui(self):
        self.layout.setSpacing(10)

        self.init_titel()
        self.init_cols_info_table()
        self.init_brief_info_table()

        self.layout.addWidget(self.title)
        
        self.layout.addStretch(0)

        self.layout.addLayout(self.brief_info)

        self.layout.addStretch(0)

        self.layout.addWidget(self.excel_cols_info)

        self.layout.addStretch(0)

        self.setLayout(self.layout)