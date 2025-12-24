import openpyxl


class ExcelConverter:
    """Класс для работы с excle файлами"""
    wb = None


    def __init__(self, filepath) -> None:
        self._open_file(filepath)


    def _open_file(self, filepath) -> None:
        """Открыть excel файл"""
        self.wb = openpyxl.load_workbook(filepath)

    
    def get_work_book(self) -> openpyxl.Workbook:
        """Получить данные в массиве"""
        return self.wb
