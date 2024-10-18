import base64
import io

import pandas as pd
import openpyxl
from openpyxl.cell.cell import Cell
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from UploadPhoto import GetImage


class Processing:
    def __init__(self):
        self.file_name = "TestFile.xlsx"


    def process(self):
        start_cr_df = CreateStoreDF()  # Класс чтения файла и разбивки на DF магазина
        pass


class CreateStoreDF:
    def __init__(self):
        self.path = None
        self.df_names = ['Магазин', 'Артикул', 'Наименование', 'Количество', 'Категория', 'Сезон', 'SKU', 'RRP',
                         'RRP Amount', 'Grade', 'Пол', 'Коллекция', 'Сумма со скидкой', 'Кол-во продаж']
        self.is_df_correct = True

    def read_file(self, path):
        """
        Creation DF to each store for all category

        :param path: string (File path)
        :return:  dict (Dict format {Store: Store DF...}
        """
        self.path = path
        df = pd.read_excel(self.path)  # Читаем загрузочный файл
        names = df.columns.to_list()  # Получаем заголовки таблицы и далее сравниваем их с необходимыми заголовками
        for header in self.df_names:
            if header not in names:
                self.is_df_correct = False
        if self.is_df_correct:
            self.process_df(df)  # Запускаем обработку файла

    def process_df(self, df):
        pass


if __name__ == "__main__":
    start = Processing()
    start.process()