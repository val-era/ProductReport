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
        self.store_df = None
        self.category_df = None

    def process(self):
        start_cr_df = CreateStoreDF()  # Класс чтения файла и разбивки на DF магазина
        self.store_df = start_cr_df.create_df(self.file_name)

        start_bs_df = CreateBSDF()   # Создание DF по Best sellers на уровне магазинов и категорий
        self.category_df = start_bs_df.create_store_bs(self.store_df)
        self.excel_file()

    def excel_file(self):
        wb = openpyxl.Workbook()
        worksheet = wb.active
        wb.remove(worksheet)
        for store in self.category_df:
            # Создание листов для записи в файл
            wb.create_sheet(title=store)
            category_df = self.category_df[store]
            ws = wb[store]
            # Составляем общую
            ttl_df = category_df["TTL"]
            ttl_list = ttl_df.values.tolist()
            header = ["TOTAL"]

            def styled_cells(header):
                for c in header:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)
                    c.fill = PatternFill(patternType='solid',
                                         fgColor='e6e6e6')
                    yield c

            ws.append(styled_cells(header))
            headers = ['Артикул', "IMG", 'Модель ANTA', 'Сезон', 'RRP', 'Пол',
                       'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', 'Moscow Net Sales rank']

            def styled_cells(headers):
                for c in headers:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)
                    c.fill = PatternFill(patternType='solid',
                                         fgColor='e6e6e6')
                    yield c

            ws.append(styled_cells(headers))
            for i in ttl_list:
                ws.append(i)

            # Пару пробелов для читаемости
            ws.append([""])
            ws.append([""])
            # Составляем табличку на обувь
            ftw_df = category_df["FTW"]
            ftw_list = ftw_df.values.tolist()
            header = ["FTW"]
            def styled_cells(header):
                for c in header:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)

                    c.fill = PatternFill(patternType='solid',
                                         fgColor='e6e6e6')
                    yield c
            ws.append(styled_cells(header))

            headers = ['Артикул', "IMG", 'Модель ANTA', 'Сезон', 'RRP', 'Пол',
                       'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', 'Moscow Net Sales rank']

            def styled_cells(headers):
                for c in headers:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)

                    c.fill = PatternFill(patternType='solid',
                                         fgColor='e6e6e6')
                    yield c
            ws.append(styled_cells(headers))
            for i in ftw_list:
                ws.append(i)

            # Пару пробелов для читаемости
            ws.append([""])
            ws.append([""])
            # Составляем табличку на одежду
            app_df = category_df["APP"]
            app_list = app_df.values.tolist()
            header = ["APP"]

            def styled_cells(header):
                for c in header:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)
                    c.fill = PatternFill(patternType='solid',
                                             fgColor='e6e6e6')
                    yield c

            ws.append(styled_cells(header))
            headers = ['Артикул', "IMG", 'Модель ANTA', 'Сезон', 'RRP', 'Пол',
                           'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', 'Moscow Net Sales rank']

            def styled_cells(headers):
                for c in headers:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)
                    c.fill = PatternFill(patternType='solid',
                                             fgColor='e6e6e6')
                    yield c
            ws.append(styled_cells(headers))
            for i in app_list:
                ws.append(i)

            # Пару пробелов для читаемости
            ws.append([""])
            ws.append([""])

            # Составляем табличку на аксессуары
            acc_df = category_df["ACC"]
            acc_list = acc_df.values.tolist()
            header = ["ACC"]

            def styled_cells(header):
                for c in header:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)
                    c.fill = PatternFill(patternType='solid',
                                         fgColor='e6e6e6')
                    yield c

            ws.append(styled_cells(header))
            headers = ['Артикул', "IMG", 'Модель ANTA',  'Сезон', 'RRP', 'Пол',
                       'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', 'Moscow Net Sales rank']

            def styled_cells(headers):
                for c in headers:
                    c = Cell(ws, column="A", row=1, value=c)
                    c.font = Font(bold=True)
                    c.fill = PatternFill(patternType='solid',
                                         fgColor='e6e6e6')
                    yield c

            ws.append(styled_cells(headers))
            for i in acc_list:
                ws.append(i)

            for col in ws.columns:
                max_length = 0
                column = get_column_letter(col[0].column)  # Get the column name
                # Since Openpyxl 2.6, the column name is  ".column_letter" as .column became the column number (1-based)
                for cell in col:
                    try:  # Necessary to avoid error on empty cells
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                ws.column_dimensions[column].width = adjusted_width
            ws.column_dimensions['B'].width = 12

            for i in range(1, ws.max_row + 1):
                if ws.cell(i, 4).value != "":
                    rd = ws.row_dimensions[i]
                    rd.height = 50

            row_numb = 1
            for cell in ws['A']:
                if cell.value not in ["Артикул", "APP", "ACC", "FTW", "TOTAL", ""]:
                    self.photo = GetImage().get_photo(cell.value)
                    if self.photo is None:
                        self.photo = GetImage().get_photo("No_photo")
                    image_bytes = base64.b64decode(self.photo[0])
                    image_buffer = io.BytesIO(image_bytes)
                    img = openpyxl.drawing.image.Image(image_buffer)
                    img.height = 60
                    img.width = 70
                    number = ws.cell(row=row_numb, column=2)
                    img.anchor = number.coordinate
                    ws.add_image(img)
                row_numb += 1

        wb.save("BS_Report.xlsx")


class CreateBSDF:

    def __init__(self):
        self.bs_df = {}

    def create_store_bs(self, main_df):
        """
        Create DF for each category in store

        :param main_df: dict (Dict format {Store: Store DF...}
        :return: dict (Dict format {Store: {Category: {Category DF}...}
        """
        for keys in main_df:
            category = {}
            store_df = main_df[keys]
            # Обработка DF для каждой категории, сортировка по продажам в штучках и вывод первых 10 рез-ов
            ftw_df = store_df.loc[store_df['Категория'] == "FTW"].sort_values(['Кол-во продаж', 'Сумма со скидкой'],
                                                                              ascending=[False, False]).head(n=10)
            ftw_df["IMG"] = ""
            app_df = store_df.loc[store_df['Категория'] == "APP"].sort_values(['Кол-во продаж', 'Сумма со скидкой'],
                                                                              ascending=[False, False]).head(n=10)
            app_df["IMG"] = ""
            acc_df = store_df.loc[store_df['Категория'] == "ACC"].sort_values(['Кол-во продаж', 'Сумма со скидкой'],
                                                                              ascending=[False, False]).head(n=10)
            acc_df["IMG"] = ""
            ttl_df = store_df.sort_values(['Сумма со скидкой', 'Кол-во продаж'],
                                          ascending=[False, False]).head(n=10)
            ttl_df["IMG"] = ""
            # Удаление из топа продаж результатов с продажами 0
            final_ftw_df = ftw_df.loc[ftw_df['Кол-во продаж'] > 0]
            final_app_df = app_df.loc[app_df['Кол-во продаж'] > 0]
            final_acc_df = acc_df.loc[acc_df['Кол-во продаж'] > 0]
            final_ttl_df = ttl_df.loc[ttl_df['Кол-во продаж'] > 0]
            # Занесение результатов по каждой категории с определенными колонками
            category['FTW'] = final_ftw_df[['Артикул', "IMG", 'Модель ANTA', 'Сезон', 'RRP', 'Пол',
                                            'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', "Rank by all stores"]]
            category['APP'] = final_app_df[['Артикул', "IMG", 'Модель ANTA', 'Сезон', 'RRP', 'Пол',
                                            'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', "Rank by all stores"]]
            category['ACC'] = final_acc_df[['Артикул', "IMG", 'Модель ANTA', 'Сезон', 'RRP', 'Пол',
                                            'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', "Rank by all stores"]]
            category['TTL'] = final_ttl_df[['Артикул', "IMG", 'Модель ANTA', 'Сезон', 'RRP', 'Пол',
                                            'Коллекция', 'Кол-во продаж', 'Сумма со скидкой', "Rank by all stores"]]

            self.bs_df[keys] = category
        return self.bs_df


class CreateStoreDF:

    def __init__(self):
        self.store_names = None
        self.df_names = ['Магазин', 'Артикул', 'Наименование', 'Количество', 'Категория', 'Сезон', 'SKU', 'RRP',
                         'RRP Amount', 'Grade', 'Пол', 'Коллекция', 'Сумма со скидкой', 'Кол-во продаж', 'Модель ANTA']
        self.is_df_correct = True
        self.store_dataframes = {}

    def create_df(self, file_name):
        """
        Creation DF to each store for all category

        :param file_name: string (File path)
        :return:  dict (Dict format {Store: Store DF...}
        """
        df = pd.read_excel(file_name)   # Читаем загрузочный файл
        names = df.columns.to_list()  # Получаем заголовки таблицы и далее сравниваем их с необходимыми заголовками
        for header in self.df_names:
            if header not in names:
                self.is_df_correct = False
        if self.is_df_correct:
            self.process_df(df)   # Запускаем обработку файла
        return self.store_dataframes

    def process_df(self, data):
        df = data
        self.store_names = df['Магазин'].unique()   # Получаем уникальные магазины из  файла
        total = df.groupby("Артикул")[['Сумма со скидкой', 'Кол-во продаж']].sum().reset_index()   # Создаем общий DF на все магазины
        total_df = pd.merge(total,
                            df[['Артикул', 'Наименование', 'Модель ANTA', 'Категория', 'Сезон', 'SKU', 'RRP',
                                'Пол', 'Коллекция']],
                            on='Артикул', how='left').drop_duplicates(subset=['Артикул'])
        total_df["Rank by all stores"] = total_df["Сумма со скидкой"].rank(ascending=False,
                                                                           method='first', na_option='bottom')
        self.store_dataframes["Все магазины"] = total_df
        for store in self.store_names:
            store_df = df.loc[df['Магазин'] == store]
            store_df_vr = pd.merge(store_df, total_df[['Артикул', "Rank by all stores"]],
                                                        on='Артикул', how='left').drop_duplicates(subset=['Артикул'])
            self.store_dataframes[store] = store_df_vr     # Заполняем словарь Ключ(Магазин)-Значение(DF магазина)


if __name__ == "__main__":
    start = Processing()
    start.process()
