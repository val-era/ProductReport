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
        pass


if __name__ == "__main__":
    start = Processing()
    start.process()