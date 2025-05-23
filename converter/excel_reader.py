import pandas as pd

class ExcelReader:
    def __init__(self, file_path):
        self.file_path = file_path

    def read_sheets(self):
        xl = pd.ExcelFile(self.file_path)
        dataframes = {sheet: xl.parse(sheet) for sheet in xl.sheet_names}
        return dataframes

