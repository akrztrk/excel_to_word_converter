from .excel_reader import ExcelReader
from .word_writer import WordWriter
import os

class ConverterController:
    def __init__(self, excel_path, output_folder):
        self.excel_path = excel_path
        self.output_folder = output_folder

    def convert_all_sheets(self):
        reader = ExcelReader(self.excel_path)
        sheet_data = reader.read_sheets()

        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]

        for sheet_name, df in sheet_data.items():
            output_path = os.path.join(
                self.output_folder, f"{base_name}_{sheet_name}.docx"
            )
            writer = WordWriter(df, output_path)
            writer.write()


