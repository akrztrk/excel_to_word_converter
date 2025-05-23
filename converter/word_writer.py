from docx import Document

class WordWriter:
    def __init__(self, dataframe, output_path):
        self.dataframe = dataframe
        self.output_path = output_path

    def write(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=len(self.dataframe.columns))

        # Header
        hdr_cells = table.rows[0].cells
        for i, column in enumerate(self.dataframe.columns):
            hdr_cells[i].text = str(column)

        # Rows
        for _, row in self.dataframe.iterrows():
            row_cells = table.add_row().cells
            for i, cell_value in enumerate(row):
                row_cells[i].text = str(cell_value)

        doc.save(self.output_path)

