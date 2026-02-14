# sheet_writer.py

from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet

class SheetWriter:
    def __init__(self, ws: Worksheet):
        self.ws = ws
    
    # TODO: Change 
    def write_headers(self, headers: list[str]):
        if self.ws.max_row == 1 and self.ws.cell(row=1, column=1).value is None:
            for col, header in enumerate(headers, start=1):
                cell = self.ws.cell(1, col, header)
                cell.font = Font(bold=True)
    
    def append_row(self, row: list):
        self.ws.append(row)
    
    def add_column_if_missing(self, header: str):
        headers = [self.ws.cell(1, col).value for col in range(1, self.ws.max_column + 1)]
        if header not in headers:
            self.ws.cell(1, self.ws.max_column + 1, header).font = Font(bold=True)