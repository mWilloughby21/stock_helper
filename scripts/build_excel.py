# build_excel.py

from pathlib import Path
from src.excel.workbook_manager import WorkbookManager
from src.services.format_excel import SetupFormatter

CSV_PATH = Path('data/input/stocks.csv')
EXCEL_PATH = Path('data/output/portfolio.xlsx')

def main():

    # Manage workbook
    wb_mgr = WorkbookManager(EXCEL_PATH)
    wb = wb_mgr.load_or_create()
    ws = wb.active
    ws.title = "Portfolio"
    ws.sheet_view.showGridLines = False
    
    # Initialed Data for Excel Sheet
    formatter = SetupFormatter(ws)
    
    formatter.center()
    formatter.get_year()
    dates = formatter.get_dates()
    for row, (date, day) in enumerate(dates, start=6):
        ws.cell(row=row, column=1, value=date)
        ws.cell(row=row, column=2, value=day[0:3])
    
    formatter.specific_cells()
    
    # Format + Save workbook
    formatter.static_format()
    wb_mgr.save()
    print(f"Portfolio Excel file saved to {EXCEL_PATH}")

if __name__ == "__main__":
    main()