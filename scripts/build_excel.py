# build_excel.py

from pathlib import Path
from src.excel.workbook_manager import WorkbookManager
from src.excel.sheet_writer import SheetWriter
from src.services.format_excel import SetupFormatter
from src.models.portfolio import Portfolio
from src.data_sources.csv_source import CSVSource

CSV_PATH = Path('data/input/stocks.csv')
EXCEL_PATH = Path('data/output/portfolio.xlsx')

def main():
    # Manage Workbook
    wb_mgr = WorkbookManager(EXCEL_PATH)
    wb = wb_mgr.load_or_create()
    ws = wb.active
    ws.title = "Portfolio"
    ws.sheet_view.showGridLines = False
    
    # Init Data for Excel Sheet
    formatter = SetupFormatter(ws)
    
    formatter.center()
    formatter.get_year()
    dates = formatter.get_dates()
    for row, (date, day) in enumerate(dates, start=8):
        ws.cell(row=row, column=1, value=date)
        ws.cell(row=row, column=2, value=day[0:3])
    
    formatter.specific_cells()
    
    # Format Init Data
    formatter.static_format()
    
    # Create Stock Info
    writer = SheetWriter(ws)
    portfolio = Portfolio()
    csv_source = CSVSource(CSV_PATH)
    lots = csv_source.load_lots()
    portfolio.add_lots(lots)
    portfolio.calc_port_pct()
    for stock in portfolio.stocks:
        cols = writer.get_block()
        writer.format_info(cols)
        writer.write_info(cols, stock)
    
    wb_mgr.save()
    print(f"Portfolio Excel file saved to {EXCEL_PATH}")

if __name__ == "__main__":
    main()