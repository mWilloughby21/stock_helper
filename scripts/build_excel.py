# build_excel.py

from pathlib import Path
from src.models.portfolio import Portfolio
from src.data_sources.csv_source import CSVSource
from src.excel.workbook_manager import WorkbookManager
from src.excel.sheet_writer import SheetWriter
from src.excel import formulas
from scripts.update_prices import fetch_curr_prices
from src.services.format_excel import PortfolioFormatter

CSV_PATH = Path('data/input/stocks.csv')
EXCEL_PATH = Path('data/output/portfolio.xlsx')

def main():
    # Load lots from CSV
    lots = CSVSource(CSV_PATH).load_lots()
    
    # Build portfolio
    portfolio = Portfolio()
    portfolio.add_lots(lots)
    
    # Fetch latest prices
    fetch_curr_prices(portfolio)
    
    # Manage workbook
    wb_mgr = WorkbookManager(EXCEL_PATH)
    wb = wb_mgr.load_or_create()
    ws = wb.active
    ws.title = "Portfolio"
    
    # Write headers
    writer = SheetWriter(ws)
    headers = ["Ticker", "Shares", "Avg Cost", "Total Cost", "Current Price", "Total Value", "Gain/Loss", "% Change"]
    writer.write_headers(headers)
    
    # Format Portfolio Sheet
    formatter = PortfolioFormatter(ws)
    formatter.static_format()
    
    # Write rows
    for stock in portfolio.stocks:
        total_value = stock.total_shares * (stock.current_price or 0)
        row = [
            stock.ticker,
            stock.total_shares,
            stock.average_cost,
            stock.total_cost,
            stock.current_price or 0,
            total_value,
            total_value - stock.total_cost,
            formulas.percent_change(f"F{ws.max_row + 1}", f"D{ws.max_row + 1}")
        ]
        writer.append_row(row)
        
    formatter.dynamic_format(2, ws.max_row)
    
    # Save workbook
    wb_mgr.save()
    print(f"Portfolio Excel file saved to {EXCEL_PATH}")

if __name__ == "__main__":
    main()