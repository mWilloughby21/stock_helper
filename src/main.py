import openpyxl
from openpyxl.utils import column_index_from_string
from helper import read_portfolio_tickers, read_portfolio_dates, fetch_closing_prices, update_close_prices
import os
import yfinance as yf
import datetime as dt


# Config constants
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "..", "data", "original.xlsx")
DATE_COL = "A"
DATE_CHECK_COL = "C"
DATE_CHECK_ROW_START = 11
DATE_CHECK_ROW_END = 260
TICKER_ROW = 8
START_COL = "AC"
END_COL = "TU"

def main():
    # Load workbook and select sheet
    wb_values = openpyxl.load_workbook(EXCEL_FILE_PATH, data_only=True, read_only=True)
    wb_formulas = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb_values["Investment Analysis"]
    update_sheet = wb_formulas["Investment Analysis"]
    
    # Set column indices
    start_idx = column_index_from_string(START_COL)
    end_idx = column_index_from_string(END_COL) + 1
    
    # Get portfolio date and tickers
    dates = read_portfolio_dates(sheet, DATE_COL, DATE_CHECK_COL, DATE_CHECK_ROW_START, DATE_CHECK_ROW_END)
    tickers = read_portfolio_tickers(sheet, TICKER_ROW, start_idx, end_idx)
    
    for date in dates.keys():
        closing_prices = fetch_closing_prices(tickers, date)
        date_row = dates[date]
        update_close_prices(update_sheet, closing_prices, date_row)
    
    # Save workbook with updated prices
    wb_formulas.save(EXCEL_FILE_PATH)
    print("Stock closing prices updated successfully.")


if __name__ == "__main__":
    main()