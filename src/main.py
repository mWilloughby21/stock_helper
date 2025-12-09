import openpyxl
from openpyxl.utils import column_index_from_string
from helper import read_portfolio_tickers, read_portfolio_dates, fetch_closing_prices, update_close_prices, time_update
import time

# Config constants
from config import EXCEL_FILE_PATH, DATE_COL, DATE_CHECK_COL, DATE_CHECK_ROW_START, DATE_CHECK_ROW_END, TICKER_ROW, START_COL, END_COL


def main():
    start_time = time.time()
    
    # Load workbook and select sheet
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb["Investment Analysis"]
    
    # Set column indices
    start_idx = column_index_from_string(START_COL)
    end_idx = column_index_from_string(END_COL) + 1
    
    # Get portfolio date and tickers
    dates = read_portfolio_dates(sheet, DATE_COL, DATE_CHECK_COL, DATE_CHECK_ROW_START, DATE_CHECK_ROW_END)
    tickers = read_portfolio_tickers(sheet, TICKER_ROW, start_idx, end_idx)
    
    for date in dates.keys():
        closing_prices = fetch_closing_prices(tickers, date)
        date_row = dates[date]
        update_close_prices(sheet, closing_prices, date_row)
        time_update(start_time)
    
    # Save workbook with updated prices
    wb.save(EXCEL_FILE_PATH)
    print("Stock closing prices updated successfully.")
    time_update(start_time)


if __name__ == "__main__":
    main()