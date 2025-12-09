import time
import openpyxl

# Import helper functions
from helper import time_update, read_portfolio_dates, fetch_closing_prices, update_close_prices

# Import config constants
from config import EXCEL_FILE_PATH, MARKET_DICT, DATE_COL, DATE_CHECK_COL, DATE_CHECK_ROW_START, DATE_CHECK_ROW_END

# Stock market index test function
def test_market_index():
    start_time = time.time()
    
    # Load workbook and select sheet
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb["Investment Analysis"]
    
    dates = read_portfolio_dates(sheet, DATE_COL, DATE_CHECK_COL, DATE_CHECK_ROW_START, DATE_CHECK_ROW_END)
    
    for date in dates.keys():
        date_row = dates[date]
        closing_prices = fetch_closing_prices(MARKET_DICT, date)
        update_close_prices(sheet, closing_prices, date_row)
        time_update(start_time)
    
    # Save workbook with updated prices
    wb.save(EXCEL_FILE_PATH)
    print("Market index prices updated successfully.")
    time_update(start_time)


test_market_index()