import openpyxl
import os
import yfinance as yf
import datetime as dt

# Config constants
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "..", "data", "test.xlsx")
DATE_COL = "A"
DATE_CHECK_COL = "C"
DATE_CHECK_ROW_START = 11
DATE_CHECK_ROW_END = 260
TICKER_ROW = 8
START_COL = "AC"
END_COL = "TU"

# This test script adds a new stock to the investment analysis Excel sheet
def test_new_stock(ticker: str, col_letter: str):
    print(f"Adding new stock {ticker} in column {col_letter}...")
    
    # Load workbook and select sheet
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb["Investment Analysis"]
    
    # Fetch historical data from yfinance
    df = yf.download(ticker, start="2025-01-01", end=dt.datetime.now().strftime("%Y-%m-%d"), interval="1d", progress=False, auto_adjust=True, group_by='ticker')
    if df.empty:
        print(f"No data found for ticker {ticker} between 2025-01-01 and {dt.datetime.now().strftime('%Y-%m-%d')}.")
        return
    else:
        print(f"Data fetched for {ticker}: {len(df)} records.")
    
    # Extract closing prices
    closing_prices = df[ticker]['Close'].round(2).tolist()
    
    # Update sheet with new closing prices
    row = 11
    for price in closing_prices:
        sheet[f"{col_letter}{row}"].value = float(price)
        print(f"Updated {col_letter}{row} with closing price {price}.")
        row += 1
    
    # Save workbook
    wb.save(EXCEL_FILE_PATH)
    print(f"New stock {ticker} added successfully in column {col_letter}.")


test_new_stock("CRWD", "UB")