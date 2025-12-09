import openpyxl
import yfinance as yf
import datetime as dt

# Config constants
from config import TEST_EXCEL_FILE_PATH, DEFAULT_START_ROW

# This test script adds a new stock to the investment analysis Excel sheet
def test_new_stock(ticker: str, col_letter: str):
    print(f"Adding new stock {ticker} in column {col_letter}...")
    
    # Load workbook and select sheet
    wb = openpyxl.load_workbook(TEST_EXCEL_FILE_PATH)
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
    row = DEFAULT_START_ROW
    for price in closing_prices:
        sheet[f"{col_letter}{row}"].value = float(price)
        print(f"Updated {col_letter}{row} with closing price {price}.")
        row += 1
    
    # Save workbook
    wb.save(TEST_EXCEL_FILE_PATH)
    print(f"New stock {ticker} added successfully in column {col_letter}.")


test_new_stock("CRWD", "UB")