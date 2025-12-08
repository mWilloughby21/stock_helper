import yfinance as yf
import datetime as dt
import time
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

#  Function to fetch stock closing prices for multiple tickers on a specific date / dict[col_letter: close]
def fetch_closing_prices(tickers: dict[str, str], date: dt.datetime) -> dict[str, float]:
    closing_prices = {}
    stock_list = []
    start = date.strftime("%Y-%m-%d")
    end = (date + dt.timedelta(days=1)).strftime("%Y-%m-%d")
    
    # Make list of tickers
    for ticker in tickers.keys():
        stock_list.append(ticker)
        
    # Fetch data in bulk
    df = yf.download(stock_list, start=start, end=end, interval="1d", progress=False, auto_adjust=True)
    
    if df.empty:
        raise ValueError(f"No data found for the provided tickers on {date.strftime('%Y-%m-%d')}")
    
    for ticker in stock_list:
        closing_prices[tickers[ticker]] = round(float(df['Close'][ticker].iloc[0]), 2)
        
    return closing_prices

# Function to validate ticker symbols
def is_valid_ticker(ticker: str) -> bool:
    if not isinstance(ticker, str) or ticker == "Index":
        return False
    
    ticker = ticker.strip().upper()
    return ticker.isalnum() and 1 <= len(ticker) <= 5

# Function to check if a date is today or in the future
def is_today(date) -> bool:
    today = dt.date.today()
    if isinstance(date, dt.datetime):
        return date.date() == today
    elif isinstance(date, dt.date):
        return date == today
    else:
        date = dt.datetime.strptime(date, "%m/%d/%y").date()
        today = dt.datetime.now().date()
        return date == today
    
# Function to check if a date is in the future
def is_future(date) -> bool:
    today = dt.date.today()
    if isinstance(date, dt.datetime):
        return date.date() > today
    elif isinstance(date, dt.date):
        return date > today
    else:
        date = dt.datetime.strptime(date, "%m/%d/%y").date()
        today = dt.datetime.now().date()
        return date > today

# Function to read portfolio tickers from Investment Analysis sheet / dict[ticker: col_letter]
def read_portfolio_tickers(sheet: Worksheet, ticker_row: int, start_idx: int, end_idx: int) -> dict[str, str]:
    
    tickers = {} 
    for col in range(start_idx, end_idx):
        col_letter = get_column_letter(col)
        cell_value = sheet[f"{col_letter}{ticker_row}"].value
        if is_valid_ticker(cell_value):
            ticker = cell_value.strip().upper()
            tickers[ticker] = col_letter
            
    return tickers

# Function to read portfolio dates from Investment Analysis sheet / dict[date: row]
def read_portfolio_dates(sheet: Worksheet, date_col: str, check_col: str, date_row_start: int, date_row_end: int) -> dict[dt.datetime, int]:
    dates = {}
    
    for row in range(date_row_start, date_row_end + 1):
        check_value = sheet[f"{check_col}{row}"].value
        
        if check_value is not None:
            continue
        
        date_value = sheet[f"{date_col}{row}"].value
        if isinstance(date_value, dt.datetime):
            date_dt = date_value
        elif isinstance(date_value, dt.date):
            date_dt = dt.datetime(date_value.year, date_value.month, date_value.day)
        else:
            date_dt = dt.datetime.strptime(date_value, "%m/%d/%y")
                
        dates[date_dt] = row
        last_date = date_dt
        
        if is_today(last_date):
            break
        
        if is_future(last_date):
            dates.pop(date_dt)
            break
        
    return dates

# Update close prices in the sheet for given tickers and date
def update_close_prices(sheet: Worksheet, closing_prices: dict[str, float], date_row: int) -> None:
    for col, price in closing_prices.items():
        sheet[f"{col}{date_row}"].value = price
        print(f"Price {price} written to {col}{date_row}")
        
def time_update(start_time: int) -> None:
    end_time = time.time()
    print(f"Total time elapsed: {end_time - start_time:.2f} seconds.")