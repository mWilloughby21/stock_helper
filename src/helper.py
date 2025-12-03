import yfinance as yf
import openpyxl
import datetime as dt
from openpyxl.utils.datetime import from_excel

# Function to fetch stock closing prices
def fetch_close_price(ticker: str, excel_date) -> float:
    
    # Excel date error checks
    if excel_date is None:
        raise ValueError("Date is empty in Excel sheet.")
    
    # Ensure excel_date is datetime
    if isinstance(excel_date, dt.datetime):
        date = excel_date
        
    elif isinstance(excel_date, (int, float)):
        date = from_excel(excel_date)
        
    elif isinstance(excel_date, str):
        possible_formats = [
            "%Y-%m-%d",  # ISO format
            "%m/%d/%Y",  # US format
            "%d/%m/%Y"   # European format
        ]
        
        for fmt in possible_formats:
            try:
                date = dt.datetime.strptime(excel_date, fmt)
                break
            except ValueError:
                continue
            
        else:
            raise ValueError(f"Date format for {excel_date} is not recognized.")
        
    else:
        raise TypeError(f"Unsupported type for excel_date: {type(excel_date)}")
    
    start = date.strftime("%Y-%m-%d")
    end = (date + dt.timedelta(days=1)).strftime("%Y-%m-%d")
    
    close = yf.download(ticker, start=start, end=end, interval='1d', progress=False)
    
    if not close.empty:
        return float(close['Close'][0])
    else:
        raise ValueError(f"No data found for {ticker} on {date}")
    
# Function to validate ticker symbols
def is_valid_ticker(ticker: str) -> bool:
    if not isinstance(ticker, str):
        return False
    
    ticker = ticker.strip().upper()
    
    if ticker == "":
        return False

    try:
        info = yf.Ticker(ticker).fast_info
        return (
            info is not None and
            isinstance(info, dict) and
            info.get("last_price") is not None
        )
    except:
        return False

# Function to read portfolio tickers from Investment Analysis sheet
def read_portfolio_tickers(file_path: str, start_col: str, end_col: str) -> list:
    wb = openpyxl.load_workbook(file_path)
    sheet = wb["Investment Analysis"]
    
    start_idx = openpyxl.utils.column_index_from_string(start_col)
    end_idx = openpyxl.utils.column_index_from_string(end_col) + 1
    
    tickers = [] # Create dictionary to store tickers with column values
    for col in range(start_idx, end_idx):
        col_letter = openpyxl.utils.get_column_letter(col)
        cell_value = sheet[f"{col_letter}8"].value
        if cell_value and is_valid_ticker(cell_value):
            tickers.append(cell_value.strip().upper())
            
    return tickers

def read_portfolio_date(file_path: str, date_col: str) -> dt.datetime:
    pass