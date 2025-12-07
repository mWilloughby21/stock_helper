import os

# Config constants

# Base directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Excel file paths
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "..", "data", "original.xlsx")
TEST_EXCEL_FILE_PATH = os.path.join(BASE_DIR, "..", "data", "test.xlsx")


# Date Columnn Settings
DATE_COL = "A"
DATE_CHECK_COL = "C"
DATE_CHECK_ROW_START = 11
DATE_CHECK_ROW_END = 260

# Ticker Row
TICKER_ROW = 8

# Start and End Columns for Tickers
START_COL = "AC"
END_COL = "TU"


# Default row to start writing closing prices
DEFAULT_START_ROW = 11