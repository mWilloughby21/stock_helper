import time
import openpyxl


# Import helper functions
from helper import time_update

# Import config constants
from config import TEST_EXCEL_FILE_PATH, DEFAULT_START_ROW, EXCEL_FILE_PATH

# 
def test_market_index():
    start_tie = time.time()
    
    # Load workbook and select sheet
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb["Investment Analysis"]
    


test_market_index()