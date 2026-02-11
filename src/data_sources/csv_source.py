# csv_source.py

from datetime import datetime
from pathlib import Path
import csv

from src.models.stock_lot import StockLot
from .base_source import BaseSource

class CSVSource(BaseSource):
    def __init__(self, path: str):
        self.path = Path(path)
    
    def load_lots(self) -> list[StockLot]:
        lots = []
        
        with open(self.path, newline='') as f:
            reader = csv.DictReader(f)
            for row in reader:
                lot = StockLot(
                    ticker=row['ticker'],
                    shares=float(row['shares']),
                    purchase_price=float(row['purchase_price']),
                    purchase_date=datetime.strptime(row['purchase_date'], '%Y-%m-%d').date()
                )
                lots.append(lot)
        return lots