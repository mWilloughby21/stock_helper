# base_source.py

from typing import List
from src.models.stock_lot import StockLot

class BaseSource:
    def load_lots(self) -> List[StockLot]:
        raise NotImplementedError("Must be implemented by subclasses")