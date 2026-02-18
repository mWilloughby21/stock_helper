# stock.py

from typing import List
from .stock_lot import StockLot

# This class represents all purchases of a single ticker aggregated together.
class Stock:
    def __init__(self, ticker: str, lots: List[StockLot]):
        self.ticker: str = ticker
        self.lots: List[StockLot] = lots
        self._current_price: float | None = None
        self.cols: List = None
    
    def __repr__(self):
        return f"Stock(ticker={self.ticker}, lots={self.lots})"
    
    @property
    def total_shares(self) -> float:
        return sum(lot.shares for lot in self.lots)
    
    @property
    def total_cost(self) -> float:
        return sum(lot.shares * lot.purchase_price for lot in self.lots)
    
    @property
    def average_cost(self) -> float:
        total_shares = self.total_shares
        return self.total_cost / total_shares if total_shares > 0 else 0.0
    
    @property
    def has_lots(self) -> bool:
        return len(self.lots) > 0
    
    @property
    def current_price(self) -> float | None:
        return self._current_price
    
    @current_price.setter
    def current_price(self, price: float | None):
        self._current_price = price
    
    # Add new lots to existing stock
    def add_lots(self, new_lots: List[StockLot]):
        self.lots.extend(new_lots)