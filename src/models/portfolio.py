# portfolio.py

from collections import defaultdict
from typing import List
from .stock import Stock
from .stock_lot import StockLot

class Portfolio:
    def __init__(self):
        self.stocks: List[Stock] = []
    
    def __repr__(self):
        return f"Portfolio(stocks={self.stocks})"
    
    @property
    def total_invested(self) -> float:
        return sum(stock.total_cost() for stock in self.stocks)
    
    def get_stock_by_ticker(self, ticker: str) -> Stock | None:
        for stock in self.stocks:
            if stock.ticker == ticker:
                return stock
        return None
    
    # Group the lots by ticker and create / update Stock objects in the portfolio
    def add_lots(self, lots: List[StockLot]):
        grouped_lots = defaultdict(list)
        
        for lot in lots:
            grouped_lots[lot.ticker].append(lot)
        
        for ticker, ticker_lots in grouped_lots.items():
            stock = self.get_stock_by_ticker(ticker)
            
            if stock is None:
                self.stocks.append(Stock(ticker, ticker_lots))
            else:
                stock.lots.extend(ticker_lots)
