# stock_lot.py

from datetime import date

# This class represents a single purchase of a stock, including the ticker, number of shares, purchase price, and purchase date.
class StockLot:
    def __init__(self, ticker: str, shares: float, purchase_price: float, purchase_date: date):
        self.ticker: str = ticker
        self.shares: float = shares
        self.purchase_price: float = purchase_price
        self.purchase_date: date = purchase_date
    
    def __repr__(self):
        return f"StockLot(ticker={self.ticker}, shares={self.shares}, purchase_price={self.purchase_price}, purchase_date={self.purchase_date})"