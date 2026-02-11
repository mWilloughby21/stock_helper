# update_prices.py

from src.models.portfolio import Portfolio
from src.data_sources.csv_source import CSVSource
from src.services.cache_service import get_cached_price, update_cache
import yfinance as yf

CSV_PATH = 'data/input/stocks.csv'

def fetch_curr_prices(portfolio: Portfolio) -> dict[str, float]:
    for stock in portfolio.stocks:
        cached_price = get_cached_price(stock.ticker)
        if cached_price is not None:
            stock.current_price = cached_price
            continue
        
        try:
            ticker_data = yf.Ticker(stock.ticker)
            price = ticker_data.fast_info.get('last_price', None)
            if price is None:
                price = ticker_data.history(period='1d')['Close'].iloc[-1]
            
            stock.current_price = price
            update_cache(stock.ticker, price)
        except Exception as e:
            print(f"Error fetching price for {stock.ticker}: {e}")
            stock.current_price = None

def print_portfolio_summary(portfolio: Portfolio):
    total_current_value = 0.0
    print("\nPortfolio Summary:")
    for stock in portfolio.stocks:
        current_value = stock.total_shares * stock.current_price if stock.current_price is not None else 0.0
        gain_loss = current_value - stock.total_cost
        total_current_value += current_value
        print(
            f"{stock.ticker}: "
            f"{stock.total_shares} shares, "
            f"avg cost ${stock.average_cost:.2f}, "
            f"current price ${stock.current_price:.2f}, "
            f"gain/loss ${gain_loss:.2f}"
        )
    print(f"\nTotal invested: ${portfolio.total_invested:.2f}")
    print(f"Total current value: ${total_current_value:.2f}")
    print(f"Total gain/loss: ${total_current_value - portfolio.total_invested:.2f}")

def main():
    # Load lots from CSV
    csv_source = CSVSource(CSV_PATH)
    lots = csv_source.load_lots()
    
    # Create portfolio and add lots
    portfolio = Portfolio()
    portfolio.add_lots(lots)
    
    # Fetch latest prices
    fetch_curr_prices(portfolio)
    
    # Print summary
    print_portfolio_summary(portfolio)

if __name__ == "__main__":
    main()