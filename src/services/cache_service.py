# cache_service.py

from pathlib import Path
import json
from datetime import datetime, date

CACHE_FILE = Path('data/cache/price_cache.json')

def load_cache() -> dict[str, dict[str, float | str]]:
    if CACHE_FILE.exists():
        with open(CACHE_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_cache(cache: dict[str, dict[str, float | str]]):
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CACHE_FILE, 'w') as f:
        json.dump(cache, f, indent=2)

def get_cached_price(ticker: str) -> float | None:
    cache = load_cache()
    data = cache.get(ticker)
    if data:
        cached_date = datetime.strptime(data['date'], '%Y-%m-%d').date()
        if cached_date == date.today():
            return data['price']
    return None

def update_cache(ticker: str, price: float):
    cache = load_cache()
    cache[ticker] = {
        'price': price,
        'date': date.today().isoformat()
    }
    save_cache(cache)