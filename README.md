# Stock Helper

A Python-based stock portfolio helper that ingests transaction data, aggregates holdings, updates prices, and generates a reusable Excel-based portfolio report.

The project is designed to start simple (CSV input, Excel output) while being intentionally structured so it can scale later to:
- Database-backed storage
- Additional analytics
- More complex financial logic

Excel is treated as a presentation layer, not the source of truth.

---

## Project Goals

- Track ~40–50 stock tickers
- Support multiple purchases (lots) per ticker
- Generate and reuse a single Excel portfolio file
- Append new data safely and predictably
- Update current stock prices via script
- Perform portfolio calculations in a controlled, auditable way
- Allow future migration from CSV → database with minimal refactoring

---

## High-Level Architecture

**Data flow:**

CSV / Database  
→ Python models & services  
→ Excel portfolio report  

**Key principle:**  
Raw input data is never modified. All derived values are calculated.

---

## Folder Structure Overview

stock-helper/
├── data/
│ ├── input/ # Raw input data (CSV for now)
│ ├── output/ # Generated Excel portfolio
│ └── cache/ # Cached price data (optional)
│
├── src/
│ ├── main.py # Program entry point
│ ├── config/ # Centralized configuration
│ ├── models/ # Portfolio domain models
│ ├── data_sources/ # CSV / DB abstractions
│ ├── excel/ # Excel creation and updates
│ ├── services/ # Price updates and calculations
│ └── utils/ # Shared utilities
│
├── scripts/ # Task-specific scripts
├── tests/ # Unit and integration tests
├── requirements.txt
├── README.md
└── .gitignore

---

## Data Model Philosophy

### Input Data (CSV / Database)

- One row represents one purchase (lot)
- Append-only
- No calculated or derived fields
- Raw data is never modified

This structure supports:
- Multiple purchases per ticker
- Accurate aggregation
- Easy migration to a database

---

### Internal Models

- **StockLot**: Represents a single purchase
- **Stock**: Aggregated view per ticker
- **Portfolio**: Collection of all stocks

All aggregation and calculations happen in Python, not in Excel.

---

### Excel Output

The Excel file is:
- Created on first run
- Reused on subsequent runs
- Fully regenerable from source data

Recommended sheets:
- **Portfolio** – One row per ticker (primary view)
- **Lots** – Transaction-level detail (optional)
- **Summary** – Totals and highlights
- **Metadata** – Last update, price source, warnings

Excel formulas may be used, but all calculations must be reproducible in Python.

---

## Scripts

- **main.py**  
  Orchestrates the full workflow: load data → build portfolio → update Excel

- **scripts/update_prices.py**  
  Updates current prices for all tickers

- **scripts/rebuild_excel.py**  
  Recreates the Excel file from source data

All scripts are designed to be safe to rerun.

---

## Configuration

All configurable values live in `src/config/`, including:
- File paths
- API keys (not committed)
- Default currency
- Price update behavior

No business logic should live in configuration files.

---

## Error Handling & Safety

- Input data is validated early
- Missing prices or partial failures are surfaced clearly
- Excel updates are idempotent (no duplication on reruns)
- Source data remains untouched

---

## Out of Scope (For Now)

The following are explicitly not handled in the current version:
- Tax calculations
- Dividends and reinvestment
- Intraday or real-time pricing
- Options or derivatives
- Performance benchmarking

The architecture allows these to be added later without major refactoring.

---

## Future Extensions

- Database-backed data source
- Historical price tracking
- Benchmark comparison
- Portfolio risk metrics
- Command-line interface (CLI)
- Additional reporting formats

---

## Design Principles

- Separation of concerns
- Explicit data flow
- Predictable file behavior
- Readability over cleverness
- Future-proofing without overengineering

---