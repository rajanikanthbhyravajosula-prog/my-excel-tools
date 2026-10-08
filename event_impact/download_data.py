"""Download daily price history for the indices, macro series and watchlist stocks.

Saves one CSV per ticker into data/prices/ (cached: existing files are skipped
unless --refresh is given). Source: Yahoo Finance via the yfinance package.
"""
import csv
import sys
import time
from pathlib import Path

import yfinance as yf

HERE = Path(__file__).parent
OUT = HERE / "data" / "prices"
START = "1997-01-01"

# Market-wide series. Yahoo's Nifty history starts Sep-2007, so Sensex covers 2004-2007.
MARKET = {
    "NIFTY": "^NSEI",
    "SENSEX": "^BSESN",
    "BANKNIFTY": "^NSEBANK",
    "INDIAVIX": "^INDIAVIX",
    "BRENT": "BZ=F",
    "WTI": "CL=F",
    "USDINR": "USDINR=X",
}

# Global indices for the country recovery comparison (local-currency price indices)
GLOBAL = {
    "G_US_SP500": "^GSPC",
    "G_US_NASDAQ": "^IXIC",
    "G_UK_FTSE100": "^FTSE",
    "G_GERMANY_DAX": "^GDAXI",
    "G_JAPAN_NIKKEI": "^N225",
    "G_HONGKONG_HANGSENG": "^HSI",
    "G_CHINA_SHANGHAI": "000001.SS",
    "G_KOREA_KOSPI": "^KS11",
    "G_TAIWAN_TAIEX": "^TWII",
    "G_BRAZIL_BOVESPA": "^BVSP",
    "G_INDONESIA_JCI": "^JKSE",
    # US rates, volatility, dollar and gold (used by us_link_study.py)
    "US_10Y": "^TNX",
    "US_3M": "^IRX",
    "US_30Y": "^TYX",
    "US_VIX": "^VIX",
    "DXY": "DX-Y.NYB",
    "GOLD": "GC=F",
}


# Watchlist symbols that Yahoo lists under a different code
YAHOO_ALIAS = {"LTIM": "LTM"}


def watchlist():
    with open(HERE / "watchlist.csv", newline="") as f:
        return [row["NSE_Code"] for row in csv.DictReader(f)]


def fetch(name, ticker, refresh):
    path = OUT / f"{name}.csv"
    if path.exists() and not refresh:
        return "cached"
    for attempt in range(4):
        try:
            df = yf.download(ticker, start=START, progress=False, auto_adjust=True, threads=False)
            if df is None or df.empty:
                return "EMPTY"
            if hasattr(df.columns, "levels"):
                df.columns = df.columns.get_level_values(0)
            df = df[["Open", "High", "Low", "Close", "Volume"]].dropna(subset=["Close"])
            df.to_csv(path, index_label="Date")
            return f"{len(df)} rows from {df.index.min().date()}"
        except Exception as e:  # rate limits: back off and retry
            time.sleep(2 ** (attempt + 1))
            err = e
    return f"FAILED ({err})"


def main():
    refresh = "--refresh" in sys.argv
    OUT.mkdir(parents=True, exist_ok=True)
    jobs = list(MARKET.items()) + list(GLOBAL.items()) + [(s, f"{YAHOO_ALIAS.get(s, s)}.NS") for s in watchlist()]
    for name, ticker in jobs:
        print(f"{name:12s} {fetch(name, ticker, refresh)}", flush=True)
        time.sleep(0.3)


if __name__ == "__main__":
    main()
