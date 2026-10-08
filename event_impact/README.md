# Event Impact Analyzer

Measures the market's **fear and greed** around news events (elections, budgets, RBI, crude, bank collapses, government crises, blackouts, pandemics, global shocks) and crisis periods since 1997 (including the 2026 Iran-war oil shock) across India and 11 world markets, and turns it into numbers for the trading model.

Open **`Event_Impact_Analyzer.xlsx`** and start with the *How_To_Use* sheet.

## What it answers

| Question | Sheet |
|---|---|
| How does Nifty react before, during and after elections, budgets, RBI rate changes, crude shocks and global shocks? | Category_Summary, Nifty_Impact |
| How do markets react to positive vs negative news, surprise vs expected? | Category_Summary (second table) |
| How long until the market settles (price recovered, India VIX normal)? | Category_Summary, Nifty_Impact |
| Who made the money: buying before, on the event day or after? R:R of each? | Category_Summary (third table) |
| Which watchlist stocks are hit hardest, and which stay resilient in fear? | Top_Lists, Stock_Scorecard |
| Which industries gain or lose for each event type? | Sector_Impact |
| At what size of crude move / crude price level does the market react? | Crude_Ranges |
| What do FII and DII flows do to Nifty? | FII_DII |
| How should I adjust my short-term and long-term targets for an upcoming event? | Target_Adjuster |
| How frightened is the market today, and what followed similar panic levels in the past? | Panic_Meter |
| How deep and long were past recessions / bear markets, and when was it best to buy? | Crisis_Periods |
| Which stocks and industries defended best in crises, and which recovered fastest? | Crisis_Scorecard, Crisis_Sectors, Top_Lists |
| How accurate are expert and analyst predictions? | Forecast_Studies, Forecast_Tracker |
| Where are we in the current (2026) crisis compared with history? | Crisis_Periods (WHERE ARE WE NOW) |
| Which countries recovered from each crisis and which did not? | Global_Recovery |
| Where did professional money move in each crisis? | Strategy_Shifts |

## Files

| File | Purpose |
|---|---|
| `events.csv` | Event calendar with news context (edit to add events) |
| `watchlist.csv` | Watchlist symbols, industry and market-cap group |
| `fii_dii_daily.csv` | Daily FII/DII data from the shares worksheet tracker |
| `download_data.py` | Downloads daily prices from Yahoo Finance into `data/prices/` |
| `analyze.py` | Measures every event for Nifty and each stock |
| `build_workbook.py` | Builds the Excel workbook (summaries are live formulas) |

## Updating

```bash
pip install yfinance pandas openpyxl
python download_data.py --refresh   # needs network access to query1/query2.finance.yahoo.com and fc.yahoo.com
python analyze.py
python build_workbook.py
```

Open the workbook in Excel and it recalculates automatically.
