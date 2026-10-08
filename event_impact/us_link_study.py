"""How US markets, US bond yields, US VIX, the dollar and gold affect Nifty.
US sessions close after the Indian close, so each US day is matched with the NEXT Indian trading day.
Run after download_data.py:  python us_link_study.py
"""
from pathlib import Path
import numpy as np
import pandas as pd

P = Path(__file__).parent / "data" / "prices"


def close(name):
    return pd.read_csv(P / f"{name}.csv", parse_dates=["Date"]).set_index("Date")["Close"].sort_index()


nifty = close("NIFTY")
n = pd.DataFrame({"Nifty": nifty})
n["N_ret"] = nifty.pct_change()
n["N_next20"] = nifty.shift(-20) / nifty - 1
n["N_next60"] = nifty.shift(-60) / nifty - 1
n["N_next250"] = nifty.shift(-250) / nifty - 1
n["N_same20"] = nifty.pct_change(20)

us = pd.DataFrame({"SP": close("G_US_SP500"), "NQ": close("G_US_NASDAQ"), "Y10": close("US_10Y"),
                   "Y3M": close("US_3M"), "VIX": close("US_VIX"), "DXY": close("DXY"), "GOLD": close("GOLD")}).ffill()
us["SP_ret"], us["NQ_ret"] = us["SP"].pct_change(), us["NQ"].pct_change()
us["VIX_chg"] = us["VIX"].pct_change()
us["Y10_20d_bps"] = (us["Y10"] - us["Y10"].shift(20)) * 100
us["Curve"] = us["Y10"] - us["Y3M"]
us["DXY_20d"] = us["DXY"].pct_change(20)
us["GOLD_20d"] = us["GOLD"].pct_change(20)
# last US session strictly before each Indian date
m = pd.merge_asof(n.reset_index(), us.reset_index().rename(columns={"Date": "US_Date"}),
                  left_on="Date", right_on="US_Date", allow_exact_matches=False).set_index("Date").dropna(subset=["SP_ret", "N_ret"])


def bucket(df, col, edges, labels, out_cols, fmt="pct"):
    g = df.groupby(pd.cut(df[col], edges, labels=labels), observed=True)
    t = g[out_cols].mean()
    t.insert(0, "Days", g.size())
    for c in out_cols:
        t[f"{c} >0"] = g[c].apply(lambda s: (s > 0).mean())
    return t


pd.set_option("display.width", 220)
print("=== 1. Overnight: last US S&P 500 move -> Nifty next Indian day")
print(f"Correlation {m['SP_ret'].corr(m['N_ret']):.2f}, beta {np.polyfit(m['SP_ret'], m['N_ret'], 1)[0]:.2f}")
t = bucket(m, "SP_ret", [-1, -0.03, -0.02, -0.01, 0, 0.01, 0.02, 0.03, 1],
           ["< -3%", "-3% to -2%", "-2% to -1%", "-1% to 0", "0 to +1%", "+1% to +2%", "+2% to +3%", "> +3%"], ["N_ret", "N_next20"])
print((t.assign(**{c: t[c] for c in t.columns})).round(4).to_string())
print("\nCorrelation of daily moves by period:")
for a, b in [("2007", "2012"), ("2013", "2019"), ("2020", "2023"), ("2024", "2026")]:
    s = m[a:b]
    print(f"  {a}-{b}: {s['SP_ret'].corr(s['N_ret']):.2f}   (Nasdaq {s['NQ_ret'].corr(s['N_ret']):.2f})")

w = pd.DataFrame({"N": nifty, "SP": us["SP"]}).ffill().resample("W-FRI").last().pct_change().dropna()
print("\nWeekly correlation by year:", {y: round(g["N"].corr(g["SP"]), 2) for y, g in w.groupby(w.index.year)})

print("\n=== 2. US 10-year yield: 20-day change -> Nifty (same 20 days, next 20 days)")
print(bucket(m, "Y10_20d_bps", [-500, -50, -25, 0, 25, 50, 500],
             ["fall >50bp", "fall 25-50bp", "fall 0-25bp", "rise 0-25bp", "rise 25-50bp", "rise >50bp"],
             ["N_same20", "N_next20"]).round(3).to_string())
print("\nUS 10-year yield LEVEL -> Nifty next 1 year")
print(bucket(m, "Y10", [0, 2, 3, 4, 4.5, 5, 10], ["<2%", "2-3%", "3-4%", "4-4.5%", "4.5-5%", ">5%"],
             ["N_next60", "N_next250"]).round(3).to_string())
print("\nUS yield curve (10Y minus 3M): inverted vs normal -> Nifty next 1 year")
print(bucket(m, "Curve", [-10, -0.5, 0, 1, 2, 10], ["inverted >0.5", "inverted 0-0.5", "flat 0-1", "normal 1-2", "steep >2"],
             ["N_next250"]).round(3).to_string())

print("\n=== 3. US VIX level -> Nifty")
print(bucket(m, "VIX", [0, 15, 20, 25, 30, 40, 100], ["<15", "15-20", "20-25", "25-30", "30-40", ">40"],
             ["N_ret", "N_next20", "N_next250"]).round(3).to_string())
sp = m[m["VIX_chg"] > 0.30]
print(f"\nUS VIX jumped >30% in a day: {len(sp)} times; Nifty next day avg {sp['N_ret'].mean():.2%}, "
      f"down {(sp['N_ret'] < 0).mean():.0%} of times; Nifty 20 days later avg {sp['N_next20'].mean():.2%}")

print("\n=== 4. US dollar index (DXY) 20-day change -> Nifty")
print(bucket(m, "DXY_20d", [-1, -0.03, -0.01, 0.01, 0.03, 1], ["dollar -3%+", "-3% to -1%", "flat", "+1% to +3%", "dollar +3%+"],
             ["N_same20", "N_next20"]).round(3).to_string())

print("\n=== 5. Gold 20-day change vs Nifty same 20 days")
print(bucket(m, "N_same20", [-1, -0.10, -0.05, 0, 0.05, 1], ["Nifty -10%+", "-10% to -5%", "-5% to 0", "0 to +5%", "+5%+"],
             ["GOLD_20d"]).round(3).to_string())

print("\n=== 6. Current readings (latest)")
last = m.iloc[-1]
print(f"US 10Y {last['Y10']:.2f}%  3M {last['Y3M']:.2f}%  curve {last['Curve']:.2f}  10Y 20d change {last['Y10_20d_bps']:.0f}bp  "
      f"US VIX {last['VIX']:.1f}  DXY {last['DXY']:.1f} (20d {last['DXY_20d']:.1%})  Gold {last['GOLD']:.0f} (20d {last['GOLD_20d']:.1%})")
peak = us["SP"]["2025-12-01":].max()
print(f"S&P 500 vs its 2026 high: {us['SP'].iloc[-1] / peak - 1:.1%};  Nifty vs its 2026 high: {nifty.iloc[-1] / nifty['2025-12-01':].max() - 1:.1%}")
print("US 10Y monthly:", us["Y10"].resample("ME").last()["2025-06":].round(2).to_dict())
