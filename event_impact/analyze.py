"""Measure how Nifty and every watchlist stock reacted around each event.

Reads data/prices/*.csv (from download_data.py), events.csv and watchlist.csv.
Writes data/nifty_impact.csv, data/stock_events.csv and data/market_daily.csv,
which build_workbook.py turns into the Excel workbook.

Event windows (trading days, T0 = first trading day on/after the event date):
  Pre20 / Pre5   : run-up before the event      close(T-1) / close(T-21 or T-6) - 1
  Day0           : event-day reaction            close(T0) / close(T-1) - 1
  Post5/20/60    : continuation after the event  close(T+n) / close(T0) - 1
  Total20        : whole move from pre-event     close(T+20) / close(T-1) - 1
  MaxFall20/Rise : worst low / best high T0..T+20 vs close(T-1)
  Recovery days  : trading days from T0 until close is back at close(T-1),
                   counted only when the index/stock fell at least 2% (0 = no real fall)
  R:R per entry  : buy at close of T-5 (Before), T0 (Event day), T+5 (After);
                   reward = best high in next 20 days, risk = worst low in next 20 days
"""
import csv
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).parent
PRICES = HERE / "data" / "prices"
OUT = HERE / "data"

H = 20               # short-term horizon (trading days)
RECOVERY_CAP = 250   # stop counting recovery after ~1 year
FALL_FOR_RECOVERY = -0.02
CRUDE_MOVE = 0.20    # Brent move in 20 trading days that counts as a crude shock
CRUDE_COOLDOWN = 60  # trading days before another shock in the same direction

INDUSTRY_GROUPS = {  # watchlist industries used for the crude-oil range analysis
    "Petroleum": ["Petroleum"],
    "Paints": ["Paints&Pigments"],
    "Tyres": ["Tyres"],
    "Aviation/Transport": ["Transport"],
    "Chemicals": ["Chemicals", "Petrochemicals"],
    "Auto": ["Auto", "Auto Ancillaries"],
    "Cement": ["Cement"],
    "Banks": ["Banks"],
    "Finance": ["Finance"],
    "IT": ["IT"],
    "FMCG": ["Personal Care", "Food Processing", "Tobacco"],
    "Pharma": ["Pharma and Health"],
}


SPIKE_UP, SPIKE_DOWN = 0.40, -0.30  # one-day moves beyond these are treated as data faults
JUMPS = {}  # stock -> positions of split/demerger-like jumps left after cleaning


def load(name, clean=False):
    path = PRICES / f"{name}.csv"
    if not path.exists():
        return None
    df = pd.read_csv(path, parse_dates=["Date"]).set_index("Date").sort_index()
    df = df[~df.index.duplicated()]
    df = df[df["Close"] > 0]
    jumps = np.array([], dtype=int)
    if clean:
        # Yahoo's older stock data has one-day spikes that reverse next day: drop those days
        for _ in range(3):
            r = df["Close"].pct_change()
            nxt = r.shift(-1)
            spike = ((r > SPIKE_UP) & (nxt < SPIKE_DOWN)) | ((r < SPIKE_DOWN) & (nxt > SPIKE_UP))
            if not spike.any():
                break
            df = df[~spike]
        # Remaining huge jumps are unadjusted splits, bonuses or demergers: remember them
        r = df["Close"].pct_change().values
        jumps = np.nonzero((r > SPIKE_UP) | (r < SPIKE_DOWN))[0]
        body_hi = df[["Open", "Close"]].max(axis=1)
        body_lo = df[["Open", "Close"]].min(axis=1)
        df["High"] = df["High"].clip(upper=body_hi * 1.15)
        df["Low"] = df["Low"].clip(lower=body_lo * 0.85)
    # Yahoo sometimes leaves High/Low at zero or outside the close; repair conservatively
    df["High"] = df[["High", "Close"]].max(axis=1)
    df["Low"] = df[["Low", "Close"]].where(df["Low"] > 0, df["Close"], axis=0).min(axis=1)
    JUMPS[name] = jumps
    return df if len(df) > 30 else None


def pos_on_or_after(index, date):
    i = index.searchsorted(pd.Timestamp(date))
    return i if i < len(index) else None


def window_metrics(df, i0):
    """All phase metrics for one price series with the event at position i0."""
    c, hi, lo = df["Close"].values, df["High"].values, df["Low"].values
    n = len(c)
    if i0 is None or i0 < 21:
        return None
    pre = c[i0 - 1]

    def ret(a, b):
        return c[b] / c[a] - 1 if b < n else np.nan

    end20 = min(i0 + H, n - 1)
    m = {
        "Pre20": c[i0 - 1] / c[i0 - 21] - 1,
        "Pre5": c[i0 - 1] / c[i0 - 6] - 1,
        "Day0": c[i0] / pre - 1,
        "Day0_Low": lo[i0] / pre - 1,
        "Post5": ret(i0, i0 + 5),
        "Post20": ret(i0, i0 + 20),
        "Post60": ret(i0, i0 + 60),
        "Total20": c[i0 + 20] / pre - 1 if i0 + 20 < n else np.nan,
        "MaxFall20": lo[i0:end20 + 1].min() / pre - 1,
        "MaxRise20": hi[i0:end20 + 1].max() / pre - 1,
    }
    # recovery: only meaningful if the close fell at least 2% below the pre-event close
    trough = i0 + int(np.argmin(c[i0:end20 + 1]))
    if c[trough] / pre - 1 > FALL_FOR_RECOVERY:
        m["Recovery_Days"] = 0
    else:
        later = np.nonzero(c[trough:min(trough + RECOVERY_CAP, n)] >= pre)[0]
        m["Recovery_Days"] = (trough - i0 + int(later[0])) if len(later) else np.nan
    for label, e in (("Before", i0 - 5), ("EventDay", i0), ("After", i0 + 5)):
        if e + H < n:
            reward = max(hi[e + 1:e + H + 1].max() / c[e] - 1, 0.0)
            risk = max(1 - lo[e + 1:e + H + 1].min() / c[e], 0.0)
            m[f"Ret20_{label}"] = c[e + H] / c[e] - 1
            m[f"Reward_{label}"] = reward
            m[f"Risk_{label}"] = risk
            m[f"RR_{label}"] = reward / max(risk, 0.01)
        else:
            for k in ("Ret20", "Reward", "Risk", "RR"):
                m[f"{k}_{label}"] = np.nan
    return m


def vix_metrics(vix, date):
    if vix is None:
        return {}
    i0 = pos_on_or_after(vix.index, date)
    if i0 is None or i0 < 10:
        return {}
    v = vix["Close"].values
    base = v[i0 - 10:i0].mean()
    end = min(i0 + H, len(v) - 1)
    peak_i = i0 + int(np.argmax(v[i0:end + 1]))
    back = np.nonzero(v[peak_i:min(peak_i + RECOVERY_CAP, len(v))] <= base * 1.05)[0]
    return {
        "VIX_Before": base,
        "VIX_Peak20": v[peak_i],
        "VIX_Change": v[peak_i] / base - 1,
        "VIX_Settle_Days": (peak_i - i0 + int(back[0])) if len(back) else np.nan,
    }


def crude_events(brent):
    """Auto-detect Brent shocks: >= +/-20% move within 20 trading days."""
    chg = brent["Close"].pct_change(H).values
    rows, last = [], {"up": -10**9, "down": -10**9}
    for k in range(1, len(chg)):
        x, prev = chg[k], chg[k - 1]
        if np.isnan(x) or np.isnan(prev):
            continue
        side = "up" if x >= CRUDE_MOVE > prev else "down" if x <= -CRUDE_MOVE < prev else None
        if side is None or k - last[side] <= CRUDE_COOLDOWN:
            continue
        last[side] = k
        d, price = brent.index[k], brent["Close"].iloc[k]
        # Brent settles after Indian market close, so India reacts next session
        rows.append({
            "Date": (d + pd.Timedelta(days=1)).date().isoformat(),
            "Category": "Crude Spike" if side == "up" else "Crude Crash",
            "News_Tone": "Negative" if side == "up" else "Positive",
            "Surprise": "Yes",
            "Headline": f"Brent {x:+.0%} in 20 days to ${price:.0f}",
            "Context": "Auto-detected from Brent price data (higher crude hurts India as a net importer).",
        })
    return rows


def main():
    nifty, sensex, vix = load("NIFTY"), load("SENSEX"), load("INDIAVIX")
    brent = load("BRENT")
    with open(HERE / "events.csv", newline="") as f:
        events = list(csv.DictReader(f))
    events += crude_events(brent)
    events.sort(key=lambda e: e["Date"])

    watch = pd.read_csv(HERE / "watchlist.csv")
    stocks = {s: load(s, clean=True) for s in watch["NSE_Code"]}
    missing = [s for s, d in stocks.items() if d is None]
    if missing:
        print("No price data for:", ", ".join(missing))

    nifty_rows, stock_rows = [], []
    for k, ev in enumerate(events, start=1):
        ev_id = f"E{k:03d}"
        # Yahoo's Nifty history starts Sep-2007; use Sensex for earlier events
        idx, idx_name = (nifty, "Nifty") if pd.Timestamp(ev["Date"]) > nifty.index[25] else (sensex, "Sensex")
        i0 = pos_on_or_after(idx.index, ev["Date"])
        m = window_metrics(idx, i0)
        if m is None:
            print("Skipped (no market data):", ev["Date"], ev["Headline"])
            continue
        t0 = idx.index[i0]
        row = {"Event_ID": ev_id, **ev, "Trading_Day": t0.date().isoformat(), "Index": idx_name}
        row.update(m)
        row.update(vix_metrics(vix, t0))
        nifty_rows.append(row)

        for s, df in stocks.items():
            if df is None:
                continue
            j0 = pos_on_or_after(df.index, t0)
            if j0 is None or (df.index[j0] - t0).days > 4:
                continue
            jumps = JUMPS[s]
            if ((jumps >= j0 - 21) & (jumps <= j0 + 60)).any():
                continue  # split/demerger-like jump inside the window would distort results
            sm = window_metrics(df, j0)
            if sm is None:
                continue
            stock_rows.append({
                "Stock": s, "Event_ID": ev_id, "Trading_Day": t0.date().isoformat(),
                "Day0": sm["Day0"], "Total20": sm["Total20"], "MaxFall20": sm["MaxFall20"],
                "MaxRise20": sm["MaxRise20"], "Recovery_Days": sm["Recovery_Days"],
                "Post60": sm["Post60"], "Reward_EventDay": sm["Reward_EventDay"],
                "Risk_EventDay": sm["Risk_EventDay"],
                "Ret20_EventDay": sm["Ret20_EventDay"],
            })

    pd.DataFrame(nifty_rows).to_csv(OUT / "nifty_impact.csv", index=False)
    sdf = pd.DataFrame(stock_rows)
    order = {s: i for i, s in enumerate(watch["NSE_Code"])}
    sdf = sdf.sort_values(["Stock", "Trading_Day"], key=lambda col: col.map(order) if col.name == "Stock" else col)
    sdf.to_csv(OUT / "stock_events.csv", index=False)

    # ---- daily market table for crude-oil ranges and FII/DII lookups ----
    d = pd.DataFrame(index=nifty.index)
    d["Nifty"] = nifty["Close"]
    d["Nifty_Ret1"] = d["Nifty"].pct_change()
    d["Nifty_Next1"] = d["Nifty"].shift(-1) / d["Nifty"] - 1
    d["Nifty_Same5"] = d["Nifty"].pct_change(5)
    d["Nifty_Next5"] = d["Nifty"].shift(-5) / d["Nifty"] - 1
    d["Nifty_Next20"] = d["Nifty"].shift(-20) / d["Nifty"] - 1
    d["VIX"] = vix["Close"].reindex(d.index, method="ffill") if vix is not None else np.nan
    b = brent["Close"].reindex(d.index, method="ffill")
    d["Brent"] = b
    d["Brent_5D"] = b.pct_change(5)
    d["Brent_20D"] = b.pct_change(20)
    ind_of = dict(zip(watch["NSE_Code"], watch["Industry"]))
    for group, inds in INDUSTRY_GROUPS.items():
        same, nxt = [], []
        for s in stocks:
            if stocks[s] is None or ind_of.get(s) not in inds:
                continue
            c = stocks[s]["Close"]
            bad = pd.Series(0, index=c.index)
            bad.iloc[JUMPS[s]] = 1
            r_same = c.pct_change(5).where(bad.rolling(5, min_periods=1).sum() == 0)
            r_next = r_same.shift(-5)
            same.append(r_same.reindex(d.index))
            nxt.append(r_next.reindex(d.index))
        if same:
            d[f"{group}_Same5"] = pd.concat(same, axis=1).mean(axis=1)
            d[f"{group}_Next5"] = pd.concat(nxt, axis=1).mean(axis=1)
    d = d[d["Brent"].notna()]
    d.to_csv(OUT / "market_daily.csv", index_label="Date")
    print(f"{len(nifty_rows)} events, {len(stock_rows)} stock-event rows, {len(d)} market days")


if __name__ == "__main__":
    main()
