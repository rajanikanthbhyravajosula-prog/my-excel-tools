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


# Crisis periods: (name, search from, search to, cause). Inside the window the code finds the
# lowest close (bottom), then the highest close before it (peak), then the date the peak is regained.
CRISES = [
    ("Asian financial crisis 1997-98", "1997-08-01", "1998-12-31",
     "Currency crises in Thailand, Korea and Indonesia; Pokhran sanctions and Russia default in 1998."),
    ("Dot-com bust and 9/11 2000-01", "2000-01-01", "2001-12-31",
     "Technology bubble burst, Ketan Parekh scandal, UTI US-64 freeze and the 9/11 attacks."),
    ("May 2006 correction", "2006-04-01", "2006-07-31",
     "Global commodity and emerging-market sell-off after a strong rally; FIIs sold heavily."),
    ("Global financial crisis 2008-09", "2007-12-01", "2009-06-30",
     "US subprime crisis, Lehman collapse and a global recession."),
    ("Euro debt crisis and US downgrade 2010-11", "2010-10-01", "2011-12-31",
     "European sovereign debt crisis, India's high inflation and rate hikes, US credit downgrade."),
    ("Taper tantrum and rupee crisis 2013", "2013-05-01", "2013-09-30",
     "Fed taper hint triggered FII outflows; the rupee fell to record lows."),
    ("China slowdown and commodity crash 2015-16", "2015-03-01", "2016-03-31",
     "Yuan devaluation, Chinese market crash and the collapse in crude and metal prices."),
    ("IL&FS / NBFC crisis 2018", "2018-08-01", "2018-11-30",
     "IL&FS defaults froze NBFC funding; crude above $85 and a weak rupee."),
    ("COVID-19 crash 2020", "2020-01-01", "2020-04-30",
     "Global pandemic and national lockdown; fastest bear market in history."),
    ("Inflation, rate hikes and Ukraine war 2021-22", "2021-10-01", "2022-07-31",
     "Global inflation, aggressive Fed hikes, Russia-Ukraine war and record FII selling."),
    ("FII exodus and US tariff correction 2024-25", "2024-09-01", "2025-04-30",
     "Record FII selling, slowing earnings and US tariff shock."),
]
YEAR = 250  # trading days in a year
MAX_1Y_GAIN = 10.0  # a 1-year gain above +1000% is treated as a data fault

def panic_meter(nifty, vix, usdinr, stocks):
    """Daily raw Panic Meter components (scored 0-100 by formulas in the workbook) and forward returns."""
    p = pd.DataFrame(index=nifty.index)
    c = nifty["Close"]
    p["VIX"] = vix["Close"].reindex(p.index, method="ffill")
    p["Drawdown"] = c / c.rolling(YEAR, min_periods=20).max() - 1
    p["Return20"] = c.pct_change(20)
    below = []
    for s, df in stocks.items():
        if df is None:
            continue
        sc = df["Close"]
        ok = pd.Series(1.0, index=sc.index)
        for j in JUMPS[s]:  # ignore 200 days after an unadjusted split/demerger
            ok.iloc[j:j + 200] = np.nan
        ma = sc.rolling(200).mean()
        below.append(((sc < ma).astype(float) * ok).where(ma.notna()).reindex(p.index))
    b = pd.concat(below, axis=1)
    p["Breadth"] = b.mean(axis=1).where(b.notna().sum(axis=1) >= 30)
    p["Rupee"] = usdinr["Close"].reindex(p.index, method="ffill").pct_change(20)
    p["Nifty_Next60"] = c.shift(-60) / c - 1
    p["Nifty_Next250"] = c.shift(-YEAR) / c - 1
    return p


def crisis_periods(nifty, sensex, vix, stocks):
    rows, srows = [], []
    for name, start, end, cause in CRISES:
        idx, idx_name = (nifty, "Nifty") if pd.Timestamp(start) > nifty.index[0] else (sensex, "Sensex")
        c = idx["Close"]
        w = c[start:end]
        if len(w) < 20:
            continue
        trough_d = w.idxmin()
        peak_d = w[:trough_d].idxmax()
        peak, trough = c[peak_d], c[trough_d]
        after = c[trough_d:]
        rec = after[after >= peak]
        rec_d = rec.index[0] if len(rec) else None
        ip, it = c.index.get_loc(peak_d), c.index.get_loc(trough_d)
        down10 = c[peak_d:trough_d][c[peak_d:trough_d] <= peak * 0.9]
        up20 = after[after >= trough * 1.2]

        def fwd(d):
            if d is None:
                return np.nan
            i = c.index.get_loc(d)
            return c.iloc[i + YEAR] / c.iloc[i] - 1 if i + YEAR < len(c) else np.nan

        d10 = down10.index[0] if len(down10) else None
        d20 = up20.index[0] if len(up20) else None
        row = {
            "Crisis": name, "Cause": cause, "Index": idx_name,
            "Peak_Date": peak_d.date().isoformat(), "Peak": peak,
            "Trough_Date": trough_d.date().isoformat(), "Trough": trough,
            "Fall": trough / peak - 1, "Days_Down": it - ip,
            "Recovery_Date": rec_d.date().isoformat() if rec_d is not None else None,
            "Days_To_Recover": (c.index.get_loc(rec_d) - it) if rec_d is not None else np.nan,
            "Ret1Y_After_Trough": fwd(trough_d),
            "Buy10_Date": d10.date().isoformat() if d10 is not None else None,
            "Buy10_FurtherFall": trough / c[d10] - 1 if d10 is not None else np.nan,
            "Buy10_Ret1Y": fwd(d10),
            "BuyTrough_Ret1Y": fwd(trough_d),
            "Buy20Up_Date": d20.date().isoformat() if d20 is not None else None,
            "Buy20Up_Ret1Y": fwd(d20),
        }
        v = vix["Close"][peak_d:trough_d + pd.Timedelta(days=30)] if vix is not None else pd.Series(dtype=float)
        row["VIX_Peak"] = v.max() if len(v) else np.nan
        rows.append(row)

        for s, df in stocks.items():
            if df is None:
                continue
            sc = df["Close"]
            jp = sc.index.searchsorted(peak_d)
            jt = sc.index.searchsorted(trough_d)
            if jp >= len(sc) or jt >= len(sc) or (sc.index[jp] - peak_d).days > 5 or jp < 1:
                continue
            if ((JUMPS[s] >= jp) & (JUMPS[s] <= jt + YEAR)).any():
                continue
            s_peak = sc.iloc[jp]
            s_low = sc.iloc[jp:jt + 61].min()
            s_after = sc.iloc[jt:]
            s_rec = s_after[s_after >= s_peak]
            srows.append({
                "Stock": s, "Crisis": name,
                "Fall_Peak_To_Trough": sc.iloc[jt] / s_peak - 1,
                "Max_Fall": s_low / s_peak - 1,
                "Vs_Index": (sc.iloc[jt] / s_peak - 1) - (trough / peak - 1),
                "Days_To_Recover": (sc.index.get_loc(s_rec.index[0]) - jt) if len(s_rec) else np.nan,
                "Recovered": 1 if len(s_rec) else 0,
                "Ret1Y_After_Trough": sc.iloc[jt + YEAR] / sc.iloc[jt] - 1 if jt + YEAR < len(sc) else np.nan,
            })
            if srows[-1]["Ret1Y_After_Trough"] > MAX_1Y_GAIN:  # more than 10x in a year: treat as a data fault
                srows[-1]["Ret1Y_After_Trough"] = np.nan
    return pd.DataFrame(rows), pd.DataFrame(srows)


def main():
    nifty, sensex, vix = load("NIFTY"), load("SENSEX"), load("INDIAVIX")
    usdinr = load("USDINR")
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
    panic = panic_meter(nifty, vix, usdinr, stocks)

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
    for col in ["Nifty_Next60", "Nifty_Next250", "Drawdown", "Return20", "Breadth", "Rupee"]:
        d[col] = panic[col]
    d = d[d["Brent"].notna()]
    d.to_csv(OUT / "market_daily.csv", index_label="Date")
    cp, cs = crisis_periods(nifty, sensex, vix, stocks)
    cp.to_csv(OUT / "crisis_periods.csv", index=False)
    cs.to_csv(OUT / "crisis_stocks.csv", index=False)
    print(f"{len(cp)} crisis periods, {len(cs)} stock-crisis rows")
    print(f"{len(nifty_rows)} events, {len(stock_rows)} stock-event rows, {len(d)} market days")


if __name__ == "__main__":
    main()
