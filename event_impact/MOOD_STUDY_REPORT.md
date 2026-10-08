# Market Moods Study — Fear, Greed and How to Use Them in Our Trading

*SLNS Vaibhavam — trading research notes. Prepared 8 October 2026. Data up to 6 October 2026.*

> **In one line:** our software already calculates targets from mathematics; this study adds a measured
> "emotion layer" — how frightened or greedy the market is, how long each mood lasts, what usually
> follows it — so that targets, stop-losses, position sizes and the timing of entry and exit can be
> adjusted with evidence instead of feeling.

This is research on historical tendencies, not investment advice. History shows probabilities, never certainties.

---

## Contents

1. [How this study came about (our discussion)](#1-how-this-study-came-about-our-discussion)
2. [What we built](#2-what-we-built)
3. [Data and method](#3-data-and-method)
4. [What the study says about moods — main findings](#4-what-the-study-says-about-moods--main-findings)
5. [The mood framework: zones, signals and time frames](#5-the-mood-framework-zones-signals-and-time-frames)
6. [Do mood rules actually make money? (back-tests)](#6-do-mood-rules-actually-make-money-back-tests)
7. [How to implement this in our trading calculations](#7-how-to-implement-this-in-our-trading-calculations)
8. [When to be away, when to enter, when to cash out — the playbook](#8-when-to-be-away-when-to-enter-when-to-cash-out--the-playbook)
9. [Where we are now (October 2026)](#9-where-we-are-now-october-2026)
10. [Other thoughts and cautions](#10-other-thoughts-and-cautions)
11. [Limitations and next steps](#11-limitations-and-next-steps)
12. [Sources](#12-sources)

---

## 1. How this study came about (our discussion)

| Step | What you asked | What we did |
|---|---|---|
| 1 | "Can we add emotions to our software?" | Proposed three ideas; you chose the trading direction. |
| 2 | Software is purely mathematical; give *marks* to the emotions of trends: crude oil moves, positive / negative news, repo rate changes, what to adjust in short and long-term targets | Designed a sentiment layer: event register → mood score → target adjuster. |
| 3 | Impact on Nifty, on our watchlist, most and least impacted shares, FII/DII, before/during/after events, settling time, who made money, elections, budgets, R:R | Built the Event Impact Analyzer on 20+ years of real price data and your 205-stock watchlist. |
| 4 | Recessions, blackouts, government falls, bank collapses, pandemics | Added these event types, a Crisis Periods study (peak → bottom → recovery) and extended history to 1997. |
| 5 | Famous studies of analyst predictions; panic level vs outcomes | Added Forecast Studies, a Forecast Tracker and a daily **Panic Meter** (0–100). |
| 6 | Splits/bonuses/buybacks, government policy changes, portfolio-manager strategies, which countries recovered, US policy, Iran–US war, Russia–Ukraine, current phase | Verified corporate-action handling, added 2025–26 events, the ongoing 2026 crisis, 12 world markets and documented strategy shifts. |
| 7 | Overall conclusions, implementation, mood changes, time frames, when to enter / stay away | This report, plus new back-tests of mood rules (section 6). |

---

## 2. What we built

All files are in the `event_impact/` folder of the `my-excel-tools` repository.

| File | Purpose |
|---|---|
| `Event_Impact_Analyzer.xlsx` | The workbook (24 sheets, about 162,000 live formulas, zero errors). Start with *How_To_Use*. |
| `events.csv` | 220 events from 1997 to 2026 with news context (add new events here). |
| `watchlist.csv` | Your 205 watchlist stocks with industry and market-cap group (no holdings or prices). |
| `fii_dii_daily.csv` | Your FII/DII tracker data (Jun–Oct 2026). |
| `download_data.py` → `analyze.py` → `build_workbook.py` | Re-run these three to refresh everything with new prices and events. |
| `mood_backtest.py` | The mood duration, transition and back-test calculations used in sections 5–6 of this report. |
| `MOOD_STUDY_REPORT.md` | This report. |

**Main workbook sheets**

| Sheet | Answers |
|---|---|
| Settings | Yellow cells: fear/greed thresholds, points scale, target weights. |
| Target_Adjuster | Pick a stock + an upcoming event → adjusted short/long-term target, stop-loss, R:R. |
| Panic_Meter | Today's panic score and what Nifty did after each panic level. |
| Category_Summary | Average behaviour per event type; positive vs negative news; fear vs greed; points. |
| Crisis_Periods | 12 crises since 1997 incl. the ongoing 2026 one, and **WHERE ARE WE NOW**. |
| Top_Lists | Most resilient / most impacted stocks; crisis defenders; recovery leaders. |
| Stock_Scorecard, Crisis_Scorecard | Every watchlist stock's behaviour in fear, greed and crises. |
| Sector_Impact, Crisis_Sectors | Industry heat maps by event type and by crisis. |
| Global_Recovery | India vs 11 world markets: fall, recovery time, rebound. |
| Strategy_Shifts | Measured sector rotation + documented fund-manager strategy changes. |
| Crude_Ranges | At what crude move / price level the market reacts. |
| FII_DII | Your flow data linked to Nifty moves. |
| Forecast_Studies, Forecast_Tracker | Research on expert accuracy; log and score any analyst's predictions. |
| Event_Calendar, Nifty_Impact, Stock_Event_Data, Crisis_Stock_Data, Global_Crisis_Data, Market_Daily | Raw measurements the formulas read. |

---

## 3. Data and method

- **Prices:** Yahoo Finance daily prices, adjusted for splits, bonuses and dividends. Nifty from Sep-2007, Sensex from Jul-1997 (used for earlier events), India VIX from 2008, Brent from 2007, USD/INR from 2003, 11 world indices from 1997, and all 205 watchlist stocks.
- **Corporate actions checked:** bonuses (Reliance, TCS, Infosys, HDFC Bank, Bajaj Finance), splits and buybacks show no artificial jumps. Demergers (e.g. Tata Motors 2025, Raymond, ABFRL) cannot be adjusted cleanly, so any event window containing such a jump is skipped. One-day data spikes are removed.
- **Event windows** (trading days, T0 = event day):
  - BEFORE: run-up over 20 and 5 days before the event.
  - DURING: event-day move and intraday worst point.
  - AFTER: next 5, 20 and 60 days; worst fall and best rise within 20 days.
  - SETTLING: days until price regains the pre-event close; days until India VIX calms.
  - R:R: buying 5 days before, on the day, or 5 days after — best gain vs worst dip over 20 days.
- **Market mood of an event:** *Fear* if Nifty fell 5% or more within 20 days, *Greed* if it rose 5% or more, otherwise *Calm*.
- **Panic Meter (daily, 0–100):** average of five scores, each scaled between a calm and a fear value: India VIX (12 → 40), Nifty fall from its 1-year high (0 → −30%), Nifty 20-day return (+5% → −15%), share of watchlist stocks below their 200-day average (20% → 90%), and USD/INR 20-day change (−1% → +4%). Fixed scales mean no hindsight goes into the score.
- **Crises:** in each crisis window the lowest close is the bottom, the highest close before it is the peak, recovery is the first close back at the peak.

---

## 4. What the study says about moods — main findings

### 4.1 Events: the market reacts to surprise, not to news

| News | Events | Avg event-day move | Avg 20-day move | Nifty higher after 20 days |
|---|---|---|---|---|
| Positive and a surprise | 34 | +1.8% | **+2.4%** | 65% |
| Positive but expected | 30 | +0.7% | +0.1% | 37% |
| Negative and a surprise | 79 | **−1.9%** | −1.0% | 49% |
| Negative but expected | 23 | +0.3% | −0.5% | 43% |

**Lesson:** expected news is already in the price. The points system must score the *gap between actual and expected*, not the news itself.

### 4.2 Fear vs greed events — who made the money?

| Mood | Events | Avg worst fall | Avg days to recover | Buy 5 days before | Buy on the day | Buy 5 days after | Winner |
|---|---|---|---|---|---|---|---|
| **Fear** | 80 | −11.0% | **63** | −6.6% | −4.1% | **−1.8%** | Sellers before the event / cash holders |
| **Greed** | 60 | −1.7% | 2 | +7.2% | **+7.4%** | +5.1% | Buyers on the event day |
| Calm | 80 | −2.5% | 24 | +0.4% | +0.4% | 0.0% | Early buyers |

**Lesson:** in a fear event, waiting is cheap and buying early is expensive; in a greed event, hesitation costs money.

### 4.3 Event types — the "personality" of each kind of news

| Event type | Events | Avg 20-day move | Avg worst fall | Days to recover | Best entry (history) | Signal |
|---|---|---|---|---|---|---|
| Central election | 7 | +4.2% | −5.8% | 17 | Buy before | Bullish |
| Exit poll | 2 | +4.8% | −2.0% | 1 | Buy before | Bullish (low confidence) |
| State election | 19 | +0.3% | −3.2% | 24 | Buy before | Neutral |
| Government crisis | 5 | +7.3% | −6.2% | 9 | **Buy 5 days after** | Bullish (rebounds) |
| Union Budget | 29 | **−1.0%** | −5.8% | 40 | Buy on the day | **Bearish** (only 34% higher after 20 days) |
| RBI rate hike | 27 | −0.6% | −5.1% | 22 | Buy 5 days after | Neutral |
| RBI rate cut | 28 | +1.8% | −4.8% | 21 | Buy 5 days after | Bullish |
| Crude spike | 13 | +3.0% | −4.0% | 12 | Buy 5 days after | Bullish (dip gets bought) |
| Crude crash | 9 | +1.1% | −4.7% | 52 | Buy on the day | Mixed (often a slowdown signal) |
| Bank collapse | 8 | +2.0% | −6.4% | 16 | Buy 5 days after | Bullish after the shock |
| Pandemic | 10 | −1.8% | **−10.8%** | **55** | Buy 5 days after | Bearish, high risk |
| Global shock | 14 | **−3.7%** | **−10.8%** | **56** | Buy on the day (least bad) | Bearish, highest risk |
| Geopolitical / war | 19 | −1.2% | −4.9% | 46 | Buy on the day | Mildly bearish |
| Domestic shock (scams, defaults) | 4 | −7.4% | −8.2% | 109 | — | Bearish (low confidence) |
| Policy reform | 7 | +1.7% | −2.6% | 12 | Buy 5 days after | Bullish |
| US policy | 7 | +1.8% | −2.2% | 36 | Buy before | Mildly bullish |

**Lessons:**
- **Wars and political crises frighten more than they hurt.** Unless oil supply is hit, the damage is short.
- **Global shocks, pandemics and domestic financial scandals** are the dangerous ones: deep (−11%) and slow to heal (2–5 months).
- **Budgets have been a "sell the news" event** in recent years; do not pre-position for a budget rally.

### 4.4 The Panic Meter — fear is a buying opportunity *if you can wait*

Daily history since 2007:

| Panic zone | Share of days | Nifty next 20 days | Nifty next 1 year | Higher after 1 year |
|---|---|---|---|---|
| 0–20 Complacent / greedy | 46% | +0.4% | +7.6% | 76% |
| 20–40 Calm | 33% | +0.8% | +11.1% | 78% |
| 40–60 Worried | 13% | +1.9% | +14.0% | 81% |
| 60–80 Fear | 5% | +1.7% | +23.7% | 79% |
| **80–100 Panic** | **2%** | +0.2% | **+58.7%** | **96%** |

**Lesson:** the more frightened the market, the better the following year. But note the next-20-day column: **the short term after panic is still rough**. Fear pays the patient investor, not the short-term trader.

### 4.5 Crises (recessions and bear markets) since 1997

| Crisis | Fall | Days falling | Days to recover | Panic peak | 1 year after bottom |
|---|---|---|---|---|---|
| Asian crisis 1997–98 | −39% | 293 | 175 | – | +72% |
| Dot-com & 9/11 2000–01 | −56% | 402 | 566 | – | +16% |
| May 2006 correction | −29% | 25 | 85 | – | +58% |
| Global financial crisis 2008 | **−60%** | 198 | 496 | 100 | **+98%** |
| Euro debt crisis 2010–11 | −28% | 271 | 457 | 84 | +32% |
| Taper tantrum 2013 | −15% | 71 | 34 | 76 | +55% |
| China slowdown 2015–16 | −23% | 241 | 255 | 70 | +29% |
| IL&FS / NBFC 2018 | −15% | 39 | 114 | 66 | +19% |
| COVID-19 2020 | −38% | 47 | 158 | 100 | **+91%** |
| Inflation & Ukraine war 2021–22 | −17% | 166 | 108 | 66 | +23% |
| FII exodus & US tariffs 2024–25 | −16% | 109 | 206 | 49 | +10% |
| **Iran war & oil shock 2026 (ongoing)** | **−15% so far** | 57 to the low | **not yet** | 77 | – |
| **Average of recovered crises** | **−30%** | **169** | **241** | **76** | **+46%** |

- Peak to full recovery takes about **20 months** on average (8 months falling plus 12 months recovering).
- **Buy early or wait?** Buying at the first 10% fall meant suffering a further −22% on average. **Waiting for a 20% bounce off the bottom paid more in 7 of 10 crises.** The bottom itself cannot be timed.

### 4.6 Stocks and sectors in fear

- **Defenders** (fall least, recover fastest in fear): Abbott India, Nestle, Alkem, HUL, Sun Pharma, Colgate, DMart, GSK Pharma, Britannia, Apollo Hospitals, Marico. These are mostly **pharma and FMCG**, with Event Beta of 0.0–0.6.
- **Most hit** (fall 2 to 2.5 times as much as Nifty): Walchandnagar, NCC, Satin Creditcare, Century Extrusions, Bandhan Bank, Quess, Welspun Corp. These are mostly **small caps, construction, metals and lenders**, with Event Beta of 1.6–2.2.
- **Recovery leaders after crisis bottoms:** **steel, metals, cement, tyres, auto** (JK Lakshmi Cement, Jindal Steel, Eicher, Granules, Aurobindo, JK Tyre). They often rose 2 to 2.7 times in the year after the bottom.
- **Rotation pattern:** in the fall, FMCG, pharma and IT held up while steel, metals and banks were hit; in the recovery, the order reversed. **Money rotates from defence to high beta at the bottom.**

### 4.7 Crude oil ranges

- Weekly Brent moves between −2% and +2% barely matter.
- A **crude collapse of more than 15% in 5 days** was bad for Nifty (−3.4% the same week) because it usually signals a global slowdown.
- **Brent above $120:** Nifty was higher 20 days later only 32% of the time. **Below $60:** about two-thirds of the time.
- In weeks when Brent jumped 10% or more, **paints (−1.4% vs Nifty) and aviation/transport (−0.8%)** lagged most, while **pharma (+1.1%) and IT (+0.8%)** did best (the weak rupee helps exporters).

### 4.8 Other countries

- **Most resilient markets:** UK, USA and Germany (average fall about 20–25%); every crisis was eventually recovered.
- **India:** falls are larger (about −29% on average) but rebounds are strong (**+46% in the year after the bottom**), and every crisis before 2026 was recovered, usually faster than the US.
- **Never recovered:** **China (Shanghai) is still 30% below its 2008 peak** and below its 2015 peak.
- **In 2026:** Korea and Taiwan did not fall at all (AI and chip boom) and the US has already recovered, while India, Hong Kong, China and Indonesia have not. This is the oil-importer pattern.

### 4.9 Experts and fund managers

- Expert forecasts are about a coin toss: market gurus scored 47% accuracy across 6,582 forecasts. Economists missed the vast majority of 153 recessions. Analysts' earnings forecasts ran almost 2 times too high over 25 years.
- **Forecasters cut targets after the fall, not before.** In April 2026, Indian brokerages cut their average Nifty target from 29,899 to 28,748, two months into the war.
- **Fund-manager fear is a contrarian signal.** BofA's cash rule treats cash above about 5% as a buy signal. In March–April 2026 managers went to cash and oil; by May they made the biggest jump into equities since 2001.
- **Structural change in India:** since March 2025, domestic institutions own more of the market than foreigners. SIP money now cushions foreign selling, which is one reason the 2026 fall stayed shallow despite record foreign outflows.

---

## 5. The mood framework: zones, signals and time frames

### 5.1 The five moods

| Zone | Panic score | What it looks like |
|---|---|---|
| 😎 Complacent / greedy | 0–20 | Near highs, VIX low, most stocks above their 200-day average, rupee stable |
| 🙂 Calm | 20–40 | Normal dips, mixed breadth |
| 😟 Worried | 40–60 | 10–15% off highs, breadth weak, rupee slipping |
| 😨 Fear | 60–80 | Fast falls, VIX up sharply, most stocks broken |
| 😱 Panic | 80–100 | Crash days, VIX 35+, 80–90% of stocks broken, rupee under pressure |

### 5.2 How long does each mood last? (5-day smoothed panic score, 2007–2026)

| Zone | Episodes | Average length | Median | Longest |
|---|---|---|---|---|
| Complacent | 77 | 28 trading days | 17 | 263 (more than a year) |
| Calm | 115 | 14 | 10 | 83 |
| Worried | 57 | 11 | 9 | 34 |
| Fear | 25 | 9 | 6 | 28 |
| Panic | 7 | 15 | 7 | 48 |

**Greed lasts long; fear is short and violent.** Complacency can last more than a year, while fear usually burns out in 1–6 weeks.

Major fear episodes (panic 60+, short gaps merged):

| Episode | Length | Panic peak | Nifty 1 year after the panic peak |
|---|---|---|---|
| Mar–Apr 2008 | 20 days | 85 | **−32%** (first fear wave of a big bear market came too early) |
| May–Jul 2008 | 46 days | 88 | +18% |
| Aug–Dec 2008 | 77 days | 100 | +36% |
| Jan–Apr 2009 | 55 days | 97 | **+82%** |
| Aug–Oct 2011 | 40 days | 84 | +13% |
| Nov 2011–Jan 2012 | 33 days | 79 | +32% |
| Aug–Sep 2013 | 13 days | 76 | +55% |
| Mar–Apr 2020 | 30 days | 100 | **+78%** |
| Mar–Apr 2026 | 8 days | 77 | not yet known |

**Time frames to plan with:**
- **Event mood:** fear events need about **2–3 months** to recover (63 days on average); greed events recover in a couple of days.
- **India VIX** settles in 2–7 weeks after most events (about 50 days after global shocks).
- **Crisis mood:** about **8 months falling and 12 months recovering**.

### 5.3 How moods change: what usually comes 20 days later

| Today ↓ / In 20 days → | Complacent | Calm | Worried | Fear | Panic |
|---|---|---|---|---|---|
| Complacent | **66%** | 27% | 6% | 1% | 0% |
| Calm | 40% | 42% | 15% | 3% | 1% |
| Worried | 14% | 47% | 25% | 12% | 2% |
| Fear | 0% | 24% | 30% | 29% | 17% |
| Panic | 0% | 0% | 39% | 25% | **35%** |

- Moods are **sticky**: complacency usually stays complacent, and panic often stays panic for another month.
- **From "Worried", 61% of the time the market calms down within a month**, and only 14% of the time it turns to fear or panic.
- There is **no direct jump from panic to complacency**: the recovery always passes through "worried" first.

### 5.4 How to identify a mood change: the early-warning checklist

| Signal | Mood worsening (greed → fear) | Mood improving (fear → calm) |
|---|---|---|
| Panic score | Rises above 40, then 60 | Falls back below 50 after being 70+ ("fear fading") |
| India VIX | Jumps 30–50% in a week | Falls back to its pre-event level |
| Breadth (watchlist stocks above 200-day average) | Falls below 50% | Climbs back above 50% |
| Nifty vs 200-day average | Closes below it | Closes back above it |
| Distance from the low | – | 20% above the crisis low = recovery confirmed |
| Rupee | 2%+ fall in 20 days | Stabilises |
| Crude | +20% in 20 days, or above $110–120 | Back below $90 |
| FII flows | Several days of heavy selling (more than ₹3,000 cr) | Selling slows; DIIs keep absorbing |
| Fund-manager cash (BofA) | Below 4% (greed) | Above 5% (fear, a contrarian buy) |
| News | Negative *surprise* | Bad news stops pushing prices lower |

**Important:** "fear fading" alone is not a buy signal in a new bear market. In early 2008 it fired three times, and Nifty was still 25–50% lower a year later. It worked in 2009, 2012, 2013 and 2020, when the fall was already 25% or more.

---

## 6. Do mood rules actually make money? (back-tests)

These are tests on the Nifty price index from Sep-2007 to Oct-2026, without dividends or costs. "Cash" earns an assumed 6% a year, like a liquid fund or FD.

| Rule | Annual return | Worst fall | Average invested |
|---|---|---|---|
| Buy and hold 100% | **9.1%** | −59.9% | 100% |
| **Trend rule:** 100% when Nifty is above its 200-day average, otherwise cash | 8.5% | **−18.5%** | 68% |
| **Half-trend rule:** 100% above the 200-day average, otherwise 50% | **9.1%** | **−33.9%** | 84% |
| Trend + panic: above the 200-day average or panic 70+ → 100%, else 50% | 8.9% | −47.0% | 86% |
| Mood scaling only (40% when complacent … 100% in fear), cash earning nothing | 5.6% | −53.4% | 56% |

Monthly SIP test (from 2007):

| SIP style | Value per ₹1 invested today |
|---|---|
| Plain monthly SIP | ₹2.78 |
| **SIP plus double extra in any month the panic score reached 60** | **₹3.21** (+15%) |

**What the back-tests teach us:**
1. **Mood alone is not a timing system.** Cutting exposure when the market is greedy lost money, because complacency lasts a long time and markets keep rising.
2. **The trend rule protects capital.** With cash in a liquid fund, the 200-day rule kept almost the whole return and **cut the worst fall from −60% to −19%**. That is the "when to stay away" rule.
3. **Fear is best used to add money in stages**, not to go all-in. The fear-boosted SIP beat the plain SIP by 15%.
4. Buying on panic inside a fresh downtrend increases risk: the 2008 example. **Panic and a deep fall (25%+) together** is the strong signal.

---

## 7. How to implement this in our trading calculations

Think of the system as **three layers on top of the existing mathematical model**:

```
Layer 1  Mathematical model (existing)  →  Base target, base stop-loss
Layer 2  Regime filter (trend)          →  How much capital is allowed in the market
Layer 3  Mood & event layer (this study) →  Adjust targets, stop-loss, size and timing
```

### 7.1 Daily inputs (all available free, after 6 pm)

Nifty close and 200-day average, India VIX, USD/INR, Brent, FII/DII net flows, the share of watchlist stocks above their 200-day average, and the calendar of upcoming events. The workbook's Market_Daily and Panic_Meter sheets already calculate the score from these.

### 7.2 Formulas to add to our software

**a) Regime (trend)**
```
Trend% = Nifty / 200-day average − 1
Regime = "Up" if Trend% > 0 else "Down"
```

**b) Mood score**
```
Panic = average of the five 0–100 scores (VIX, drawdown, 20-day return, breadth, rupee)
Zone  = Complacent <20 | Calm 20–40 | Worried 40–60 | Fear 60–80 | Panic 80+
```

**c) Capital allowed in equities (the half-trend rule, with fear add-ons)**
```
Regime Up                          → 100% of the trading capital
Regime Down and Panic < 60         → 50% (the rest in a liquid fund)
Regime Down and Panic ≥ 60         → add back in 3–4 tranches (each after a further 5% fall
                                      or 2 weeks), favouring defenders first
Recovery confirmed (Nifty 20% off
the low, or back above 200-day)    → 100%, rotate into recovery leaders
```

**d) Event adjustment of targets** (the Target_Adjuster sheet already does this)
```
Expected event move = stock's own average 20-day move for this event type
                      (or Event Beta × Nifty's average, if fewer than 5 past events)
Adjusted short-term target = Base ST target × (1 + Expected event move × 1.00)
Adjusted long-term target  = Base LT target × (1 + Expected event move × 0.25)
```

**e) Mood adjustment of long-term targets** (optional; based on the 1-year returns in section 4.4)
```
Long-term expected return after each zone: Complacent 7.6% | Calm 11% | Worried 14% | Fear 24% | Panic 59%
Mood factor = (Zone 1-year return − 11.7% overall average) × 0.25
LT target   = Base LT target × (1 + Mood factor)
→ about −1% when complacent, +3% in fear, +12% in panic
```

**f) Stop-loss and position size**
```
Stop-loss = CMP × (1 + stock's average worst 20-day fall for the upcoming event type × buffer)
          (wider when VIX is high, so normal swings don't stop you out)
Position size (shares) = Rupees you accept to lose per trade / (CMP − Stop-loss)
→ in fear the stop-loss is wider, so the position automatically becomes smaller
```

**g) Event points** (in Category_Summary)
```
Direction points (−10 to +10) = average 20-day move after the event type ÷ 0.5%
Risk points (0 to 10)         = average worst fall ÷ 1%
Use the "Surprise?" field: halve the points if the outcome was fully expected
```

### 7.3 Routine

| When | Action |
|---|---|
| **Daily** (after 6 pm) | Update the FII/DII tracker; check the Panic score, Nifty vs its 200-day average and the zone. |
| **Weekly** | Re-run the three scripts; review Top_Lists and your holdings' Event Beta; check upcoming events in the calendar. |
| **Before every known event** (budget, RBI, elections, Fed) | Run Target_Adjuster for each holding; follow the event's "best entry" from section 4.3. |
| **Monthly** | Log brokerage and analyst targets in Forecast_Tracker; review SIP top-ups if panic reached 60. |
| **After each crisis** | Add the crisis to `analyze.py` (CRISES list) and review which stocks defended and led. |

---

## 8. When to be away, when to enter, when to cash out — the playbook

### 8.1 Mood × trend decision table

| | **Uptrend** (Nifty above 200-day average) | **Downtrend** (below 200-day average) |
|---|---|---|
| **Complacent (0–20)** | Stay invested; **cash out in stages**: trim positions that reached targets, trail stop-losses tighter. Avoid chasing small caps. | Rare. Be careful: a rally inside a bear market. |
| **Calm (20–40)** | Fully invested; buy dips in quality stocks. | Hold 50%; no new aggressive positions. |
| **Worried (40–60)** | Stay invested; check stop-losses. | **Stay mostly away (50%)**, keep cash in a liquid fund, prepare the buy list, continue SIPs. |
| **Fear (60–80)** | Opportunity: add to quality on dips. | **Start entering in tranches**: defenders first (pharma, FMCG). Double SIP. |
| **Panic (80+)** | Very rare; buy aggressively. | **Accumulate in 3–4 tranches**, especially if the fall is already 25%+. High-beta leaders can be added gradually. Historically 96% of cases were higher one year later. |

### 8.2 Entry rules
1. **The strongest entry is panic and a deep fall together** (panic 70+ and Nifty 25%+ below its high).
2. **Stagger, never all at once.** The first fear wave of a big bear market can come too early, as in March 2008.
3. **"Recovery confirmed" entry for the safety-first investor:** Nifty 20% above its crisis low, or back above its 200-day average. In 7 of 10 crises this beat buying at the first 10% fall.
4. **For greed events** (positive surprises, elections with clear mandates): buy on the event day. Waiting has cost more than it saved.
5. **For fear events:** wait about 5 days. Buying 5 days after was the least bad entry for fear events, pandemics, bank collapses and RBI hikes.

### 8.3 When to be away
- Nifty below its 200-day average **and** panic below 60. You're in a slow decline where cash in a liquid fund pays almost as much with far less pain.
- **Before Union Budgets** (only 34% were higher 20 days later) and around **global shocks and pandemics**. These are the deepest and slowest-healing event types.
- When **Brent is above $120** or crude has risen more than 20% in a month, reduce exposure to paints, aviation and oil marketing companies; pharma and IT have held up best in such weeks.
- When foreign selling is heavy **and** the rupee is breaking down: the 2013 and 2026 pattern.

### 8.4 How to "cash it": taking profits
- Fear events reward the buyer *after* the shock. Greed events reward the buyer *on* the day and **the seller into strength**: greed episodes reverse within days.
- In the **complacent** zone (46% of all days), expected 1-year returns are the lowest (+7.6%). Book partial profits on stocks that hit their targets, and rotate gains into defenders or liquid funds.
- After a crisis bottom, **high-beta recovery leaders** (steel, metals, cement, tyres) can rise 100%+ in a year. Take profits in stages once Nifty is back near its old peak, because the rotation reverses.

---

## 9. Where we are now (October 2026)

| Measure | Now (6 Oct 2026) | Meaning |
|---|---|---|
| Nifty | 22,776 | −13.5% from the 2 Jan 2026 peak of 26,329 |
| Lowest close of this crisis | 22,331 (30 Mar 2026) | Nifty is only 2% above it, **testing the low** |
| Nifty vs its 200-day average (about 24,320) | −6.4% | **Downtrend** |
| Watchlist stocks below their 200-day average | 68% | Broad weakness |
| Panic score | **45, Worried** (March peak 77) | Fear already peaked once |
| India VIX | 13.6 | Calm volatility despite weak prices: a "slow grind" |
| Brent | $101.5 | Iran war, Hormuz still restricted |
| Rupee | 96.7 per $ | Record low; record foreign selling in 2026 |
| RBI | Repo raised to 5.50% on 7 Oct 2026 | Rate-hike events: best entry historically 5 days after |

**Playbook reading: "Downtrend + Worried"**, which calls for staying partly away (about 50%), keeping cash in a liquid fund, continuing SIPs and preparing the buy list.

What would change the reading:
- **Panic rises to 60+** while the March low breaks: start staggered buying, defenders first.
- **Nifty above its 200-day average (about 24,300)** or **20% above the low (about 26,800)**: the recovery is confirmed, so return to fully invested and rotate to recovery leaders.
- **A durable Iran ceasefire and Brent back under $90** is the most likely trigger for an improving mood. The April 2026 ceasefire produced a +3.8% day.
- Compared with history, this crisis is **shallow** (−15% against an average of −30%, and 82% of past crises fell deeper). It's also **long** (185 trading days since the peak against 169 days falling on average). Domestic SIP money is cushioning the fall.

---

## 10. Other thoughts and cautions

1. **Use moods to size and time, never to predict.** The Panic Meter tells us *where* the crowd's emotion is, not *what* tomorrow brings. Combine it with the trend.
2. **Keep the mathematical model in charge of stock selection**, and let the mood layer decide *how much* and *when*.
3. **Our own emotions are the biggest risk.** Losses feel about twice as painful as gains (loss aversion). Written rules, like the tables above, protect us from panic-selling at the bottom and over-buying at the top.
4. **Event types with few samples** (exit polls, domestic shocks, RBI pauses) carry low confidence. Treat their points as hints.
5. **Structural changes matter.** DIIs now own more of the market than FIIs; F&O taxes rose in 2024 and 2026; India's index weight changes. Past reactions may be milder or stronger in future, so refresh the study every quarter.
6. **Use Forecast_Tracker to learn whose forecasts to trust.** Log every brokerage or TV prediction; after a year, the sheet tells us whose hit rate is better than a coin toss.
7. **Liquid funds and FDs are not "doing nothing".** At about 6%, they make the trend rule almost as profitable as holding through crashes, with a third of the pain.
8. **The same thinking applies to the function-hall business.** Events and moods (wedding seasons, auspicious dates, economic slowdowns, fuel prices) can be logged the same way: record bookings against such events and measure the impact. Your Excel skills transfer directly.

---

## 11. Limitations and next steps

**Limitations**
- Event dates and news summaries were compiled during this study. Please verify them, especially events before 2004 and in 2026. Those marked "Please verify" have weaker sources.
- Prices are from Yahoo Finance and are price indices, without dividends for indices. Early data (1997–2006) is noisier and some stock windows are skipped.
- India VIX only exists from 2008, so the Panic Meter can't score crises before then.
- FII/DII history is only Jun–Oct 2026. Long history from NSE or NSDL would strengthen the flow analysis.
- Back-tests ignore taxes, brokerage, STT and slippage.

**Suggested next steps**
1. Download long-term FII/DII history from NSE or NSDL and paste it into the FII_DII sheet.
2. Add the October 2026 RBI hike and future events as they happen, and re-run the scripts each week.
3. Connect your holdings (quantity and buy price, kept private) to Target_Adjuster to get event-adjusted targets for each holding.
4. Back-test the playbook on individual watchlist stocks, not just Nifty.
5. Start logging brokerage targets in Forecast_Tracker.

---

## 12. Sources

**Research on forecasts and behaviour**
- Cowles (1933), *Can Stock Market Forecasters Forecast?* — https://economics.yale.edu/sites/default/files/2022-08/cowles-forecasters33.pdf
- Tetlock (2005), *Expert Political Judgment* — https://www.journalofaccountancy.com/issues/2006/mar/bewareexpertpredictions.html
- CXO Advisory, Guru Grades — https://www.cxoadvisory.com/gurus/
- An, Jalles & Loungani (2018), IMF WP 18/39 — https://www.imf.org/en/publications/wp/issues/2018/03/05/how-well-do-economists-forecast-recessions-45672
- McKinsey (2010), *Equity analysts: Still too bullish* — https://www.mckinsey.com/capabilities/strategy-and-corporate-finance/our-insights/equity-analysts-still-too-bullish
- SPIVA India Year-End 2024 — https://www.spglobal.com/spdji/en/spiva/article/spiva-india-year-end-2024/
- BofA cash rule — https://www.ii.co.uk/analysis-commentary/professional-investor-pessimism-triggers-buy-signal-ii529538
- Kahneman & Tversky (1979), *Prospect Theory* — https://doi.org/10.2307/1914185

**2025–2026 events and strategy shifts**
- 2026 Iran war — https://www.britannica.com/event/2026-Iran-war and https://en.wikipedia.org/wiki/Timeline_of_the_2026_Iran_war
- Markets on 2 Mar 2026 — https://www.business-standard.com/markets/news/stock-market-crash-us-iran-war-sensex-nifty-oil-fii-rupee-why-are-markets-falling-today-126030400156_1.html
- Ceasefire rally, 8 Apr 2026 — https://www.newsonair.gov.in/indian-markets-rally-over-3-on-us-israel-iran-ceasefire
- Budget 2026 STT sell-off — https://www.businesstoday.in/amp/union-budget/story/budget-2026-stock-market-selloff-investors-lose-rs-10-lakh-crore-as-sensex-nifty-crash-514005-2026-02-01
- India–US trade deal — https://www.whitehouse.gov/briefings-statements/2026/02/united-states-india-joint-statement/
- US 50% tariff on India — https://www.india-briefing.com/news/us-india-tariff-50-percent-new-rules-impact-exporters-39458.html/
- RBI hike, 7 Oct 2026 — https://www.forbesindia.com/article/news/rbi-mpc-meeting-live-rate-hike-in-focus-as-inflation-oil-rise-liveblog/2999124/1
- 2026 state election results — https://www.outlookmoney.com/invest/west-bengal-tamil-nadu-election-results-exit-poll-2026-day-sensex-surges-1000-points-as-vote-counting-begins-nifty-50-above-24200
- Nifty target cuts, Apr 2026 — https://www.business-standard.com/amp/markets/news/iran-war-impact-nifty-target-cut-2026-brokerages-flag-oil-inflation-risks-india-126042801604_1.html
- DIIs overtake FPIs — https://www.business-standard.com/markets/news/diis-surpass-fpis-in-ownership-of-nse-listed-firms-in-march-2025-125050200426_1.html
- 2026 FPI outflows — https://www.kotakneo.com/news/market-news/fpi-outflow-2026-nears-30-billion-record-foreign-selling/
- BofA surveys 2026 — https://www.scmp.com/business/china-business/article/3347009/iran-conflict-drives-fund-managers-slash-risk-and-hoard-cash-bofa-survey-shows , https://www.bloomberg.com/news/articles/2026-04-14/investors-slash-growth-views-by-most-in-four-years-bofa-says , https://www.axios.com/2026/05/20/fund-managers-stocks-bofa
- Hartnett on Iran war winners and losers — https://www.investing.com/news/stock-market-news/bofas-hartnett-flags-asset-winners-and-losers-from-prolonged-iran-war-4546287
- Russia–Ukraine status 2026 — https://www.cbsnews.com/news/russia-ukraine-war-no-rush-for-peace-moscow-says-despite-trump-push/

**Price data:** Yahoo Finance (via the `yfinance` package). All other figures in this report are calculated in `Event_Impact_Analyzer.xlsx` and `mood_backtest.py`.
