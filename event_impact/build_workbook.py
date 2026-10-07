"""Build Event_Impact_Analyzer.xlsx from the CSVs produced by analyze.py.

Raw measurements (per event, per stock) are written as values; every summary,
ranking, score and the target adjuster are live Excel formulas that refer to
them, so changing a setting (yellow cell) recalculates the whole workbook.
"""
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).parent
DATA = HERE / "data"
OUT = HERE / "Event_Impact_Analyzer.xlsx"
FII_SOURCE = HERE / "fii_dii_daily.csv"  # exported from the FII_DII_Daily sheet of the shares worksheet
FII_ROWS = 1000  # pre-filled formula rows on the FII_DII sheet

FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color="1F3864")
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_OUT = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PCT = '0.0%;[Red]-0.0%;"-"'
NUM1 = '0.0;[Red]-0.0;"-"'
NUM0 = '#,##0;[Red]-#,##0;"-"'
RS = '"₹"#,##0.00'
CR = '#,##0;[Red](#,##0);"-"'

RED_FILL = PatternFill("solid", fgColor="F8CBAD")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
GREY_FILL = PatternFill("solid", fgColor="EDEDED")


def scale3():
    return ColorScaleRule(start_type="min", start_color="F8696B", mid_type="num", mid_value=0,
                          mid_color="FFFFFF", end_type="max", end_color="63BE7B")


def header(ws, row, labels, col=1, fill=FILL_HEAD, font=F_HEAD, height=30):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=text)
        c.font, c.fill, c.border = font, fill, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def title(ws, text, note=None):
    ws["A1"] = text
    ws["A1"].font = F_TITLE
    if note:
        ws["A2"] = note
        ws["A2"].font = F_NOTE


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def style_range(ws, rng, fmt=None, font=F_BASE, border=True):
    for row in ws[rng]:
        for c in row:
            c.font = font
            if fmt:
                c.number_format = fmt
            if border:
                c.border = BOX


def mood_rules(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Fear"'], fill=RED_FILL))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Greed"'], fill=GREEN_FILL))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Calm"'], fill=GREY_FILL))


def nz(v):
    return None if pd.isna(v) else v


# ---------------------------------------------------------------- settings
SETTINGS = [  # (row, label, value, format, explanation)
    (4, "Fear threshold: Nifty worst fall within 20 days of the event", -0.05, PCT,
     "An event is tagged FEAR when Nifty fell at least this much (and the fall was bigger than the rise)."),
    (5, "Greed threshold: Nifty best rise within 20 days of the event", 0.05, PCT,
     "An event is tagged GREED when Nifty rose at least this much. Everything else is CALM."),
    (6, "Points scale: average move worth 1 point", 0.005, PCT,
     "0.5% means an average 20-day move of +3% scores +6 points (capped at +/-10)."),
    (7, "Minimum events before trusting a stock's own history", 5, "0",
     "Below this, the Target Adjuster uses Event Beta x Nifty instead of the stock's own average."),
    (8, "Short-term target weight", 1.0, "0.00",
     "1.00 = apply the full expected event move to your short-term target."),
    (9, "Long-term target weight", 0.25, "0.00",
     "Long-term targets should follow fundamentals; 0.25 applies only a quarter of the event move."),
    (10, "Stop-loss buffer multiplier", 1.0, "0.00",
     "1.00 = stop-loss at the typical worst dip after this kind of event; 1.2 = 20% wider."),
    (11, "FII/DII 'big flow' threshold (Rs crore)", 3000, "#,##0",
     "Used on the FII_DII sheet to split big buying days from big selling days."),
]
S = {k: f"Settings!$B${r}" for k, r in zip(
    ["fear", "greed", "scale", "minev", "stw", "ltw", "slm", "fii"], [r for r, *_ in SETTINGS])}


def build_settings(wb):
    ws = wb.create_sheet("Settings")
    title(ws, "Settings — change the yellow cells; the whole workbook recalculates")
    header(ws, 3, ["Setting", "Value", "What it means"])
    for r, label, val, fmt, expl in SETTINGS:
        ws.cell(r, 1, label).font = F_BASE
        c = ws.cell(r, 2, val)
        c.font, c.fill, c.number_format, c.border = F_INPUT, FILL_INPUT, fmt, BOX
        ws.cell(r, 3, expl).font = F_NOTE
    widths(ws, {"A": 58, "B": 12, "C": 100})


# ---------------------------------------------------------------- calendar + nifty impact
NI_COLS = [  # (csv column, header, format)
    ("Event_ID", "Event ID", None), ("Trading_Day", "Trading day (T0)", "dd-mmm-yyyy"),
    ("Category", "Category", None), ("News_Tone", "News tone", None), ("Surprise", "Surprise?", None),
    ("Headline", "Headline", None), ("Index", "Index used", None),
    ("Pre20", "BEFORE: 20-day run-up", PCT), ("Pre5", "BEFORE: 5-day run-up", PCT),
    ("Day0", "DURING: event-day move", PCT), ("Day0_Low", "DURING: intraday worst", PCT),
    ("Post5", "AFTER: next 5 days", PCT), ("Post20", "AFTER: next 20 days", PCT),
    ("Post60", "AFTER: next 60 days", PCT), ("Total20", "Total move T-1 to T+20", PCT),
    ("MaxFall20", "Worst fall within 20 days", PCT), ("MaxRise20", "Best rise within 20 days", PCT),
    ("Recovery_Days", "Days to recover pre-event level", NUM0),
    ("VIX_Before", "India VIX before", NUM1), ("VIX_Peak20", "VIX peak (20 days)", NUM1),
    ("VIX_Change", "VIX jump", PCT), ("VIX_Settle_Days", "Days for VIX to settle", NUM0),
    ("Ret20_Before", "Buy 5 days BEFORE: 20-day return", PCT),
    ("Reward_Before", "Buy BEFORE: best gain", PCT), ("Risk_Before", "Buy BEFORE: worst dip", PCT),
    ("Ret20_EventDay", "Buy ON EVENT DAY: 20-day return", PCT),
    ("Reward_EventDay", "Buy ON EVENT DAY: best gain", PCT), ("Risk_EventDay", "Buy ON EVENT DAY: worst dip", PCT),
    ("Ret20_After", "Buy 5 days AFTER: 20-day return", PCT),
    ("Reward_After", "Buy AFTER: best gain", PCT), ("Risk_After", "Buy AFTER: worst dip", PCT),
]
NI = {name: L(i + 1) for i, (name, _, _) in enumerate(NI_COLS)}
NI["Mood"] = L(len(NI_COLS) + 1)
NI["Up20"] = L(len(NI_COLS) + 2)


def build_calendar(wb, ni):
    ws = wb.create_sheet("Event_Calendar")
    title(ws, "Event Calendar — the news behind each event",
          "Dates and news summaries were compiled for this analysis; please verify them. Crude Spike/Crash rows "
          "are detected automatically from Brent prices. To add an event, add a row to events.csv and re-run the scripts.")
    cols = ["Event ID", "Event date", "Trading day (T0)", "Category", "News tone", "Surprise?", "Headline",
            "What happened / what the market expected"]
    header(ws, 4, cols)
    for r, row in enumerate(ni.itertuples(index=False), start=5):
        vals = [row.Event_ID, pd.Timestamp(row.Date).to_pydatetime(), pd.Timestamp(row.Trading_Day).to_pydatetime(),
                row.Category, row.News_Tone, row.Surprise, row.Headline, row.Context]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell.font, cell.border = F_BASE, BOX
            if c in (2, 3):
                cell.number_format = "dd-mmm-yyyy"
            if c == 8:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:H{4 + len(ni)}"
    widths(ws, {"A": 9, "B": 12, "C": 12, "D": 17, "E": 10, "F": 9, "G": 42, "H": 90})


def build_nifty_impact(wb, ni):
    ws = wb.create_sheet("Nifty_Impact")
    heads = [h for _, h, _ in NI_COLS] + ["Market mood (formula)", "Higher after 20 days? (1=yes)"]
    header(ws, 1, heads, height=45)
    n = len(ni)
    for r, row in enumerate(ni.to_dict("records"), start=2):
        for c, (name, _, fmt) in enumerate(NI_COLS, start=1):
            v = row[name]
            if name == "Trading_Day":
                v = pd.Timestamp(v).to_pydatetime()
            cell = ws.cell(r, c, nz(v))
            cell.font, cell.border = F_BASE, BOX
            if fmt:
                cell.number_format = fmt
        f, g = NI["MaxFall20"], NI["MaxRise20"]
        ws[f"{NI['Mood']}{r}"] = (f'=IF({f}{r}="","",IF(AND({f}{r}<={S["fear"]},-{f}{r}>={g}{r}),"Fear",'
                                  f'IF({g}{r}>={S["greed"]},"Greed","Calm")))')
        t = NI["Total20"]
        ws[f"{NI['Up20']}{r}"] = f'=IF({t}{r}="","",IF({t}{r}>0,1,0))'
        for k in ("Mood", "Up20"):
            ws[f"{NI[k]}{r}"].font, ws[f"{NI[k]}{r}"].border = F_BASE, BOX
    mood_rules(ws, f"{NI['Mood']}2:{NI['Mood']}{n + 1}")
    for k in ("Day0", "Total20", "Post60"):
        ws.conditional_formatting.add(f"{NI[k]}2:{NI[k]}{n + 1}", scale3())
    ws.freeze_panes = "G2"
    ws.auto_filter.ref = f"A1:{NI['Up20']}{n + 1}"
    for i in range(1, len(heads) + 1):
        ws.column_dimensions[L(i)].width = 11
    widths(ws, {"C": 16, "F": 40, NI["Mood"]: 12})
    return n + 1  # last data row


def ni_rng(key, last):
    return f"Nifty_Impact!${NI[key]}$2:${NI[key]}${last}"


# ---------------------------------------------------------------- category summary
def build_category_summary(wb, categories, last):
    ws = wb.create_sheet("Category_Summary")
    title(ws, "Category Summary — how Nifty behaved around each type of event (all formulas)",
          "Averages over every event in the category. Phases: BEFORE = run-up into the event, DURING = event day, "
          "AFTER = following days. R:R = average best gain / average worst dip in the 20 days after buying.")
    heads = ["Category", "Events", "BEFORE: avg 20-day run-up", "DURING: avg event-day move",
             "AFTER: avg next 5 days", "Avg total move (T-1 to T+20)", "AFTER: avg next 60 days",
             "% of events Nifty higher after 20 days", "Avg worst fall (20 days)", "Avg best rise (20 days)",
             "Avg days to recover", "Avg VIX jump", "Avg days for VIX to settle", "Fear events", "Greed events",
             "Buy BEFORE: avg 20-day return", "Buy BEFORE: R:R", "Buy ON DAY: avg 20-day return", "Buy ON DAY: R:R",
             "Buy AFTER: avg 20-day return", "Buy AFTER: R:R", "Best entry (history)",
             "Direction points (-10 to +10)", "Risk points (0 to 10)", "Confidence", "Signal"]
    header(ws, 4, heads, height=60)
    cat = ni_rng("Category", last)

    def avg(key, r):
        return f'=IFERROR(AVERAGEIFS({ni_rng(key, last)},{cat},$A{r}),"")'

    def rr(label, r):
        return (f'=IFERROR(AVERAGEIFS({ni_rng("Reward_" + label, last)},{cat},$A{r})/'
                f'AVERAGEIFS({ni_rng("Risk_" + label, last)},{cat},$A{r}),"")')

    for r, name in enumerate(categories, start=5):
        ws.cell(r, 1, name)
        ws.cell(r, 2, f"=COUNTIF({cat},$A{r})")
        for c, key in zip(range(3, 14), ["Pre20", "Day0", "Post5", "Total20", "Post60", "Up20", "MaxFall20",
                                         "MaxRise20", "Recovery_Days", "VIX_Change", "VIX_Settle_Days"]):
            ws.cell(r, c, avg(key, r))
        ws.cell(r, 14, f'=COUNTIFS({cat},$A{r},{ni_rng("Mood", last)},"Fear")')
        ws.cell(r, 15, f'=COUNTIFS({cat},$A{r},{ni_rng("Mood", last)},"Greed")')
        ws.cell(r, 16, avg("Ret20_Before", r))
        ws.cell(r, 17, rr("Before", r))
        ws.cell(r, 18, avg("Ret20_EventDay", r))
        ws.cell(r, 19, rr("EventDay", r))
        ws.cell(r, 20, avg("Ret20_After", r))
        ws.cell(r, 21, rr("After", r))
        ws.cell(r, 22, f'=IF(P{r}="","",IF(AND(P{r}>=R{r},P{r}>=T{r}),"Buy before",'
                       f'IF(R{r}>=T{r},"Buy on event day","Buy 5 days after")))')
        ws.cell(r, 23, f'=IF(F{r}="","",MAX(-10,MIN(10,ROUND(F{r}/{S["scale"]},0))))')
        ws.cell(r, 24, f'=IF(I{r}="","",MAX(0,MIN(10,ROUND(-I{r}/(2*{S["scale"]}),0))))')
        ws.cell(r, 25, f'=IF(B{r}>=10,"High",IF(B{r}>=5,"Medium","Low"))')
        ws.cell(r, 26, f'=IF(W{r}="","",IF(W{r}>=2,"Bullish bias",IF(W{r}<=-2,"Bearish bias","Neutral")))')
    end = 4 + len(categories)
    style_range(ws, f"A5:Z{end}")
    for col in "CDEFGHIJLPRT":
        style_range(ws, f"{col}5:{col}{end}", PCT)
    for col in "KM":
        style_range(ws, f"{col}5:{col}{end}", NUM0)
    for col in "QSU":
        style_range(ws, f"{col}5:{col}{end}", "0.00")
    for col in "DFG":
        ws.conditional_formatting.add(f"{col}5:{col}{end}", scale3())
    ws.conditional_formatting.add(f"W5:W{end}", scale3())
    for r in range(5, end + 1):
        ws.cell(r, 1).font = F_BOLD

    # ---- news tone x surprise
    r0 = end + 3
    ws.cell(r0 - 1, 1, "How the market reacts to POSITIVE vs NEGATIVE news (and whether it was a surprise)").font = F_BOLD
    header(ws, r0, ["News tone", "Surprise?", "Events", "Avg event-day move", "Avg total move (20 days)",
                    "% higher after 20 days", "Avg worst fall", "Avg days to recover"], height=45)
    tone, sur = ni_rng("News_Tone", last), ni_rng("Surprise", last)
    r = r0 + 1
    for t in ["Positive", "Negative", "Mixed", "Neutral"]:
        for s in ["Yes", "No", "Partly"]:
            ws.cell(r, 1, t)
            ws.cell(r, 2, s)
            ws.cell(r, 3, f"=COUNTIFS({tone},$A{r},{sur},$B{r})")
            for c, key in zip(range(4, 9), ["Day0", "Total20", "Up20", "MaxFall20", "Recovery_Days"]):
                ws.cell(r, c, f'=IFERROR(AVERAGEIFS({ni_rng(key, last)},{tone},$A{r},{sur},$B{r}),"")')
            r += 1
    style_range(ws, f"A{r0 + 1}:H{r - 1}")
    for col in "DEFG":
        style_range(ws, f"{col}{r0 + 1}:{col}{r - 1}", PCT)
    style_range(ws, f"H{r0 + 1}:H{r - 1}", NUM0)
    ws.conditional_formatting.add(f"E{r0 + 1}:E{r - 1}", scale3())

    # ---- mood summary: who made the money
    m0 = r + 2
    ws.cell(m0 - 1, 1, "FEAR vs GREED events — who made the money? (buyer's average 20-day result by entry timing)").font = F_BOLD
    header(ws, m0, ["Market mood", "Events", "Avg worst fall", "Avg days to recover", "Avg days for VIX to settle",
                    "Buy BEFORE: 20-day return", "Buy ON DAY: 20-day return", "Buy AFTER: 20-day return",
                    "Buy ON DAY: R:R", "Who made the money"], height=45)
    mood = ni_rng("Mood", last)
    for i, m in enumerate(["Fear", "Greed", "Calm"]):
        rr_ = m0 + 1 + i
        ws.cell(rr_, 1, m)
        ws.cell(rr_, 2, f'=COUNTIF({mood},$A{rr_})')
        for c, key in zip(range(3, 9), ["MaxFall20", "Recovery_Days", "VIX_Settle_Days", "Ret20_Before",
                                        "Ret20_EventDay", "Ret20_After"]):
            ws.cell(rr_, c, f'=IFERROR(AVERAGEIFS({ni_rng(key, last)},{mood},$A{rr_}),"")')
        ws.cell(rr_, 9, f'=IFERROR(AVERAGEIFS({ni_rng("Reward_EventDay", last)},{mood},$A{rr_})/'
                        f'AVERAGEIFS({ni_rng("Risk_EventDay", last)},{mood},$A{rr_}),"")')
        ws.cell(rr_, 10, f'=IF(G{rr_}="","",IF(MAX(F{rr_}:H{rr_})<0,"Sellers before the event / cash holders",'
                         f'IF(AND(H{rr_}>=G{rr_},H{rr_}>=F{rr_}),"Patient buyers who waited 5 days",'
                         f'IF(G{rr_}>=F{rr_},"Buyers on the event day","Early buyers (positioned before)"))))')
    style_range(ws, f"A{m0 + 1}:J{m0 + 3}")
    for col in "CFGH":
        style_range(ws, f"{col}{m0 + 1}:{col}{m0 + 3}", PCT)
    for col in "DE":
        style_range(ws, f"{col}{m0 + 1}:{col}{m0 + 3}", NUM0)
    style_range(ws, f"I{m0 + 1}:I{m0 + 3}", "0.00")
    mood_rules(ws, f"A{m0 + 1}:A{m0 + 3}")

    ws.freeze_panes = "B5"
    widths(ws, {"A": 18, "V": 17, "Y": 11, "Z": 13, "J": 13})
    for i in range(2, 27):
        if L(i) not in ("V", "Y", "Z", "J"):
            ws.column_dimensions[L(i)].width = 11
    return f"Category_Summary!$A$5:$A${end}"


# ---------------------------------------------------------------- stock level
SE_COLS = ["Stock", "Event ID", "Trading day", "Category", "Industry", "Cap", "Market mood",
           "Stock event-day move", "Nifty event-day move", "Excess vs Nifty (day)",
           "Stock total move (20 days)", "Nifty total move (20 days)", "Excess vs Nifty (20 days)",
           "Stock worst fall (20 days)", "Stock best rise (20 days)", "Days to recover",
           "Stock next 60 days", "Buy on day: best gain", "Buy on day: worst dip", "Buy on day: 20-day return"]


def build_stock_events(wb, se, last_ni):
    ws = wb.create_sheet("Stock_Event_Data")
    header(ws, 1, SE_COLS, height=45)
    for r, row in enumerate(se.itertuples(index=False), start=2):
        vals = {1: row.Stock, 2: row.Event_ID, 3: pd.Timestamp(row.Trading_Day).to_pydatetime(),
                4: row.Category, 5: row.Industry, 6: row.Cap,
                8: nz(row.Day0), 9: nz(row.Nifty_Day0), 11: nz(row.Total20), 12: nz(row.Nifty_Total20),
                14: nz(row.MaxFall20), 15: nz(row.MaxRise20), 16: nz(row.Recovery_Days), 17: nz(row.Post60),
                18: nz(row.Reward_EventDay), 19: nz(row.Risk_EventDay), 20: nz(row.Ret20_EventDay)}
        for c, v in vals.items():
            ws.cell(r, c, v)
        ws.cell(r, 7, f'=INDEX({ni_rng("Mood", last_ni)},MATCH(B{r},{ni_rng("Event_ID", last_ni)},0))')
        ws.cell(r, 10, f'=IF(OR(H{r}="",I{r}=""),"",H{r}-I{r})')
        ws.cell(r, 13, f'=IF(OR(K{r}="",L{r}=""),"",K{r}-L{r})')
    last = len(se) + 1
    for col, fmt in (("C", "dd-mmm-yyyy"), ("P", NUM0)):
        for c in ws[f"{col}2:{col}{last}"]:
            c[0].number_format = fmt
    for col in "HIJKLMNOQRST":
        for c in ws[f"{col}2:{col}{last}"]:
            c[0].number_format = PCT
    mood_rules(ws, f"G2:G{last}")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:T{last}"
    for i in range(1, 21):
        ws.column_dimensions[L(i)].width = 11
    widths(ws, {"D": 16, "E": 18})
    return last


def se_rng(col, last):
    return f"Stock_Event_Data!${col}$2:${col}${last}"


def build_scorecard(wb, watch, last_se):
    ws = wb.create_sheet("Stock_Scorecard")
    title(ws, "Stock Scorecard — how each watchlist stock behaves around events (all formulas)",
          "FEAR columns use only events where Nifty fell (see Settings). Event Beta = how much the stock moves for "
          "every 1% Nifty move around events. Lower 'score' = better rank.")
    heads = ["Stock", "Name", "Industry", "Cap", "Events measured", "First data row", "Fear events",
             "Avg event-day excess vs Nifty", "FEAR: avg 20-day move", "FEAR: avg excess vs Nifty",
             "FEAR: avg worst fall", "FEAR: avg days to recover", "GREED: avg 20-day move",
             "GREED: avg excess vs Nifty", "Event Beta", "Event R:R (buy on day)",
             "Resilience score", "Resilience rank (1 = best)", "Impact score", "Impact rank (1 = most hit)"]
    header(ws, 4, heads, height=60)
    stk, mood = se_rng("A", last_se), se_rng("G", last_se)
    n = len(watch)
    end = 4 + n
    for r, w in enumerate(watch.itertuples(index=False), start=5):
        ws.cell(r, 1, w.NSE_Code).font = F_BOLD
        ws.cell(r, 2, w.Name)
        ws.cell(r, 3, w.Industry)
        ws.cell(r, 4, w.Cap)
        ws.cell(r, 5, f"=COUNTIF({stk},$A{r})")
        ws.cell(r, 6, f'=IFERROR(MATCH($A{r},{stk},0)+1,"")')
        ws.cell(r, 7, f'=COUNTIFS({stk},$A{r},{mood},"Fear")')
        ws.cell(r, 8, f'=IFERROR(AVERAGEIFS({se_rng("J", last_se)},{stk},$A{r}),"")')
        for c, col in ((9, "K"), (10, "M"), (11, "N"), (12, "P")):
            ws.cell(r, c, f'=IFERROR(AVERAGEIFS({se_rng(col, last_se)},{stk},$A{r},{mood},"Fear"),"")')
        for c, col in ((13, "K"), (14, "M")):
            ws.cell(r, c, f'=IFERROR(AVERAGEIFS({se_rng(col, last_se)},{stk},$A{r},{mood},"Greed"),"")')
        ws.cell(r, 15, f'=IF(E{r}<3,"",IFERROR(SLOPE(INDEX(Stock_Event_Data!$K:$K,F{r}):INDEX(Stock_Event_Data!$K:$K,F{r}+E{r}-1),'
                       f'INDEX(Stock_Event_Data!$L:$L,F{r}):INDEX(Stock_Event_Data!$L:$L,F{r}+E{r}-1)),""))')
        ws.cell(r, 16, f'=IFERROR(AVERAGEIFS({se_rng("R", last_se)},{stk},$A{r})/AVERAGEIFS({se_rng("S", last_se)},{stk},$A{r}),"")')
        ok = f'OR(G{r}<{S["minev"]},J{r}="",K{r}="",L{r}="",O{r}="")'
        ws.cell(r, 17, f'=IF({ok},"",RANK(J{r},$J$5:$J${end},0)+RANK(K{r},$K$5:$K${end},0)'
                       f'+RANK(L{r},$L$5:$L${end},1)+ROW()/100000)')
        ws.cell(r, 18, f'=IF(Q{r}="","",RANK(Q{r},$Q$5:$Q${end},1))')
        ws.cell(r, 19, f'=IF({ok},"",RANK(K{r},$K$5:$K${end},1)+RANK(O{r},$O$5:$O${end},0)+ROW()/100000)')
        ws.cell(r, 20, f'=IF(S{r}="","",RANK(S{r},$S$5:$S${end},1))')
    style_range(ws, f"A5:T{end}")
    for r in range(5, end + 1):
        ws.cell(r, 1).font = F_BOLD
    for col in "HIJKMN":
        style_range(ws, f"{col}5:{col}{end}", PCT)
    style_range(ws, f"L5:L{end}", NUM0)
    for col in "OP":
        style_range(ws, f"{col}5:{col}{end}", "0.00")
    for col in "QS":
        style_range(ws, f"{col}5:{col}{end}", "0.0")
    for col in "JKN":
        ws.conditional_formatting.add(f"{col}5:{col}{end}", scale3())
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:T{end}"
    for i in range(1, 21):
        ws.column_dimensions[L(i)].width = 11
    widths(ws, {"B": 20, "C": 18, "F": 7})
    ws.column_dimensions["F"].hidden = True
    return end


def build_top_lists(wb, end_sc, top=25):
    ws = wb.create_sheet("Top_Lists")
    title(ws, "Top Lists — most resilient vs most impacted stocks in FEAR events (formulas)",
          "Resilient = fell less than Nifty, smaller dips, quicker recovery. Most impacted = deepest dips and highest Event Beta. "
          "Only stocks with enough FEAR events (Settings) are ranked.")
    sc = lambda c: f"Stock_Scorecard!${c}$5:${c}${end_sc}"  # noqa: E731
    blocks = [("MOST RESILIENT — performed best even in fear", "Q", 4),
              ("MOST IMPACTED — hit hardest in fear", "S", 8 + top)]
    for label, score_col, r0 in blocks:
        ws.cell(r0 - 1, 1, label).font = F_BOLD
        header(ws, r0, ["Rank", "Stock", "Name", "Industry", "Cap", "Fear events", "FEAR: avg 20-day move",
                        "FEAR: excess vs Nifty", "FEAR: avg worst fall", "FEAR: days to recover",
                        "GREED: avg 20-day move", "Event Beta"], height=45)
        for k in range(1, top + 1):
            r = r0 + k
            ws.cell(r, 1, k)
            match = f"MATCH(SMALL({sc(score_col)},$A{r}),{sc(score_col)},0)"
            for c, col in zip(range(2, 13), "ABCDGIJKLMO"):
                ws.cell(r, c, f'=IFERROR(INDEX({sc(col)},{match}),"")')
        style_range(ws, f"A{r0 + 1}:L{r0 + top}")
        for col in "GHIK":
            style_range(ws, f"{col}{r0 + 1}:{col}{r0 + top}", PCT)
        style_range(ws, f"J{r0 + 1}:J{r0 + top}", NUM0)
        style_range(ws, f"L{r0 + 1}:L{r0 + top}", "0.00")
    widths(ws, {"A": 6, "B": 13, "C": 20, "D": 18, "E": 10})
    for col in "FGHIJKL":
        ws.column_dimensions[col].width = 12


def build_sector(wb, industries, categories, last_se):
    ws = wb.create_sheet("Sector_Impact")
    title(ws, "Sector Impact — average 20-day move of watchlist stocks in each industry, by event type (formulas)",
          "Green = rose on average after this type of event, red = fell. Built from Stock_Event_Data.")
    heads = ["Industry", "Stock-events"] + categories + ["FEAR: avg 20-day move", "FEAR: avg worst fall",
                                                        "FEAR: days to recover", "GREED: avg 20-day move"]
    header(ws, 4, heads, height=60)
    ind, cat, mood = se_rng("E", last_se), se_rng("D", last_se), se_rng("G", last_se)
    nc = len(categories)
    for r, name in enumerate(industries, start=5):
        ws.cell(r, 1, name).font = F_BOLD
        ws.cell(r, 2, f"=COUNTIF({ind},$A{r})")
        for j in range(nc):
            col = L(3 + j)
            ws.cell(r, 3 + j, f'=IFERROR(AVERAGEIFS({se_rng("K", last_se)},{ind},$A{r},{cat},{col}$4),"")')
        base = 3 + nc
        ws.cell(r, base, f'=IFERROR(AVERAGEIFS({se_rng("K", last_se)},{ind},$A{r},{mood},"Fear"),"")')
        ws.cell(r, base + 1, f'=IFERROR(AVERAGEIFS({se_rng("N", last_se)},{ind},$A{r},{mood},"Fear"),"")')
        ws.cell(r, base + 2, f'=IFERROR(AVERAGEIFS({se_rng("P", last_se)},{ind},$A{r},{mood},"Fear"),"")')
        ws.cell(r, base + 3, f'=IFERROR(AVERAGEIFS({se_rng("K", last_se)},{ind},$A{r},{mood},"Greed"),"")')
    end = 4 + len(industries)
    style_range(ws, f"A5:{L(base + 3)}{end}", PCT)
    style_range(ws, f"B5:B{end}", NUM0)
    style_range(ws, f"{L(base + 2)}5:{L(base + 2)}{end}", NUM0)
    for r in range(5, end + 1):
        ws.cell(r, 1).font = F_BOLD
    ws.conditional_formatting.add(f"C5:{L(base - 1)}{end}", scale3())
    for k in (0, 1, 3):
        c = L(base + k)
        ws.conditional_formatting.add(f"{c}5:{c}{end}", scale3())
    ws.freeze_panes = "C5"
    widths(ws, {"A": 22, "B": 9})
    for i in range(3, base + 4):
        ws.column_dimensions[L(i)].width = 11


# ---------------------------------------------------------------- crude + daily data
def build_market_daily(wb, md):
    ws = wb.create_sheet("Market_Daily")
    cols = list(md.columns)
    header(ws, 1, cols, height=45)
    for r, row in enumerate(md.itertuples(index=False), start=2):
        for c, v in enumerate(row, start=1):
            if c == 1:
                v = pd.Timestamp(v).to_pydatetime()
            ws.cell(r, c, nz(v))
    last = len(md) + 1
    for c, name in enumerate(cols, start=1):
        fmt = ("dd-mmm-yyyy" if name == "Date" else "#,##0.00" if name in ("Nifty", "Brent")
               else "0.00" if name == "VIX" else PCT)
        for cell in ws[f"{L(c)}2:{L(c)}{last}"]:
            cell[0].number_format = fmt
    ws.freeze_panes = "B2"
    for i in range(1, len(cols) + 1):
        ws.column_dimensions[L(i)].width = 11
    return {name: L(i + 1) for i, name in enumerate(cols)}, last


def build_crude(wb, mdc, last_md):
    ws = wb.create_sheet("Crude_Ranges")
    title(ws, "Crude Oil Ranges — at what level of crude move does the market react? (formulas over Market_Daily)",
          "Change the yellow bucket limits to test your own ranges. 'Same 5 days' = during the crude move; "
          "'next' = what followed. Industry columns are equal-weight averages of your watchlist stocks.")
    md = lambda k: f"Market_Daily!${mdc[k]}$2:${mdc[k]}${last_md}"  # noqa: E731
    groups = [k[:-6] for k in mdc if k.endswith("_Same5") and k != "Nifty_Same5"]

    # ---- table A: 5-day Brent change
    ws["A4"] = "A. Brent change over 5 trading days"
    ws["A4"].font = F_BOLD
    heads = ["From", "To", "Days", "Nifty same 5 days", "Nifty next 5 days", "Nifty next 20 days",
             "% Nifty higher after 20 days", "Avg India VIX"] + [f"{g} same 5 days" for g in groups]
    header(ws, 5, heads, height=45)
    buckets = [(-1, -0.15), (-0.15, -0.10), (-0.10, -0.05), (-0.05, -0.02), (-0.02, 0.02),
               (0.02, 0.05), (0.05, 0.10), (0.10, 0.15), (0.15, 5)]
    for i, (lo, hi) in enumerate(buckets):
        r = 6 + i
        for c, v in ((1, lo), (2, hi)):
            cell = ws.cell(r, c, v)
            cell.font, cell.fill, cell.number_format, cell.border = F_INPUT, FILL_INPUT, PCT, BOX
        crit = f'{md("Brent_5D")},">="&$A{r},{md("Brent_5D")},"<"&$B{r}'
        ws.cell(r, 3, f"=COUNTIFS({crit})")
        for c, key in ((4, "Nifty_Same5"), (5, "Nifty_Next5"), (6, "Nifty_Next20"), (8, "VIX")):
            ws.cell(r, c, f'=IFERROR(AVERAGEIFS({md(key)},{crit}),"")')
        ws.cell(r, 7, f'=IFERROR(COUNTIFS({crit},{md("Nifty_Next20")},">0")/C{r},"")')
        for j, g in enumerate(groups):
            ws.cell(r, 9 + j, f'=IFERROR(AVERAGEIFS({md(g + "_Same5")},{crit}),"")')
    endA = 5 + len(buckets)
    lastc = L(8 + len(groups))
    style_range(ws, f"C6:{lastc}{endA}", PCT)
    style_range(ws, f"C6:C{endA}", NUM0)
    style_range(ws, f"H6:H{endA}", "0.0")
    ws.conditional_formatting.add(f"D6:F{endA}", scale3())
    ws.conditional_formatting.add(f"I6:{lastc}{endA}", scale3())

    # ---- table B: Brent price level
    r0 = endA + 3
    ws.cell(r0 - 1, 1, "B. Brent price level ($ per barrel)").font = F_BOLD
    heads = ["From $", "To $", "Days", "Nifty next 20 days", "% Nifty higher after 20 days", "Avg India VIX"] + \
            [f"{g} next 5 days" for g in groups]
    header(ws, r0, heads, height=45)
    levels = [(0, 40), (40, 60), (60, 80), (80, 100), (100, 120), (120, 200)]
    for i, (lo, hi) in enumerate(levels):
        r = r0 + 1 + i
        for c, v in ((1, lo), (2, hi)):
            cell = ws.cell(r, c, v)
            cell.font, cell.fill, cell.number_format, cell.border = F_INPUT, FILL_INPUT, "0", BOX
        crit = f'{md("Brent")},">="&$A{r},{md("Brent")},"<"&$B{r}'
        ws.cell(r, 3, f"=COUNTIFS({crit})")
        ws.cell(r, 4, f'=IFERROR(AVERAGEIFS({md("Nifty_Next20")},{crit}),"")')
        ws.cell(r, 5, f'=IFERROR(COUNTIFS({crit},{md("Nifty_Next20")},">0")/C{r},"")')
        ws.cell(r, 6, f'=IFERROR(AVERAGEIFS({md("VIX")},{crit}),"")')
        for j, g in enumerate(groups):
            ws.cell(r, 7 + j, f'=IFERROR(AVERAGEIFS({md(g + "_Next5")},{crit}),"")')
    endB = r0 + len(levels)
    lastb = L(6 + len(groups))
    style_range(ws, f"C{r0 + 1}:{lastb}{endB}", PCT)
    style_range(ws, f"C{r0 + 1}:C{endB}", NUM0)
    style_range(ws, f"F{r0 + 1}:F{endB}", "0.0")
    ws.conditional_formatting.add(f"D{r0 + 1}:D{endB}", scale3())
    ws.conditional_formatting.add(f"G{r0 + 1}:{lastb}{endB}", scale3())
    ws.cell(endB + 2, 1, "Reading tip: compare the extreme rows with the middle (-2% to +2%) row. The range where "
                         "the numbers clearly change is the level where the market starts reacting to crude.").font = F_NOTE
    widths(ws, {"A": 9, "B": 9})
    for i in range(3, 9 + len(groups)):
        ws.column_dimensions[L(i)].width = 12


def build_fii(wb, mdc, last_md):
    ws = wb.create_sheet("FII_DII")
    title(ws, "FII / DII flows vs Nifty — your daily tracker data (formulas)",
          "Copied from your FII_DII_Daily sheet (Jun-Oct 2026). Paste more history in the yellow columns A-E "
          f"(up to {FII_ROWS} rows); Nifty columns fill in automatically.")
    header(ws, 4, ["Date", "FII net (Rs cr)", "DII net (Rs cr)", "PCR Nifty", "India VIX", "Nifty close",
                   "Nifty move same day", "Nifty move next day", "FII cumulative", "DII cumulative",
                   "Who absorbed?"], height=45)
    rows = []
    if FII_SOURCE.exists():
        fii = pd.read_csv(FII_SOURCE, parse_dates=["Date"])
        rows = [(d.to_pydatetime(), a, b, c, e) for d, a, b, c, e in fii.itertuples(index=False)]
    md = lambda k: f"Market_Daily!${mdc[k]}$2:${mdc[k]}${last_md}"  # noqa: E731
    end = 4 + FII_ROWS
    for i in range(FII_ROWS):
        r = 5 + i
        if i < len(rows):
            d, fii, dii, pcr, vix = (nz(v) for v in rows[i])
            for c, v in ((1, d), (2, fii), (3, dii), (4, pcr), (5, vix)):
                ws.cell(r, c, v)
        for c in range(1, 6):
            ws.cell(r, c).fill, ws.cell(r, c).font = FILL_INPUT, F_INPUT
        ws.cell(r, 6, f'=IF(A{r}="","",IFERROR(INDEX({md("Nifty")},MATCH(A{r},{md("Date")},0)),""))')
        ws.cell(r, 7, f'=IF(A{r}="","",IFERROR(INDEX({md("Nifty_Ret1")},MATCH(A{r},{md("Date")},0)),""))')
        ws.cell(r, 8, f'=IF(A{r}="","",IFERROR(INDEX({md("Nifty_Next1")},MATCH(A{r},{md("Date")},0)),""))')
        ws.cell(r, 9, f'=IF(B{r}="","",SUM($B$5:B{r}))')
        ws.cell(r, 10, f'=IF(C{r}="","",SUM($C$5:C{r}))')
        ws.cell(r, 11, f'=IF(OR(B{r}="",C{r}=""),"",IF(AND(B{r}<0,C{r}>0),"DII bought FII selling",'
                       f'IF(AND(B{r}>0,C{r}<0),"FII bought DII selling",IF(B{r}>0,"Both buying","Both selling"))))')
    for col, fmt in (("A", "dd-mmm-yyyy"), ("B", CR), ("C", CR), ("D", "0.00"), ("E", "0.00"), ("F", "#,##0.00"),
                     ("G", PCT), ("H", PCT), ("I", CR), ("J", CR)):
        for c in ws[f"{col}5:{col}{end}"]:
            c[0].number_format = fmt
    ws.freeze_panes = "B5"

    # ---- summary block
    rng = lambda c: f"${c}$5:${c}${end}"  # noqa: E731
    th = S["fii"]
    items = [
        ("Days with data", f'=COUNT({rng("B")})', "0"),
        ("Total FII net (Rs cr)", f'=SUM({rng("B")})', CR),
        ("Total DII net (Rs cr)", f'=SUM({rng("C")})', CR),
        ("Correlation: FII flow vs Nifty same day", f'=IFERROR(CORREL({rng("B")},{rng("G")}),"")', "0.00"),
        ("Correlation: FII flow vs Nifty NEXT day", f'=IFERROR(CORREL({rng("B")},{rng("H")}),"")', "0.00"),
        ("Correlation: DII flow vs Nifty same day", f'=IFERROR(CORREL({rng("C")},{rng("G")}),"")', "0.00"),
        ("Nifty avg move on BIG FII BUY days", f'=IFERROR(AVERAGEIFS({rng("G")},{rng("B")},">="&{th}),"")', PCT),
        ("Nifty avg move on BIG FII SELL days", f'=IFERROR(AVERAGEIFS({rng("G")},{rng("B")},"<="&-{th}),"")', PCT),
        ("Nifty avg NEXT-day move after BIG FII SELL", f'=IFERROR(AVERAGEIFS({rng("H")},{rng("B")},"<="&-{th}),"")', PCT),
        ("Nifty avg move on BIG DII BUY days", f'=IFERROR(AVERAGEIFS({rng("G")},{rng("C")},">="&{th}),"")', PCT),
        ("Days DII absorbed FII selling", f'=COUNTIF({rng("K")},"DII bought FII selling")', "0"),
        ("Days FII absorbed DII selling", f'=COUNTIF({rng("K")},"FII bought DII selling")', "0"),
        ("Nifty avg move when DII absorbed FII selling",
         f'=IFERROR(AVERAGEIFS({rng("G")},{rng("K")},"DII bought FII selling"),"")', PCT),
    ]
    ws["M4"] = "Summary"
    ws["M4"].font, ws["M4"].fill = F_HEAD, FILL_HEAD
    ws["N4"].fill = FILL_HEAD
    for i, (label, f, fmt) in enumerate(items):
        r = 5 + i
        ws.cell(r, 13, label).font = F_BASE
        c = ws.cell(r, 14, f)
        c.font, c.number_format, c.border = F_BOLD, fmt, BOX
    ws.cell(5 + len(items) + 1, 13, "Long-term FII/DII history (2007 onwards) is needed to link flows to each event. "
                                    "Download it from NSE/NSDL and paste here.").font = F_NOTE
    widths(ws, {"A": 12, "B": 12, "C": 12, "D": 9, "E": 9, "F": 11, "G": 11, "H": 11, "I": 12, "J": 12,
                "K": 22, "L": 3, "M": 44, "N": 12})


# ---------------------------------------------------------------- target adjuster
def build_adjuster(wb, last_se, end_sc, cat_list, example):
    ws = wb.create_sheet("Target_Adjuster")
    title(ws, "Target Adjuster — adjust your mathematical targets for the emotion of an upcoming event",
          "Enter the yellow cells. Your software's targets stay the base; this sheet adds the historical event effect.")
    stk, cat = se_rng("A", last_se), se_rng("D", last_se)
    sc = lambda c: f"Stock_Scorecard!${c}$5:${c}${end_sc}"  # noqa: E731
    cs = lambda c: f"Category_Summary!${c}$5:${c}${cat_list.split('$')[-1]}"  # noqa: E731
    inputs = [("Stock (NSE code)", example["stock"], None), ("Upcoming event type", example["category"], None),
              ("Current market price (CMP)", example["cmp"], RS),
              ("Base short-term target (from your model)", example["st"], RS),
              ("Base long-term target (from your model)", example["lt"], RS)]
    header(ws, 3, ["INPUTS", "Value"])
    for i, (label, v, fmt) in enumerate(inputs):
        r = 4 + i
        ws.cell(r, 1, label).font = F_BASE
        c = ws.cell(r, 2, v)
        c.font, c.fill, c.border = F_INPUT, FILL_INPUT, BOX
        if fmt:
            c.number_format = fmt
    ws["C6"] = "Example values (last close and model-style targets): replace with your own."
    ws["C6"].font = F_NOTE
    dv1 = DataValidation(type="list", formula1=f"={sc('A')}", allow_blank=False)
    dv2 = DataValidation(type="list", formula1=f"={cat_list}", allow_blank=False)
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    dv1.add("B4")
    dv2.add("B5")

    crit = f"{stk},$B$4,{cat},$B$5"
    rows = [
        ("HISTORY — this stock in this event type", None, None),
        ("Times measured", f"=COUNTIFS({crit})", "0"),
        ("Avg event-day move", f'=IFERROR(AVERAGEIFS({se_rng("H", last_se)},{crit}),"")', PCT),
        ("Avg total move (20 days)", f'=IFERROR(AVERAGEIFS({se_rng("K", last_se)},{crit}),"")', PCT),
        ("Avg worst fall (20 days)", f'=IFERROR(AVERAGEIFS({se_rng("N", last_se)},{crit}),"")', PCT),
        ("Avg days to recover", f'=IFERROR(AVERAGEIFS({se_rng("P", last_se)},{crit}),"")', "0"),
        ("R:R when bought on event day", f'=IFERROR(AVERAGEIFS({se_rng("R", last_se)},{crit})/AVERAGEIFS({se_rng("S", last_se)},{crit}),"")', "0.00"),
        ("MARKET — Nifty in this event type", None, None),
        ("Nifty avg total move (20 days)", f'=IFERROR(INDEX({cs("F")},MATCH($B$5,{cs("A")},0)),"")', PCT),
        ("Nifty avg worst fall (20 days)", f'=IFERROR(INDEX({cs("I")},MATCH($B$5,{cs("A")},0)),"")', PCT),
        ("Direction points", f'=IFERROR(INDEX({cs("W")},MATCH($B$5,{cs("A")},0)),"")', "0"),
        ("Risk points", f'=IFERROR(INDEX({cs("X")},MATCH($B$5,{cs("A")},0)),"")', "0"),
        ("Best entry (history)", f'=IFERROR(INDEX({cs("V")},MATCH($B$5,{cs("A")},0)),"")', None),
        ("Stock Event Beta", f'=IFERROR(INDEX({sc("O")},MATCH($B$4,{sc("A")},0)),"")', "0.00"),
        ("RESULT", None, None),
        ("Method used", f'=IF(B10>={S["minev"]},"Stock\'s own history","Event Beta x Nifty (too few stock events)")', None),
        ("Expected event move", f'=IFERROR(IF(B10>={S["minev"]},B12,B22*B17),0)', PCT),
        ("Expected worst dip", f'=IFERROR(IF(B10>={S["minev"]},B13,B22*B18),0)', PCT),
        ("Adjusted short-term target", f'=B7*(1+B25*{S["stw"]})', RS),
        ("Adjusted long-term target", f'=B8*(1+B25*{S["ltw"]})', RS),
        ("Suggested stop-loss", f'=B6*(1+MIN(B26,0)*{S["slm"]})', RS),
        ("R:R from CMP (target gain / stop-loss risk)", '=IFERROR((B27-B6)/(B6-B29),"")', "0.00"),
        ("Short-term target change vs base", '=IFERROR(B27/B7-1,"")', PCT),
    ]
    for i, (label, f, fmt) in enumerate(rows):
        r = 9 + i
        if f is None:
            header(ws, r, [label, ""], fill=FILL_SUB, font=F_BOLD, height=18)
            continue
        ws.cell(r, 1, label).font = F_BASE
        c = ws.cell(r, 2, f)
        c.font, c.border = F_BOLD, BOX
        if fmt:
            c.number_format = fmt
        if r >= 24:
            c.fill = FILL_OUT
    ws["C25"] = "Positive = event historically lifted this stock; negative = it pulled it down."
    ws["C29"] = "Based on the typical worst dip after this event type (Settings: stop-loss buffer)."
    for a in ("C25", "C29"):
        ws[a].font = F_NOTE
    widths(ws, {"A": 46, "B": 22, "C": 80})


# ---------------------------------------------------------------- how to use
def build_howto(wb, n_events, n_stocks, first, last):
    ws = wb.create_sheet("How_To_Use", 0)
    title(ws, "Event Impact Analyzer — measuring the market's FEAR and GREED around news")
    lines = [
        ("What this does", None),
        ("", f"Measures how Nifty and your {n_stocks} watchlist stocks reacted to {n_events} events between "
             f"{first} and {last}: elections, budgets, RBI rate decisions, crude oil shocks, global and "
             "geopolitical shocks. It turns market emotion into numbers your mathematical model can use."),
        ("Phases measured", None),
        ("BEFORE", "Run-up in the 20 and 5 trading days before the event (did the market anticipate it?)."),
        ("DURING", "Move on the event day (T0) and the intraday worst point."),
        ("AFTER", "Moves over the next 5, 20 and 60 days; worst fall and best rise within 20 days."),
        ("SETTLING", "Days until the price is back at the pre-event close, and days until India VIX settles."),
        ("R:R", "If you bought 5 days before / on the event day / 5 days after: best gain vs worst dip over 20 days."),
        ("Market mood", "FEAR / GREED / CALM per event, from Nifty's reaction (thresholds on the Settings sheet)."),
        ("Sheets", None),
        ("Settings", "Yellow input cells: fear/greed thresholds, points scale, target weights."),
        ("Event_Calendar", "Every event with the news context and whether it was a surprise."),
        ("Nifty_Impact", "One row per event with all phase measurements for Nifty (Sensex before Sep-2007)."),
        ("Category_Summary", "Averages per event type, positive vs negative news, fear vs greed and who made money. "
                             "Includes Direction and Risk POINTS."),
        ("Stock_Scorecard", "Each watchlist stock: behaviour in fear and greed, Event Beta, resilience and impact ranks."),
        ("Top_Lists", "Top 25 most resilient stocks and top 25 most impacted stocks."),
        ("Sector_Impact", "Industry x event-type heat map."),
        ("Crude_Ranges", "At what size of crude move (and what crude price level) Nifty and industries react."),
        ("FII_DII", "Your daily FII/DII data linked to Nifty moves; paste more history to extend it."),
        ("Target_Adjuster", "Pick a stock and an upcoming event: get adjusted short/long-term targets and stop-loss."),
        ("Stock_Event_Data, Market_Daily", "Raw measurements that the formulas read (do not edit)."),
        ("Colour code", None),
        ("Yellow fill, blue text", "Inputs you can change."),
        ("Black text", "Formulas or measured data; do not type over them."),
        ("Green / red shading", "Higher / lower values; FEAR rows red, GREED rows green."),
        ("Data and limits", None),
        ("Prices", "Yahoo Finance daily prices (split-adjusted), downloaded with download_data.py. Yahoo's Nifty "
                   "history starts Sep-2007, so the 2004-2007 events use Sensex. One-day data spikes are removed, "
                   "and stock windows containing an unadjusted split/demerger jump are skipped."),
        ("Event dates", "Compiled for this analysis from public knowledge: please verify, especially recent events. "
                        "If an event fell on a holiday or weekend, T0 is the next trading day."),
        ("FII/DII", "Only your Jun-Oct 2026 tracker data is available here; long history must be downloaded from NSE/NSDL."),
        ("Caution", "History shows tendencies, not certainties. Small categories (few events) carry low confidence."),
        ("Updating", None),
        ("", "Add new events to events.csv, then run: python download_data.py --refresh, python analyze.py, "
             "python build_workbook.py."),
    ]
    r = 3
    for a, b in lines:
        if b is None:
            r += 1
            ws.cell(r, 1, a).font = Font(name=FONT, size=11, bold=True, color="1F3864")
        else:
            ws.cell(r, 1, a).font = F_BOLD
            c = ws.cell(r, 2, b)
            c.font = F_BASE
            c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    widths(ws, {"A": 30, "B": 120})


# ---------------------------------------------------------------- main
def main():
    ni = pd.read_csv(DATA / "nifty_impact.csv")
    se = pd.read_csv(DATA / "stock_events.csv")
    md = pd.read_csv(DATA / "market_daily.csv")
    md.columns = [c.replace("/", "_") for c in md.columns]
    watch = pd.read_csv(HERE / "watchlist.csv")

    meta = watch.set_index("NSE_Code")
    se["Industry"] = se["Stock"].map(meta["Industry"])
    se["Cap"] = se["Stock"].map(meta["Cap"])
    ev = ni.set_index("Event_ID")
    se["Category"] = se["Event_ID"].map(ev["Category"])
    se["Nifty_Day0"] = se["Event_ID"].map(ev["Day0"])
    se["Nifty_Total20"] = se["Event_ID"].map(ev["Total20"])

    order = ["Central Election", "Exit Poll", "State Election", "Union Budget", "Interim Budget",
             "RBI Rate Hike", "RBI Rate Cut", "RBI Pause", "Crude Spike", "Crude Crash",
             "Global Shock", "Geopolitical", "Domestic Shock", "Policy Shock"]
    categories = [c for c in order if c in set(ni["Category"])] + \
                 sorted(set(ni["Category"]) - set(order))
    industries = sorted(watch["Industry"].dropna().unique())

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_settings(wb)
    build_calendar(wb, ni)
    last_ni = build_nifty_impact(wb, ni)
    cat_list = build_category_summary(wb, categories, last_ni)
    last_se = build_stock_events(wb, se, last_ni)
    end_sc = build_scorecard(wb, watch, last_se)
    build_top_lists(wb, end_sc)
    build_sector(wb, industries, categories, last_se)
    mdc, last_md = build_market_daily(wb, md)
    build_crude(wb, mdc, last_md)
    build_fii(wb, mdc, last_md)

    px = pd.read_csv(DATA / "prices" / "RELIANCE.csv")
    cmp_ = round(float(px["Close"].iloc[-1]), 2)
    build_adjuster(wb, last_se, end_sc, cat_list,
                   {"stock": "RELIANCE", "category": "RBI Rate Hike", "cmp": cmp_,
                    "st": round(cmp_ * 1.08, 2), "lt": round(cmp_ * 1.25, 2)})
    first = pd.Timestamp(ni["Trading_Day"].min()).strftime("%b-%Y")
    last = pd.Timestamp(ni["Trading_Day"].max()).strftime("%b-%Y")
    build_howto(wb, len(ni), len(watch), first, last)

    order_sheets = ["How_To_Use", "Settings", "Target_Adjuster", "Category_Summary", "Top_Lists",
                    "Stock_Scorecard", "Sector_Impact", "Crude_Ranges", "FII_DII", "Event_Calendar",
                    "Nifty_Impact", "Stock_Event_Data", "Market_Daily"]
    wb._sheets = [wb[n] for n in order_sheets]
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
    wb.save(OUT)
    print(f"Saved {OUT.name}: {len(ni)} events, {len(se)} stock-event rows")


if __name__ == "__main__":
    main()
