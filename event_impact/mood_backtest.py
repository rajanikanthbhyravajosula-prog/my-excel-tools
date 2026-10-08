"""Mood study for MOOD_STUDY_REPORT.md: how long each Panic Meter zone lasts, how zones change,
fear episodes, and back-tests of simple mood / trend rules on Nifty (no dividends or costs).
Run after analyze.py:  python mood_backtest.py
"""
import pandas as pd, numpy as np
from pathlib import Path
md=pd.read_csv(Path(__file__).parent / "data" / "market_daily.csv",parse_dates=["Date"]).set_index("Date")
S={"VIX":(12,40),"Drawdown":(0,-0.30),"Return20":(0.05,-0.15),"Breadth":(0.20,0.90),"Rupee":(-0.01,0.04)}
sc=pd.DataFrame({k:((md[k]-a)/(b-a)).clip(0,1)*100 for k,(a,b) in S.items()})
p=sc.mean(axis=1).where(sc.notna().sum(axis=1)>=3)
md["Panic"]=p
ps=p.rolling(5,min_periods=1).mean()  # smoothed to avoid daily flicker
zones=pd.cut(ps,[-1,20,40,60,80,101],labels=["Complacent","Calm","Worried","Fear","Panic"])
# run lengths
runs=(zones!=zones.shift()).cumsum()
r=pd.DataFrame({"z":zones,"g":runs}).dropna().groupby("g").agg(z=("z","first"),n=("z","size"))
print("== zone episode lengths (trading days, 5-day smoothed)")
print(r.groupby("z",observed=True)["n"].describe()[["count","mean","50%","max"]].round(0))
# time spent per visit of 'fear or worse' episodes (>=60), with gaps <10 days merged
hi=(ps>=60).astype(int)
ep=[];start=None;last=None
for d,v in hi.items():
    if v:
        if start is None or (last is not None and (md.index.get_loc(d)-md.index.get_loc(last))>10):
            if start is not None: ep.append((start,last))
            start=d
        last=d
if start is not None: ep.append((start,last))
print("== FEAR episodes (panic>=60, gaps<10d merged)")
for a,b in ep:
    ia,ib=md.index.get_loc(a),md.index.get_loc(b)
    nn=md["Nifty"]; pk=p[a:b].idxmax()
    f250=nn.iloc[min(ib+250,len(nn)-1)]/nn.iloc[ib]-1 if ib+250<len(nn) else np.nan
    print(a.date(),b.date(),"days",ib-ia+1,"peakPanic",round(p[a:b].max()),pk.date(),"Nifty at peak panic vs 1y later",
          round(nn.iloc[min(md.index.get_loc(pk)+250,len(nn)-1)]/nn[pk]-1,3) if md.index.get_loc(pk)+250<len(nn) else None)
# transition: zone today -> zone 20 days later
z20=zones.shift(-20)
print("== transition zone -> zone 20d later (%)")
print((pd.crosstab(zones,z20,normalize="index")*100).round(0))
# backtests
n=md["Nifty"]; ret=n.pct_change().fillna(0)
def stats(w,name):
    eq=(1+w.shift(1).fillna(0)*ret).cumprod()
    yrs=len(eq)/250; cagr=eq.iloc[-1]**(1/yrs)-1; dd=(eq/eq.cummax()-1).min()
    print(f"{name:55s} CAGR {cagr:6.1%}  MaxDD {dd:6.1%}  avg exposure {w.mean():.0%}  final x{eq.iloc[-1]:.2f}")
print("== backtests on Nifty price index", n.index[0].date(), n.index[-1].date(), "(no dividends, no costs)")
stats(pd.Series(1.0,index=n.index),"Buy and hold 100%")
w=pd.Series(np.select([ps>=60,ps>=40,ps>=20],[1.0,0.8,0.6],0.4),index=n.index); stats(w,"Mood scaling: 40% complacent .. 100% fear")
w2=pd.Series(np.nan,index=n.index); w2[ps>=60]=1.0; w2[ps<15]=0.5; w2=w2.ffill().fillna(0.5); stats(w2,"Switch: 100% when panic>=60, 50% when panic<15")
ma=n.rolling(200).mean(); w3=(n>ma).astype(float); stats(w3,"Trend only: 100% above 200DMA else 0")
w4=pd.Series(np.where(n>ma,1.0,np.where(ps>=60,1.0,0.3)),index=n.index); stats(w4,"Trend + fear: above 200DMA or panic>=60 -> 100%, else 30%")
# fear fading signal: panic crosses below 50 after having been >=70 within last 40 days
peak40=p.rolling(40).max(); sig=(p<50)&(p.shift(1)>=50)&(peak40>=70)
sd=sig[sig].index
print("== 'fear fading' signals: panic falls back below 50 after >=70 in last 40 days")
for d in sd:
    i=n.index.get_loc(d)
    out=[round(n.iloc[i+h]/n.iloc[i]-1,3) if i+h<len(n) else None for h in (20,60,250)]
    print(d.date(), "Nifty", round(n.iloc[i]), "next 20/60/250d", out)
# current
print("== current", p.index[-1].date(), "panic", round(p.iloc[-1],1), "smoothed", round(ps.iloc[-1],1), sc.iloc[-1].round(0).to_dict())
print("recent panic monthly max:", p.resample("ME").max()["2026-01":].round(0).to_dict())
print("\n== with cash earning 6% a year (liquid fund / FD assumption)")
cash=(1.06)**(1/250)-1
def stats2(w,name):
    w=w.shift(1).fillna(0)
    eq=(1+w*ret+(1-w)*cash).cumprod()
    yrs=len(eq)/250; cagr=eq.iloc[-1]**(1/yrs)-1; dd=(eq/eq.cummax()-1).min()
    print(f"{name:60s} CAGR {cagr:6.1%}  MaxDD {dd:6.1%}  avg exposure {w.mean():.0%}")
stats2(pd.Series(1.0,index=n.index),"Buy and hold 100%")
stats2(w3,"Trend only: 100% above 200DMA else cash")
stats2(pd.Series(np.where(n>ma,1.0,0.5),index=n.index),"Trend half: 100% above 200DMA else 50%")
stats2(pd.Series(np.where(n>ma,1.0,np.where(ps>=70,1.0,0.5)),index=n.index),"Trend+panic: above 200DMA or panic>=70 -> 100%, else 50%")
stats2(pd.Series(np.where(n>ma,1.0,np.where(ps>=70,1.0,0.0)),index=n.index),"Trend+panic: above 200DMA or panic>=70 -> 100%, else cash")
# SIP style: monthly fixed investment vs 'double SIP when panic>=60'
print("\n== Monthly SIP of 1 unit vs 'SIP + extra 2 units when panic>=60 that month' (value per unit invested)")
m=n.resample("ME").last(); pm=ps.resample("ME").max()
units=0;inv=0;units2=0;inv2=0
for d,px in m.items():
    units+=1/px; inv+=1
    k=3 if pm[d]>=60 else 1
    units2+=k/px; inv2+=k
print("plain SIP: value/invested", round(units*n.iloc[-1]/inv,2), " fear-boosted SIP:", round(units2*n.iloc[-1]/inv2,2))
