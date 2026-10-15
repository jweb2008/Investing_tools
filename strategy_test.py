"""
Strategy test: development / holdout check for the growth and momentum
profiles before either changes the live watchlist.

  py strategy_test.py --download   refresh cached history for both universes first
                                   (the momentum universe is larger: ~20-30 min the first time)
  py strategy_test.py              run the test on cached history (a few minutes)

Rules are fixed in advance (below). Each strategy is tried with two exit
rules on the DEVELOPMENT period only (signals before 2025). The exit with the
best development Avg R is chosen, and only that choice is then scored ONCE on
the HOLDOUT period (2025 onward), which played no part in choosing it.

  Development pass  Avg R >= +0.10 and profit factor >= 1.2 (100+ trades)
  Holdout pass      Avg R > 0 and profit factor >= 1.1
  A strategy goes live only if it passes both.

Strategies
  growth_current   today's live rules: uptrend, within 5% below the pivot,
                   2+ of 3 contraction checks. Buy-stop at the pivot for 5 days.
  growth_demand    uptrend, within 5% below the pivot, PLUS buyers showing up:
                   up/down volume ratio >= 1.2 and either a pocket pivot or the
                   RS line at a 52-week high in the last 10 days. No contraction
                   checks (no reward for dry volume). Buy-stop at the pivot.
  mom_breakout     strong stock breaking out on volume: close > 20 SMA > 50 SMA,
                   RS rating 80+, within 10% of the 52-week high, close above the
                   prior 20-day high on 1.5x+ volume, closing in the top 40% of the
                   day's range. Enter next day's open.
  mom_pullback     leader pulling back to the 10-day EMA: RS 80+, up 10%+ in a
                   month, above the 50 SMA, a new 20-day high in the last 10 days,
                   today's low within 1% of the 10-day EMA, closing above it and
                   up on the day, on lighter volume. Enter next day's open.

Exits
  growth: "target" (live rules: conservative target, 20 days) and "trail20"
          (breakeven at +1R, then a close below the 20-day SMA, 60 days)
  momentum: "ema10" (breakeven at +1R, then a close below the 10-day EMA,
          10 days) and "2R" (fixed 2R target, 10 days)
  Momentum stop: the signal day's low, no wider than 2 ATR and no tighter than
  1 ATR below entry. A gap below the stop at the open skips the trade.

Every trade is charged 0.1% slippage on entry and on exit.

Same limits as backtest.py: today's listed stocks only (survivorship bias,
stronger for small caps), no earnings dates, no commissions.
"""
import argparse
import datetime as dt
import os
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

import backtest
import rs
import scanner
import signals

HERE = Path(__file__).parent
GROWTH_UNIVERSE = HERE / "data" / "universe.txt"
MOMENTUM_UNIVERSE = HERE / "data" / "universe_momentum.txt"
REPORT = HERE / "Strategy Test.xlsx"
TRADES_CSV = HERE / "data" / "strategy_trades.csv"

HOLDOUT_START = pd.Timestamp("2025-01-01")
SLIPPAGE = 0.001
MIN_DEV_TRADES = 100
DEV_PASS = {"avg_r": 0.10, "pf": 1.2}
HOLD_PASS = {"avg_r": 0.0, "pf": 1.1}

STRATEGIES = {
    "growth_current": {"kind": "growth", "exits": ["target", "trail20"]},
    "growth_demand": {"kind": "growth", "exits": ["target", "trail20"]},
    "mom_breakout": {"kind": "momentum", "exits": ["ema10", "2R"]},
    "mom_pullback": {"kind": "momentum", "exits": ["ema10", "2R"]},
}
MOM_EXITS = {"ema10": {"target_r": None, "max_hold": 10},
             "2R": {"target_r": 2.0, "max_hold": 10}}


# ------------------------------------------------------------ data
def universes():
    growth = backtest.load_symbols(GROWTH_UNIVERSE)
    if MOMENTUM_UNIVERSE.exists():
        mom = backtest.load_symbols(MOMENTUM_UNIVERSE)
    else:
        print("WARNING: no momentum universe yet (py build_momentum_universe.py). "
              "Momentum strategies will use the growth universe.")
        mom = growth
    return growth, mom


def prepare():
    growth, mom = universes()
    spy = backtest.load_history("SPY")
    if spy is None:
        sys.exit("No cached history. Run: py strategy_test.py --download")
    spy_close = spy["close"]
    syms = sorted(set(growth) | set(mom))
    print(f"Computing indicators for {len(syms)} stocks...")
    data, comp = {}, {}
    for s in syms:
        df = backtest.load_history(s)
        if df is None or len(df) < scanner.MIN_BARS:
            continue
        ind = backtest.indicator_frame(df)
        dem = signals.demand_features(df, spy_close)
        mf = signals.momentum_features(df)
        data[s] = (df, ind, dem, mf)
        j = pd.concat([df["close"].rename("s"), spy_close.rename("m")], axis=1, join="inner")
        rel = {k: (1 + j["s"].pct_change(n)) / (1 + j["m"].pct_change(n)) - 1 for k, n in rs.LOOKBACKS.items()}
        comp[s] = rs.WEIGHTS["rs_3m"] * rel["rs_3m"] + rs.WEIGHTS["rs_6m"] * rel["rs_6m"]
    if not data:
        sys.exit("No usable history. Run: py strategy_test.py --download")
    rating = (pd.DataFrame(comp).rank(axis=1, pct=True) * 98 + 1).round()
    start = spy.index.min() + pd.Timedelta(days=400)
    return {"data": data, "rating": rating, "growth": set(growth), "momentum": set(mom),
            "start": start, "last": spy.index.max()}


# ------------------------------------------------------------ signals
def growth_mask(name, ind, dem):
    if name == "growth_current":
        return ind["required_ok"] & (ind["checks"] >= scanner.MIN_CHECKS)
    return (ind["required_ok"] & (dem["updown_vol"] >= signals.UPDOWN_MIN)
            & (dem["pp_recent"] | dem["rs_line_high"]))


def momentum_mask(name, df, ind, dem, mf, rsr):
    base = ind["passes_filters"] & (rsr >= 80)
    c = df["close"]
    if name == "mom_breakout":
        return (base & (c > mf["sma20"]) & (mf["sma20"] > mf["sma50"]) & (mf["pct_from_high"] >= 0.90)
                & (c > mf["high20_prior"]) & (dem["rvol"] >= 1.5) & (mf["range_pos"] >= 0.6))
    return (base & (mf["ret21"] >= 0.10) & (c > mf["sma50"]) & mf["new20_recent"]
            & (df["low"] <= mf["ema10"] * 1.01) & (c > mf["ema10"]) & (c > df["open"])
            & (dem["rvol"] < 1.2))


# ------------------------------------------------------------ simulation
def simulate_growth(sym, df, ind, mask, start, exit_mode):
    ind2 = ind.copy()
    ind2["required_ok"] = mask.fillna(False)
    trades = backtest.simulate(sym, df, ind2, start, exit_mode)
    for t in trades:
        t["Status"] = "Signal"
    return trades


def simulate_momentum(sym, df, ind, mf, mask, start, exit_mode):
    cfg = MOM_EXITS[exit_mode]
    o, h, l, c = (df[k].to_numpy() for k in ("open", "high", "low", "close"))
    ema = mf["ema10"].to_numpy()
    atr = ind["atr14"].to_numpy()
    sig = mask.fillna(False).to_numpy()
    dates = df.index
    n, trades, i = len(df), [], 0
    while i < n - 1:
        if not sig[i] or dates[i] < start:
            i += 1
            continue
        e_i = i + 1
        entry = o[e_i]
        stop = max(l[i], entry - 2 * atr[i])
        if entry - stop < atr[i]:
            stop = entry - atr[i]
        if entry <= stop:             # gapped below the stop: skip
            i += 1
            continue
        risk = entry - stop
        target = entry + cfg["target_r"] * risk if cfg["target_r"] else np.inf
        cur_stop, active, best = stop, False, entry
        exit_i = exit_px = how = None
        for k in range(e_i, min(e_i + cfg["max_hold"], n)):
            best = max(best, h[k])
            if l[k] <= cur_stop:
                exit_i, exit_px = k, min(o[k], cur_stop) if k > e_i else cur_stop
                how = "stop" if cur_stop < entry else "breakeven"
                break
            if h[k] >= target and (k > e_i or c[k] >= target):
                exit_i, exit_px, how = k, max(o[k], target) if k > e_i else target, "target"
                break
            if cfg["target_r"] is None:
                if active and c[k] < ema[k]:
                    exit_i, exit_px, how = k, c[k], "trail"
                    break
                if not active and h[k] >= entry + risk:
                    active, cur_stop = True, max(cur_stop, entry)
        if exit_i is None:
            exit_i = min(e_i + cfg["max_hold"] - 1, n - 1)
            exit_px, how = c[exit_i], "time" if exit_i < n - 1 else "open"
        trades.append({
            "Symbol": sym, "SignalDate": dates[i].date(), "EntryDate": dates[e_i].date(),
            "ExitDate": dates[exit_i].date(), "Exit": how, "ExitRule": exit_mode,
            "Entry": round(entry, 2), "Stop": round(stop, 2),
            "Target": round(target, 2) if np.isfinite(target) else None,
            "ExitPrice": round(exit_px, 2), "R": round((exit_px - entry) / risk, 3),
            "ReturnPct": round((exit_px - entry) / entry * 100, 2),
            "MaxR": round((best - entry) / risk, 2), "HoldDays": int(exit_i - e_i + 1),
        })
        i = exit_i + 1
    return trades


def run_strategy(ctx, name, exit_mode):
    kind = STRATEGIES[name]["kind"]
    pool = ctx["growth"] if kind == "growth" else ctx["momentum"]
    out = []
    for s, (df, ind, dem, mf) in ctx["data"].items():
        if s not in pool:
            continue
        if kind == "growth":
            out += simulate_growth(s, df, ind, growth_mask(name, ind, dem), ctx["start"], exit_mode)
        else:
            rsr = ctx["rating"][s].reindex(df.index) if s in ctx["rating"] else pd.Series(np.nan, df.index)
            out += simulate_momentum(s, df, ind, mf, momentum_mask(name, df, ind, dem, mf, rsr),
                                     ctx["start"], exit_mode)
    t = pd.DataFrame(out)
    if t.empty:
        return t
    t["Strategy"] = name
    # slippage on both sides, expressed in R
    risk = t["Entry"] - t["Stop"]
    t["R"] = t["R"] - (t["Entry"] * SLIPPAGE + t["ExitPrice"] * SLIPPAGE) / risk
    t["Period"] = np.where(pd.to_datetime(t["SignalDate"]) < HOLDOUT_START, "Development", "Holdout")
    return t


# ------------------------------------------------------------ evaluation
def metrics(g, years):
    r = g["R"]
    if len(r) == 0:
        return pd.Series(dtype=float)
    pos, neg = r[r > 0].sum(), -r[r < 0].sum()
    cum = g.sort_values("ExitDate")["R"].cumsum()
    return pd.Series({
        "Trades": len(r), "Trades / yr": len(r) / years,
        "Win %": (r > 0).mean() * 100, "Avg R": r.mean(), "Median R": r.median(),
        "Profit factor": pos / neg if neg > 0 else np.nan,
        "Avg win R": r[r > 0].mean(), "Avg loss R": r[r < 0].mean(),
        "Total R / yr": r.sum() / years,
        "Max drawdown R": (cum - cum.cummax()).min(),
        "Avg hold days": g["HoldDays"].mean(),
    })


def evaluate(ctx):
    dev_years = (HOLDOUT_START - ctx["start"]).days / 365.25
    hold_years = max((ctx["last"] - HOLDOUT_START).days / 365.25, 0.1)
    allt, dev_rows, hold_rows, verdicts = [], {}, {}, []
    for name, spec in STRATEGIES.items():
        best = None
        for ex in spec["exits"]:
            t0 = time.time()
            t = run_strategy(ctx, name, ex)
            print(f"  {name:<15} {ex:<8} {len(t):>6} trades  ({time.time() - t0:.0f}s)")
            if t.empty:
                continue
            allt.append(t)
            d, h = t[t["Period"] == "Development"], t[t["Period"] == "Holdout"]
            dm, hm = metrics(d, dev_years), metrics(h, hold_years)
            dev_rows[(name, ex)] = dm
            hold_rows[(name, ex)] = hm
            if len(d) >= MIN_DEV_TRADES and (best is None or dm["Avg R"] > dev_rows[best]["Avg R"]):
                best = (name, ex)
        if best is None:
            verdicts.append({"Strategy": name, "Chosen exit": "", "Verdict": "Not enough trades"})
            continue
        dm, hm = dev_rows[best], hold_rows[best]
        dev_ok = dm["Avg R"] >= DEV_PASS["avg_r"] and dm["Profit factor"] >= DEV_PASS["pf"]
        hold_ok = len(hm) and hm["Avg R"] > HOLD_PASS["avg_r"] and hm["Profit factor"] >= HOLD_PASS["pf"]
        verdict = ("PASSED" if dev_ok and hold_ok else
                   "Failed holdout" if dev_ok else "Failed development")
        verdicts.append({
            "Strategy": name, "Chosen exit": best[1],
            "Dev trades": int(dm["Trades"]), "Dev Avg R": dm["Avg R"], "Dev PF": dm["Profit factor"],
            "Holdout trades": int(hm["Trades"]) if len(hm) else 0,
            "Holdout Avg R": hm.get("Avg R", np.nan), "Holdout PF": hm.get("Profit factor", np.nan),
            "Holdout Total R / yr": hm.get("Total R / yr", np.nan),
            "Verdict": verdict,
        })
    t = pd.concat(allt) if allt else pd.DataFrame()
    dev = pd.DataFrame(dev_rows).T
    hold = pd.DataFrame(hold_rows).T
    for df_ in (dev, hold):
        if not df_.empty:
            df_.index.names = ["Strategy", "Exit"]
    return pd.DataFrame(verdicts), dev, hold, t


def write_report(verdict, dev, hold, path, meta):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    sheets = {"Verdict": verdict.set_index("Strategy"), "Development": dev, "Holdout (all exits)": hold}
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.round(2).to_excel(xw, sheet_name=name, startrow=3)
            ws = xw.sheets[name]
            ws["A1"] = f"Strategy test: {name}"
            ws["A1"].font = Font(name="Arial", size=14, bold=True)
            ws["A2"] = meta
            ws["A2"].font = Font(name="Arial", size=9, italic=True, color="555555")
            for row in ws.iter_rows(min_row=4):
                for cell in row:
                    cell.font = Font(name="Arial", size=10, bold=cell.row == 4,
                                     color="FFFFFF" if cell.row == 4 else "000000")
                    if cell.row == 4:
                        cell.fill = PatternFill("solid", start_color="1F4E78")
                        cell.alignment = Alignment(horizontal="center", wrap_text=True)
                    elif isinstance(cell.value, str) and cell.value in ("PASSED", "Failed holdout",
                                                                       "Failed development"):
                        cell.fill = PatternFill("solid", start_color="C6EFCE" if cell.value == "PASSED"
                                                else "F8CBAD")
            for c in range(1, df.shape[1] + 3):
                ws.column_dimensions[get_column_letter(c)].width = 15
            ws.row_dimensions[4].height = 30
        notes = pd.DataFrame({"How to read": [l for l in __doc__.strip().splitlines()]})
        notes.to_excel(xw, sheet_name="Rules", index=False)
        xw.sheets["Rules"].column_dimensions["A"].width = 110


def main():
    ap = argparse.ArgumentParser(description="Development / holdout test for growth and momentum")
    ap.add_argument("--download", action="store_true", help="refresh cached history for both universes")
    ap.add_argument("--open", action="store_true", help="open the report when done")
    args = ap.parse_args()
    if args.download:
        g, m = universes()
        backtest.download(symbols=sorted(set(g) | set(m)))
    ctx = prepare()
    print(f"Development: signals {ctx['start'].date()} to {(HOLDOUT_START - pd.Timedelta(days=1)).date()}. "
          f"Holdout: {HOLDOUT_START.date()} to {ctx['last'].date()}.")
    verdict, dev, hold, t = evaluate(ctx)
    TRADES_CSV.parent.mkdir(exist_ok=True)
    if not t.empty:
        t.to_csv(TRADES_CSV, index=False)
    meta = (f"Growth universe {len(ctx['growth'])}, momentum universe {len(ctx['momentum'])} stocks. "
            f"Today's listed stocks only (survivorship bias). 0.1% slippage per side. "
            f"Generated {dt.date.today()}.")
    try:
        write_report(verdict, dev, hold, REPORT, meta)
        where = REPORT
    except PermissionError:
        where = REPORT.with_name(f"Strategy Test ({dt.datetime.now():%H%M}).xlsx")
        write_report(verdict, dev, hold, where, meta)
    pd.set_option("display.width", 220)
    print(f"\n{meta}\n")
    print(verdict.round(2).to_string(index=False))
    print(f"\nFull detail: {where.name}   Every trade: data/{TRADES_CSV.name}")
    if args.open and os.name == "nt":
        os.startfile(where)


if __name__ == "__main__":
    main()
