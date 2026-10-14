"""
Run the breakout scanner from the command line.

  py scan.py                 scan watchlist.txt
  py scan.py --universe      scan universe.txt (S&P 500 + 400 + Nasdaq-100, built weekly)
  py scan.py AAPL            one ticker, with full indicator detail
  py scan.py NVDA AMD        several tickers
  py scan.py --all           show every stock's score, not just 60+
  py scan.py --source yahoo  use Yahoo instead of Schwab for this run

Read-only. Pulls daily price history only.
"""
import argparse
import sys
import time
from pathlib import Path

import pandas as pd

import account
import earnings
import options
import levels
import scan_log
import scanner
from data import DataError, get_source, source_name

HERE = Path(__file__).parent
WATCHLIST = HERE / "watchlist.txt"


def load_watchlist(path=WATCHLIST):
    if not path.exists():
        sys.exit(f"{path.name} not found. Add tickers to it, or run: py scan.py AAPL")
    out = []
    for line in path.read_text().splitlines():
        line = line.split("#", 1)[0].strip().upper()
        if line:
            out.append(line)
    return list(dict.fromkeys(out))  # dedupe, keep order


def print_detail(sym, r):
    print(f"\n{sym}  ({r['date']})")
    print(f"  Close            {r['close']:>10.2f}")
    print(f"  50 SMA           {r['sma50']:>10.2f}")
    print(f"  200 SMA          {r['sma200']:>10.2f}")
    print(f"  Pivot (30d high) {r['pivot']:>10.2f}   {r['pct_below_pivot']:.1f}% below")
    print(f"  Base low (30d)   {r['base_low']:>10.2f}")
    print(f"  BB width pctile  {r['bb_width_pctile']:>9.0f}%   (6-month rank, lower = tighter)")
    print(f"  ATR% 10d / 40d   {r['atr_pct_10']:>6.2f} / {r['atr_pct_prior40']:.2f}   ratio {r['atr_ratio']:.2f}")
    print(f"  ATR(14)          {r['atr14']:>10.2f}")
    print(f"  Volume 10d / 50d {r['vol_10'] / 1e6:>7.2f}M / {r['vol_50'] / 1e6:.2f}M")
    print(f"  Today's volume   {r['today_vol_x']:>9.2f}x 50-day average")
    print("  Score")
    for k, w in scanner.WEIGHTS.items():
        p = r[f"pts_{k}"]
        tag = "(required)" if k in scanner.REQUIRED else ""
        print(f"    {'+' if p else '-'} {k:<13} {p:>3} / {w}  {tag}")
    print(f"    TOTAL         {r['score']:>3} / 100")
    if scanner.is_setup(r):
        print(f"  SETUP: uptrend and near pivot, {r['checks']} of 3 contraction checks")
    elif not r["required_ok"]:
        print("  Not a setup: missing a required check")
    else:
        print(f"  Not a setup: only {r['checks']} of 3 contraction checks "
              f"(needs {scanner.MIN_CHECKS})")
    if r["breakout"]:
        print("  ** BREAKOUT: closed above pivot on 1.5x+ volume **")
    print_levels(r)
    e = r.get("earn")
    if e and e.get("date"):
        warn = {"danger": "  << inside a week", "warn": "  << inside a typical swing hold"}.get(e["level"], "")
        est = " (estimated)" if e["estimate"] else ""
        print(f"  Next earnings    {e['date']:%a %m/%d/%Y}{est}, {e['days']} days out{warn}")
    elif e is not None:
        print("  Next earnings    not scheduled / not available")
    print_options(r)


def _pct(x, entry):
    return f"{(x - entry) / entry * 100:+.1f}%"


def print_levels(r):
    e = r["entry"]
    print("  Levels (planning references, not predictions)")
    print(f"    Entry          {e:>10.2f}   ({'close' if r['breakout'] else 'pivot'})")
    print(f"    Stop           {r['stop']:>10.2f}   {_pct(r['stop'], e)}  ({r['stop_method']})")
    print(f"    Risk/share     {r['risk_per_share']:>10.2f}   {r['risk_pct']:.1f}% of entry")
    print(f"    Measured move  {r['target_measured']:>10.2f}   {_pct(r['target_measured'], e)}")
    res = r["resistance"]
    print(f"    Resistance     {res:>10.2f}   {_pct(res, e)}" if res
          else "    Resistance           none   (no swing high above entry in the last year)")
    print(f"    2R / 3R        {r['target_2r']:>10.2f} / {r['target_3r']:.2f}")
    print(f"    3x ATR         {r['target_atr']:>10.2f}   {_pct(r['target_atr'], e)}")
    flag = "" if r["rr_ok"] else "   << under 2:1"
    print(f"    Target         {r['target']:>10.2f}   {_pct(r['target'], e)}  "
          f"(lower of measured move and resistance)")
    print(f"    Reward:risk    {r['reward_risk']:>10.1f} : 1{flag}")
    if r["shares"] is not None:
        why = "capped by available cash" if r["cash_limited"] else "risk budget / risk per share"
        print(f"    Position size  {r['shares']:>10d} shares  ({why})")


def add_earnings(results, min_checks, single=False):
    """Look up next earnings for setups and breakouts (or the one ticker checked)."""
    want = [r["symbol"] for r in results
            if single or r["breakout"] or scanner.is_setup(r, min_checks)]
    info, failed = {}, 0
    if want and results:
        try:
            info, failed = earnings.next_earnings(want, pd.Timestamp(results[0]["date"]).date())
        except Exception as e:
            print(f"Earnings dates unavailable ({str(e)[:80]}).")
    if failed:
        print(f"Earnings dates: {failed} of {len(want)} lookups failed (Yahoo); those show blank.")
    for r in results:
        r["earn"] = info.get(r["symbol"])


def add_options(results, src, min_checks, single=False):
    """Swing call, debit spread and LEAPS ideas for setups and breakouts (Schwab only)."""
    if getattr(src, "client", None) is None:
        print("Options skipped: option chains need Schwab data.")
        return
    want = [r for r in results if single or r["breakout"] or scanner.is_setup(r, min_checks)]
    if len(want) > 1:
        print(f"Looking up option chains for {len(want)} setups...")
    fails = 0
    for r in want:
        try:
            r.update(options.ideas(src, r["symbol"], r, (r.get("earn") or {}).get("date")))
        except DataError:
            raise
        except Exception as e:
            r["opt_error"] = str(e)[:80]
        if r.get("opt_error"):
            fails += 1
    if fails:
        print(f"Options: {fails} of {len(want)} chain lookups had problems (blank in the table).")


def print_options(r):
    if not r.get("opt_contract") and not r.get("leaps_contract"):
        if r.get("opt_error"):
            print(f"  Options          unavailable ({r['opt_error']})")
        return
    f = lambda v, d=2: "n/a" if v is None or v != v else f"{v:,.{d}f}"
    if r.get("opt_contract"):
        print("  Swing call (planning estimates, not quotes)")
        print(f"    Contract       {r['opt_contract']}   {r['opt_dte']} days")
        print(f"    Bid / ask      {f(r['opt_bid'])} / {f(r['opt_ask'])}   mid {f(r['opt_mid'])}  "
              f"(spread {f(r['opt_spread_pct'], 1)}%)")
        print(f"    Delta / IV     {f(r['opt_delta'])} / {f(r['opt_iv'] * 100 if r['opt_iv'] == r['opt_iv'] else None, 0)}%"
              f"   20-day realized {f(r['hv20'] * 100 if r['hv20'] == r['hv20'] else None, 0)}%   OI {r['opt_oi']:,}")
        print(f"    Breakeven      {f(r['opt_breakeven'])}  ({f(r['opt_be_pct'], 1)}% above close)")
        print(f"    Est. value     at entry {f(r['opt_val_entry'])}, at target {f(r['opt_val_target'])}, "
              f"at stop {f(r['opt_val_stop'])}  (after ~{options.HOLD_DAYS} days)")
        print(f"    Option R:R     {f(r['opt_rr'], 1)} : 1")
        if r.get("opt_tradeable") is False:
            print(f"    Option Score     0 / 100  NOT TRADEABLE ({r['opt_gate_reason']})")
        elif "opt_score" in r:
            print(f"    Option Score   {r['opt_score']:>3} / 100  (payoff {r['opt_pts_payoff']}/35, "
                  f"vol price {r['opt_pts_vol']}/25, cost {r['opt_pts_cost']}/25, "
                  f"liquidity {r['opt_pts_liquidity']}/15)")
        if r.get("opt_flags"):
            print(f"    Flags          {r['opt_flags']}")
        if r.get("spr_debit"):
            print(f"  Debit spread     buy {r['opt_strike']:.2f} / sell {r['spr_short_strike']:.2f} call, "
                  f"debit {f(r['spr_debit'])}, max gain {f(r['spr_max_gain'])} "
                  f"({f(r['spr_max_rr'], 1)}:1), est. value at target {f(r['spr_val_target'])}")
    if r.get("leaps_contract"):
        print(f"  LEAPS            {r['leaps_contract']}   {r['leaps_dte']} days, mid {f(r['leaps_mid'])}, "
              f"delta {f(r['leaps_delta'])}, {f(r['leaps_extrinsic_pct'], 0)}% time value, "
              f"breakeven {f(r['leaps_breakeven'])}")


def market_check(src):
    """SPY and QQQ: True if the last close is above the 50-day SMA, None if unavailable."""
    out = {}
    for sym in ("SPY", "QQQ"):
        try:
            c = src.history(sym)["close"]
            out[sym] = bool(c.iloc[-1] > c.rolling(50).mean().iloc[-1]) if len(c) >= 50 else None
        except DataError:
            raise
        except Exception:
            out[sym] = None
    return out


def describe_market(m):
    word = {True: "above", False: "BELOW", None: "unavailable vs"}
    line = f"Market: SPY {word[m.get('SPY')]} 50-day, QQQ {word[m.get('QQQ')]} 50-day"
    if False in m.values():
        line += "  (caution: breakouts fail more often in a weak market)"
    return line


def main():
    ap = argparse.ArgumentParser(description="Breakout setup scanner (daily bars)")
    ap.add_argument("tickers", nargs="*", help="tickers to scan (default: watchlist.txt)")
    ap.add_argument("--universe", action="store_true",
                    help="scan universe.txt (built weekly by build_universe.py)")
    ap.add_argument("--all", action="store_true", help="show every stock scored, not just setups")
    ap.add_argument("--source", help="schwab or yahoo (overrides DATA_SOURCE)")
    ap.add_argument("--min-checks", type=int, default=scanner.MIN_CHECKS, choices=[1, 2, 3],
                    help="contraction checks a setup needs (default 2)")
    ap.add_argument("--no-log", action="store_true", help="don't write to the scan log")
    ap.add_argument("--no-options", action="store_true", help="skip the option chain lookups")
    args = ap.parse_args()

    if args.universe:
        uni = HERE / "data" / "universe.txt"
        if not uni.exists():
            sys.exit("universe.txt not found. Run: py build_universe.py")
        age = (time.time() - uni.stat().st_mtime) / 86400
        if age > 8:
            print(f"WARNING: universe.txt is {age:.0f} days old. Run: py build_universe.py")
        tickers = load_watchlist(uni)
        list_name = "universe"
    else:
        list_name = "custom" if args.tickers else "watchlist"
        tickers = [t.upper() for t in args.tickers] or load_watchlist()
    try:
        src = get_source(args.source)
    except DataError as e:
        sys.exit(f"SCAN FAILED: {e}")

    market = {}
    if len(tickers) > 1:
        try:
            market = market_check(src)
        except DataError as e:
            sys.exit(f"SCAN FAILED: {e}")
        print(describe_market(market))

    sizing, note = account.sizing_account(src)
    print(account.describe(sizing) if sizing else note)

    print(f"Scanning {len(tickers)} ticker(s) with {src.name} data...")
    started = time.time()
    results, skipped, errors = [], [], []
    for i, sym in enumerate(tickers, 1):
        try:
            df = src.history(sym)
        except DataError as e:
            sys.exit(f"\nSCAN FAILED: {e}")
        except Exception as e:
            errors.append((sym, str(e)[:100]))
            continue
        reason = scanner.check_filters(df)
        if reason:
            skipped.append((sym, reason))
            continue
        r = scanner.score(df)
        r["hv20"] = options.realized_vol(df["close"])
        results.append({"symbol": sym, **r, **levels.compute(df, r, sizing)})
        if len(tickers) > 20 and i % 100 == 0:
            print(f"  {i}/{len(tickers)}...")

    add_earnings(results, args.min_checks, single=len(tickers) == 1)
    if not args.no_options:
        add_options(results, src, args.min_checks, single=len(tickers) == 1)

    if len(tickers) == 1 and results:
        print_detail(tickers[0], results[0])
    elif results:
        df = pd.DataFrame(results).sort_values(["score", "pct_below_pivot"], ascending=[False, True])
        df["setup"] = [scanner.is_setup(r, args.min_checks) for r in df.to_dict("records")]
        df["R:R"] = [f"{x:.1f}" + ("" if ok else " <2") for x, ok in zip(df["reward_risk"], df["rr_ok"])]
        df["Earnings"] = [earnings.label(e) for e in df["earn"]]
        shown = df if args.all else df[df["setup"]]
        cols = {"symbol": "Symbol", "score": "Score", "checks": "Chk", "close": "Close",
                "pivot": "Pivot", "pct_below_pivot": "% Below", "stop": "Stop",
                "risk_pct": "Risk%", "target": "Target", "target_pct": "Tgt%", "R:R": "R:R",
                "Earnings": "Earnings"}
        if df["shares"].notna().any():
            df["Shares"] = [f"{int(s)}{'*' if c else ''}" if s == s and s is not None else ""
                            for s, c in zip(df["shares"], df["cash_limited"])]
            shown = df if args.all else df[df["setup"]]
            cols["Shares"] = "Shares"
        if len(shown):
            title = ("All scores" if args.all else
                     f"{len(shown)} setups (uptrend, within 5% of pivot, "
                     f"{args.min_checks}+ of 3 contraction checks)")
            print(f"\n{title}:")
            print(shown[list(cols)].rename(columns=cols).to_string(
                index=False, float_format=lambda x: f"{x:.2f}"))
        else:
            print("\nNo setups today.")
        near = df[df["required_ok"] & ~df["setup"]]
        if len(near) and not args.all:
            print(f"\n{len(near)} near misses (uptrend and near pivot, fewer checks) "
                  "are in the scan log, not shown here.")
        brk = df[df["breakout"]]
        if len(brk):
            print("\nBreakouts (uptrend, close above pivot on 1.5x+ volume; entry = close):")
            bcols = {"symbol": "Symbol", "close": "Close", "pivot": "Pivot", "today_vol_x": "Vol x",
                     "stop": "Stop", "risk_pct": "Risk%", "target": "Target", "R:R": "R:R",
                     "Earnings": "Earnings"}
            print(brk[list(bcols)].rename(columns=bcols).to_string(
                index=False, float_format=lambda x: f"{x:.2f}"))
        if "opt_contract" in df and not args.all:
            o = shown[shown["opt_contract"].notna()].copy() if "opt_contract" in shown else shown.iloc[0:0]
            if len(o):
                o["opt_oi"] = o["opt_oi"].astype("Int64")
                o["OptScore"] = [("no trade" if t is False else int(s)) if s == s else ""
                                 for s, t in zip(o["opt_score"], o["opt_tradeable"])]
                o = o.sort_values("opt_score", ascending=False)
            if len(o):
                print("\nSwing calls by Option Score (estimates; full detail with  py scan.py SYMBOL):")
                ocols = {"symbol": "Symbol", "OptScore": "Opt Score", "opt_contract": "Call", "opt_mid": "Mid",
                         "opt_delta": "Delta", "opt_be_pct": "BE %", "opt_spread_pct": "B/A %",
                         "opt_oi": "OI", "opt_iv_hv": "IV/HV", "opt_rr": "Opt R:R", "opt_flags": "Flags"}
                print(o[list(ocols)].rename(columns=ocols).to_string(
                    index=False, float_format=lambda x: f"{x:.2f}"))
        flagged = shown[[bool(e) and e.get("level") != "" for e in shown["earn"]]]
        if len(flagged):
            print(f"\nEarnings within {earnings.WARN_DAYS} days: {len(flagged)} of {len(shown)} setups "
                  f"(! = within {earnings.DANGER_DAYS} days, * = within {earnings.WARN_DAYS}, "
                  "~ = estimated date). A gap through your stop is the main risk on a multi-week hold.")
        if len(shown) and shown["cash_limited"].any():
            print("* = share count capped by available cash, not by the risk budget.")
        if not args.all and len(shown) and (~shown["rr_ok"]).any():
            print(f"\n<2 = reward:risk under 2:1 to the conservative target "
                  f"({(~shown['rr_ok']).sum()} of {len(shown)} setups).")

    for sym, reason in skipped:
        print(f"  skipped {sym}: {reason}")
    for sym, err in errors:
        print(f"  ERROR   {sym}: {err}")

    print(f"\n{len(results)} scored, {len(skipped)} filtered out, {len(errors)} errors "
          f"in {time.time() - started:.0f}s")
    if tickers and not results and errors and len(errors) == len(tickers):
        sys.exit("SCAN FAILED: every ticker errored. Check your connection or Schwab status.")

    # Log list scans (universe or watchlist), not one-off ticker checks
    if results and list_name != "custom" and not args.no_log:
        if len(errors) > 0.2 * len(tickers):
            print("Scan log NOT updated: too many errors for a complete day's record.")
        else:
            try:
                rows = scan_log.build_rows(results, market, args.min_checks, list_name, src.name)
                print(scan_log.append(rows))
            except Exception as e:
                print(f"Scan log NOT updated: {e}")


if __name__ == "__main__":
    main()
