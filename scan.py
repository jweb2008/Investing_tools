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
    args = ap.parse_args()

    if args.universe:
        uni = HERE / "universe.txt"
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
        results.append({"symbol": sym, **scanner.score(df)})
        if len(tickers) > 20 and i % 100 == 0:
            print(f"  {i}/{len(tickers)}...")

    if len(tickers) == 1 and results:
        print_detail(tickers[0], results[0])
    elif results:
        df = pd.DataFrame(results).sort_values(["score", "pct_below_pivot"], ascending=[False, True])
        df["setup"] = [scanner.is_setup(r, args.min_checks) for r in df.to_dict("records")]
        shown = df if args.all else df[df["setup"]]
        cols = {"symbol": "Symbol", "score": "Score", "checks": "Checks", "close": "Close",
                "pivot": "Pivot",
                "pct_below_pivot": "% Below", "bb_width_pctile": "BB %ile",
                "atr_ratio": "ATR ratio", "breakout": "Breakout"}
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
            print("\nBreakouts (uptrend, close above pivot on 1.5x+ volume): "
                  + ", ".join(brk["symbol"]))

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
