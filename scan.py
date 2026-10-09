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
        print(f"    {'+' if p else '-'} {k:<13} {p:>3} / {w}")
    print(f"    TOTAL         {r['score']:>3} / 100")
    if r["breakout"]:
        print("  ** BREAKOUT: closed above pivot on 1.5x+ volume **")


def main():
    ap = argparse.ArgumentParser(description="Breakout setup scanner (daily bars)")
    ap.add_argument("tickers", nargs="*", help="tickers to scan (default: watchlist.txt)")
    ap.add_argument("--universe", action="store_true",
                    help="scan universe.txt (built weekly by build_universe.py)")
    ap.add_argument("--all", action="store_true", help="show every score, not just 60+")
    ap.add_argument("--source", help="schwab or yahoo (overrides DATA_SOURCE)")
    ap.add_argument("--min-score", type=int, default=scanner.MIN_SCORE)
    args = ap.parse_args()

    if args.universe:
        uni = HERE / "universe.txt"
        if not uni.exists():
            sys.exit("universe.txt not found. Run: py build_universe.py")
        age = (time.time() - uni.stat().st_mtime) / 86400
        if age > 8:
            print(f"WARNING: universe.txt is {age:.0f} days old. Run: py build_universe.py")
        tickers = load_watchlist(uni)
    else:
        tickers = [t.upper() for t in args.tickers] or load_watchlist()
    try:
        src = get_source(args.source)
    except DataError as e:
        sys.exit(f"SCAN FAILED: {e}")

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
        df = pd.DataFrame(results).sort_values("score", ascending=False)
        shown = df if args.all else df[df["score"] >= args.min_score]
        cols = {"symbol": "Symbol", "score": "Score", "close": "Close", "pivot": "Pivot",
                "pct_below_pivot": "% Below", "bb_width_pctile": "BB %ile",
                "atr_ratio": "ATR ratio", "breakout": "Breakout"}
        if len(shown):
            title = "All scores" if args.all else f"Setups scoring {args.min_score}+"
            print(f"\n{title}:")
            print(shown[list(cols)].rename(columns=cols).to_string(
                index=False, float_format=lambda x: f"{x:.2f}"))
        else:
            print(f"\nNo setups scored {args.min_score}+.")
        brk = df[df["breakout"]]
        if len(brk):
            print("\nBreakouts today (close above pivot on 1.5x+ volume): "
                  + ", ".join(brk["symbol"]))

    for sym, reason in skipped:
        print(f"  skipped {sym}: {reason}")
    for sym, err in errors:
        print(f"  ERROR   {sym}: {err}")

    print(f"\n{len(results)} scored, {len(skipped)} filtered out, {len(errors)} errors "
          f"in {time.time() - started:.0f}s")
    if tickers and not results and errors and len(errors) == len(tickers):
        sys.exit("SCAN FAILED: every ticker errored. Check your connection or Schwab status.")


if __name__ == "__main__":
    main()
