"""
Builds the scan universe from index membership. Run weekly.

  py build_universe.py              fetch members, apply filters (~10 minutes)
  py build_universe.py --no-filter  fetch members only (fast)

Universe: S&P 500 + S&P 400 (midcap) + Nasdaq-100, about 1,000 stocks.
Membership comes from Wikipedia's index tables (free, updated as changes
happen). Each stock is then checked against the scanner's filters (price $5+,
500K+ average volume, 1 year of history) so the daily scan only pulls stocks
that can actually qualify.

Writes
  universe.txt       symbols that passed, one per line (what py scan.py --universe reads)
  universe_info.csv  every member with sector, index, and pass/skip reason

Read-only. Pulls index lists and daily price history only.
"""
import argparse
import datetime as dt
import io
import re
import sys
import time
import urllib.request
from pathlib import Path

import pandas as pd

import scanner
from data import DataError, get_source

HERE = Path(__file__).parent
OUT_TXT = HERE / "universe.txt"
OUT_INFO = HERE / "universe_info.csv"

INDEXES = {
    "SP500": "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies",
    "SP400": "https://en.wikipedia.org/wiki/List_of_S%26P_400_companies",
    "NDX": "https://en.wikipedia.org/wiki/Nasdaq-100",
}
SYMBOL_COLS = ("Symbol", "Ticker")
SECTOR_COLS = ("GICS Sector", "ICB Industry")


def fetch_members(name, url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (investing_tools scanner)"})
    html = urllib.request.urlopen(req, timeout=30).read().decode("utf-8")
    found = []
    for t in pd.read_html(io.StringIO(html)):
        # flatten two-row headers and drop footnote marks like "ICB Industry[14]"
        if isinstance(t.columns, pd.MultiIndex):
            t.columns = [c[-1] for c in t.columns]
        t.columns = [re.sub(r"\[.*?\]", "", str(c)).strip() for c in t.columns]
        sym_col = next((c for c in t.columns if c.split(" ")[0] in SYMBOL_COLS), None)
        if not sym_col:
            continue
        sec_col = next((c for c in t.columns
                        if any(c.startswith(s) for s in SECTOR_COLS + ("GICS", "ICB", "Sector", "Industry"))),
                       None)
        out = pd.DataFrame({
            "symbol": t[sym_col].astype(str).str.strip().str.upper(),
            "sector": t[sec_col].astype(str).str.strip() if sec_col else "Unknown",
            "index": name,
        })
        out = out[out["symbol"].str.fullmatch(r"[A-Z]{1,5}(\.[A-Z])?")]
        found.append((sec_col is not None, len(out), out))
    # the constituents table: a realistic number of valid symbols, preferring
    # one that also has a sector column
    good = [f for f in found if f[1] >= 90]
    if not good:
        sizes = ", ".join(str(f[1]) for f in found) or "none"
        raise RuntimeError(f"no members table found (symbol tables with sizes: {sizes})")
    return max(good, key=lambda f: (f[0], f[1]))[2]


def build_member_list():
    frames = []
    for name, url in INDEXES.items():
        try:
            m = fetch_members(name, url)
            print(f"  {name}: {len(m)} members")
            frames.append(m)
        except Exception as e:
            print(f"  {name}: FAILED ({e})")
    if not frames:
        sys.exit("BUILD FAILED: no index lists could be downloaded. universe.txt was left unchanged.")
    allm = pd.concat(frames, ignore_index=True)
    # one row per symbol; keep the GICS sector (S&P lists) over ICB when both exist
    allm["_pref"] = allm["index"].map({"SP500": 0, "SP400": 1, "NDX": 2})
    indexes = allm.groupby("symbol")["index"].apply(lambda s: "+".join(sorted(set(s))))
    out = (allm.sort_values("_pref").drop_duplicates("symbol")
           .drop(columns=["_pref", "index"]).set_index("symbol"))
    out["indexes"] = indexes
    return out.sort_index().reset_index(), len(frames) == len(INDEXES)


def main():
    ap = argparse.ArgumentParser(description="Build the weekly scan universe")
    ap.add_argument("--no-filter", action="store_true", help="skip the price/volume/history check")
    ap.add_argument("--source", help="schwab or yahoo (overrides DATA_SOURCE)")
    args = ap.parse_args()

    print("Downloading index members...")
    members, complete = build_member_list()
    print(f"{len(members)} unique stocks")

    members["status"] = "ok"
    if not args.no_filter:
        try:
            src = get_source(args.source)
        except DataError as e:
            sys.exit(f"BUILD FAILED: {e}")
        print(f"Checking filters with {src.name} data (about {len(members) * 0.6 / 60:.0f} minutes)...")
        started = time.time()
        for i, row in members.iterrows():
            try:
                reason = scanner.check_filters(src.history(row["symbol"]))
            except DataError as e:
                sys.exit(f"\nBUILD FAILED: {e}\nuniverse.txt was left unchanged.")
            except Exception as e:
                reason = f"error: {str(e)[:80]}"
            members.at[i, "status"] = reason or "ok"
            if (i + 1) % 100 == 0:
                print(f"  {i + 1}/{len(members)}  ({time.time() - started:.0f}s)")

    ok = members[members["status"] == "ok"]
    errs = members["status"].str.startswith("error")
    if len(ok) < 0.5 * len(members):
        sys.exit(f"BUILD FAILED: only {len(ok)} of {len(members)} passed, which looks like a data "
                 "problem rather than real filtering. universe.txt was left unchanged.")

    members["built"] = dt.date.today().isoformat()
    members.to_csv(OUT_INFO, index=False)
    OUT_TXT.write_text(f"# built {dt.date.today()} from {', '.join(INDEXES)}; {len(ok)} stocks\n"
                       + "\n".join(ok["symbol"]) + "\n")

    print(f"\n{len(ok)} stocks in universe.txt "
          f"({(members['status'] != 'ok').sum() - errs.sum()} filtered out, {errs.sum()} errors)")
    if errs.any():
        print("Errors (often renamed or delisted tickers): "
              + ", ".join(members.loc[errs, "symbol"].head(20)))
    if not complete:
        print("WARNING: at least one index list failed to download, so the universe is incomplete.")


if __name__ == "__main__":
    main()
