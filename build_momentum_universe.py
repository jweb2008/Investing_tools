"""
Builds the broad universe for the momentum profile. Run weekly.

  py build_momentum_universe.py

Source: the free Nasdaq Trader symbol directories, which list every stock
on Nasdaq, NYSE, NYSE American, NYSE Arca and Cboe.
  https://www.nasdaqtrader.com/dynamic/SymDir/nasdaqlisted.txt
  https://www.nasdaqtrader.com/dynamic/SymDir/otherlisted.txt

Kept: common stocks and ADRs. Dropped: ETFs, test issues, warrants, units,
rights, preferreds, notes and other non-common securities.

Then a cheap prefilter with Schwab quotes (about 200 symbols per request,
read-only): last price $5+ and roughly $20M+ average daily dollar volume
(price x 10-day average volume). The scanner's full filters (50-day dollar
volume, 1 year of history) still run on every stock at scan time.

Writes
  data/universe_momentum.txt       symbols that passed
  data/universe_momentum_info.csv  every candidate with exchange and why it was dropped
"""
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
DATA_DIR = HERE / "data"
OUT_TXT = DATA_DIR / "universe_momentum.txt"
OUT_INFO = DATA_DIR / "universe_momentum_info.csv"
GROWTH_UNIVERSE = DATA_DIR / "universe.txt"

SOURCES = {
    "nasdaq": "https://www.nasdaqtrader.com/dynamic/SymDir/nasdaqlisted.txt",
    "other": "https://www.nasdaqtrader.com/dynamic/SymDir/otherlisted.txt",
}
EXCLUDE_NAME = re.compile(
    r"\b(?:warrants?|units?|rights?|preferred|depositary shares? representing.*preferred|"
    r"notes? due|debentures?|subordinated|trust preferred|perpetual|"
    r"closed[- ]end|fund|etn|exchange[- ]traded)\b", re.I)
QUOTE_BATCH = 200
MIN_PRICE = scanner.MIN_PRICE
MIN_DOLLAR_VOLUME = scanner.MIN_DOLLAR_VOLUME


def _download(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (investing_tools scanner)"})
    return urllib.request.urlopen(req, timeout=60).read().decode("utf-8", errors="replace")


def listed_stocks():
    frames = []
    for name, url in SOURCES.items():
        text = _download(url)
        lines = [l for l in text.splitlines() if l and not l.startswith("File Creation Time")]
        df = pd.read_csv(io.StringIO("\n".join(lines)), sep="|", dtype=str).fillna("")
        if name == "nasdaq":
            df = df.rename(columns={"Symbol": "symbol", "Security Name": "name"})
            df["exchange"] = "NASDAQ"
        else:
            df = df.rename(columns={"ACT Symbol": "symbol", "Security Name": "name"})
            df["exchange"] = df["Exchange"].map({"N": "NYSE", "A": "NYSE American", "P": "NYSE Arca",
                                                 "Z": "Cboe", "V": "IEX"}).fillna(df["Exchange"])
        df = df[(df["ETF"] != "Y") & (df["Test Issue"] != "Y")]
        frames.append(df[["symbol", "name", "exchange"]])
        print(f"  {name}: {len(df)} listed (ETFs and test issues removed)")
    allm = pd.concat(frames, ignore_index=True).drop_duplicates("symbol")
    allm["symbol"] = allm["symbol"].str.strip().str.upper()
    # common stock symbols only: 1-5 letters, optional .A/.B class suffix
    ok_sym = allm["symbol"].str.fullmatch(r"[A-Z]{1,5}(\.[A-Z])?")
    ok_name = ~allm["name"].str.contains(EXCLUDE_NAME) & ~allm["name"].str.contains("%", regex=False)
    allm["status"] = "listed"
    allm.loc[~ok_sym, "status"] = "not a common stock symbol"
    allm.loc[ok_sym & ~ok_name, "status"] = "not common stock (warrant/unit/preferred/fund)"
    return allm


def quote_prefilter(src, symbols):
    """{symbol: (price, est dollar volume) or None} using Schwab quotes."""
    client = getattr(src, "client", None)
    if client is None:
        raise DataError("The quote prefilter needs Schwab data (DATA_SOURCE=schwab).")
    out = {}
    for i in range(0, len(symbols), QUOTE_BATCH):
        batch = symbols[i:i + QUOTE_BATCH]
        src._pace()
        r = client.get_quotes([s.replace(".", "/") for s in batch],
                              fields=[client.Quote.Fields.QUOTE, client.Quote.Fields.FUNDAMENTAL])
        if r.status_code == 401:
            raise DataError("Schwab login expired. Run: py auth_setup.py")
        if r.status_code != 200:
            for s in batch:
                out[s] = None
            continue
        data = r.json() or {}
        for s in batch:
            q = data.get(s.replace(".", "/")) or {}
            quote, fund = q.get("quote") or {}, q.get("fundamental") or {}
            price = quote.get("lastPrice") or quote.get("closePrice") or quote.get("mark")
            vol = fund.get("avg10DaysVolume") or fund.get("avg1YearVolume") or quote.get("totalVolume")
            out[s] = (float(price), float(price) * float(vol)) if price and vol else None
        if (i // QUOTE_BATCH) % 10 == 9:
            print(f"  quotes {min(i + QUOTE_BATCH, len(symbols))}/{len(symbols)}")
    return out


def main():
    print("Downloading the US listed-stock directories...")
    try:
        allm = listed_stocks()
    except Exception as e:
        sys.exit(f"BUILD FAILED: could not download the symbol directories ({e}). "
                 "universe_momentum.txt was left unchanged.")
    cands = allm[allm["status"] == "listed"]["symbol"].tolist()
    print(f"{len(cands)} common stocks and ADRs. Checking price and volume with Schwab quotes...")
    try:
        src = get_source("schwab")
        q = quote_prefilter(src, cands)
    except DataError as e:
        sys.exit(f"BUILD FAILED: {e}")

    def verdict(s):
        v = q.get(s)
        if v is None:
            return "no quote"
        price, dv = v
        if price < MIN_PRICE:
            return f"price under ${MIN_PRICE:.0f}"
        if dv < MIN_DOLLAR_VOLUME:
            return f"dollar volume under ${MIN_DOLLAR_VOLUME / 1e6:.0f}M"
        return "ok"

    mask = allm["status"] == "listed"
    allm.loc[mask, "status"] = allm.loc[mask, "symbol"].map(verdict)
    ok = sorted(set(allm.loc[allm["status"] == "ok", "symbol"]))
    no_quote = (allm["status"] == "no quote").sum()
    if len(ok) < 500:
        sys.exit(f"BUILD FAILED: only {len(ok)} stocks passed, which looks like a data problem "
                 "rather than real filtering. universe_momentum.txt was left unchanged.")

    # make sure every growth-universe stock is included too
    if GROWTH_UNIVERSE.exists():
        growth = {l.split("#")[0].strip() for l in GROWTH_UNIVERSE.read_text().splitlines()} - {""}
        ok = sorted(set(ok) | growth)

    DATA_DIR.mkdir(exist_ok=True)
    allm["built"] = dt.date.today().isoformat()
    allm.to_csv(OUT_INFO, index=False)
    OUT_TXT.write_text(f"# built {dt.date.today()} from Nasdaq Trader directories; {len(ok)} stocks\n"
                       + "\n".join(ok) + "\n")
    print(f"\n{len(ok)} stocks in {OUT_TXT.name} ({no_quote} had no quote).")


if __name__ == "__main__":
    t0 = time.time()
    main()
    print(f"Done in {time.time() - t0:.0f}s")
