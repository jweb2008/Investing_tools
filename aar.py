"""
After Action Review (AAR): grades how trades actually turned out, using the
Excel tracker as the source of truth.

  py aar.py            build aar.xlsx from the newest tracker
  py aar.py --tracker "2025-26_Trading_Tracker_10.9.11.xlsx"   use a specific file

Read-only. Reads the tracker (table Table14), fills_cache.csv and
data/scan_log.csv. Never opens the tracker in Excel, never writes to it, and
never talks to Schwab orders. See AAR_DESIGN.md.

Build step 1 (this version)
  - Tracker rows -> trades: one continuous holding per account and symbol.
    Scale-ins and partial exits are rows inside one trade.
  - Trade net % = sum of recorded Net Profit / sum of recorded TOTAL INVESTMENT
    (Total Sale Value for sold-to-open trades, matching the tracker's % CHANGE).
    Nothing is recalculated from prices.
  - Style from your Action column (Swing, Long, Hedge, Day, ...), with
    aar_tags.csv overrides. Sold-to-open trades are always Hedge.
  - Outcome: Win / Scratch / Loss with a symmetric scratch band.
  - Decision ID: the same contract or stock entered on the same day in several
    accounts is one decision.
Not yet: matching to scanner flags, grading by setup, price paths.

Output: aar.xlsx with tables Trades, TradeRows, Review, RunInfo (Power Query).

Settings (.env, all optional)
  AAR_SCRATCH_STOCK    scratch band for stocks, in %       (default 2)
  AAR_SCRATCH_OPTION   scratch band for options, in %      (default 10)
  AAR_LONG_SWING_DAYS  flag Swing trades held longer than this many trading
                       days on the Review sheet            (default 40)
  AAR_OUTPUT           output file                         (default aar.xlsx)
  TRACKER_DIR, TRACKER_PREFIX, TRACKER_TABLE   same as update_tracker.py
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

HERE = Path(__file__).parent
TRACKER_DIR = Path(os.getenv("TRACKER_DIR", str(HERE)))
PREFIX = os.getenv("TRACKER_PREFIX", "2025-26_Trading_Tracker_")
TABLE = os.getenv("TRACKER_TABLE", "Table14")
FILLS = Path(os.getenv("TRADE_LOG_PATH", str(HERE / "trade_log.xlsx"))).with_name("fills_cache.csv")
SCAN_LOG = HERE / "data" / "scan_log.csv"
TAGS = HERE / "aar_tags.csv"
OUTPUT = Path(os.getenv("AAR_OUTPUT", str(HERE / "aar.xlsx")))

SCRATCH = {"Stock": float(os.getenv("AAR_SCRATCH_STOCK", "2")),
           "Option": float(os.getenv("AAR_SCRATCH_OPTION", "10"))}
LONG_SWING_DAYS = int(os.getenv("AAR_LONG_SWING_DAYS", "40"))

# Action column value -> style. Anything not listed keeps its own name and is
# reported on the Review sheet so you can decide where it belongs.
# STC and BTO mark the front or back end of a hedge.
STYLE_MAP = {
    "swing": "Swing",
    "long": "Long term", "long term": "Long term", "leaps": "Long term",
    "hedge": "Hedge",
    "stc": "Hedge", "bto": "Hedge",   # front or back end of a hedge
    "day": "Day",
}

# NYSE full-day closures, so hold periods count trading days.
NYSE_HOLIDAYS = pd.to_datetime([
    "2023-01-02", "2023-01-16", "2023-02-20", "2023-04-07", "2023-05-29", "2023-06-19",
    "2023-07-04", "2023-09-04", "2023-11-23", "2023-12-25",
    "2024-01-01", "2024-01-15", "2024-02-19", "2024-03-29", "2024-05-27", "2024-06-19",
    "2024-07-04", "2024-09-02", "2024-11-28", "2024-12-25",
    "2025-01-01", "2025-01-09", "2025-01-20", "2025-02-17", "2025-04-18", "2025-05-26",
    "2025-06-19", "2025-07-04", "2025-09-01", "2025-11-27", "2025-12-25",
    "2026-01-01", "2026-01-19", "2026-02-16", "2026-04-03", "2026-05-25", "2026-06-19",
    "2026-07-03", "2026-09-07", "2026-11-26", "2026-12-25",
    "2027-01-01", "2027-01-18", "2027-02-15", "2027-03-26", "2027-05-31", "2027-06-18",
    "2027-07-05", "2027-09-06", "2027-11-25", "2027-12-24",
]).values.astype("datetime64[D]")


def trading_days(start, end):
    """Trading days from start to end (same day = 0)."""
    if pd.isna(start) or pd.isna(end):
        return None
    a = np.datetime64(pd.Timestamp(start).date(), "D")
    b = np.datetime64(pd.Timestamp(end).date(), "D")
    return int(np.busday_count(a, b, holidays=NYSE_HOLIDAYS))


# ------------------------------------------------------------------ symbols
def canon_root(root):
    """EOSE1 (adjusted contract) -> EOSE, same as update_tracker.py."""
    r = re.sub(r"\d+$", "", str(root).strip().upper())
    return r or str(root).strip().upper()


def norm_sym(s):
    """Same contract however the date or strike was typed."""
    t = re.sub(r"\s+", " ", str(s or "")).strip().upper()
    m = re.match(r"^(\S+) (\d{1,2})/(\d{1,2})/(\d{2}|\d{4}) ([\d.]+) ([CP])$", t)
    if m:
        root, mm, dd, yy, k, cp = m.groups()
        yy = yy if len(yy) == 4 else "20" + yy
        try:
            k = f"{float(k):.2f}"
        except ValueError:
            pass
        t = f"{canon_root(root)} {int(mm):02d}/{int(dd):02d}/{yy} {k} {cp}"
    return t


def underlying(sym):
    s = str(sym or "").strip()
    return canon_root(s.split(" ")[0]) if s else ""


def option_expiry(sym):
    m = re.match(r"^\S+ (\d{2})/(\d{2})/(\d{4}) [\d.]+ [CP]$", norm_sym(sym))
    return pd.Timestamp(f"{m.group(3)}-{m.group(1)}-{m.group(2)}") if m else None


def blank(v):
    return v is None or (isinstance(v, float) and np.isnan(v)) or v is pd.NaT \
        or (isinstance(v, str) and not v.strip())


def num(v):
    try:
        return 0.0 if blank(v) else float(v)
    except (TypeError, ValueError):
        return 0.0


# ------------------------------------------------------------------- inputs
def find_latest_tracker():
    files = [p for p in TRACKER_DIR.glob(f"{PREFIX}*.xlsx") if not p.name.startswith("~$")]
    if not files:
        sys.exit(f"No tracker found in {TRACKER_DIR} matching {PREFIX}*.xlsx")
    return max(files, key=lambda p: p.stat().st_mtime)


def read_tracker(path):
    """Table rows with the values Excel last saved (formulas come in as results).
    Reads a temporary copy, so it works even while the tracker is open in Excel."""
    import openpyxl
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        shutil.copy2(path, tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True)
        ws = next((s for s in wb.worksheets if TABLE in s.tables), None)
        if ws is None:
            sys.exit(f"Couldn't find the table {TABLE} in {path.name}. "
                     "If you renamed it, add TRACKER_TABLE=<new name> to .env.")
        cells = list(ws[ws.tables[TABLE].ref])
        wb.close()
    finally:
        os.remove(tmp)
    headers = [c.value for c in cells[0]]
    df = pd.DataFrame([[c.value for c in r] for r in cells[1:]], columns=headers)
    df.insert(0, "ExcelRow", [r[0].row for r in cells[1:]])
    for c in ("Opened Date", "Closed Date"):
        df[c] = pd.to_datetime(df[c], errors="coerce")
    needed = {"Account", "Opened Date", "Stock Or Option", "SYMBOL", "Closed Date",
              "TOTAL INVESTMENT", "Total Sale Value", "Net Profit", "Open/Closed"}
    missing = needed - set(df.columns)
    if missing:
        sys.exit(f"{TABLE} is missing columns: {', '.join(sorted(missing))}")
    for c in ("Action", "Sync ID", "Buy Quantity", "Sell Quantity",
              "Purchase PRICE PER SHARE", "% CHANGE"):
        if c not in df.columns:
            df[c] = None
    keep = ~df[["Account", "SYMBOL", "Opened Date", "Closed Date"]].apply(
        lambda r: all(blank(v) for v in r), axis=1)
    return df[keep].reset_index(drop=True), path.name


def short_open_fills():
    """Fill IDs that opened a position by selling (sold to open)."""
    if not FILLS.exists():
        return set()
    f = pd.read_csv(FILLS, dtype=str, usecols=["FillID", "Side", "PositionEffect"])
    return set(f.loc[(f["Side"] == "SELL") & (f["PositionEffect"] == "OPENING"), "FillID"])


def load_tags():
    """aar_tags.csv: Account, Symbol, OpenedDate (optional), Style."""
    if not TAGS.exists():
        return []
    t = pd.read_csv(TAGS, dtype=str).fillna("")
    out = []
    for r in t.itertuples(index=False):
        d = pd.to_datetime(r.OpenedDate, errors="coerce") if r.OpenedDate.strip() else None
        out.append((r.Account.strip(), norm_sym(r.Symbol), d, r.Style.strip()))
    return out


# ------------------------------------------------------------- rows -> trades
def prepare_rows(trk, short_fills):
    r = trk.copy()
    r["Asset"] = np.where(r["Stock Or Option"].astype(str).str.strip().str.lower()
                          .str.startswith("opt"), "Option", "Stock")
    r["Symbol"] = r["SYMBOL"].map(norm_sym)
    r["Underlying"] = r["SYMBOL"].map(underlying)
    r["Account"] = r["Account"].astype(str).str.strip()
    r["ActionRaw"] = r["Action"].map(lambda v: "" if blank(v) else str(v).strip())

    def is_short(row):
        sync = "" if blank(row["Sync ID"]) else str(row["Sync ID"])
        m = re.match(r"^[^:]+:(.*?)>", sync)
        if m:
            opener = m.group(1)
            if opener in short_fills or re.search(r"seed:\d+:S", opener):
                return True
        buy, sell = num(row["Buy Quantity"]), num(row["Sell Quantity"])
        if buy < 0:                       # old "Hedge" convention
            return True
        return buy == 0 and sell > 0 and num(row["TOTAL INVESTMENT"]) == 0
    r["Short"] = r.apply(is_short, axis=1)
    r["Closed"] = r["Open/Closed"].astype(str).str.strip().str.lower().eq("closed")
    # each row's own style, so a swing lot on top of a long-term core holding
    # in the same account stays a separate trade
    r["RowStyle"] = [("Hedge" if sh else STYLE_MAP.get(a.lower(), a) if a else "")
                     for sh, a in zip(r["Short"], r["ActionRaw"])]
    return r


def chain_trades(rows):
    """Assign TradeID: rows of one account+symbol+style whose holding periods
    touch or overlap form one trade (flat to open to flat)."""
    rows = rows.copy()
    rows["TradeID"] = None
    n = 0
    for _, g in rows.groupby(["Account", "Symbol", "RowStyle"], sort=False):
        g = g.sort_values(["Opened Date", "Closed Date"], na_position="first")
        cur_end, cur_open, cur_id = None, False, None
        for idx, row in g.iterrows():
            o, c = row["Opened Date"], row["Closed Date"]
            if pd.isna(o):                # can't place it in time: its own trade
                n += 1
                rows.at[idx, "TradeID"] = f"T{n:04d}"
                continue
            joins = cur_id is not None and (cur_open or (cur_end is not None and o <= cur_end))
            if not joins:
                n += 1
                cur_id, cur_end, cur_open = f"T{n:04d}", None, False
            rows.at[idx, "TradeID"] = cur_id
            if row["Closed"] and not pd.isna(c):
                cur_end = c if cur_end is None else max(cur_end, c)
            else:
                cur_open = True
    return rows


def style_for(trade_rows, tags):
    first = trade_rows.sort_values("Opened Date", na_position="first").iloc[0]
    acct, sym, opened = first["Account"], first["Symbol"], first["Opened Date"]
    if trade_rows["Short"].any():
        return "Hedge", "sold to open"
    for t_acct, t_sym, t_date, t_style in tags:
        if t_acct.lower() == acct.lower() and t_sym == sym and (
                t_date is None or (not pd.isna(opened) and t_date.date() == opened.date())):
            return t_style, "aar_tags.csv"
    actions = [a for a in trade_rows.sort_values("Opened Date", na_position="first")["ActionRaw"] if a]
    if actions:
        a = actions[0]
        return STYLE_MAP.get(a.lower(), a), "Action column"
    if first["Asset"] == "Option" and not pd.isna(opened):
        exp = option_expiry(sym)
        if exp is not None and (exp - opened).days >= 365:
            return "Long term", "LEAPS (1y+ to expiry)"
    return "Untagged", ""


def outcome(pct, asset):
    if pct is None or pd.isna(pct):
        return ""
    band = SCRATCH[asset]
    return "Win" if pct > band else ("Loss" if pct < -band else "Scratch")


def build_trades(rows, tags):
    out, review = [], []
    for tid, g in rows.groupby("TradeID", sort=False):
        g = g.sort_values(["Opened Date", "Closed Date"], na_position="first")
        first = g.iloc[0]
        style, src = style_for(g, tags)
        closed = bool(g["Closed"].all())
        invest = g["TOTAL INVESTMENT"].map(num).sum()
        sale = g["Total Sale Value"].map(num).sum()
        net = g["Net Profit"].map(num).sum()
        base = invest if invest > 0 else sale
        pct = round(net / base * 100, 2) if closed and base else None
        entry, exit_ = g["Opened Date"].min(), g["Closed Date"].max() if closed else pd.NaT
        long_rows = g[~g["Short"]]
        qty = long_rows["Buy Quantity"].map(num)
        px = long_rows["Purchase PRICE PER SHARE"].map(num)
        avg_px = round((qty * px).sum() / qty.sum(), 4) if qty.sum() > 0 else None
        acts = sorted({a for a in g["ActionRaw"] if a})
        out.append({
            "TradeID": tid, "Account": first["Account"], "Asset": first["Asset"],
            "Style": style, "StyleFrom": src, "Underlying": first["Underlying"],
            "Symbol": first["Symbol"], "Direction": "Short" if g["Short"].any() else "Long",
            "Status": "Closed" if closed else "Open",
            "EntryDate": entry, "ExitDate": exit_,
            "HoldTradingDays": trading_days(entry, exit_) if closed else pd.NA,
            "Rows": len(g),
            "Entries": int(g["Opened Date"].dropna().dt.normalize().nunique()),
            "Exits": int(g.loc[g["Closed"], "Closed Date"].dropna().dt.normalize().nunique()),
            "AvgEntryPrice": avg_px,
            "TotalInvestment": round(invest, 2), "TotalSaleValue": round(sale, 2),
            "NetProfit": round(net, 2) if closed else None,
            "NetPct": pct, "Outcome": outcome(pct, first["Asset"]) if closed else "",
            "ActionValues": ", ".join(acts),
            "Pre2026": bool(not pd.isna(entry) and entry < pd.Timestamp("2026-01-01")),
        })
        if len(acts) > 1:
            review.append(("Mixed Action values in one trade", tid, first["Account"],
                           first["Symbol"], f"{', '.join(acts)}; using '{style}'"))
        if g["Short"].any() and acts and STYLE_MAP.get(acts[0].lower()) != "Hedge":
            review.append(("Sold to open but Action is not Hedge", tid, first["Account"],
                           first["Symbol"], f"Action '{acts[0]}'; graded as Hedge"))
        if pd.isna(first["Opened Date"]):
            review.append(("No Opened Date", tid, first["Account"], first["Symbol"],
                           f"Excel row {int(first['ExcelRow'])}; can't be matched or timed"))
    trades = pd.DataFrame(out)
    trades["HoldTradingDays"] = trades["HoldTradingDays"].astype("Int64")

    # one decision = same symbol, same entry day, same style, across accounts
    key = (trades["Symbol"] + "|" + trades["EntryDate"].dt.strftime("%Y-%m-%d").fillna(trades["TradeID"])
           + "|" + trades["Style"])
    trades["DecisionID"] = [f"D{i + 1:04d}" for i in pd.factorize(key)[0]]
    trades["Accounts"] = trades.groupby("DecisionID")["Account"].transform("count")

    for t in trades.itertuples(index=False):
        if t.Style == "Untagged":
            review.append(("No style", t.TradeID, t.Account, t.Symbol,
                           "Add Swing/Long/Hedge/Day in the Action column or aar_tags.csv"))
        elif t.Style not in ("Swing", "Long term", "Hedge", "Day"):
            review.append((f"Action '{t.Style}' not mapped to a style", t.TradeID, t.Account,
                           t.Symbol, "Reported as its own group for now"))
        if t.Style == "Swing" and t.Status == "Closed" and (t.HoldTradingDays or 0) > LONG_SWING_DAYS:
            review.append(("Swing held a long time", t.TradeID, t.Account, t.Symbol,
                           f"{t.HoldTradingDays} trading days; Long term?"))
    review = pd.DataFrame(review, columns=["Issue", "TradeID", "Account", "Symbol", "Detail"])
    return trades, review


def decision_view(trades):
    """Closed trades rolled up to one row per decision (combined across accounts)."""
    c = trades[trades["Status"] == "Closed"]
    g = c.groupby("DecisionID")
    d = g.agg(Asset=("Asset", "first"), Style=("Style", "first"), Symbol=("Symbol", "first"),
              EntryDate=("EntryDate", "first"), Accounts=("Account", "count"),
              Invest=("TotalInvestment", "sum"), Sale=("TotalSaleValue", "sum"),
              Net=("NetProfit", "sum")).reset_index()
    base = np.where(d["Invest"] > 0, d["Invest"], d["Sale"])
    d["NetPct"] = np.round(np.where(base != 0, d["Net"] / np.where(base == 0, 1, base) * 100, np.nan), 2)
    d["Outcome"] = [outcome(p, a) for p, a in zip(d["NetPct"], d["Asset"])]
    return d


# ------------------------------------------------------------------ summary
def summarize(df, by):
    rows = []
    for key, g in df.groupby(by, sort=True):
        key = key if isinstance(key, tuple) else (key,)
        p = g["NetPct"].astype(float)
        wins, losses = p[g["Outcome"] == "Win"], p[g["Outcome"] == "Loss"]
        n = len(g)
        rows.append({**dict(zip(by, key)), "N": n,
                     "Win%": round((g["Outcome"] == "Win").mean() * 100, 1),
                     "Scratch%": round((g["Outcome"] == "Scratch").mean() * 100, 1),
                     "Loss%": round((g["Outcome"] == "Loss").mean() * 100, 1),
                     "AvgPct": round(p.mean(), 2), "MedianPct": round(p.median(), 2),
                     "AvgWinPct": round(wins.mean(), 2) if len(wins) else None,
                     "AvgLossPct": round(losses.mean(), 2) if len(losses) else None,
                     "Note": "too few to read" if n < 20 else ""})
    return pd.DataFrame(rows)


# ------------------------------------------------------------------- output
def write_excel(sheets):
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=OUTPUT.parent)
    os.close(fd)
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df = df.copy()
            for c in df.columns:  # dates without the time part
                if pd.api.types.is_datetime64_any_dtype(df[c]):
                    df[c] = df[c].dt.date
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for i, col in enumerate(df.columns, 1):
                width = max([len(str(col))] + [len(str(v)) for v in df[col].head(300)]) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(width, 40)
            ws.freeze_panes = "A2"
            if len(df):
                ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
                tbl = Table(displayName=name.replace(" ", ""), ref=ref)
                tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
                ws.add_table(tbl)
    try:
        os.replace(tmp, OUTPUT)
    except PermissionError:
        os.remove(tmp)
        sys.exit(f"{OUTPUT.name} is open in Excel. Close it and run again.")


def scan_log_info():
    if not SCAN_LOG.exists():
        return "not found"
    s = pd.read_csv(SCAN_LOG, usecols=["ScanDate", "Status"])
    return (f"{len(s)} rows, {s['ScanDate'].min()} to {s['ScanDate'].max()} "
            f"({', '.join(f'{v} {k}' for k, v in s['Status'].value_counts().items())})")


def main():
    ap = argparse.ArgumentParser(description="After Action Review: grade trades from the tracker")
    ap.add_argument("--tracker", help="tracker file (default: newest in TRACKER_DIR)")
    args = ap.parse_args()

    path = Path(args.tracker) if args.tracker else find_latest_tracker()
    if not path.is_absolute():
        path = TRACKER_DIR / path
    trk, name = read_tracker(path)
    rows = chain_trades(prepare_rows(trk, short_open_fills()))
    trades, review = build_trades(rows, load_tags())
    decisions = decision_view(trades)

    summary = summarize(decisions, ["Style", "Asset"])
    summary.insert(0, "Report", "Closed decisions by style")

    row_cols = ["TradeID", "ExcelRow", "Account", "Opened Date", "Stock Or Option", "Action",
                "SYMBOL", "Buy Quantity", "Purchase PRICE PER SHARE", "TOTAL INVESTMENT",
                "Closed Date", "Sell Quantity", "Total Sale Value", "% CHANGE", "Net Profit",
                "Open/Closed", "Sync ID", "Short", "RowStyle"]
    info = pd.DataFrame({"Item": ["Run", "Tracker", "Tracker rows", "Trades", "Decisions (closed)",
                                  "Scan log", "Scratch band stock %", "Scratch band option %",
                                  "Build step"],
                         "Value": [dt.datetime.now().strftime("%Y-%m-%d %H:%M"), name, len(rows),
                                   len(trades), len(decisions), scan_log_info(),
                                   SCRATCH["Stock"], SCRATCH["Option"],
                                   "1: trades and styles (no scanner matching yet)"]})
    write_excel({"Summary": summary, "Trades": trades, "Decisions": decisions,
                 "TradeRows": rows[row_cols].sort_values(["TradeID", "Opened Date"]),
                 "Review": review, "RunInfo": info})

    print(f"After Action Review  ({name})")
    print(f"  {len(rows)} tracker rows -> {len(trades)} trades "
          f"({(trades['Status'] == 'Closed').sum()} closed) -> {len(decisions)} closed decisions")
    print("\n  Closed decisions by style (a decision = same entry across accounts):")
    show = summarize(decisions, ["Style", "Asset"])
    if len(show):
        print(show[["Style", "Asset", "N", "Win%", "Scratch%", "Loss%", "AvgPct", "MedianPct",
                    "Note"]].to_string(index=False))
    print(f"\n  Review items: {len(review)}"
          + (f" ({', '.join(f'{v} {k.lower()}' for k, v in review['Issue'].value_counts().items())})"
             if len(review) else ""))
    print(f"  -> {OUTPUT}")


if __name__ == "__main__":
    main()
