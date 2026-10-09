"""
Updates your trading tracker from Schwab fills.

  py update_tracker.py            pull fills, build a review report (no tracker changes)
  py update_tracker.py --apply    same, then save a NEW tracker version with the changes
  py update_tracker.py --no-sync  skip the Schwab pull, use fills already in fills_cache.csv

How it works
  1. Runs sync_trades.py to refresh fills_cache.csv (unless --no-sync).
  2. Groups fills into positions in the tracker's layout: one row per position,
     open side in the Buy columns, close side in the Sell columns. Sold-to-open
     options use a negative Buy Quantity (your "Hedge" convention). Expired
     options close with the Sell columns blank.
  3. Compares those rows to table Table14 on the "New 2026" sheet:
       Synced        row already carries a Sync ID; refreshed from Schwab
       Fill gaps     matching row found; only its EMPTY cells get filled
       Mismatch      matching row found but a filled-in value disagrees; left alone
       New           no matching row; appended
       Needs open    a 2026 close of something bought before 2026; open side blank
       Tracker only  2026 row in the tracker with no Schwab match; left alone
  4. Writes Review_<version>.xlsx so you can see every change.
  5. With --apply: opens the newest tracker in Excel itself (so charts, pivots and
     formatting are untouched), makes the changes, refreshes, and saves it as the
     next version, e.g. 2025-26_Trading_Tracker_10.6.1.xlsx. The original is never
     modified. Your Action column is never overwritten.

Settings (.env)
  TRACKER_DIR     folder holding the tracker files (default: this folder)
  TRACKER_PREFIX  file name before the date (default: 2025-26_Trading_Tracker_)
  ACCOUNT_NAMES   last 4 of each account = tracker name, e.g. 1234=Roth,5678=Emma
"""
import datetime as dt
import os
import re
import sys
from collections import defaultdict, deque
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

HERE = Path(__file__).parent
TRACKER_DIR = Path(os.getenv("TRACKER_DIR", str(HERE)))
PREFIX = os.getenv("TRACKER_PREFIX", "2025-26_Trading_Tracker_")
CACHE = Path(os.getenv("TRADE_LOG_PATH", str(HERE / "trade_log.xlsx"))).with_name("fills_cache.csv")
SYNC_START = pd.Timestamp(os.getenv("SYNC_START", "2026-01-01"))
TABLE = os.getenv("TRACKER_TABLE", "Table14")  # found by name, on whatever sheet it lives
SYNC_COL = "Sync ID"

# columns this script writes. Calculated columns (TOTAL INVESTMENT, Total Sale
# Value, % CHANGE, Net Profit, Neg $ Loss) and Action are left to the tracker/you.
DATA_COLS = ["Account", "Opened Date", "Stock Or Option", "SYMBOL", "Buy Quantity",
             "Purchase PRICE PER SHARE", "Fees & Comm", "Closed Date", "Sell Quantity",
             "Sell Price Per Share", "Sale Fees & Comm", "Open/Closed"]
NUM_COLS = {"Buy Quantity", "Purchase PRICE PER SHARE", "Fees & Comm", "Sell Quantity",
            "Sell Price Per Share", "Sale Fees & Comm"}
DATE_COLS = {"Opened Date", "Closed Date"}


# ------------------------------------------------------------------ helpers
def account_map():
    raw = os.getenv("ACCOUNT_NAMES", "")
    m = dict(p.strip().split("=", 1) for p in raw.split(",") if "=" in p)
    return {k.strip()[-4:]: v.strip() for k, v in m.items()}


def fmt_symbol(sym, asset):
    """Schwab OCC option symbol -> tracker style 'IREN 11/28/2025 55.00 C'."""
    if asset != "OPTION":
        return sym.strip()
    m = re.match(r"^(\S+)\s*(\d{2})(\d{2})(\d{2})([CP])(\d{8})$", sym.strip())
    if not m:
        return sym.strip()
    root, yy, mm, dd, cp, strike = m.groups()
    return f"{root} {mm}/{dd}/20{yy} {int(strike) / 1000:.2f} {cp}"


def option_expiry(sym):
    m = re.match(r"^\S+\s*(\d{2})(\d{2})(\d{2})[CP]\d{8}$", str(sym).strip())
    return pd.Timestamp(f"20{m.group(1)}-{m.group(2)}-{m.group(3)}") if m else None


def norm_sym(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().upper()


def is_blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and not v.strip()) \
        or v is pd.NaT


def same(a, b, col):
    if col in DATE_COLS:
        return pd.Timestamp(a).date() == pd.Timestamp(b).date()
    if col in NUM_COLS:
        tol = 0.02 if "Fees" in col else 0.0051
        try:
            return abs(float(a) - float(b)) <= tol
        except (TypeError, ValueError):
            return False
    return norm_sym(a) == norm_sym(b)


# ------------------------------------------------------- tracker files/versions
def find_latest_tracker():
    files = [p for p in TRACKER_DIR.glob(f"{PREFIX}*.xlsx") if not p.name.startswith("~$")]
    if not files:
        sys.exit(f"No tracker found in {TRACKER_DIR} matching {PREFIX}*.xlsx")
    return max(files, key=lambda p: p.stat().st_mtime)


def next_version_name(today=None):
    today = today or dt.date.today()
    base = f"{PREFIX}{today.month}.{today.day}."
    used = [int(m.group(1)) for p in TRACKER_DIR.glob(f"{base}*.xlsx")
            if (m := re.match(re.escape(base) + r"(\d+)\.xlsx$", p.name))]
    return f"{base}{max(used, default=0) + 1}.xlsx"


def read_tracker(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=False)
    ws = next((s for s in wb.worksheets if TABLE in s.tables), None)
    if ws is None:
        sys.exit(f"Couldn't find the table {TABLE} in {path.name}. If you renamed it, "
                 f"add TRACKER_TABLE=<new name> to .env.")
    tbl = ws.tables[TABLE]
    (c1, r1), (c2, r2) = [openpyxl.utils.cell.coordinate_to_tuple(x)[::-1] for x in tbl.ref.split(":")]
    headers = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
    rows = []
    for r in range(r1 + 1, r2 + 1):
        vals = {h: ws.cell(r, c1 + i).value for i, h in enumerate(headers)}
        vals["_row"] = r
        rows.append(vals)
    df = pd.DataFrame(rows)
    if SYNC_COL not in df.columns:
        df[SYNC_COL] = None
    for c in DATE_COLS:
        df[c] = pd.to_datetime(df[c], errors="coerce")
    df["_empty"] = df[DATA_COLS].apply(lambda r: all(is_blank(v) for v in r), axis=1)
    return df, headers


def convert_old_shorts(trk):
    """Rows from the old layout, sold-to-open entered as a negative Buy Quantity,
    move to the sell side. P&L is unchanged. Returns review actions to write."""
    acts = []
    for idx, t in trk.iterrows():
        b = pd.to_numeric(t["Buy Quantity"], errors="coerce")
        if pd.isna(b) or b >= 0 or not is_blank(t["Sell Quantity"]):
            continue
        new = {"Sell Quantity": -float(b), "Sell Price Per Share": t["Purchase PRICE PER SHARE"],
               "Sale Fees & Comm": t["Fees & Comm"], "Buy Quantity": None,
               "Purchase PRICE PER SHARE": None, "Fees & Comm": None}
        for k, v in new.items():
            trk.at[idx, k] = v
        vals = {c: trk.at[idx, c] for c in DATA_COLS}
        acts.append(dict(status="Converted", idx=idx, row=t["_row"], values=vals, overwrite=False,
                         apply=new, detail="sold to open: moved from Buy to Sell columns"))
    return acts


# ------------------------------------------------------ fills -> positions
MATCH_METHOD = os.getenv("MATCH_METHOD", "swing").lower()  # "swing" or "fifo"


def pick_lot(book, qty):
    """Which open lot a closing fill pairs with.

    swing (default): a lot with exactly the closing quantity, newest first;
    otherwise the newest lot (LIFO). This keeps a swing trade paired with its
    own entry even when a longer-term position in the same stock is also open.
    fifo: always the oldest lot.
    """
    if MATCH_METHOD == "fifo":
        return 0
    for i in range(len(book) - 1, -1, -1):
        if abs(book[i]["qty"] - qty) < 1e-9:
            return i
    return len(book) - 1


def build_positions(fills, names):
    """Match fills into tracker-style position rows (see pick_lot)."""
    f = fills.copy()
    f["Account"] = f["Account"].astype(str).str[-4:].map(lambda a: names.get(a, f"...{a}"))
    f["OrderID"] = f["OrderID"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    f["DateTime"] = pd.to_datetime(f["DateTime"])
    f["Date"] = f["DateTime"].dt.normalize()

    # merge partial fills of one order into a single execution (volume-weighted price)
    f["_grp"] = f.apply(lambda r: f"{r.Account}|{r.Symbol}|{r.Side}|{r.OrderID}|{r.Date.date()}"
                        if r.OrderID else r.FillID, axis=1)
    agg = (f.assign(_px=f.Price * f.Qty)
           .groupby("_grp", sort=False)
           .agg(FillID=("FillID", "min"), FillIDs=("FillID", lambda v: frozenset(v)),
                Account=("Account", "first"), Symbol=("Symbol", "first"),
                AssetType=("AssetType", "first"), Side=("Side", "first"),
                PositionEffect=("PositionEffect", "first"), Source=("Source", "first"),
                DateTime=("DateTime", "min"), Date=("Date", "first"),
                Qty=("Qty", "sum"), _px=("_px", "sum"), Fees=("Fees", "sum"))
           .reset_index(drop=True))
    agg["Price"] = (agg["_px"] / agg["Qty"]).round(4)

    lots = defaultdict(list)
    out = []

    def row(acct, sym, asset, lot=None, close=None, n=0.0):
        r = {c: None for c in DATA_COLS}
        r.update({"Account": acct, "SYMBOL": fmt_symbol(sym, asset),
                  "Stock Or Option": "Option" if asset == "OPTION" else "Stock"})
        short = (lot["signed"] < 0) if lot else (close is not None and close.Side == "BUY")
        # long:  open = Buy columns,  close = Sell columns
        # short (sold to open, e.g. covered calls): open = Sell columns, close = Buy columns
        open_cols = ("Sell Quantity", "Sell Price Per Share", "Sale Fees & Comm") if short else \
                    ("Buy Quantity", "Purchase PRICE PER SHARE", "Fees & Comm")
        close_cols = ("Buy Quantity", "Purchase PRICE PER SHARE", "Fees & Comm") if short else \
                     ("Sell Quantity", "Sell Price Per Share", "Sale Fees & Comm")
        if lot:
            r["Opened Date"] = lot["date"]
            r.update(dict(zip(open_cols, (n, lot["price"], round(lot["fees"] * n / lot["orig"], 2)))))
        removed = expired = False
        if close is not None:
            r["Closed Date"] = close.Date
            removed = close.Source == "RECEIVE_AND_DELIVER" and close.Price == 0
            exp = option_expiry(sym)
            # a true expiration is removed on or just after the expiry date;
            # anything earlier is a contract adjustment (e.g. EOSE -> EOSE1)
            expired = removed and exp is not None and close.Date >= exp - pd.Timedelta(days=3)
            if expired:
                r["Closed Date"] = min(exp, close.Date)
            if not removed:
                r.update(dict(zip(close_cols, (n, close.Price, round(close.Fees * n / close.Qty, 2)))))
        r["Open/Closed"] = "Closed" if close is not None else "Open"
        r[SYNC_COL] = f"{acct}:{lot['fid'] if lot else 'pre'}>{close.FillID if close is not None else ''}"
        r["_open_ids"] = lot["fids"] if lot else frozenset({"pre"})
        r["_close_ids"] = close.FillIDs if close is not None else frozenset({""})
        r["_needs_open"] = lot is None
        adjusted = ((close is not None and removed and not expired)
                    or (lot is not None and lot.get("src") == "RECEIVE_AND_DELIVER"))
        if adjusted:
            r["_flag"] = "CHECK: contract adjusted or delivered by a corporate action; enter by hand"
        elif close is not None and close.Source == "RECEIVE_AND_DELIVER" and not expired:
            r["_flag"] = "assignment/exercise, check prices"
        else:
            r["_flag"] = ""
        return r

    for x in agg.sort_values(["DateTime", "FillID"]).itertuples(index=False):
        key = (x.Account, x.Symbol)
        signed = x.Qty if x.Side == "BUY" else -x.Qty
        remaining = x.Qty
        book = lots[key]
        while remaining > 1e-9 and book and (book[0]["signed"] > 0) != (signed > 0):
            i = pick_lot(book, remaining)
            lot = book[i]
            n = min(remaining, lot["qty"])
            out.append(row(x.Account, x.Symbol, x.AssetType, lot, x, n))
            lot["qty"] -= n
            remaining -= n
            if lot["qty"] <= 1e-9:
                del book[i]
        if remaining > 1e-9:
            closing = x.PositionEffect == "CLOSING" or x.Source == "RECEIVE_AND_DELIVER"
            if closing:  # close of a position opened before SYNC_START
                out.append(row(x.Account, x.Symbol, x.AssetType, None, x, remaining))
            else:
                book.append({"qty": remaining, "orig": remaining, "signed": signed,
                             "price": x.Price, "fees": x.Fees, "date": x.Date, "fid": x.FillID, "fids": x.FillIDs,
                             "asset": x.AssetType, "src": x.Source})

    for (acct, sym), book in lots.items():
        for lot in book:
            out.append(row(acct, sym, lot["asset"], lot, None, lot["qty"]))
    return pd.DataFrame(out)


# ------------------------------------------------------------ reconciliation
def reconcile(pos, trk):
    """Returns list of actions: dict(status, row (excel row or None), values, detail)."""
    actions = []
    claimed = set()
    by_sync = {s: i for i, s in trk[SYNC_COL].items() if not is_blank(s)}
    by_parts = {}
    for s_, i in by_sync.items():
        if ":" in str(s_) and ">" in str(s_):
            acct_, rest = str(s_).split(":", 1)
            o_, c_ = rest.split(">", 1)
            by_parts[(acct_, o_, c_)] = i

    def find_synced(p, open_only=False):
        """Tracker row whose Sync ID matches this position, accepting any partial
        fill of the same order (older versions labeled by whichever piece came first)."""
        closes = {""} if open_only else p["_close_ids"]
        for o_ in p["_open_ids"]:
            for c_ in closes:
                i = by_parts.get((p["Account"], o_, c_))
                if i is not None:
                    return i
        return None
    live = trk[~trk["_empty"]]

    def qty_of(row):
        for col in ("Buy Quantity", "Sell Quantity"):
            v = pd.to_numeric(row[col], errors="coerce") if not isinstance(row, dict) else row[col]
            if not is_blank(v):
                return abs(float(v))
        return None

    def candidates(p, mode):
        c = live[(live["Account"].astype(str) == p["Account"]) &
                 (live["SYMBOL"].map(norm_sym) == norm_sym(p["SYMBOL"])) &
                 (live[SYNC_COL].map(is_blank)) & (~live.index.isin(claimed))]
        qty = qty_of(p)
        if mode == "close":
            c = c[c["Closed Date"].dt.normalize() == pd.Timestamp(p["Closed Date"])]
        else:
            c = c[c["Opened Date"].dt.normalize() == pd.Timestamp(p["Opened Date"])]
        if qty is None or c.empty:
            return c
        return c[c.apply(lambda t: qty_of(t) is not None and abs(qty_of(t) - qty) < 1e-6, axis=1)]

    for _, p in pos.iterrows():
        sid = p[SYNC_COL]
        vals = {c: p[c] for c in DATA_COLS} | {SYNC_COL: sid}
        note = p["_flag"]
        if note.startswith("CHECK"):
            idx = find_synced(p)
            if idx is not None:
                claimed.add(idx)
            actions.append(dict(status="Check", idx=idx, row=None if idx is None else trk.loc[idx, "_row"],
                                values=vals, overwrite=False, apply={}, detail=note))
            continue

        # 1) already synced (or the same position's open-only row from a past sync)
        idx = find_synced(p)
        if idx is None and p["Open/Closed"] == "Closed" and not p["_needs_open"]:
            idx = find_synced(p, open_only=True)
        if idx is not None and idx not in claimed:
            claimed.add(idx)
            t = trk.loc[idx]
            changed = [c for c in DATA_COLS
                       if not (is_blank(vals[c]) and is_blank(t[c]))
                       and (is_blank(vals[c]) != is_blank(t[c]) or not same(vals[c], t[c], c))]
            if changed or t[SYNC_COL] != sid:
                actions.append(dict(status="Synced", idx=idx, row=t["_row"], values=vals, overwrite=True,
                                    detail="updated: " + ", ".join(changed or ["Sync ID"])))
            continue

        # 2) match an existing hand-entered row
        cand = pd.DataFrame()
        if p["Open/Closed"] == "Closed":
            cand = candidates(p, "close")
        if cand.empty and not is_blank(p["Opened Date"]):
            cand = candidates(p, "open")
        if not cand.empty:
            idx = cand.index[0]
            claimed.add(idx)
            t = trk.loc[idx]
            conflicts = [f"{c}: tracker {t[c]} vs Schwab {vals[c]}" for c in DATA_COLS
                         if not is_blank(t[c]) and not is_blank(vals[c]) and not same(t[c], vals[c], c)
                         and c != "Open/Closed"]
            fills = {c: vals[c] for c in DATA_COLS if is_blank(t[c]) and not is_blank(vals[c])}
            if not is_blank(t["Open/Closed"]) and not same(t["Open/Closed"], vals["Open/Closed"], "Open/Closed"):
                fills["Open/Closed"] = vals["Open/Closed"]  # status is factual, take Schwab's
            if conflicts:
                actions.append(dict(status="Mismatch", idx=idx, row=t["_row"], values=vals, overwrite=False,
                                    apply={}, detail="; ".join(conflicts)))
            else:
                fills[SYNC_COL] = sid
                status = "Fill gaps" if len(fills) > 1 else "Match"
                detail = ("filled: " + ", ".join(k for k in fills if k != SYNC_COL)) if len(fills) > 1 else ""
                if p["_needs_open"]:
                    detail = (detail + "; " if detail else "") + "opened before 2026, open side from tracker"
                actions.append(dict(status=status, idx=idx, row=t["_row"], values=vals, overwrite=False,
                                    apply=fills, detail=(detail + (" | " + note if note else "")).strip(" |")))
            continue

        # 3) nothing matched: new row
        status = "Needs open" if p["_needs_open"] else "New"
        detail = "opened before 2026: fill in the open side" if p["_needs_open"] else ""
        actions.append(dict(status=status, idx=None, row=None, values=vals, overwrite=True,
                            detail=(detail + (" | " + note if note else "")).strip(" |")))

    # tracker rows from 2026 that Schwab never matched
    accts = set(pos["Account"]) if len(pos) else set()
    for idx, t in live.iterrows():
        if idx in claimed:
            continue
        recent = any(not pd.isna(t[c]) and t[c] >= SYNC_START for c in DATE_COLS)
        if recent and str(t["Account"]) in accts:
            actions.append(dict(status="Tracker only", idx=idx, row=t["_row"],
                                values={c: t[c] for c in DATA_COLS}, overwrite=False, apply={},
                                detail="no matching Schwab fill; check for typos or a missing trade"))
    return actions


# --------------------------------------------------------------- outputs
def write_review(actions, path):
    order = ["New", "Needs open", "Fill gaps", "Mismatch", "Converted", "Check", "Tracker only", "Synced", "Match"]
    rows = [{"Status": a["status"], "Tracker Row": a["row"], **{c: a["values"].get(c) for c in DATA_COLS},
             "Detail": a["detail"]} for a in actions]
    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=["Status", "Tracker Row"] + DATA_COLS + ["Detail"])
    df["_o"] = df["Status"].map({s: i for i, s in enumerate(order)})
    df = df.sort_values(["_o", "Account", "Closed Date", "Opened Date"]).drop(columns="_o")
    summary = df["Status"].value_counts().reindex(order, fill_value=0).rename("Count").reset_index()
    summary.columns = ["Status", "Count"]
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="mm/dd/yyyy") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        df.to_excel(xw, sheet_name="Review", index=False)
        for ws in xw.sheets.values():
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    max(len(str(c.value or "")) for c in col[:200]) + 2, 60)
    return summary


EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def to_serial(v):
    """Whole-day Excel date number. Writing serials (not datetimes) keeps Windows
    from shifting the value by your UTC offset, which left 6:00/7:00 AM times."""
    return float((pd.Timestamp(v).normalize() - EXCEL_EPOCH).days)


def to_com(v, col):
    if is_blank(v):
        return None
    if col in DATE_COLS:
        return to_serial(v)
    if col in NUM_COLS:
        return float(v)
    return str(v)


SORT_COL = "Sort Date"


PCT_FORMULA = ('=IFERROR([@[Net Profit]]/IF([@[TOTAL INVESTMENT]]>0,[@[TOTAL INVESTMENT]],'
               '[@[Total Sale Value]]),"")')


def set_pct_change(lo, names):
    """% CHANGE = return on cost for buys; for sold-to-open hedges (no cost) it's
    the share of the premium you kept. Avoids #DIV/0! on hedge rows."""
    needed = {"% CHANGE", "Net Profit", "TOTAL INVESTMENT", "Total Sale Value"}
    if needed.issubset(names):
        lo.ListColumns("% CHANGE").DataBodyRange.Formula = PCT_FORMULA


def sort_table(ws, lo):
    """Closed trades first, oldest to newest by Closed Date; then open
    positions, oldest holding first by Opened Date (unknown open dates first).
    Blank rows sink to the bottom. A position joins the closed block once it
    closes. Uses a small helper column at the end of the table."""
    names = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]

    # strip stray times from date cells (left by earlier versions of this script)
    for col in ("Opened Date", "Closed Date"):
        rng = lo.ListColumns(col).DataBodyRange
        vals = rng.Value2  # raw serial numbers, no time zone conversion
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        fixed = tuple(((float(int(v)) if isinstance(v, float) else v),) for (v,) in vals)
        rng.Value2 = fixed

    if SORT_COL not in names:
        lo.ListColumns.Add().Name = SORT_COL
    body = lo.ListColumns(SORT_COL).DataBodyRange
    body.Formula = ('=IF([@[Open/Closed]]="Closed",[@[Closed Date]],'
                    'IF([@[Open/Closed]]="Open",N([@[Opened Date]]),""))')
    body.NumberFormat = "mm/dd/yyyy"

    xlAscending, xlSortOnValues, xlYes = 1, 0, 1
    sort = lo.Sort
    sort.SortFields.Clear()
    xlDescending = 2
    open_first = os.getenv("OPEN_ON_TOP", "no").lower() in ("1", "yes", "true", "y")
    # "Closed" sorts before "Open"; descending flips open positions to the top
    sort.SortFields.Add(lo.ListColumns("Open/Closed").DataBodyRange, xlSortOnValues,
                        xlDescending if open_first else xlAscending)
    sort.SortFields.Add(lo.ListColumns(SORT_COL).DataBodyRange, xlSortOnValues, xlAscending)
    sort.Header = xlYes
    sort.Apply()


def apply_with_excel(src, dst, actions, trk):
    try:
        import win32com.client as win32
    except ImportError:
        sys.exit("Applying changes needs Excel automation. Run: py -m pip install pywin32")

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Open(str(src.resolve()), UpdateLinks=0, ReadOnly=True)
        lo = ws = None
        for sh in wb.Worksheets:
            for t in sh.ListObjects:
                if t.Name == TABLE:
                    lo, ws = t, sh
        if lo is None:
            raise SystemExit(f"Couldn't find the table {TABLE} in {src.name}.")
        names = [lo.ListColumns(i).Name for i in range(1, lo.ListColumns.Count + 1)]
        if SYNC_COL not in names:
            col = lo.ListColumns.Add()
            col.Name = SYNC_COL
            names.append(SYNC_COL)
        colnum = {n: lo.ListColumns(n).Range.Column for n in names}

        def write(r, values):
            for c, v in values.items():
                if c in colnum:
                    ws.Cells(r, colnum[c]).Value = to_com(v, c)

        empty_rows = deque(sorted(trk.loc[trk["_empty"], "_row"].tolist()))
        changed = 0
        for a in actions:
            if a["status"] in ("Mismatch", "Tracker only", "Check"):
                continue
            if a["row"] is not None:
                if a["overwrite"]:
                    write(a["row"], a["values"])  # Synced rows: Schwab is the source
                elif a.get("apply"):
                    write(a["row"], a["apply"])   # hand-entered rows: blanks only
                else:
                    continue
            else:
                if empty_rows:
                    r = empty_rows.popleft()
                else:
                    r = lo.ListRows.Add().Range.Row
                write(r, a["values"])
            changed += 1

        set_pct_change(lo, names)
        sort_table(ws, lo)
        wb.RefreshAll()
        xl.CalculateUntilAsyncQueriesDone()
        wb.SaveAs(str(dst.resolve()), FileFormat=51)  # 51 = .xlsx
        return changed
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()


# ------------------------------------------------------------------ main
def main():
    args = set(sys.argv[1:])
    if "--no-sync" not in args:
        import sync_trades
        sync_trades.main()
        print()
    if not CACHE.exists():
        sys.exit("No fills_cache.csv yet. Run: py sync_trades.py")

    names = account_map()
    fills = pd.read_csv(CACHE, dtype={"Account": str, "OrderID": str, "FillID": str})
    unmapped = sorted({a[-4:] for a in fills["Account"].astype(str)} - set(names))
    if unmapped:
        sys.exit("Add these accounts to ACCOUNT_NAMES in .env, using the names in your tracker:\n"
                 + "\n".join(f"  ...{a}" for a in unmapped)
                 + "\nExample: ACCOUNT_NAMES=1234=Roth,5678=Emma")

    src = find_latest_tracker()
    dst = TRACKER_DIR / next_version_name()
    print(f"Tracker: {src.name}")
    trk, _ = read_tracker(src)
    pos = build_positions(fills, names)
    actions = convert_old_shorts(trk) + reconcile(pos, trk)

    review = HERE / f"Review_{dst.stem.replace(PREFIX, '')}.xlsx"
    summary = write_review(actions, review)
    print("\n".join(f"  {s:<13}{n}" for s, n in summary.itertuples(index=False)))
    print(f"\nReview file: {review.name}")

    if "--apply" in args:
        todo = [a for a in actions if a["status"] not in ("Mismatch", "Tracker only", "Check", "Match")
                or (a["status"] == "Match" and len(a.get("apply", {})) > 1)]
        if not todo and "--force" not in args:
            print(f"No new trades or changes. No new version saved; {src.name} is still current.")
            return
        n = apply_with_excel(src, dst, actions, trk)
        print(f"Saved {dst.name} with {n} rows added or updated. {src.name} was not changed.")
        # open the new version in Excel (turn off with OPEN_AFTER_UPDATE=no in .env, or --no-open)
        if (os.getenv("OPEN_AFTER_UPDATE", "yes").lower() in ("1", "yes", "true", "y")
                and "--no-open" not in args and hasattr(os, "startfile")):
            try:
                os.startfile(str(dst.resolve()))
            except OSError as e:
                print(f"Couldn't open {dst.name} automatically: {e}")
    else:
        print("Nothing in the tracker was changed. Look over the review, then run with --apply.")


if __name__ == "__main__":
    main()
