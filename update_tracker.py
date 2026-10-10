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
    return f"{canon_root(root)} {mm}/{dd}/20{yy} {int(strike) / 1000:.2f} {cp}"


def canon_root(root):
    """Schwab renames adjusted option contracts (EOSE -> EOSE1 after a reorg),
    even on trades made before it. They're the same contract here."""
    r = re.sub(r"\d+$", "", str(root).strip().upper())
    return r or str(root).strip().upper()


def underlying(sym):
    """Ticker a symbol belongs to: 'EOSE 01/15/2027 10.00 C' -> 'EOSE'."""
    return canon_root(str(sym).strip().split(" ")[0]) if str(sym).strip() else ""


def option_expiry(sym):
    m = re.match(r"^\S+\s*(\d{2})(\d{2})(\d{2})[CP]\d{8}$", str(sym).strip())
    if m:
        return pd.Timestamp(f"20{m.group(1)}-{m.group(2)}-{m.group(3)}")
    m = re.match(r"^\S+ (\d{1,2})/(\d{1,2})/(\d{2}|\d{4}) [\d.]+ [CP]$", str(sym).strip().upper())
    if m:
        yy = m.group(3) if len(m.group(3)) == 4 else "20" + m.group(3)
        return pd.Timestamp(f"{yy}-{int(m.group(1)):02d}-{int(m.group(2)):02d}")
    return None


def norm_sym(s):
    t = re.sub(r"\s+", " ", str(s or "")).strip().upper()
    m = re.match(r"^(\S+) (\d{1,2})/(\d{1,2})/(\d{2}|\d{4}) ([\d.]+) ([CP])$", t)
    if m:  # option: same contract however the date/strike was typed
        root, mm, dd, yy, k, cp = m.groups()
        yy = yy if len(yy) == 4 else "20" + yy
        try:
            k = f"{float(k):.2f}"
        except ValueError:
            pass
        t = f"{canon_root(root)} {int(mm):02d}/{int(dd):02d}/{yy} {k} {cp}"
    return t


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
    wbv = openpyxl.load_workbook(path, data_only=True)  # computed values
    ws = next((s for s in wb.worksheets if TABLE in s.tables), None)
    if ws is None:
        sys.exit(f"Couldn't find the table {TABLE} in {path.name}. If you renamed it, "
                 f"add TRACKER_TABLE=<new name> to .env.")
    tbl = ws.tables[TABLE]
    (c1, r1), (c2, r2) = [openpyxl.utils.cell.coordinate_to_tuple(x)[::-1] for x in tbl.ref.split(":")]
    headers = [ws.cell(r1, c).value for c in range(c1, c2 + 1)]
    rows = []
    for r in range(r1 + 1, r2 + 1):
        vals = {}
        for i, h in enumerate(headers):
            v = ws.cell(r, c1 + i).value
            if isinstance(v, str) and v.startswith("=") and h in DATA_COLS:
                v = wbv[ws.title].cell(r, c1 + i).value  # what Excel shows
            vals[h] = v
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
        z = 0 if str(t["Open/Closed"]).strip().lower() == "closed" else None
        new = {"Sell Quantity": -float(b), "Sell Price Per Share": t["Purchase PRICE PER SHARE"],
               "Sale Fees & Comm": t["Fees & Comm"], "Buy Quantity": z,
               "Purchase PRICE PER SHARE": z, "Fees & Comm": z}
        for k, v in new.items():
            trk.at[idx, k] = v
        vals = {c: trk.at[idx, c] for c in DATA_COLS}
        acts.append(dict(status="Converted", idx=idx, row=t["_row"], values=vals, overwrite=False,
                         apply=new, detail="sold to open: moved from Buy to Sell columns"))
    return acts


# ------------------------------------------------- 2025 lots (starting inventory)
SEED_RE = re.compile(r"^seed:(\d{8}):(S?)([\d.]+)(#\d+)?$")


def _num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if f != f else f


def _fmt(x):
    return f"{x:.6f}".rstrip("0").rstrip(".")


def _split_factor(splits, acct, sym, since, until):
    """Product of split ratios for this holding between its buy date and the
    row's close date (or now, if still open). Row price x factor = the price
    as bought; row shares / factor = the shares as bought."""
    f = 1.0
    for ev in splits or []:
        r = ev.get("Ratio") or 0
        if (ev["Account"] == acct and ev["SYMBOL"] == sym and r > 0 and abs(r - 1) > 1e-9
                and since < pd.Timestamp(ev["Date"]) and (until is None or pd.Timestamp(ev["Date"]) <= until)):
            f *= r
    return f


def collect_seeds(trk, splits=None):
    """Lots you held on Jan 1: rows opened before SYNC_START that are still open
    (no Closed Date) with their cost filled in. Sales in 2026 draw from these.

    Each lot gets a Sync ID like  Roth:seed:20251126:1030>  (date and original
    quantity), shared by every row the lot is later split into. Price and fees
    come from the lot's open row (or its first row once it's fully sold), so a
    price fix on that row carries to the others on the next update.
    Returns ({(account, symbol): [lots]}, {tracker index: Sync ID to write})."""
    by_key = {}
    pending = {}
    order = []
    for idx, t in trk.iterrows():
        if t["_empty"]:
            continue
        od = t["Opened Date"]
        if pd.isna(od) or od >= SYNC_START:
            continue
        acct, sym = str(t["Account"]), norm_sym(t["SYMBOL"])
        asset = "OPTION" if str(t["Stock Or Option"]).strip().lower() == "option" else "EQUITY"
        sid = "" if is_blank(t[SYNC_COL]) else str(t[SYNC_COL])
        b, sq = _num(t["Buy Quantity"]), _num(t["Sell Quantity"])

        if ":seed:" in sid:
            key = sid.split(":", 1)[1].split(">", 1)[0]
            m = SEED_RE.match(key)
            if not m:
                continue
            short, qty = m.group(2) == "S", float(m.group(3))
            anchor = sid.endswith(">")  # the lot's open remainder row
            k = (acct, sym, key)
            if k not in by_key:
                order.append(k)
                by_key[k] = dict(qty=qty, short=short, date=od, asset=asset, sym=str(t["SYMBOL"]).strip(),
                                 action=t.get("Action"), price=None, fees=None, anchor=False)
            lot = by_key[k]
            if lot["price"] is None or (anchor and not lot["anchor"]):
                pc, fc = (("Sell Price Per Share", "Sale Fees & Comm") if short
                          else ("Purchase PRICE PER SHARE", "Fees & Comm"))
                # rows show split-adjusted shares and price; undo that to get the
                # lot as bought (otherwise every run would apply the split again)
                until = None if pd.isna(t["Closed Date"]) else t["Closed Date"]
                fac = 1.0 if (short or asset != "EQUITY") else _split_factor(splits, acct, sym, od, until)
                rq = ((sq if short else b) or 0) / fac
                fee = _num(t[fc]) or 0
                lot.update(price=(_num(t[pc]) or 0) * fac, anchor=anchor,
                           fees=fee * qty / rq if rq else fee)  # back to the whole lot's fees
            continue

        if sid:
            o_part = sid.split(":", 1)[1].split(">", 1)[0] if ":" in sid and ">" in sid else ""
            if o_part in ("pre", "") or o_part.startswith("seed:"):
                continue  # a script row; its open side is already accounted for
            # a script row opened from a 2026 fill can't have a 2025 open date:
            # you retyped it into a 2025 lot, so treat it as one
        if str(t["Open/Closed"]).strip().lower() == "closed" and not sid:
            continue
        if not pd.isna(t["Closed Date"]):
            continue
        if b and b > 0 and not (sq and sq > 0):
            short, qty, price, fees = False, b, _num(t["Purchase PRICE PER SHARE"]), _num(t["Fees & Comm"]) or 0
        elif sq and sq > 0 and not (b and b > 0):
            short, qty, price, fees = True, sq, _num(t["Sell Price Per Share"]), _num(t["Sale Fees & Comm"]) or 0
        else:
            continue
        if price is None:
            continue  # cost not entered yet
        key = f"seed:{od:%Y%m%d}:{'S' if short else ''}{_fmt(qty)}"
        n = 2
        while (acct, sym, key) in by_key:
            key = f"seed:{od:%Y%m%d}:{'S' if short else ''}{_fmt(qty)}#{n}"
            n += 1
        k = (acct, sym, key)
        order.append(k)
        by_key[k] = dict(qty=qty, short=short, date=od, asset=asset, sym=str(t["SYMBOL"]).strip(),
                         action=t.get("Action"), price=price, fees=fees, anchor=True)
        pending[idx] = f"{acct}:{key}>"

    seeds = defaultdict(list)
    for k in order:
        acct, sym, key = k
        L = by_key[k]
        seeds[(acct, sym)].append({
            "qty": L["qty"], "orig": L["qty"], "signed": -L["qty"] if L["short"] else L["qty"],
            "price": L["price"] or 0, "fees": L["fees"] or 0, "date": L["date"], "fid": key,
            "fids": frozenset({key}), "asset": L["asset"], "src": "SEED", "sym": L["sym"],
            "action": L["action"]})
    for book in seeds.values():
        book.sort(key=lambda l: l["date"])
    return seeds, pending


CARRY = CACHE.with_name("carry_cache.csv")
HISTORY_DIR = HERE / "history"

# Schwab website export actions -> (asset, side, effect). Expired/assigned/
# exercised legs are $0 removals; the sign of Quantity says which side.
_HIST_ACTIONS = {
    "buy": ("EQUITY", "BUY", "OPENING"), "sell": ("EQUITY", "SELL", "CLOSING"),
    "buy to open": ("OPTION", "BUY", "OPENING"), "sell to close": ("OPTION", "SELL", "CLOSING"),
    "sell to open": ("OPTION", "SELL", "OPENING"), "buy to close": ("OPTION", "BUY", "CLOSING"),
}
_HIST_REMOVALS = {"expired", "assigned", "exchange or exercise"}


def _money(v):
    v = str(v or "").replace("$", "").replace(",", "").strip()
    try:
        return float(v)
    except ValueError:
        return None


def load_history(names):
    """Schwab transaction exports (schwab.com > History > Export) saved in the
    'history' folder, e.g. Roth_XXX946_Transactions_20261009.csv. Only trades
    before SYNC_START are used, to work out the lots held on Jan 1. Returns
    (fills, covered) where covered = {(account, ticker)} the exports cover."""
    rows, covered = [], set()
    if not HISTORY_DIR.exists():
        return pd.DataFrame(), covered
    by_tail = {k[-3:]: v for k, v in names.items()}
    for path in sorted(HISTORY_DIR.glob("*.csv")):
        m = re.search(r"XXX(\d{3,4})", path.name)
        acct = by_tail.get(m.group(1)[-3:]) if m else None
        if acct is None:
            print(f"  (skipping {path.name}: can't tell which account it is; "
                  "keep Schwab's file name, e.g. Roth_XXX946_...)")
            continue
        try:
            raw = pd.read_csv(path, dtype=str).fillna("")
        except Exception as e:
            print(f"  (skipping {path.name}: {e})")
            continue
        raw = raw.iloc[::-1].reset_index(drop=True)  # exports are newest first
        for i, r in raw.iterrows():
            d = re.findall(r"\d{2}/\d{2}/\d{4}", r.get("Date", ""))
            if not d:
                continue
            when = pd.to_datetime(d[-1], format="%m/%d/%Y")  # "as of" date when given
            act = r.get("Action", "").strip().lower()
            sym = norm_sym(r.get("Symbol", ""))
            q = _money(r.get("Quantity"))
            if not sym or q is None or q == 0:
                continue
            is_opt = bool(re.match(r"^\S+ \d{2}/\d{2}/\d{4} [\d.]+ [CP]$", sym))
            covered.add((acct, underlying(sym)))
            if when >= SYNC_START:
                continue
            if act in _HIST_ACTIONS:
                asset, side, eff = _HIST_ACTIONS[act]
                asset = "OPTION" if is_opt else asset
                price, src = _money(r.get("Price")) or 0.0, "TRADE"
            elif act in _HIST_REMOVALS and is_opt:
                # +qty closes a short (bought back at $0), -qty closes a long
                asset, side, eff, price, src = "OPTION", ("BUY" if q > 0 else "SELL"), "CLOSING", 0.0, \
                    "RECEIVE_AND_DELIVER"
            else:
                continue  # dividends, journals, reorg notices, etc.
            rows.append({"FillID": f"H{path.stem[-15:]}-{i}", "Account": acct,
                         "DateTime": when + pd.Timedelta(seconds=i), "Date": when.date(),
                         "Symbol": sym, "Underlying": underlying(sym), "AssetType": asset,
                         "Side": side, "PositionEffect": eff, "Qty": abs(q), "Price": price,
                         "Fees": _money(r.get("Fees & Comm")) or 0.0, "NetAmount": _money(r.get("Amount")) or 0.0,
                         "Source": src, "OrderID": ""})
    return pd.DataFrame(rows), covered


def carry_lots(names):
    """Lots still open on SYNC_START according to Schwab's 2025 fills (oldest
    sold first, like your account setting). None if the 2025 pull hasn't run."""
    hist, covered = load_history(names)
    if not CARRY.exists() and hist.empty:
        return None, covered
    parts = []
    if CARRY.exists():
        f = pd.read_csv(CARRY, dtype={"Account": str, "OrderID": str, "FillID": str})
        if not f.empty:
            f["Account"] = f["Account"].astype(str).str[-4:].map(lambda a: names.get(a, f"...{a}"))
            f["_u"] = [underlying(fmt_symbol(s, a)) for s, a in zip(f.Symbol, f.AssetType)]
            f = f[[(a, u) not in covered for a, u in zip(f.Account, f["_u"])]].drop(columns="_u")
            parts.append(f)
    if not hist.empty:
        parts.append(hist)
    if not parts:
        return {}, covered
    f = pd.concat(parts, ignore_index=True)
    if f.empty:
        return {}, covered
    f["DateTime"] = pd.to_datetime(f["DateTime"])
    # names are already applied; pass an identity map so build_positions keeps them
    ident = {a: a for a in f["Account"].astype(str).unique()}
    f["Account"] = f["Account"].astype(str)
    _, lots = build_positions(f, _Ident(ident), fifo=True, return_lots=True)
    seeds = defaultdict(list)
    unknown = []
    for (acct, sym), book in lots.items():
        used = set()
        for l in sorted(book, key=lambda l: l["date"]):
            if l["qty"] <= 1e-9:
                continue
            if l.get("src") == "UNKNOWN":
                unknown.append({"Account": acct, "SYMBOL": sym, "Shares": l["qty"]})
                continue
            short = l["signed"] < 0
            key = f"seed:{l['date']:%Y%m%d}:{'S' if short else ''}{_fmt(l['qty'])}"
            n = 2
            while key in used:
                key = f"seed:{l['date']:%Y%m%d}:{'S' if short else ''}{_fmt(l['qty'])}#{n}"
                n += 1
            used.add(key)
            fees = l["fees"] * l["qty"] / l["orig"] if l["orig"] else l["fees"]
            seeds[(acct, sym)].append({
                "qty": l["qty"], "orig": l["qty"], "signed": -l["qty"] if short else l["qty"],
                "price": l["price"], "fees": round(fees, 2), "date": l["date"], "fid": key,
                "fids": frozenset({key}), "asset": l["asset"], "src": "SEED",
                "sym": fmt_symbol(l.get("sym", sym), l["asset"]), "action": None})
    seeds = dict(seeds)
    seeds["__unknown__"] = unknown
    return seeds, covered


class _Ident(dict):
    """Account map for already-named accounts: build_positions looks names up
    by the last 4 characters, so map those back to the full name."""
    def get(self, k, default=None):
        for name in self:
            if str(name)[-4:] == k:
                return name
        return default


def merge_seeds(carry, mine, trk, pending):
    """Schwab's Jan 1 lots win where they exist; your hand-entered lots fill in
    anything older than Schwab's history. A hand row for the same lot (same
    account, symbol and buy date) is adopted and corrected to Schwab's numbers."""
    report = []
    schwab, covered = carry if isinstance(carry, tuple) else (carry, set())
    if schwab is None:
        return mine, pending, report
    schwab = dict(schwab)
    for u in schwab.pop("__unknown__", []):
        report.append({"Account": u["Account"], "SYMBOL": u["SYMBOL"], "Bought": None,
                       "Schwab qty": u["Shares"], "Schwab price": None, "Your qty": None, "Your price": None,
                       "Result": "shares bought before Schwab's history reaches: enter this lot by hand"})
    pend_by_key = {}
    for idx, sid in pending.items():
        acct, rest = sid.split(":", 1)
        pend_by_key[(acct, norm_sym(trk.at[idx, "SYMBOL"]), rest.split(">", 1)[0])] = idx
    merged, new_pending = defaultdict(list), {}
    for k in sorted(set(schwab) | set(mine)):
        S, T = [dict(x) for x in schwab.get(k, [])], list(mine.get(k, []))
        adopted, matched = set(), set()
        for t in T:
            idx = pend_by_key.get((k[0], k[1], t["fid"]))
            same_day = [x for x in S if x["date"].normalize() == t["date"].normalize()]
            if same_day:
                x = next((y for y in same_day if abs(y["qty"] - t["qty"]) < 1e-6), same_day[0])
                matched.add(x["fid"])
                if idx is not None and x["fid"] not in adopted:
                    new_pending[idx] = f"{k[0]}:{x['fid']}>"
                    adopted.add(x["fid"])
                if not is_blank(t.get("action")):
                    x["action"] = t["action"]
                    for y in S:
                        if y["fid"] == x["fid"]:
                            y["action"] = t["action"]
                same = abs(x["qty"] - t["qty"]) < 1e-6 and abs((x["price"] or 0) - (t["price"] or 0)) < 0.005
                report.append({"Account": k[0], "SYMBOL": k[1], "Bought": t["date"],
                               "Schwab qty": x["qty"], "Schwab price": x["price"],
                               "Your qty": t["qty"], "Your price": t["price"],
                               "Result": "matches your row" if same else "your row corrected to Schwab's numbers"})
                continue
            if (k[0], underlying(k[1])) in covered:
                # your export covers this ticker's whole history and has no such lot
                report.append({"Account": k[0], "SYMBOL": k[1], "Bought": t["date"], "Schwab qty": None,
                               "Schwab price": None, "Your qty": t["qty"], "Your price": t["price"],
                               "Result": "NOT in Schwab's history: likely a duplicate, delete your row",
                               "Row": trk.at[idx, "_row"] if idx is not None else None})
                continue
            merged[k].append(t)
            if idx is not None:
                new_pending[idx] = pending[idx]
            report.append({"Account": k[0], "SYMBOL": k[1], "Bought": t["date"], "Schwab qty": None,
                           "Schwab price": None, "Your qty": t["qty"], "Your price": t["price"],
                           "Result": "your lot, not in Schwab's 2025 history (kept as entered)"})
        for x in S:
            merged[k].append(x)
            if x["fid"] not in matched:
                report.append({"Account": k[0], "SYMBOL": k[1], "Bought": x["date"], "Schwab qty": x["qty"],
                               "Schwab price": x["price"], "Your qty": None, "Your price": None,
                               "Result": "added from Schwab's 2025 history"})
        merged[k].sort(key=lambda l: l["date"])
    return merged, new_pending, report


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


def build_positions(fills, names, seeds=None, fifo=False, return_lots=False):
    """Match fills into tracker-style position rows (see pick_lot). seeds: your
    2025 lots, used as the starting inventory. fifo: always pair with the oldest
    lot (used to work out what was still held on Jan 1). return_lots: also
    return the lots still open at the end."""
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

    # reverse/forward splits: Schwab removes the old shares and delivers the new
    # ones at $0 in the same minute. Convert the open lots instead of treating
    # it as a sale; same-count swaps (a new security ID) change nothing.
    splits = {}      # FillID -> event
    corp = agg[(agg.AssetType.isin(["EQUITY", "OPTION"])) & (agg.Source == "RECEIVE_AND_DELIVER")
               & (agg.Price == 0)]
    if len(corp):
        corp = corp.assign(_k=[norm_sym(fmt_symbol(s, a)) for s, a in zip(corp.Symbol, corp.AssetType)],
                           _t=corp.DateTime.dt.floor("min"))
        for (acct, sym, t), g in corp.groupby(["Account", "_k", "_t"]):
            if g.AssetType.iloc[0] == "EQUITY":
                out_, in_ = g[g.Side == "SELL"], g[g.Side == "BUY"]
            else:
                out_, in_ = g[g.PositionEffect == "CLOSING"], g[g.PositionEffect == "OPENING"]
            if len(out_) == 1 and len(in_) == 1:
                a_, b_ = float(out_.Qty.iloc[0]), float(in_.Qty.iloc[0])
                if g.AssetType.iloc[0] == "OPTION" and abs(a_ - b_) > 1e-9:
                    continue  # an option adjustment that changes the count: leave for review
                ev = {"key": (acct, sym), "ratio": b_ / a_ if a_ else 1.0, "old": a_, "new": b_,
                      "done": False, "date": t.normalize()}
                splits[out_.FillID.iloc[0]] = ev
                splits[in_.FillID.iloc[0]] = ev
    split_log = []

    lots = defaultdict(list)
    for k, book in (seeds or {}).items():
        lots[k] = [dict(l) for l in book]
    out = []

    # covered calls assigned: Schwab removes the calls at $0 and sells the shares
    # at the strike. Find those share sales so they come off the oldest shares
    # first and get tagged "Assigned".
    assigned_ids = {}  # share-sale fill -> expiration date of the assigned calls
    removals = []
    for x in agg.itertuples(index=False):
        m = re.match(r"^(\S+?)\d?\s*(\d{2})(\d{2})(\d{2})C(\d{8})$", str(x.Symbol).strip())
        if x.AssetType == "OPTION" and x.Source == "RECEIVE_AND_DELIVER" and x.Price == 0 and m:
            exp = pd.Timestamp(f"20{m.group(2)}-{m.group(3)}-{m.group(4)}")
            removals.append((x.Account, m.group(1), exp, int(m.group(5)) / 1000, x.Qty))
    for x in agg.itertuples(index=False):
        if x.AssetType != "EQUITY" or x.Side != "SELL":
            continue
        for acct, root, exp, strike, n in removals:
            if (acct == x.Account and root == str(x.Symbol).strip() and abs(x.Price - strike) < 0.005
                    and exp - pd.Timedelta(days=1) <= x.Date <= exp + pd.Timedelta(days=4)
                    and x.Qty <= n * 100 + 1e-9):
                assigned_ids[x.FillID] = exp
                break

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
                # expired worthless: close side shows a $0 close, not blanks
                r.update(dict(zip(close_cols, (0, 0, 0))))
            if not removed:
                r.update(dict(zip(close_cols, (n, close.Price, round(close.Fees * n / close.Qty, 2)))))
        r["Open/Closed"] = "Closed" if close is not None else "Open"
        r[SYNC_COL] = f"{acct}:{lot['fid'] if lot else 'pre'}>{close.FillID if close is not None else ''}"
        r["_open_ids"] = lot["fids"] if lot else frozenset({"pre"})
        r["_close_ids"] = close.FillIDs if close is not None else frozenset({""})
        r["_needs_open"] = lot is None
        r["_short"] = short
        r["_assigned"] = close is not None and close.FillID in assigned_ids
        if r["_assigned"]:
            # Schwab posts assignments the next morning; date them on expiration
            r["Closed Date"] = min(assigned_ids[close.FillID], close.Date)
        r["_action"] = lot.get("action") if lot else None
        r["_open_cols"] = ("Opened Date",) + open_cols
        adjusted = ((close is not None and removed and not expired)
                    or (lot is not None and lot.get("src") == "RECEIVE_AND_DELIVER"))
        adjusted = adjusted and not (lot is not None and lot.get("src") == "SEED" and close is None)
        if adjusted:
            r["_flag"] = "CHECK: contract adjusted or delivered by a corporate action; enter by hand"
        elif close is not None and close.Source == "RECEIVE_AND_DELIVER" and not expired:
            r["_flag"] = "assignment/exercise, check prices"
        else:
            r["_flag"] = ""
        return r

    for x in agg.sort_values(["DateTime", "FillID"]).itertuples(index=False):
        key = (x.Account, norm_sym(fmt_symbol(x.Symbol, x.AssetType)))
        ev = splits.get(x.FillID)
        if ev is not None:
            if not ev["done"]:
                ev["done"] = True
                held = sum(l["qty"] for l in lots[ev["key"]] if l["signed"] > 0)
                if fifo and ev["old"] - held > 1e-9 and ev["key"][1].isalpha():
                    # shares held from before the history we have: unknown cost
                    lots[ev["key"]].insert(0, {
                        "qty": ev["old"] - held, "orig": ev["old"] - held, "signed": ev["old"] - held,
                        "price": None, "fees": 0, "date": pd.Timestamp("1900-01-01"), "fid": "unknown",
                        "fids": frozenset({"unknown"}), "asset": "EQUITY", "src": "UNKNOWN", "sym": ev["key"][1]})
                if abs(ev["ratio"] - 1) > 1e-9:
                    for l in lots[ev["key"]]:
                        if l["signed"] > 0 and l["asset"] == "EQUITY":
                            l["qty"] *= ev["ratio"]
                            l["orig"] *= ev["ratio"]
                            l["signed"] = l["qty"]
                            if l["price"] is not None:
                                l["price"] = l["price"] / ev["ratio"]
                split_log.append({"Account": ev["key"][0], "SYMBOL": ev["key"][1],
                                  "Date": ev["date"], "Old shares": ev["old"], "New shares": ev["new"],
                                  "Ratio": round(ev["ratio"], 6), "Known lots before": held})
            continue
        signed = x.Qty if x.Side == "BUY" else -x.Qty
        remaining = x.Qty
        book = lots[key]
        is_assigned = x.FillID in assigned_ids
        while remaining > 1e-9 and book and (book[0]["signed"] > 0) != (signed > 0):
            i = 0 if (is_assigned or fifo) else pick_lot(book, remaining)  # assignments: oldest first
            lot = book[i]
            n = min(remaining, lot["qty"])
            out.append(row(x.Account, x.Symbol, x.AssetType, lot, x, n))
            lot["qty"] -= n
            remaining -= n
            if lot["qty"] <= 1e-9:
                del book[i]
        if remaining > 1e-9:
            closing = x.PositionEffect == "CLOSING" or x.Source == "RECEIVE_AND_DELIVER"
            if (closing and x.AssetType == "EQUITY" and x.Source == "RECEIVE_AND_DELIVER"
                    and x.Price == 0 and abs(remaining - x.Qty) < 1e-9):
                # shares removed for $0 that we never had a lot for (e.g. rights
                # that expired worthless): nothing gained or lost, nothing to record
                split_log.append({"Account": x.Account, "SYMBOL": key[1], "Date": x.Date,
                                  "Old shares": x.Qty, "New shares": 0, "Ratio": 0,
                                  "Known lots before": 0})
                continue
            if closing:  # close of a position opened before SYNC_START
                out.append(row(x.Account, x.Symbol, x.AssetType, None, x, remaining))
            else:
                book.append({"qty": remaining, "orig": remaining, "signed": signed,
                             "price": x.Price, "fees": x.Fees, "date": x.Date, "fid": x.FillID, "fids": x.FillIDs,
                             "asset": x.AssetType, "src": x.Source, "sym": x.Symbol})

    for (acct, sym), book in lots.items():
        for lot in book:
            if lot.get("src") == "UNKNOWN":
                continue
            out.append(row(acct, lot.get("sym", sym), lot["asset"], lot, None, lot["qty"]))
    df = pd.DataFrame(out)
    df.attrs["splits"] = split_log
    return (df, lots) if return_lots else df


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
            if not is_blank(v) and float(v) != 0:
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
        tag = {"Action": "Assigned"} if p.get("_assigned") else {}
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
            if p["_needs_open"]:
                # opened before Schwab's data starts: the open side is yours
                # (entered by hand), so only the closing side is ever written
                keep = set(p["_open_cols"])
                upd = {c: vals[c] for c in DATA_COLS if c not in keep and not is_blank(vals[c])
                       and (is_blank(t[c]) or not same(vals[c], t[c], c))}
                if tag and "Action" in t and is_blank(t["Action"]):
                    upd |= tag
                if t[SYNC_COL] != sid:
                    upd[SYNC_COL] = sid
                if upd:
                    actions.append(dict(status="Synced", idx=idx, row=t["_row"], values=vals,
                                        overwrite=False, apply=upd,
                                        detail="updated: " + ", ".join(upd)))
                continue
            changed = [c for c in DATA_COLS
                       if not (is_blank(vals[c]) and is_blank(t[c]))
                       and (is_blank(vals[c]) != is_blank(t[c]) or not same(vals[c], t[c], c))]
            if tag and "Action" in t and is_blank(t["Action"]):
                vals = vals | tag
                changed.append("Action")
            if changed or t[SYNC_COL] != sid or t.get("_sid_pending") is True:
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

        # 2b) a pre-2026 close you split into several rows (one per 2025 buy):
        # rows with the same account, contract and close date that add up
        if p["_needs_open"] and p["Open/Closed"] == "Closed":
            grp = live[(live["Account"].astype(str) == p["Account"]) &
                       (live["SYMBOL"].map(norm_sym) == norm_sym(p["SYMBOL"])) &
                       (live[SYNC_COL].map(is_blank)) & (~live.index.isin(claimed)) &
                       (live["Closed Date"].dt.normalize() == pd.Timestamp(p["Closed Date"]))]
            q = qty_of(p)
            if len(grp) > 1 and q and abs(sum(qty_of(t) or 0 for _, t in grp.iterrows()) - q) < 1e-6:
                for gi, t in grp.iterrows():
                    claimed.add(gi)
                    fix = {"Open/Closed": "Closed"} if is_blank(t["Open/Closed"]) else {}
                    actions.append(dict(status="Match", idx=gi, row=t["_row"], values=vals,
                                        overwrite=False, apply=fix,
                                        detail=f"part of a {q:g} sale you split into {len(grp)} rows"))
                continue

        # 3) nothing matched: new row
        status = "Needs open" if p["_needs_open"] else "New"
        detail = "opened before 2026: fill in the open side" if p["_needs_open"] else ""
        inherit = {} if is_blank(p.get("_action")) else {"Action": p.get("_action")}
        actions.append(dict(status=status, idx=None, row=None, values=vals | inherit | tag, overwrite=True,
                            detail=(detail + (" | " + note if note else "")).strip(" |")))

    # tracker rows from 2026 that Schwab never matched
    accts = set(pos["Account"]) if len(pos) else set()
    for idx, t in live.iterrows():
        if idx in claimed:
            continue
        if not is_blank(t[SYNC_COL]) and t.get("_sid_pending") is not True:
            # a script row the current pairing no longer produces (its trade is
            # now on other rows, e.g. after 2025 lots were added)
            actions.append(dict(status="Stale", idx=idx, row=t["_row"],
                                values={c: t[c] for c in DATA_COLS}, overwrite=False, apply={},
                                detail="script row replaced by the current pairing; removed on --apply"))
            continue
        recent = any(not pd.isna(t[c]) and t[c] >= SYNC_START for c in DATE_COLS)
        if recent and str(t["Account"]) in accts:
            actions.append(dict(status="Tracker only", idx=idx, row=t["_row"],
                                values={c: t[c] for c in DATA_COLS}, overwrite=False, apply={},
                                detail="no matching Schwab fill; check for typos or a missing trade"))
    return actions


# --------------------------------------------------------------- outputs
MAX_STALE = 75
POSITIONS = CACHE.with_name("positions_cache.csv")


def holdings_check(pos, names):
    """Tracker's open positions vs what Schwab says you hold right now."""
    if not POSITIONS.exists() or pos is None or pos.empty:
        return None
    sch = pd.read_csv(POSITIONS, dtype={"Account": str})
    if sch.empty:
        return None
    sch["Account"] = sch["Account"].str[-4:].map(lambda a: names.get(a, f"...{a}"))
    sch["SYMBOL"] = [norm_sym(fmt_symbol(s, a)) for s, a in zip(sch["Symbol"], sch["AssetType"])]
    sch["Schwab"] = sch["Long"].fillna(0) - sch["Short"].fillna(0)
    sch = sch.groupby(["Account", "SYMBOL"], as_index=False)["Schwab"].sum()

    op = pos[pos["Open/Closed"] == "Open"].copy()
    op["SYMBOL"] = op["SYMBOL"].map(norm_sym)
    op["Tracker"] = [(_num(b) or 0) - (_num(s) or 0)
                     for b, s in zip(op["Buy Quantity"], op["Sell Quantity"])]
    trk = op.groupby(["Account", "SYMBOL"], as_index=False)["Tracker"].sum()

    m = trk.merge(sch, on=["Account", "SYMBOL"], how="outer").fillna(0)
    m["Difference"] = m["Tracker"] - m["Schwab"]
    m["Status"] = m["Difference"].map(lambda d: "OK" if abs(d) < 1e-6 else
                                      ("tracker has MORE than Schwab" if d > 0 else
                                       "tracker has LESS (missing 2025 lots?)"))
    return m.sort_values(["Status", "Account", "SYMBOL"])[
        ["Status", "Account", "SYMBOL", "Tracker", "Schwab", "Difference"]]


def write_review(actions, path, holdings=None, carry=None, splits=None):
    order = ["New", "Needs open", "Fill gaps", "Mismatch", "Converted", "Stale", "Check", "Tracker only",
             "Synced", "Match"]
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
        if holdings is not None:
            holdings.to_excel(xw, sheet_name="Holdings Check", index=False)
        if carry:
            pd.DataFrame(carry).sort_values(["Result", "Account", "SYMBOL"]).to_excel(
                xw, sheet_name="Carry-forward", index=False)
        if splits:
            pd.DataFrame(splits).to_excel(xw, sheet_name="Splits", index=False)
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


def write_tlh_sheet(wb, frames):
    """Replace the workbook's TLH sheet with the latest tax-loss view, sections
    stacked top to bottom."""
    sheets = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    if "TLH" in sheets:
        ws = wb.Worksheets("TLH")
        ws.Cells.Clear()
    else:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "TLH"

    def cell(v):
        if v is None or (isinstance(v, float) and v != v) or v is pd.NaT:
            return None
        if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
            return to_serial(v)
        if hasattr(v, "item"):
            v = v.item()
        return v

    r = 1
    titles = {"Summary": "Where you stand (taxable accounts)", "Candidates": "Harvest candidates (open lots at a loss)",
              "Wash Watch": "Wash-sale watch (bought in the last 30 days, all accounts)",
              "Realized": "Realized this year (oldest-shares-first)", "Notes": "Notes"}
    for name in ("Summary", "Candidates", "Wash Watch", "Realized", "Notes"):
        df = frames.get(name)
        ws.Cells(r, 1).Value = titles[name]
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 13
        r += 1
        if df is None or df.empty:
            ws.Cells(r, 1).Value = "(none)"
            r += 2
            continue
        cols = list(df.columns)
        data = [tuple(cols)] + [tuple(cell(v) for v in row) for row in df.itertuples(index=False)]
        rng = ws.Range(ws.Cells(r, 1), ws.Cells(r + len(data) - 1, len(cols)))
        rng.Value = tuple(data)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, len(cols))).Font.Bold = True
        for j, c in enumerate(cols, 1):
            body = ws.Range(ws.Cells(r + 1, j), ws.Cells(r + len(data) - 1, j))
            if df[c].dtype.kind == "M" or c in ("Opened", "Closed", "Bought", "Turns long-term") \
                    or c.startswith("Selling this ticker"):
                body.NumberFormat = "mm/dd/yyyy"
            elif df[c].dtype.kind == "f" and c not in ("Qty",):
                body.NumberFormat = "$#,##0.00;[Red]($#,##0.00)"
        r += len(data) + 1
    ws.Columns.AutoFit()


def apply_with_excel(src, dst, actions, trk, tlh_frames=None):
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

        stale = sorted({int(a["row"]) for a in actions if a["status"] == "Stale"}, reverse=True)
        if stale and len(stale) <= MAX_STALE:
            header = lo.HeaderRowRange.Row
            for r in stale:
                lo.ListRows(r - header).Delete()
            changed += len(stale)
        elif stale:
            print(f"Not removing {len(stale)} stale rows (more than {MAX_STALE}); "
                  "something changed broadly, so check the review first.")

        set_pct_change(lo, names)
        sort_table(ws, lo)
        if tlh_frames:
            try:
                write_tlh_sheet(wb, tlh_frames)
            except Exception as e:  # never let the TLH view block the update
                print(f"(TLH sheet skipped: {e})")
        wb.RefreshAll()
        xl.CalculateUntilAsyncQueriesDone()
        wb.SaveAs(str(dst.resolve()), FileFormat=51)  # 51 = .xlsx
        return changed
    finally:
        # by here the new version is saved (or an error is already raised);
        # never let closing Excel hide either one
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            xl.Quit()
        except Exception:
            pass


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
    converted = convert_old_shorts(trk)
    splits = build_positions(fills, names).attrs.get("splits", [])
    seeds, pending = collect_seeds(trk, splits)
    seeds, pending, carry = merge_seeds(carry_lots(names), seeds, trk, pending)
    trk["_sid_pending"] = False
    for idx, sid in pending.items():
        trk.at[idx, SYNC_COL] = sid
        trk.at[idx, "_sid_pending"] = True
    pos = build_positions(fills, names, seeds)
    actions = converted + reconcile(pos, trk)

    holdings = holdings_check(pos, names)
    tlh_frames = None
    try:
        import tlh
        tlh_frames = tlh.build(fills, names, seeds)
        tlh.write(tlh_frames, HERE / "TLH_Report.xlsx")
        s0 = tlh_frames["Summary"].iloc[0]
        print(f"  TLH: realized {s0['Realized net']:+,.2f} this year; "
              f"{len(tlh_frames['Candidates'])} lots at a loss "
              f"({s0['Unrealized losses (short-term)'] + s0['Unrealized losses (long-term)']:,.2f}) "
              "-> TLH_Report.xlsx")
    except PermissionError:
        print("  (TLH_Report.xlsx is open in Excel; close it to refresh)")
    except Exception as e:
        print(f"  (TLH view skipped: {e})")
    review = HERE / f"Review_{dst.stem.replace(PREFIX, '')}.xlsx"
    summary = write_review(actions, review, holdings, carry, pos.attrs.get("splits"))
    print("\n".join(f"  {s:<13}{n}" for s, n in summary.itertuples(index=False)))
    if carry:
        c = pd.Series([r["Result"] for r in carry]).value_counts()
        print("  Carry-forward (lots held on Jan 1): "
              + "; ".join(f"{n} {r}" for r, n in c.items()))
    elif pending:
        print(f"  ({len(pending)} of your 2025 lots picked up as starting inventory)")
    for sp in pos.attrs.get("splits", []):
        if sp["Ratio"] == 0:
            print(f"  Removed for $0 (nothing to record): {sp['Account']} {sp['SYMBOL']} "
                  f"{sp['Old shares']:g} on {sp['Date']:%m/%d}")
        elif abs(sp["Ratio"] - 1) > 1e-9:
            print(f"  Split: {sp['Account']} {sp['SYMBOL']} {sp['Old shares']:g} -> {sp['New shares']:g} "
                  f"on {sp['Date']:%m/%d}")
    if holdings is not None:
        off = holdings[holdings["Status"] != "OK"]
        print(f"  Holdings check: {len(holdings) - len(off)} match Schwab, {len(off)} differ"
              + (" (see the Holdings Check sheet)" if len(off) else ""))
    print(f"\nReview file: {review.name}")

    if "--apply" in args:
        todo = [a for a in actions if a["status"] not in ("Mismatch", "Tracker only", "Check", "Match")
                or (a["status"] == "Match" and set(a.get("apply", {})) - {SYNC_COL})]
        if not todo and "--force" not in args:
            print(f"No new trades or changes. No new version saved; {src.name} is still current.")
            return
        n = apply_with_excel(src, dst, actions, trk, tlh_frames)
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
