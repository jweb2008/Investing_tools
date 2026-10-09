"""
Trade feed: pulls your filled trades from Schwab into trade_log.xlsx.

  python sync_trades.py

Read-only. It only requests transaction history; it has no ability to place orders.

Each run:
  1. Pulls TRADE and RECEIVE_AND_DELIVER (expirations/assignments) transactions
     from SYNC_START (default Jan 1 2026) to now, in 30-day windows.
  2. Merges them into fills_cache.csv, keyed on Schwab's activity ID, so a fill
     is never duplicated and older fills are kept even after Schwab stops
     serving them.
  3. Rebuilds trade_log.xlsx from the cache:
       Fills           one row per execution
       Closed Trades   entries matched to exits (FIFO) with P&L and hold time
       Open Positions  lots still open
       Unmatched       closes with no matching entry (opened before SYNC_START)
       Sync Info       when it ran and what came back
"""
import datetime as dt
import os
import sys
import tempfile
from collections import defaultdict, deque
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

HERE = Path(__file__).parent
TOKEN_PATH = os.getenv("SCHWAB_TOKEN_PATH", str(HERE / "schwab_token.json"))
OUTPUT = Path(os.getenv("TRADE_LOG_PATH", str(HERE / "trade_log.xlsx")))
CACHE = OUTPUT.with_name("fills_cache.csv")
SYNC_START = dt.date.fromisoformat(os.getenv("SYNC_START", "2026-01-01"))
WINDOW_DAYS = 30

FILL_COLS = ["FillID", "Account", "DateTime", "Date", "Symbol", "Underlying",
             "AssetType", "Side", "PositionEffect", "Qty", "Price", "Fees",
             "NetAmount", "Source", "OrderID"]


# ---------------------------------------------------------------- Schwab pull
def get_client():
    from schwab import auth
    if not os.path.exists(TOKEN_PATH):
        sys.exit("No Schwab token found. Run: python auth_setup.py")
    return auth.client_from_token_file(
        TOKEN_PATH, os.environ["SCHWAB_APP_KEY"], os.environ["SCHWAB_APP_SECRET"])


def pull_transactions(client):
    """Returns (list of (last4, txn dict), list of status notes)."""
    r = client.get_account_numbers()
    if r.status_code == 401:
        sys.exit("Schwab login expired. Run: python auth_setup.py")
    r.raise_for_status()
    accounts = r.json()

    types = [client.Transactions.TransactionType.TRADE,
             client.Transactions.TransactionType.RECEIVE_AND_DELIVER]
    now = dt.datetime.now(dt.timezone.utc)
    start = dt.datetime.combine(SYNC_START, dt.time(), dt.timezone.utc)

    out, notes = [], []
    for acct in accounts:
        last4 = acct["accountNumber"][-4:]
        w_start = start
        while w_start < now:
            w_end = min(w_start + dt.timedelta(days=WINDOW_DAYS), now)
            resp = client.get_transactions(acct["hashValue"], start_date=w_start,
                                           end_date=w_end, transaction_types=types)
            label = f"...{last4} {w_start:%Y-%m-%d} to {w_end:%Y-%m-%d}"
            if resp.status_code == 401:
                sys.exit("Schwab login expired. Run: python auth_setup.py")
            if resp.status_code != 200:
                notes.append(f"FAILED {label}: HTTP {resp.status_code} {resp.text[:150]}")
            else:
                txns = resp.json() or []
                out += [(last4, t) for t in txns]
                notes.append(f"ok     {label}: {len(txns)} transactions")
            w_start = w_end
    return out, notes


# -------------------------------------------------------------------- parsing
def parse(last4, t):
    """One Schwab transaction -> list of fill rows (one per traded instrument)."""
    if t.get("status") not in (None, "VALID"):
        return []
    items = t.get("transferItems", [])
    fee_total = -sum(i.get("cost", 0) for i in items
                     if i.get("feeType") or i.get("instrument", {}).get("assetType") == "CURRENCY")
    legs = [i for i in items if i.get("instrument", {}).get("assetType") != "CURRENCY"
            and not i.get("feeType") and i.get("amount")]
    gross = sum(abs(i.get("cost", 0)) for i in legs) or 1
    when = pd.to_datetime(t.get("time") or t.get("tradeDate"), utc=True)
    when = when.tz_convert("America/New_York").tz_localize(None)

    rows = []
    for n, leg in enumerate(legs):
        inst = leg["instrument"]
        qty = float(leg["amount"])
        fees = round(fee_total * abs(leg.get("cost", 0)) / gross, 4) if len(legs) > 1 else round(fee_total, 4)
        rows.append({
            "FillID": f"{t['activityId']}-{n}",
            "Account": last4,
            "DateTime": when,
            "Date": when.date(),
            "Symbol": inst.get("symbol", ""),
            "Underlying": inst.get("underlyingSymbol", inst.get("symbol", "")),
            "AssetType": inst.get("assetType", ""),
            "Side": "BUY" if qty > 0 else "SELL",
            "PositionEffect": leg.get("positionEffect", ""),
            "Qty": abs(qty),
            "Price": float(leg.get("price", 0) or 0),
            "Fees": fees,
            # signed cash for this leg: negative = money out
            "NetAmount": round(float(leg.get("cost", 0)) - fees, 2),
            "Source": t.get("type", ""),
            "OrderID": str(t.get("orderId", "")),
        })
    return rows


# ------------------------------------------------------------- FIFO matching
def match_trades(fills: pd.DataFrame):
    lots = defaultdict(deque)  # (acct, symbol) -> deque of open lots
    closed, unmatched = [], []
    for f in fills.sort_values(["DateTime", "FillID"]).itertuples(index=False):
        key = (f.Account, f.Symbol)
        signed = f.Qty if f.Side == "BUY" else -f.Qty
        cash_pu = f.NetAmount / f.Qty if f.Qty else 0  # cash per unit incl. fees
        remaining = f.Qty
        book = lots[key]
        # close against opposite-direction lots first
        while remaining > 1e-9 and book and (book[0]["signed"] > 0) != (signed > 0):
            lot = book[0]
            n = min(remaining, lot["qty"])
            long_ = lot["signed"] > 0
            pnl = (lot["cash_pu"] + cash_pu) * n
            basis = abs(lot["cash_pu"]) * n
            closed.append({
                "Account": f.Account, "Symbol": f.Symbol, "Underlying": f.Underlying,
                "AssetType": f.AssetType, "Direction": "Long" if long_ else "Short",
                "Qty": n, "EntryDate": lot["date"], "ExitDate": f.Date,
                "EntryPrice": lot["price"], "ExitPrice": f.Price,
                "PnL": round(pnl, 2),
                "ReturnPct": round(pnl / basis * 100, 2) if basis else None,
                "HoldDays": (pd.Timestamp(f.Date) - pd.Timestamp(lot["date"])).days,
                "EntryFillID": lot["fill"], "ExitFillID": f.FillID,
            })
            lot["qty"] -= n
            remaining -= n
            if lot["qty"] <= 1e-9:
                book.popleft()
        if remaining > 1e-9:
            if f.PositionEffect == "CLOSING" or f.Source == "RECEIVE_AND_DELIVER":
                unmatched.append({**f._asdict(), "UnmatchedQty": remaining,
                                  "Reason": "close with no entry since SYNC_START"})
            else:
                book.append({"qty": remaining, "signed": signed, "cash_pu": cash_pu,
                             "price": f.Price, "date": f.Date, "fill": f.FillID})

    open_rows = [{"Account": a, "Symbol": s, "Direction": "Long" if l["signed"] > 0 else "Short",
                  "Qty": l["qty"], "EntryDate": l["date"], "EntryPrice": l["price"],
                  "CostBasis": round(abs(l["cash_pu"]) * l["qty"], 2), "EntryFillID": l["fill"]}
                 for (a, s), book in lots.items() for l in book]
    return pd.DataFrame(closed), pd.DataFrame(open_rows), pd.DataFrame(unmatched)


# --------------------------------------------------------------------- output
def load_cache():
    if CACHE.exists():
        df = pd.read_csv(CACHE, dtype={"Account": str, "OrderID": str, "FillID": str})
        df["DateTime"] = pd.to_datetime(df["DateTime"])
        df["Date"] = df["DateTime"].dt.date
        df["OrderID"] = df["OrderID"].fillna("")
        return df
    return pd.DataFrame(columns=FILL_COLS)


def write_excel(sheets: dict):
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=OUTPUT.parent)
    os.close(fd)
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for i, col in enumerate(df.columns, 1):
                width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(width, 40)
            ws.freeze_panes = "A2"
            if len(df):  # named Excel tables make Power Query pickup easy
                ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
                tbl = Table(displayName=name.replace(" ", ""), ref=ref)
                tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
                ws.add_table(tbl)
    try:
        os.replace(tmp, OUTPUT)
    except PermissionError:
        os.remove(tmp)
        sys.exit(f"{OUTPUT.name} is open in Excel. Close it and run again "
                 "(the fill cache was still updated).")


def main():
    client = get_client()
    print(f"Pulling fills since {SYNC_START}...")
    txns, notes = pull_transactions(client)
    for n in notes:
        print("  " + n)

    new = pd.DataFrame([row for last4, t in txns for row in parse(last4, t)], columns=FILL_COLS)
    cache = load_cache()
    before = len(cache)
    fills = (pd.concat([cache, new], ignore_index=True)
             .drop_duplicates("FillID", keep="last")
             .sort_values("DateTime").reset_index(drop=True))
    fills.to_csv(CACHE, index=False)
    added = len(fills) - before

    closed, open_pos, unmatched = match_trades(fills) if len(fills) else (pd.DataFrame(),) * 3
    info = pd.DataFrame({"Item": ["Last sync", "Sync start", "Total fills", "New this run",
                                  "Closed trades", "Open lots", "Unmatched closes"]
                                 + [f"Window {i + 1}" for i in range(len(notes))],
                         "Value": [dt.datetime.now().strftime("%Y-%m-%d %H:%M"), str(SYNC_START),
                                   len(fills), added, len(closed), len(open_pos), len(unmatched)]
                                  + notes})
    write_excel({"Fills": fills, "Closed Trades": closed, "Open Positions": open_pos,
                 "Unmatched": unmatched, "Sync Info": info})

    print(f"\n{len(fills)} fills total ({added} new) -> {OUTPUT}")
    print(f"{len(closed)} closed trades, {len(open_pos)} open lots, {len(unmatched)} unmatched closes")
    if any(n.startswith("FAILED") for n in notes):
        print("\nSome date windows failed (see above). Schwab may limit how far back it serves "
              "history. Paste the FAILED lines to Claude.")


if __name__ == "__main__":
    main()
