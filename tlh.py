"""
Tax-loss harvesting view, built from the same data as the tracker.

Uses your Schwab lot setting (oldest shares sold first) rather than the
journal's swing pairing, so realized numbers line up with the 1099.

  Summary     realized YTD (short/long term) and unrealized gains/losses,
              household total and per account
  Candidates  open lots sitting at a loss, with wash-sale status
  Realized    every taxable close this year, with possible wash sales
  Wash Watch  purchases in the last 30 days, all accounts (incl. the Roth)

Planning tool only, not tax advice. Schwab's Realized/Unrealized Gain/Loss
pages and your 1099 are the record.
"""
import datetime as dt
import os

import pandas as pd

import update_tracker as U

NON_TAXABLE = {a.strip() for a in os.getenv("NON_TAXABLE_ACCOUNTS", "Roth").split(",") if a.strip()}
WASH_DAYS = 30


def _n(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if f != f else f


def _mult(row):
    return 100 if str(row.get("Stock Or Option", "")).strip().lower() == "option" else 1


def _long_term(opened, closed, short, is_option):
    if short:
        return False  # gains/losses on options you wrote are always short-term
    if pd.isna(opened) or pd.isna(closed):
        return False
    return pd.Timestamp(closed) > pd.Timestamp(opened) + pd.DateOffset(years=1)


def _prices(names):
    """Current price per share/contract-share from the last holdings snapshot."""
    if not U.POSITIONS.exists():
        return {}, None
    p = pd.read_csv(U.POSITIONS, dtype={"Account": str})
    if "MarketValue" not in p.columns or p.empty:
        return {}, None
    p["Account"] = p["Account"].str[-4:].map(lambda a: names.get(a, f"...{a}"))
    out = {}
    for r in p.itertuples(index=False):
        qty = (r.Long or 0) - (r.Short or 0)
        mv = _n(r.MarketValue)
        if not qty or mv is None:
            continue
        mult = 100 if r.AssetType == "OPTION" else 1
        sym = U.norm_sym(U.fmt_symbol(r.Symbol, r.AssetType))
        out[(r.Account, sym)] = abs(mv / (qty * mult))
        out.setdefault(("*", sym), abs(mv / (qty * mult)))
    return out, str(p["AsOf"].iloc[0]) if "AsOf" in p.columns else None


def _buys(fills, names):
    """Every purchase (shares or calls), all accounts, for the wash-sale checks."""
    f = fills.copy()
    f["Account"] = f["Account"].astype(str).str[-4:].map(lambda a: names.get(a, f"...{a}"))
    f["DateTime"] = pd.to_datetime(f["DateTime"])
    f["Date"] = f["DateTime"].dt.normalize()
    f["SYM"] = [U.norm_sym(U.fmt_symbol(s, a)) for s, a in zip(f.Symbol, f.AssetType)]
    f["UND"] = f["SYM"].map(U.underlying)
    is_call = f["SYM"].str.endswith(" C")
    keep = (f.Side == "BUY") & (f.PositionEffect == "OPENING") & ((f.AssetType == "EQUITY") | is_call)
    return f[keep]


def build(fills, names, seeds, today=None):
    today = pd.Timestamp(today or dt.date.today()).normalize()
    year = today.year
    pos = U.build_positions(fills, names, seeds, fifo=True)
    prices, asof = _prices(names)
    buys = _buys(fills, names)

    # buys that were sold in the same sale don't wash each other
    sold_together = {}
    for _, r in pos.iterrows():
        for cid in (r.get("_close_ids") or ()):
            if cid:
                sold_together.setdefault(cid, set()).update(r.get("_open_ids") or ())

    def wash_hits(sym, is_option, around, exclude_ids, before_only=False):
        """Replacement purchases within 30 days. Shares: the same shares or calls
        on them. An option: the same contract."""
        lo = around - pd.Timedelta(days=WASH_DAYS)
        hi = around if before_only else around + pd.Timedelta(days=WASH_DAYS)
        same = (buys.SYM == sym) if is_option else (buys.UND == U.underlying(sym))
        return buys[same & (buys.Date >= lo) & (buys.Date <= hi) & (~buys.FillID.isin(exclude_ids))]

    realized, cands = [], []
    for _, r in pos.iterrows():
        acct = str(r["Account"])
        if acct in NON_TAXABLE:
            continue
        m = _mult(r)
        short = bool(r.get("_short", False))  # sold to open
        is_opt = m == 100
        und = U.underlying(r["SYMBOL"])
        bq, bp, bf = _n(r["Buy Quantity"]) or 0, _n(r["Purchase PRICE PER SHARE"]), _n(r["Fees & Comm"]) or 0
        sq, sp, sf = _n(r["Sell Quantity"]) or 0, _n(r["Sell Price Per Share"]), _n(r["Sale Fees & Comm"]) or 0
        opened, closed = r["Opened Date"], r["Closed Date"]
        own = set(r.get("_open_ids") or ())

        if r["Open/Closed"] == "Closed":
            if pd.isna(closed) or pd.Timestamp(closed).year != year:
                continue
            missing = r.get("_needs_open", False)
            proceeds = sq * (sp or 0) * m - sf
            cost = bq * (bp or 0) * m + bf
            gain = None if missing else round(proceeds - cost, 2)
            lt = _long_term(opened, closed, short, is_opt)
            note = ""
            if gain is not None and gain < 0:
                excl = set(own)
                for cid in (r.get("_close_ids") or ()):
                    excl |= sold_together.get(cid, set())
                hits = wash_hits(U.norm_sym(r["SYMBOL"]), is_opt, pd.Timestamp(closed), excl)
                if len(hits):
                    h = hits.sort_values("Date").iloc[0]
                    note = (f"possible wash sale: {h.Account} bought {U.fmt_symbol(h.Symbol, h.AssetType)} "
                            f"{h.Date:%m/%d}")
            realized.append({"Account": acct, "Symbol": r["SYMBOL"], "Opened": opened, "Closed": closed,
                             "Qty": max(bq, sq), "Proceeds": round(proceeds, 2),
                             "Cost": None if missing else round(cost, 2), "Gain/Loss": gain,
                             "Term": "Long" if lt else "Short",
                             "Note": "cost not entered (opened before 2026)" if missing else note})
            continue

        # open lot
        px = prices.get((acct, U.norm_sym(r["SYMBOL"]))) or prices.get(("*", U.norm_sym(r["SYMBOL"])))
        if short:
            qty, basis = sq, sq * (sp or 0) * m - sf          # premium received
            value = None if px is None else qty * px * m      # cost to buy back
            unreal = None if value is None else round(basis - value, 2)
        else:
            qty, basis = bq, bq * (bp or 0) * m + bf
            value = None if px is None else qty * px * m
            unreal = None if value is None else round(value - basis, 2)
        lt_date = None if (short or pd.isna(opened)) else pd.Timestamp(opened) + pd.DateOffset(years=1) + pd.Timedelta(days=1)
        lt = bool(lt_date is not None and today >= lt_date)
        cands.append({"Account": acct, "Symbol": r["SYMBOL"], "Opened": opened, "Qty": qty,
                      "Side": "Sold (short)" if short else "Owned",
                      "Basis": round(basis, 2), "Price now": px, "Value now": None if value is None else round(value, 2),
                      "Unrealized": unreal, "Term": "Long" if lt else "Short",
                      "Turns long-term": None if (lt or lt_date is None) else lt_date,
                      "_own": own, "_und": und, "_sym": U.norm_sym(r["SYMBOL"]), "_opt": is_opt})

    # wash-sale status for each open lot if sold today
    for c in cands:
        hits = wash_hits(c["_sym"], c["_opt"], today, c["_own"], before_only=True)
        if len(hits):
            h = hits.sort_values("Date").iloc[-1]
            c["Wash-sale check"] = (f"BLOCKED until {h.Date + pd.Timedelta(days=WASH_DAYS + 1):%m/%d}: "
                                    f"{h.Account} bought {U.fmt_symbol(h.Symbol, h.AssetType)} {h.Date:%m/%d}")
        else:
            c["Wash-sale check"] = "clear (don't buy back in ANY account, Roth included, for 31 days)"
    cand = pd.DataFrame(cands)
    real = pd.DataFrame(realized)

    # summary
    def sums(df, col, term=None, sign=None):
        if df.empty:
            return 0.0
        d = df.dropna(subset=[col])
        if term:
            d = d[d["Term"] == term]
        if sign == "neg":
            d = d[d[col] < 0]
        elif sign == "pos":
            d = d[d[col] > 0]
        return round(float(d[col].sum()), 2)

    rows = []
    accts = sorted(set(real.get("Account", pd.Series(dtype=str))) | set(cand.get("Account", pd.Series(dtype=str))))
    for a in ["All taxable accounts"] + accts:
        R = real if a == "All taxable accounts" else real[real["Account"] == a] if not real.empty else real
        C = cand if a == "All taxable accounts" else cand[cand["Account"] == a] if not cand.empty else cand
        Cl = C[~C["Side"].eq("Sold (short)")] if not C.empty else C
        rows.append({
            "Account": a,
            "Realized short-term": sums(R, "Gain/Loss", "Short"),
            "Realized long-term": sums(R, "Gain/Loss", "Long"),
            "Realized net": sums(R, "Gain/Loss"),
            "Unrealized losses (short-term)": sums(Cl, "Unrealized", "Short", "neg"),
            "Unrealized losses (long-term)": sums(Cl, "Unrealized", "Long", "neg"),
            "Unrealized gains": sums(Cl, "Unrealized", None, "pos"),
            "Sales missing cost": int(R["Gain/Loss"].isna().sum()) if not R.empty else 0,
        })
    summary = pd.DataFrame(rows)

    if not cand.empty:
        cand = cand.drop(columns=["_own", "_und", "_sym", "_opt"], errors="ignore")
        losers = cand[cand["Unrealized"].fillna(0) < 0].sort_values("Unrealized")
    else:
        losers = cand
    watch = buys[buys.Date >= today - pd.Timedelta(days=WASH_DAYS)].sort_values("Date", ascending=False)
    watch = pd.DataFrame({"Bought": watch.Date, "Account": watch.Account,
                          "Symbol": [U.fmt_symbol(s, a) for s, a in zip(watch.Symbol, watch.AssetType)],
                          "Qty": watch.Qty, "Price": watch.Price,
                          "Selling this ticker at a loss is a wash sale until":
                              watch.Date + pd.Timedelta(days=WASH_DAYS + 1)})
    notes = pd.DataFrame({"Notes": [
        f"Prices as of {asof}" if asof else "No current prices yet (run a sync).",
        f"Taxable accounts: everything except {', '.join(sorted(NON_TAXABLE))}; they net together as one household.",
        "Realized uses oldest-shares-first (your Schwab setting), not the journal's swing pairing.",
        "Wash sales: buying the same shares (or calls on them), or the same option contract, in ANY account within "
        "30 days before or after a loss sale defers the loss; a buy in the Roth disallows it permanently. "
        "Schwab only checks within one account.",
        "Net capital losses beyond gains offset up to $3,000 of ordinary income a year; the rest carries forward.",
        "Planning tool only, not tax advice. Your 1099 and Schwab's Gain/Loss pages are the record.",
    ]})
    return {"Summary": summary, "Candidates": losers, "Realized": real.sort_values("Closed") if not real.empty else real,
            "Wash Watch": watch, "Notes": notes}


def write(frames, path):
    with pd.ExcelWriter(path, engine="openpyxl", datetime_format="mm/dd/yyyy") as xw:
        for name, df in frames.items():
            (df if len(df) else pd.DataFrame({"": ["(none)"]})).to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    max(len(str(c.value or "")) for c in col[:300]) + 2, 70)
