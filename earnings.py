"""
Next earnings date per stock, for the earnings warning.

Source: Yahoo Finance via yfinance (free; Schwab's API does not reliably
provide earnings dates). Only the shortlist (setups and breakouts) is looked
up, and results are cached in data/earnings_cache.json for a few days, so a
normal evening adds a few dozen quick requests at most.

When Yahoo shows a date range (an unconfirmed estimate), the first date of
the range is used and the date is marked as an estimate.

Warning levels, counted in calendar days from the scan date
  WARN_DAYS (21)   earnings inside a typical swing hold: flagged
  DANGER_DAYS (7)  earnings within a week: flagged strongly

Any failure (Yahoo down, unknown symbol) leaves the date blank and never
stops the scan.
"""
import datetime as dt
import json
import time
from pathlib import Path

HERE = Path(__file__).parent
CACHE = HERE / "data" / "earnings_cache.json"
CACHE_DAYS = 3          # re-check a symbol's date after this many days
WARN_DAYS = 21
DANGER_DAYS = 7
PAUSE = 0.25            # seconds between Yahoo requests


def _load_cache():
    try:
        return json.loads(CACHE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_cache(cache):
    try:
        CACHE.parent.mkdir(exist_ok=True)
        CACHE.write_text(json.dumps(cache, indent=0), encoding="utf-8")
    except OSError:
        pass


def _fetch(symbol):
    """Returns (list of ISO dates, ok). Empty list = Yahoo has no date."""
    import yfinance as yf
    cal = yf.Ticker(symbol.replace(".", "-")).calendar or {}
    dates = cal.get("Earnings Date") or []
    return [d.isoformat() for d in dates if isinstance(d, dt.date)]


def next_earnings(symbols, as_of: dt.date):
    """{symbol: {"date": date or None, "estimate": bool, "days": int or None, "level": str}}"""
    cache = _load_cache()
    today = dt.date.today().isoformat()
    out, failed = {}, 0
    for sym in symbols:
        c = cache.get(sym)
        fresh = c and (dt.date.fromisoformat(today) - dt.date.fromisoformat(c["checked"])).days < CACHE_DAYS
        if not fresh:
            try:
                dates = _fetch(sym)
                c = {"dates": dates, "checked": today}
                cache[sym] = c
            except Exception:
                failed += 1
                c = c or {"dates": [], "checked": "1970-01-01"}
            time.sleep(PAUSE)
        future = sorted(d for d in c["dates"] if dt.date.fromisoformat(d) >= as_of)
        if future:
            d = dt.date.fromisoformat(future[0])
            days = (d - as_of).days
            level = "danger" if days <= DANGER_DAYS else "warn" if days <= WARN_DAYS else ""
            out[sym] = {"date": d, "estimate": len(c["dates"]) > 1, "days": days, "level": level}
        else:
            out[sym] = {"date": None, "estimate": False, "days": None, "level": ""}
    _save_cache(cache)
    return out, failed


def label(info):
    """Short text for terminal output, e.g. '10/21 (8d)!' or '11/04~ (24d)'."""
    if not info or not info.get("date"):
        return ""
    d = info["date"]
    mark = "!" if info["level"] == "danger" else "*" if info["level"] == "warn" else ""
    est = "~" if info["estimate"] else ""
    return f"{d.month}/{d.day}{est} ({info['days']}d){mark}"
