"""
Daily price history for the scanner.

DATA_SOURCE=schwab (default) uses the Schwab Trader API through schwab-py,
with the same .env and schwab_token.json as the trade feed.
DATA_SOURCE=yahoo uses yfinance (no login, good for cross-checking).

Read-only. This module only requests price history. It never touches
accounts or orders.
"""
import datetime as dt
import os
import time
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

HERE = Path(__file__).parent
TOKEN_PATH = os.getenv("SCHWAB_TOKEN_PATH", str(HERE / "schwab_token.json"))

# ~2 years of daily bars: 200 SMA plus a 6-month band width lookback needs
# roughly 330 trading days, and the history filter needs 1 full year.
HISTORY_DAYS = 2 * 365 + 10

# Schwab allows 120 requests/minute. Stay comfortably under it.
MIN_SECONDS_BETWEEN_CALLS = 60 / 100

COLUMNS = ["open", "high", "low", "close", "volume"]


def drop_incomplete_bar(df):
    """Drop today's bar while the session is still open (before 4:15 PM ET).

    A partial bar understates volume and gives an intraday close, which skews
    the volume dry-up and breakout checks. Mid-session scans therefore use
    the last completed day.
    """
    if df.empty:
        return df
    now = pd.Timestamp.now(tz="America/New_York")
    if df.index[-1].date() == now.date() and now.time() < dt.time(16, 15):
        return df.iloc[:-1]
    return df


class DataError(Exception):
    """A problem that should stop the whole scan (login expired, no token)."""


def source_name():
    return os.getenv("DATA_SOURCE", "schwab").strip().lower()


class SchwabSource:
    name = "schwab"

    def __init__(self):
        from schwab import auth
        if not os.path.exists(TOKEN_PATH):
            raise DataError("No Schwab token found. Run: py auth_setup.py")
        try:
            self.client = auth.client_from_token_file(
                TOKEN_PATH, os.environ["SCHWAB_APP_KEY"], os.environ["SCHWAB_APP_SECRET"])
        except KeyError as e:
            raise DataError(f"Missing {e.args[0]} in .env")
        except Exception as e:
            raise DataError(f"Could not load the Schwab token ({e}). Run: py auth_setup.py")
        self._last_call = 0.0

    def _pace(self):
        wait = MIN_SECONDS_BETWEEN_CALLS - (time.monotonic() - self._last_call)
        if wait > 0:
            time.sleep(wait)
        self._last_call = time.monotonic()

    def history(self, symbol):
        end = dt.datetime.now(dt.timezone.utc)
        start = end - dt.timedelta(days=HISTORY_DAYS)
        for attempt in range(3):
            self._pace()
            try:
                # Schwab writes class shares as BRK/B, not BRK.B
                r = self.client.get_price_history_every_day(
                    symbol.replace(".", "/"), start_datetime=start, end_datetime=end,
                    need_extended_hours_data=False)
            except Exception as e:
                msg = str(e).lower()
                if "refresh" in msg or "token" in msg or "invalid_grant" in msg:
                    raise DataError("Schwab login expired. Run: py auth_setup.py")
                raise
            if r.status_code == 401:
                raise DataError("Schwab login expired. Run: py auth_setup.py")
            if r.status_code == 429:  # rate limited: back off and retry
                time.sleep(15 * (attempt + 1))
                continue
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:120]}")
            candles = (r.json() or {}).get("candles") or []
            if not candles:
                return pd.DataFrame(columns=COLUMNS)
            df = pd.DataFrame(candles)
            idx = (pd.to_datetime(df["datetime"], unit="ms", utc=True)
                   .dt.tz_convert("America/New_York").dt.normalize().dt.tz_localize(None))
            df.index = idx
            df.index.name = "date"
            return drop_incomplete_bar(df[COLUMNS].astype(float).sort_index())
        raise RuntimeError("rate limited repeatedly")


class YahooSource:
    name = "yahoo"

    def __init__(self):
        import yfinance  # noqa: F401  (fail early if not installed)

    def history(self, symbol):
        import yfinance as yf
        # auto_adjust=False keeps split-adjusted but not dividend-adjusted
        # prices, which is what Schwab and thinkorswim show.
        df = yf.Ticker(symbol.replace(".", "-")).history(
            period="2y", interval="1d", auto_adjust=False, actions=False)
        if df is None or df.empty:
            return pd.DataFrame(columns=COLUMNS)
        df = df.rename(columns=str.lower)
        df.index = pd.to_datetime(df.index).tz_localize(None).normalize()
        df.index.name = "date"
        return drop_incomplete_bar(df[COLUMNS].astype(float).sort_index())


def get_source(name=None):
    name = (name or source_name()).lower()
    if name == "schwab":
        return SchwabSource()
    if name in ("yahoo", "yfinance"):
        return YahooSource()
    raise DataError(f"Unknown DATA_SOURCE '{name}'. Use schwab or yahoo.")
