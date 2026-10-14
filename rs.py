"""
Relative strength vs SPY.

In a choppy or uncertain market, the stocks that hold up best tend to be the
first to break out when the market turns. This ranks setups by how they have
performed against SPY.

  RS 3m / RS 6m   stock return minus SPY return over 63 / 126 trading days,
                  as relative performance: (1 + stock) / (1 + SPY) - 1
  RS composite    0.6 x RS 3m + 0.4 x RS 6m (recent quarter weighted more)
  RS rating       1 to 99 percentile of the composite across every stock
                  scored in a universe scan (99 = stronger than 99% of them).
                  Only computed when at least MIN_FOR_RATING stocks were
                  scored; a small watchlist scan shows RS 3m / 6m only.

RS is used to rank and is shown on the watchlist. It is not part of the
setup score, so it does not change which stocks qualify.
"""
import numpy as np
import pandas as pd

LOOKBACKS = {"rs_3m": 63, "rs_6m": 126}
WEIGHTS = {"rs_3m": 0.6, "rs_6m": 0.4}
MIN_FOR_RATING = 200


def _ret(close, n):
    if len(close) <= n:
        return np.nan
    return float(close.iloc[-1] / close.iloc[-1 - n] - 1)


def raw(close: pd.Series, spy_close: pd.Series) -> dict:
    """RS 3m / 6m / composite for one stock, measured on the dates both have."""
    if spy_close is None or len(spy_close) == 0:
        return {k: np.nan for k in (*LOOKBACKS, "rs_composite")}
    joined = pd.concat([close.rename("s"), spy_close.rename("m")], axis=1, join="inner").dropna()
    out = {}
    for key, n in LOOKBACKS.items():
        s, m = _ret(joined["s"], n), _ret(joined["m"], n)
        out[key] = (1 + s) / (1 + m) - 1 if s == s and m == m else np.nan
    vals = [out[k] * w for k, w in WEIGHTS.items() if out[k] == out[k]]
    out["rs_composite"] = sum(vals) / sum(w for k, w in WEIGHTS.items() if out[k] == out[k]) \
        if vals else np.nan
    return out


def add_ratings(results):
    """Adds rs_rating (1-99) to each result dict when the scan is big enough."""
    comp = pd.Series([r.get("rs_composite", np.nan) for r in results], dtype=float)
    ok = comp.notna()
    if ok.sum() < MIN_FOR_RATING:
        for r in results:
            r["rs_rating"] = np.nan
        return False
    pct = comp[ok].rank(pct=True)
    ratings = (pct * 98 + 1).round().astype(int)
    for i, r in enumerate(results):
        r["rs_rating"] = int(ratings[i]) if ok[i] else np.nan
    return True
