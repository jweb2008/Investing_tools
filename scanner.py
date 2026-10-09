"""
Breakout setup scoring. Daily bars only.

Finds stocks coiling just under resistance in an uptrend. Each stock gets a
0 to 100 score; setups at MIN_SCORE or higher are reported. This builds a
shortlist for chart review. It is not a buy signal.

Score components
  uptrend      25  close > 50 SMA > 200 SMA
  near_pivot   25  close within 5% below the pivot (prior 30-day high)
  bb_squeeze   20  Bollinger band width in the bottom 25% of the last 6 months
  atr_contract 15  10-day ATR% under 80% of the prior 40 days
  vol_dryup    15  10-day average volume below the 50-day average

Separate flag
  breakout     close above the pivot on 1.5x+ the 50-day average volume

Filters (stock is skipped, not scored)
  price $5+, 50-day average volume 500K+, 1 year of history
"""
import numpy as np
import pandas as pd

MIN_SCORE = 60

WEIGHTS = {"uptrend": 25, "near_pivot": 25, "bb_squeeze": 20,
           "atr_contract": 15, "vol_dryup": 15}

PIVOT_DAYS = 30
NEAR_PIVOT_PCT = 5.0
BB_LEN, BB_STD = 20, 2.0
BB_LOOKBACK = 126          # ~6 months of trading days
BB_PCTILE = 25.0
ATR_RECENT, ATR_PRIOR, ATR_RATIO = 10, 40, 0.80
VOL_SHORT, VOL_LONG = 10, 50
BREAKOUT_VOL_MULT = 1.5

MIN_PRICE = 5.0
MIN_AVG_VOLUME = 500_000
MIN_BARS = 252             # 1 year of trading days


def indicators(df: pd.DataFrame) -> dict:
    """Raw indicator values for the latest bar. df: daily open/high/low/close/volume."""
    c, h, l, v = df["close"], df["high"], df["low"], df["volume"]
    last = df.index[-1]

    sma50 = c.rolling(50).mean()
    sma200 = c.rolling(200).mean()

    # pivot = highest high of the 30 sessions BEFORE today, so a close above it
    # today counts as a breakout
    pivot = h.iloc[-(PIVOT_DAYS + 1):-1].max()
    base_low = l.iloc[-(PIVOT_DAYS + 1):-1].min()

    mid = c.rolling(BB_LEN).mean()
    sd = c.rolling(BB_LEN).std(ddof=0)
    width = (2 * BB_STD * sd) / mid  # (upper - lower) / middle
    recent_w = width.iloc[-BB_LOOKBACK:]
    bb_pctile = float((recent_w <= width.iloc[-1]).mean() * 100)

    prev_c = c.shift(1)
    tr = pd.concat([h - l, (h - prev_c).abs(), (l - prev_c).abs()], axis=1).max(axis=1)
    tr_pct = tr / c * 100
    atr_recent = tr_pct.iloc[-ATR_RECENT:].mean()
    atr_prior = tr_pct.iloc[-(ATR_RECENT + ATR_PRIOR):-ATR_RECENT].mean()
    atr14 = tr.iloc[-14:].mean()

    vol_short = v.iloc[-VOL_SHORT:].mean()
    vol_long = v.iloc[-VOL_LONG:].mean()
    vol_long_prior = v.iloc[-(VOL_LONG + 1):-1].mean()  # excludes today, for breakout volume

    close = float(c.iloc[-1])
    return {
        "date": last.date() if hasattr(last, "date") else last,
        "close": close,
        "sma50": float(sma50.iloc[-1]),
        "sma200": float(sma200.iloc[-1]),
        "pivot": float(pivot),
        "base_low": float(base_low),
        "pct_below_pivot": float((pivot - close) / pivot * 100),
        "bb_width": float(width.iloc[-1]),
        "bb_width_pctile": bb_pctile,
        "atr_pct_10": float(atr_recent),
        "atr_pct_prior40": float(atr_prior),
        "atr_ratio": float(atr_recent / atr_prior) if atr_prior else np.nan,
        "atr14": float(atr14),
        "vol_10": float(vol_short),
        "vol_50": float(vol_long),
        "today_vol_x": float(v.iloc[-1] / vol_long_prior) if vol_long_prior else np.nan,
    }


def check_filters(df: pd.DataFrame):
    """Returns None if the stock passes, otherwise the reason it was skipped."""
    if df is None or len(df) < MIN_BARS:
        return f"under 1 year of history ({0 if df is None else len(df)} bars)"
    if df["close"].iloc[-1] < MIN_PRICE:
        return f"price under ${MIN_PRICE:.0f}"
    if df["volume"].iloc[-VOL_LONG:].mean() < MIN_AVG_VOLUME:
        return "average volume under 500K"
    return None


def score(df: pd.DataFrame) -> dict:
    """Score one stock. Returns indicators, component points, total, and flags."""
    ind = indicators(df)
    pts = {
        "uptrend": ind["close"] > ind["sma50"] > ind["sma200"],
        "near_pivot": 0 <= ind["pct_below_pivot"] <= NEAR_PIVOT_PCT,
        "bb_squeeze": ind["bb_width_pctile"] <= BB_PCTILE,
        "atr_contract": ind["atr_ratio"] < ATR_RATIO,
        "vol_dryup": ind["vol_10"] < ind["vol_50"],
    }
    pts = {k: WEIGHTS[k] if ok else 0 for k, ok in pts.items()}
    breakout = ind["close"] > ind["pivot"] and ind["today_vol_x"] >= BREAKOUT_VOL_MULT
    return {**ind, **{f"pts_{k}": p for k, p in pts.items()},
            "score": sum(pts.values()), "breakout": bool(breakout)}
