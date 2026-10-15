"""
Demand and momentum signals, computed for every bar at once (vectorized).

Used by strategy_test.py now, and by the live scanner once a strategy passes
its holdout test. Every value at bar i uses only data through bar i.

Demand features (signs that buyers are accumulating before a breakout)
  updown_vol      50-day volume on up days / volume on down days
                  (above ~1.2 = more volume on up days)
  pocket_pivot    up day whose volume beats every down day of the prior 10
  pp_recent       a pocket pivot within the last 10 days
  rs_line_high    stock / SPY ratio at a 52-week high within the last 10 days
  rvol            today's volume / prior 50-day average

Momentum features
  ema10, sma20, sma50, ret21 (1-month return), pct_from_high (vs 52-week high),
  range_pos (where the close sits in the day's range, 0 = low, 1 = high),
  high20_prior (prior 20-day high)
"""
import numpy as np
import pandas as pd

UPDOWN_MIN = 1.2


def demand_features(df: pd.DataFrame, spy_close: pd.Series) -> pd.DataFrame:
    c, v = df["close"], df["volume"]
    up = c > c.shift(1)
    down = c < c.shift(1)
    out = pd.DataFrame(index=df.index)
    up_vol = v.where(up, 0.0).rolling(50).sum()
    down_vol = v.where(down, 0.0).rolling(50).sum()
    out["updown_vol"] = up_vol / down_vol.replace(0, np.nan)
    max_down_10 = v.where(down, 0.0).shift(1).rolling(10).max()
    out["pocket_pivot"] = up & (v > max_down_10)
    out["pp_recent"] = out["pocket_pivot"].astype(int).rolling(10).max().fillna(0).astype(bool)
    spy = spy_close.reindex(df.index).ffill()
    rs_line = c / spy
    at_high = rs_line >= rs_line.rolling(252).max()
    out["rs_line_high"] = at_high.astype(int).rolling(10).max().fillna(0).astype(bool)
    out["rvol"] = v / v.shift(1).rolling(50).mean()
    return out


def momentum_features(df: pd.DataFrame) -> pd.DataFrame:
    c, h, l = df["close"], df["high"], df["low"]
    out = pd.DataFrame(index=df.index)
    out["ema10"] = c.ewm(span=10, adjust=False).mean()
    out["sma20"] = c.rolling(20).mean()
    out["sma50"] = c.rolling(50).mean()
    out["ret21"] = c / c.shift(21) - 1
    out["pct_from_high"] = c / h.rolling(252).max()
    rng = (h - l).replace(0, np.nan)
    out["range_pos"] = ((c - l) / rng).fillna(0.5)
    out["high20_prior"] = h.shift(1).rolling(20).max()
    out["new20_recent"] = (c > out["high20_prior"]).astype(int).rolling(10).max().fillna(0).astype(bool)
    return out
