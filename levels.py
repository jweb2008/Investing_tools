"""
Risk levels and price targets for a setup. Daily bars only.

These are planning references for sizing and trade management, not
predictions. Every level is logged so the scorecard can measure how often
price actually reached each one.

Entry
  Setup     the pivot (a buy-stop at the breakout level)
  Breakout  the close (already through the pivot)

Stop (tighter of the two, but never closer than 1 ATR to entry)
  base low  lowest low of the 30 sessions before today
  ATR stop  entry minus 2 x ATR(14)

Targets
  measured move  pivot + base depth (pivot minus base low)
  2R / 3R        entry + 2x / 3x the risk per share
  resistance     nearest prior swing high above entry in the last year
                 (none = no overhead supply, "blue sky")
  ATR projection entry + 3 x ATR(14)

Conservative target = lower of measured move and resistance.
Reward-to-risk is measured to the conservative target; under 2:1 is flagged.

Position size (optional): set ACCOUNT_SIZE and RISK_PCT in .env, e.g.
ACCOUNT_SIZE=25000 and RISK_PCT=1, to get a suggested share count.
"""
import math
import os

import numpy as np
import pandas as pd

ATR_STOP_MULT = 2.0
MIN_STOP_ATR = 1.0
ATR_TARGET_MULT = 3.0
SWING_WINDOW = 5          # bars on each side that a swing high must exceed
RESISTANCE_LOOKBACK = 252
MIN_RR = 2.0


def _env_float(name):
    try:
        v = os.getenv(name, "").strip()
        return float(v) if v else None
    except ValueError:
        return None


def nearest_resistance(df: pd.DataFrame, above: float):
    """Nearest confirmed swing high above a price, within the last year."""
    h = df["high"].iloc[-RESISTANCE_LOOKBACK:].to_numpy()
    w = SWING_WINDOW
    levels = [h[i] for i in range(w, len(h) - w)
              if h[i] == h[i - w:i + w + 1].max() and h[i] > above * 1.005]
    return float(min(levels)) if levels else None


def compute(df: pd.DataFrame, r: dict) -> dict:
    """r is scanner.score() output for the same df."""
    atr = r["atr14"]
    entry = r["close"] if r["breakout"] else r["pivot"]

    atr_stop = entry - ATR_STOP_MULT * atr
    stop, method = (r["base_low"], "base low") if r["base_low"] >= atr_stop else (atr_stop, "2x ATR")
    if entry - stop < MIN_STOP_ATR * atr:
        stop, method = entry - MIN_STOP_ATR * atr, "1x ATR min"
    risk = entry - stop

    measured = r["pivot"] + (r["pivot"] - r["base_low"])
    resistance = nearest_resistance(df, entry)
    conservative = min(measured, resistance) if resistance else measured
    rr = (conservative - entry) / risk if risk > 0 else np.nan

    out = {
        "entry": entry,
        "stop": stop,
        "stop_method": method,
        "atr_stop": atr_stop,
        "risk_per_share": risk,
        "risk_pct": risk / entry * 100,
        "target_measured": measured,
        "target_2r": entry + 2 * risk,
        "target_3r": entry + 3 * risk,
        "resistance": resistance,
        "target_atr": entry + ATR_TARGET_MULT * atr,
        "target": conservative,
        "target_pct": (conservative - entry) / entry * 100,
        "reward_risk": rr,
        "rr_ok": bool(rr >= MIN_RR) if not np.isnan(rr) else False,
        "shares": None,
    }

    account, risk_pct = _env_float("ACCOUNT_SIZE"), _env_float("RISK_PCT")
    if account and risk_pct and risk > 0:
        by_risk = math.floor(account * risk_pct / 100 / risk)
        by_cash = math.floor(account / entry)          # never more than the account can buy
        out["shares"] = max(min(by_risk, by_cash), 0)
    return out
