"""
Option ideas for each setup: a swing call, the same idea as a debit call
spread, and a LEAPS line. Read-only Schwab option chain data; nothing here
can place an order.

The stock setup is the signal; the option is just the vehicle. Every number
is a planning estimate, not a quote you can count on at the open.

Swing call
  Expiration  closest to SWING_DTE_TARGET days, within SWING_DTE_MIN..MAX
  Strike      delta closest to SWING_DELTA (slightly in the money)
  Liquidity   prefers open interest >= MIN_OI with a real bid
  Estimates   option value at the setup's entry, target and stop, repriced
              with Black-Scholes at the option's own implied volatility,
              assuming HOLD_DAYS calendar days pass before target or stop
  Option R:R  (value at target - value at entry) / (value at entry - value at stop)

Debit call spread (learning comparison)
  Buy the swing call, sell the call at the first strike at or above the
  stock target, same expiration. Shows net debit, max gain, max R:R, and
  estimated value at the target.

LEAPS
  First expiration at least LEAPS_DTE_MIN days out, strike with delta
  closest to LEAPS_DELTA (deep in the money "stock replacement").
  Extrinsic % = share of the premium that is time value (what decays).

Option Score (0 to 100, swing call only)
  Judges the contract as a vehicle for the setup; the stock score judges the
  setup itself, so nothing here repeats a stock criterion.
    payoff     35  option R:R            2.0+ = 35, 1.5+ = 25, 1.0+ = 12
    vol price  25  IV / 20d realized vol 1.0x or less = 25, 1.2x = 17, 1.4x = 8
    cost       25  bid/ask % of mid      3% or less = 25, 6% = 17, 10% = 8
    liquidity  15  open interest         1,000+ = 15, 500+ = 10, 100+ = 5
  Gate: open interest under 100, no bid, or bid/ask over 15% scores 0 and is
  marked not tradeable. Theta and breakeven are already inside option R:R;
  delta and DTE decide which contract is picked, so neither is scored.
  Earnings before expiration is a flag only.

Flags
  wide       bid/ask spread over WIDE_SPREAD_PCT of the mid
  thin       open interest under MIN_OI
  pricey IV  implied volatility over IV_HV_RICH x the stock's 20-day realized
  earnings   an earnings date falls before expiration
"""
import datetime as dt
import math

import numpy as np

SWING_DTE_MIN, SWING_DTE_MAX, SWING_DTE_TARGET = 30, 60, 45
SWING_DELTA = 0.65
LEAPS_DTE_MIN, LEAPS_DTE_MAX = 365, 800
LEAPS_DELTA = 0.80
MIN_OI = 100
WIDE_SPREAD_PCT = 10.0
IV_HV_RICH = 1.3
GATE_SPREAD_PCT = 15.0
SCORE_WEIGHTS = {"payoff": 35, "vol": 25, "cost": 25, "liquidity": 15}
HOLD_DAYS = 14          # calendar days assumed to reach the target or the stop
RATE = 0.04             # risk-free rate for repricing (approximate)


# ------------------------------------------------------------ Black-Scholes
def _ncdf(x):
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def bs_call(S, K, days, iv, r=RATE):
    """Call value. iv as a fraction (0.35), days = calendar days to expiration."""
    if S <= 0 or K <= 0:
        return np.nan
    t = max(days, 0) / 365
    if t == 0 or iv <= 0:
        return max(S - K, 0.0)
    d1 = (math.log(S / K) + (r + iv * iv / 2) * t) / (iv * math.sqrt(t))
    d2 = d1 - iv * math.sqrt(t)
    return S * _ncdf(d1) - K * math.exp(-r * t) * _ncdf(d2)


def realized_vol(close, days=20):
    lr = np.log(close / close.shift(1)).dropna().iloc[-days:]
    return float(lr.std() * math.sqrt(252)) if len(lr) >= days // 2 else np.nan


# ------------------------------------------------------------ option score
def _tier(value, steps, higher_is_better=True):
    """steps: list of (threshold, points), best first."""
    if value is None or value != value:
        return 0
    for threshold, pts in steps:
        if (value >= threshold) if higher_is_better else (value <= threshold):
            return pts
    return 0


def score_option(o):
    """Option Score for the swing call in o (output of ideas()). Returns dict of fields."""
    reasons = []
    if o.get("opt_oi", 0) < MIN_OI:
        reasons.append(f"open interest under {MIN_OI}")
    if not o.get("opt_bid", 0) > 0:
        reasons.append("no bid")
    sp = o.get("opt_spread_pct")
    if sp is None or sp != sp or sp > GATE_SPREAD_PCT:
        reasons.append(f"bid/ask over {GATE_SPREAD_PCT:.0f}%")
    pts = {
        "payoff": _tier(o.get("opt_rr"), [(2.0, 35), (1.5, 25), (1.0, 12)]),
        "vol": _tier(o.get("opt_iv_hv"), [(1.0, 25), (1.2, 17), (1.4, 8)], higher_is_better=False),
        "cost": _tier(sp, [(3.0, 25), (6.0, 17), (10.0, 8)], higher_is_better=False),
        "liquidity": _tier(o.get("opt_oi"), [(1000, 15), (500, 10), (100, 5)]),
    }
    tradeable = not reasons
    return {
        "opt_score": sum(pts.values()) if tradeable else 0,
        "opt_tradeable": tradeable,
        "opt_gate_reason": "; ".join(reasons),
        **{f"opt_pts_{k}": v for k, v in pts.items()},
    }


# ------------------------------------------------------------ chain parsing
def _contracts(chain):
    """Flatten Schwab's callExpDateMap into a list of dicts."""
    out = []
    for exp_key, strikes in (chain.get("callExpDateMap") or {}).items():
        for _, lst in strikes.items():
            for c in lst:
                bid, ask = c.get("bid") or 0.0, c.get("ask") or 0.0
                mid = (bid + ask) / 2 if ask > 0 else (c.get("mark") or c.get("last") or 0.0)
                iv = c.get("volatility")
                iv = iv / 100 if isinstance(iv, (int, float)) and 0 < iv < 1000 else np.nan
                delta = c.get("delta")
                delta = delta if isinstance(delta, (int, float)) and -1 <= delta <= 1 else np.nan
                exp = exp_key.split(":")[0]
                out.append({
                    "symbol": c.get("symbol", "").strip(),
                    "desc": c.get("description", ""),
                    "exp": dt.date.fromisoformat(exp),
                    "dte": int(c.get("daysToExpiration") or 0),
                    "strike": float(c.get("strikePrice") or 0),
                    "bid": float(bid), "ask": float(ask), "mid": float(mid),
                    "delta": delta, "iv": iv,
                    "oi": int(c.get("openInterest") or 0),
                    "volume": int(c.get("totalVolume") or 0),
                })
    return [c for c in out if c["strike"] > 0 and c["mid"] > 0]


def _pick_expiration(cons, lo, hi, target=None):
    exps = sorted({(c["exp"], c["dte"]) for c in cons if lo <= c["dte"] <= hi}, key=lambda x: x[1])
    if not exps:
        return None
    if target is None:
        return exps[0][0]
    return min(exps, key=lambda x: abs(x[1] - target))[0]


def _pick_strike(cons, exp, want_delta):
    pool = [c for c in cons if c["exp"] == exp and not np.isnan(c["delta"])]
    if not pool:
        return None
    liquid = [c for c in pool if c["oi"] >= MIN_OI and c["bid"] > 0]
    return min(liquid or pool, key=lambda c: abs(c["delta"] - want_delta))


def _label(c):
    """Tracker style, e.g. 'ANET 11/20/2026 210.00 C'."""
    return f"{c['root']} {c['exp']:%m/%d/%Y} {c['strike']:.2f} C"


def _spread_pct(c):
    return (c["ask"] - c["bid"]) / c["mid"] * 100 if c["mid"] > 0 and c["ask"] > 0 else np.nan


# ------------------------------------------------------------ main entry
def fetch_chain(src, symbol, lo, hi, strike_count):
    client = src.client
    src._pace()
    today = dt.date.today()
    r = client.get_option_chain(
        symbol.replace(".", "/"),
        contract_type=client.Options.ContractType.CALL,
        strike_count=strike_count,
        from_date=today + dt.timedelta(days=lo),
        to_date=today + dt.timedelta(days=hi))
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}")
    return r.json() or {}


def ideas(src, symbol, r, earn_date=None):
    """Option ideas for one setup. r = scanner + levels output (with hv20). Returns a dict."""
    out = {"opt_error": None}
    S = r["close"]
    hv = r.get("hv20") or np.nan
    root = symbol.replace(".", "/")

    # ---- swing call + spread
    try:
        cons = _contracts(fetch_chain(src, symbol, SWING_DTE_MIN, SWING_DTE_MAX, 24))
        for c in cons:
            c["root"] = root
        exp = _pick_expiration(cons, SWING_DTE_MIN, SWING_DTE_MAX, SWING_DTE_TARGET)
        c = _pick_strike(cons, exp, SWING_DELTA) if exp else None
    except Exception as e:
        cons, c = [], None
        out["opt_error"] = f"swing chain: {str(e)[:60]}"

    if c:
        iv = c["iv"] if not np.isnan(c["iv"]) else hv
        later = max(c["dte"] - HOLD_DAYS, 0)
        v_entry = bs_call(r["entry"], c["strike"], c["dte"], iv)
        v_target = bs_call(r["target"], c["strike"], later, iv)
        v_stop = bs_call(r["stop"], c["strike"], later, iv)
        risk = v_entry - v_stop
        flags = []
        sp = _spread_pct(c)
        if sp > WIDE_SPREAD_PCT:
            flags.append("wide")
        if c["oi"] < MIN_OI:
            flags.append("thin")
        if hv == hv and not np.isnan(c["iv"]) and c["iv"] > IV_HV_RICH * hv:
            flags.append("pricey IV")
        earn_before = bool(earn_date and earn_date <= c["exp"])
        if earn_before:
            flags.append("earnings")
        out.update({
            "opt_contract": _label(c), "opt_symbol": c["symbol"], "opt_exp": c["exp"],
            "opt_dte": c["dte"], "opt_strike": c["strike"], "opt_bid": c["bid"], "opt_ask": c["ask"],
            "opt_mid": c["mid"], "opt_delta": c["delta"], "opt_iv": c["iv"],
            "opt_oi": c["oi"], "opt_volume": c["volume"], "opt_spread_pct": sp,
            "opt_breakeven": c["strike"] + c["mid"],
            "opt_be_pct": (c["strike"] + c["mid"] - S) / S * 100,
            "opt_iv_hv": c["iv"] / hv if hv == hv and hv > 0 and not np.isnan(c["iv"]) else np.nan,
            "opt_val_entry": v_entry, "opt_val_target": v_target, "opt_val_stop": v_stop,
            "opt_rr": (v_target - v_entry) / risk if risk > 0 else np.nan,
            "opt_earn_before_exp": earn_before,
            "opt_flags": ", ".join(flags),
        })
        out.update(score_option(out))

        # debit spread: short leg at the first strike at/above the stock target
        shorts = sorted((x for x in cons if x["exp"] == c["exp"] and x["strike"] > c["strike"]
                         and x["strike"] >= r["target"]), key=lambda x: x["strike"])
        if shorts:
            s = shorts[0]
            debit = c["mid"] - s["mid"]
            width = s["strike"] - c["strike"]
            if debit > 0:
                s_iv = s["iv"] if not np.isnan(s["iv"]) else iv
                v_t = v_target - bs_call(r["target"], s["strike"], later, s_iv)
                out.update({
                    "spr_short_strike": s["strike"], "spr_debit": debit,
                    "spr_max_gain": width - debit, "spr_max_rr": (width - debit) / debit,
                    "spr_val_target": v_t,
                })

    # ---- LEAPS
    try:
        lcons = _contracts(fetch_chain(src, symbol, LEAPS_DTE_MIN, LEAPS_DTE_MAX, 30))
        for x in lcons:
            x["root"] = root
        lexp = _pick_expiration(lcons, LEAPS_DTE_MIN, LEAPS_DTE_MAX)
        lc = _pick_strike(lcons, lexp, LEAPS_DELTA) if lexp else None
    except Exception as e:
        lc = None
        out["opt_error"] = (out["opt_error"] + "; " if out["opt_error"] else "") + f"LEAPS chain: {str(e)[:60]}"
    if lc:
        intrinsic = max(S - lc["strike"], 0)
        out.update({
            "leaps_contract": _label(lc), "leaps_exp": lc["exp"], "leaps_dte": lc["dte"],
            "leaps_strike": lc["strike"], "leaps_mid": lc["mid"], "leaps_delta": lc["delta"],
            "leaps_iv": lc["iv"], "leaps_oi": lc["oi"], "leaps_spread_pct": _spread_pct(lc),
            "leaps_breakeven": lc["strike"] + lc["mid"],
            "leaps_extrinsic_pct": (lc["mid"] - intrinsic) / lc["mid"] * 100 if lc["mid"] else np.nan,
        })
    return out
