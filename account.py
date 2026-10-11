"""
Account value and available cash for position sizing, read from Schwab.

Read-only: one balance request per scan (get_accounts). It never touches
orders, and nothing here is written to the scan log or the repo.

Which account sizes the trades
  SIZING_ACCOUNT in .env, by tracker name or last 4 digits, e.g.
  SIZING_ACCOUNT=Roth or SIZING_ACCOUNT=1234 (names come from ACCOUNT_NAMES).
  With only one account, that one is used automatically.

Manual override
  ACCOUNT_SIZE in .env, if set, is used instead of Schwab (for example when
  scanning with Yahoo data).

Sizing
  risk budget  = account value x RISK_PCT (default 1%)
  shares       = risk budget / risk per share, capped by available cash
"""
import os

from dotenv import load_dotenv

load_dotenv()


def _float(name, default=None):
    try:
        v = os.getenv(name, "").strip()
        return float(v) if v else default
    except ValueError:
        return default


def _names():
    raw = os.getenv("ACCOUNT_NAMES", "")
    m = dict(p.strip().split("=", 1) for p in raw.split(",") if "=" in p)
    return {k.strip()[-4:]: v.strip() for k, v in m.items()}


def _balances(acct):
    """Account value and cash available to buy stock, from Schwab's balance block."""
    b = acct.get("currentBalances", {}) or {}
    value = b.get("liquidationValue")
    if value is None:
        value = (acct.get("initialBalances", {}) or {}).get("accountValue")
    # cash accounts (IRAs) report cashAvailableForTrading; margin accounts
    # report availableFunds. Margin buying power is deliberately not used.
    cash = b.get("cashAvailableForTrading")
    if cash is None:
        cash = b.get("availableFunds")
    return value, cash


def list_accounts(client):
    r = client.get_accounts()
    r.raise_for_status()
    names = _names()
    out = []
    for a in r.json():
        sa = a.get("securitiesAccount", {})
        last4 = str(sa.get("accountNumber", ""))[-4:]
        value, cash = _balances(sa)
        out.append({"last4": last4, "name": names.get(last4, ""), "type": sa.get("type", ""),
                    "value": value, "cash": cash})
    return out


def sizing_account(src):
    """Returns (info dict or None, message). Never raises."""
    risk_pct = _float("RISK_PCT", 1.0)
    manual = _float("ACCOUNT_SIZE")
    if manual:
        return ({"label": "ACCOUNT_SIZE in .env", "value": manual, "cash": manual,
                 "risk_pct": risk_pct}, None)

    client = getattr(src, "client", None)
    if client is None:
        return None, "Sizing off: needs Schwab data (or ACCOUNT_SIZE in .env)."
    try:
        accts = list_accounts(client)
    except Exception as e:
        return None, f"Sizing off: could not read account balances ({str(e)[:80]})."
    if not accts:
        return None, "Sizing off: no accounts returned."

    want = os.getenv("SIZING_ACCOUNT", "").strip()
    if want:
        pick = [a for a in accts if want.lower() in (a["name"].lower(), a["last4"])]
        if not pick:
            return None, f"Sizing off: SIZING_ACCOUNT={want} matched no account."
        a = pick[0]
    elif len(accts) == 1:
        a = accts[0]
    else:
        choices = ", ".join(f"{a['name'] or '...' + a['last4']}" for a in accts)
        return None, (f"Sizing off: {len(accts)} accounts found ({choices}). "
                      "Set SIZING_ACCOUNT in .env to the one you swing trade in.")

    if not a["value"]:
        return None, f"Sizing off: no balance reported for ...{a['last4']}."
    label = f"{a['name'] + ' ' if a['name'] else ''}...{a['last4']}"
    return ({"label": label, "value": float(a["value"]),
             "cash": float(a["cash"]) if a["cash"] is not None else float(a["value"]),
             "risk_pct": risk_pct}, None)


def describe(info):
    budget = info["value"] * info["risk_pct"] / 100
    return (f"Sizing: {info['label']}  value ${info['value']:,.0f}, cash ${info['cash']:,.0f}, "
            f"{info['risk_pct']:g}% risk = ${budget:,.0f} per trade")
