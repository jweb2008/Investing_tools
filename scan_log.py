"""
Scan log: every setup, near-miss and breakout the scanner finds, one row per
stock per day, for the scorecard to grade against your trade tracker later.

  data/scan_log.csv   master record (the source of truth, append-only)
  scanner_log.xlsx    rebuilt from the CSV each run, as Excel table "ScanLog"
                      for Power Query

Rows are keyed on ScanDate + Symbol. Rerunning a scan on the same day
replaces that day's rows for those symbols rather than duplicating them.

Status values
  Setup      uptrend + near pivot + enough contraction checks (the shortlist)
  Near miss  uptrend + near pivot, but too few contraction checks
  Breakout   uptrend, closed above the pivot on 1.5x+ volume
"""
import datetime as dt
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd

HERE = Path(__file__).parent
CSV_PATH = HERE / "data" / "scan_log.csv"
XLSX_PATH = HERE / "scanner_log.xlsx"
UNIVERSE_INFO = HERE / "data" / "universe_info.csv"

COLUMNS = [
    "ScanDate", "Symbol", "Status", "Score", "Checks",
    "Uptrend", "NearPivot", "BBSqueeze", "ATRContract", "VolDryup",
    "Close", "Pivot", "PctBelowPivot", "BaseLow", "SMA50", "SMA200",
    "BBWidthPctile", "ATRRatio", "ATR14", "Vol10", "Vol50", "TodayVolX",
    "Entry", "Stop", "StopMethod", "RiskPerShare", "RiskPct",
    "TargetMeasured", "Target2R", "Target3R", "Resistance", "TargetATR",
    "Target", "RewardRisk", "RROk",
    "SPYAbove50", "QQQAbove50", "Sector", "MinChecks", "List", "Source", "LoggedAt",
]


def status_of(r, min_checks):
    if r["breakout"]:
        return "Breakout"
    if r["required_ok"] and r["checks"] >= min_checks:
        return "Setup"
    if r["required_ok"]:
        return "Near miss"
    return None


def _sectors():
    try:
        info = pd.read_csv(UNIVERSE_INFO, usecols=["symbol", "sector"])
        return dict(zip(info["symbol"], info["sector"]))
    except Exception:
        return {}


def _r(x, nd=2):
    return None if x is None or (isinstance(x, float) and x != x) else round(float(x), nd)


def build_rows(results, market, min_checks, list_name, source):
    sectors = _sectors()
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    rows = []
    for r in results:
        st = status_of(r, min_checks)
        if not st:
            continue
        rows.append({
            "ScanDate": pd.Timestamp(r["date"]).date(),
            "Symbol": r["symbol"],
            "Status": st,
            "Score": r["score"],
            "Checks": r["checks"],
            "Uptrend": r["pts_uptrend"],
            "NearPivot": r["pts_near_pivot"],
            "BBSqueeze": r["pts_bb_squeeze"],
            "ATRContract": r["pts_atr_contract"],
            "VolDryup": r["pts_vol_dryup"],
            "Close": round(r["close"], 2),
            "Pivot": round(r["pivot"], 2),
            "PctBelowPivot": round(r["pct_below_pivot"], 2),
            "BaseLow": round(r["base_low"], 2),
            "SMA50": round(r["sma50"], 2),
            "SMA200": round(r["sma200"], 2),
            "BBWidthPctile": round(r["bb_width_pctile"], 1),
            "ATRRatio": round(r["atr_ratio"], 3),
            "ATR14": round(r["atr14"], 2),
            "Vol10": int(r["vol_10"]),
            "Vol50": int(r["vol_50"]),
            "TodayVolX": round(r["today_vol_x"], 2),
            "Entry": _r(r.get("entry")),
            "Stop": _r(r.get("stop")),
            "StopMethod": r.get("stop_method"),
            "RiskPerShare": _r(r.get("risk_per_share")),
            "RiskPct": _r(r.get("risk_pct")),
            "TargetMeasured": _r(r.get("target_measured")),
            "Target2R": _r(r.get("target_2r")),
            "Target3R": _r(r.get("target_3r")),
            "Resistance": _r(r.get("resistance")),
            "TargetATR": _r(r.get("target_atr")),
            "Target": _r(r.get("target")),
            "RewardRisk": _r(r.get("reward_risk")),
            "RROk": r.get("rr_ok"),
            "SPYAbove50": market.get("SPY"),
            "QQQAbove50": market.get("QQQ"),
            "Sector": sectors.get(r["symbol"], ""),
            "MinChecks": min_checks,
            "List": list_name,
            "Source": source,
            "LoggedAt": now,
        })
    return pd.DataFrame(rows, columns=COLUMNS)


def load():
    if not CSV_PATH.exists():
        return pd.DataFrame(columns=COLUMNS)
    df = pd.read_csv(CSV_PATH)
    df["ScanDate"] = pd.to_datetime(df["ScanDate"]).dt.date
    return df


def _write_xlsx(df):
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=XLSX_PATH.parent)
    os.close(fd)
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="ScanLog", index=False)
        ws = xw.sheets["ScanLog"]
        for i, col in enumerate(df.columns, 1):
            width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) + 2
            ws.column_dimensions[get_column_letter(i)].width = min(width, 24)
        ws.freeze_panes = "C2"
        if len(df):
            ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
            tbl = Table(displayName="ScanLog", ref=ref)
            tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(tbl)
    try:
        os.replace(tmp, XLSX_PATH)
        return True
    except PermissionError:
        os.remove(tmp)
        return False


def append(new: pd.DataFrame):
    """Merge new rows into the log. Returns a one-line summary."""
    if new.empty:
        return "Scan log: nothing to add."
    log = load()
    merged = (pd.concat([log, new], ignore_index=True)
              .reindex(columns=COLUMNS)
              .drop_duplicates(["ScanDate", "Symbol"], keep="last")
              .sort_values(["ScanDate", "Score", "Symbol"], ascending=[False, False, True])
              .reset_index(drop=True))
    CSV_PATH.parent.mkdir(exist_ok=True)
    merged.to_csv(CSV_PATH, index=False)
    counts = new["Status"].value_counts()
    summary = (f"Scan log: {len(new)} rows for {new['ScanDate'].iloc[0]} "
               f"({', '.join(f'{v} {k.lower()}' for k, v in counts.items())}); "
               f"{len(merged)} rows total")
    if not _write_xlsx(merged):
        summary += (f"\n  {XLSX_PATH.name} is open in Excel, so it was not refreshed. "
                    "The log itself was saved; the workbook updates on the next run.")
    return summary


if __name__ == "__main__":
    # py scan_log.py  -> rebuild scanner_log.xlsx from scan_log.csv
    df = load()
    print(f"{len(df)} rows" if _write_xlsx(df) else f"Close {XLSX_PATH.name} first.")
    sys.exit(0)
