# Trade Results Review: Design Brief

Working name: "scorecard" (to be renamed; see Open questions).

Handoff document for developing the trade results review, the separate
component that grades how the breakout scanner's setups actually performed,
using the trade tracker as the source of truth.

## Purpose

Answer, with real results instead of opinion:
- Do higher-scoring setups win more often and by more?
- Which score components actually predict winners?
- Do scanner-found trades beat discretionary trades?
- Is the 2-of-3 contraction rule the right cutoff?
- Which price target method fits my trades best?

## Ground rules (decided)

- The Excel trade tracker is the source of truth for results. Use its net %
  exactly as recorded. Never recalculate P&L.
- Stocks and options are always reported separately, never blended. Options
  returns run much larger and would skew everything.
- Options match to setups on the underlying ticker.
- For options, also show how the underlying stock moved over the same holding
  period, to separate setup quality from option factors (strike, decay, IV).
- Do not tune scanner weights until roughly 50+ trades per category.
- Read-only. Nothing here touches Schwab orders. Do not break the trade feed
  (sync_trades.py, update_tracker.py, backup.py, auth_setup.py,
  run_update.bat), and do not rename or replace those scripts.
- Free tools only. Must run on Windows now and an old laptop homelab later.

## Matching rule (decided, refine if needed)

A trade counts as a scanner trade if it was entered within about 5 trading
days after the scanner flagged that ticker, linked to the most recent flag
before entry. Unmatched trades are labeled "discretionary."

## Reports wanted

- Win rate and average return by score band (80s, 90s, 100)
- Which score components predict winners (squeeze, ATR contraction,
  volume dry-up)
- Setups vs near misses (near misses are logged on purpose to test the
  2-of-3 cutoff)
- Results with SPY/QQQ above vs below their 50-day
- Pre-breakout entries vs entries after a triggered breakout
- Flagged setups taken vs passed (tests my chart review)
- Holding period
- R-multiples and target hit rates (once stops and targets are logged,
  roadmap Phase 4)
- Backtest results, labeled separately from live (roadmap Phase 6b)

## Data sources

### Scan log (scanner side)
- Master record: `data\scan_log.csv`
- Excel copy: `scanner_log.xlsx`, table `ScanLog`, rebuilt after each scan
- One row per stock per day, keyed on ScanDate + Symbol
- Status values: Setup, Near miss, Breakout
- Columns: ScanDate, Symbol, Status, Score, Checks, Uptrend, NearPivot,
  BBSqueeze, ATRContract, VolDryup, Close, Pivot, PctBelowPivot, BaseLow,
  SMA50, SMA200, BBWidthPctile, ATRRatio, ATR14, Vol10, Vol50, TodayVolX,
  Entry through RROk (see below), SPYAbove50, QQQAbove50, Sector, MinChecks,
  List, Source, LoggedAt
- First live scan logged: 2026-10-09
- Risk levels logged from 2026-10-12 on: Entry, Stop, StopMethod,
  RiskPerShare, RiskPct, TargetMeasured, Target2R, Target3R, Resistance,
  TargetATR, Target (conservative), RewardRisk, RROk

### Trade tracker (results side, source of truth)
- File: `2025-26_Trading_Tracker_<version>.xlsx` (newest version in the main
  folder; older versions in `Tracker Archive\`)
- Table `Table14` on the "New 2026" sheet, one row per position
- Columns written by the trade feed: Account, Opened Date, Stock Or Option,
  SYMBOL, Buy Quantity, Purchase PRICE PER SHARE, Fees & Comm, Closed Date,
  Sell Quantity, Sell Price Per Share, Sale Fees & Comm, Open/Closed, Sync ID
- Calculated in the tracker: TOTAL INVESTMENT, Total Sale Value, % CHANGE,
  Net Profit, Neg $ Loss. Action column is mine and never overwritten.
- "Stock Or Option" is the asset type column for the stock/option split
- Option symbols look like `IREN 11/28/2025 55.00 C`; the underlying is the
  first word
- Sold-to-open options use a negative Buy Quantity ("Hedge" convention)

### Supporting
- `trade_log.xlsx` from sync_trades.py: Fills, Closed Trades (FIFO matched,
  with ReturnPct, HoldDays, AssetType, Underlying), Open Positions
- Daily price history for underlying moves: Schwab via `data.py`
  (`get_source().history(symbol)`)

## Environment

- Repo: github.com/jweb2008/investing_tools (see ROADMAP.md, Phase 7)
- Folder: `C:\Users\jweb2\Desktop\Investing`
- Python 3.14 on Windows, run with `py`
- Excel tracker already uses Power Query

## Open questions

1. Name. Candidates: After Action Review (AAR), Trade Debrief, Setup Report
   Card, Edge Report, Results Review.
2. Delivery: Excel dashboard via Power Query, a small Python report, or both.
3. Match on the tracker (Table14) or on trade_log.xlsx Closed Trades? The
   tracker is the source of truth for net %, but Closed Trades has hold days
   and cleaner per-lot matching.
4. How to handle partial exits and scale-ins.
5. Whether the matching window (about 5 trading days) is right for how I
   actually enter.
6. What counts as a "win" (any positive net %, or above a threshold).
