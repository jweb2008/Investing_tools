# After Action Review (AAR): Design

Proposed design for the trade results review (formerly "scorecard"). Builds on
SCORECARD_BRIEF.md, which keeps the ground rules. This document records the
decisions made on the open questions and how the tool will work. Nothing is
built yet.

Status: Approved for build, 2026-10-10.

## Decisions

| # | Question | Decision |
|---|----------|----------|
| 1 | Name | After Action Review. Script `aar.py`, output `aar.xlsx`. |
| 2 | Delivery | Python does the matching and grading and writes `aar.xlsx` as named Excel tables. The Excel dashboard reads them through Power Query. A short summary also prints to the console. |
| 3 | Base data | Tracker Table14 is the base and its net % is used as recorded. Sync ID links each row to exact fills in `fills_cache.csv`. `trade_log.xlsx` Closed Trades is not used (FIFO pairing differs from the tracker's swing pairing). |
| 4 | Partial exits, scale-ins | Graded per trade, not per tracker row. Each trade is tagged Swing, Long term, or Hedge. Swing trades are the main focus. |
| 5 | Matching window | 1 to 5 trading days after ScanDate. Tuned later from the actual entry lag. |
| 6 | Win definition | Three outcomes: Win, Scratch, Loss. A small gain is a Scratch, not a Win. |
| 7 | Hedges | Sold-to-open options go to their own Hedge bucket, outside setup grading. |
| 8 | Accounts | All accounts count. Account and Source (Scanner or Discretionary) are filters on every table. |

Change from the brief: score bands are replaced by check combinations. The
score can only be 80, 85, or 100 for a setup (50, 65, or 70 for a near miss),
and 85 covers two different setups. See Reports.

## Inputs (all read-only)

| Input | Used for |
|-------|----------|
| Newest tracker, Table14 | Positions, dates, Net Profit, TOTAL INVESTMENT, % CHANGE, Open/Closed |
| `data/scan_log.csv` | Flags: status, score, checks, raw values, market filter, risk levels (from 10/12) |
| `fills_cache.csv` | Entry and exit details for rows with a Sync ID, and short (sold-to-open) detection |
| Daily bars via `data.py` | Trading-day calendar (SPY), underlying moves for options, forward outcomes for flags |
| `aar_tags.csv` (new, yours) | Manual Long term or Swing overrides (see Trade style) |

Reading the tracker: openpyxl with `data_only=True`, so the calculated
columns come in as the values Excel last saved. The tracker is never opened
in Excel by this tool and never written to. `find_latest_tracker()` and the
TRACKER_* settings are reused from `update_tracker.py` by import only; that
file is not changed. If the tracker is open in Excel, AAR reads a temporary
copy.

Price history is cached in `data/aar_bars/` (CSV, gitignored) so a rerun
only pulls what is new. Only tickers that were flagged or traded are fetched.
AAR runs on its own schedule, staggered from the trade feed and the scan so
they never refresh the token at the same moment. `run_update.bat` is not
touched; AAR gets its own `run_aar.bat`.

## Step 1: Build trades from tracker rows

A trade is one continuous holding in one account: it starts when the
position goes from flat to open and ends when it is flat again.

- Grouped by Account + SYMBOL as written in the tracker. For options that is
  the full contract, so a roll to a new strike or expiry is a new trade.
- Scale-ins and partial exits become rows inside one trade.
- Trade entry date = first Opened Date. Exit date = last Closed Date.
- Trade net % = sum of the rows' recorded Net Profit / sum of their recorded
  TOTAL INVESTMENT. These are the tracker's own numbers combined, not a
  recalculation from prices.
- Hold period in trading days, from the SPY calendar.
- Open trades are listed and matched but not graded until closed.
- Rows that can't be placed (no Opened Date, "Needs open" rows from before
  2026) go to the Review sheet instead of being guessed.

## Step 2: Trade style

Every trade gets one style. Only Swing trades feed the setup grading.

The tracker's Action column already carries the style for almost every row
(Swing, Long, Hedge, Day, plus STC and BTO), so it is the main source.

| Style | How it's set (first match wins) |
|-------|----------------------------------|
| Hedge | Opened by selling (sold-to-open option), from fills via Sync ID or a sell-side open with no cost. Overrides the Action value. |
| From `aar_tags.csv` | A matching override line (optional file you edit). |
| From the Action column | Swing, Long (Long term), Hedge, Day. STC and BTO are the front or back end of a hedge, so they map to Hedge. Unknown values keep their own group and show on the Review sheet. |
| Long term | Option with a year or more to expiry at open (LEAPS), when Action is blank. |
| Untagged | Nothing above; listed on the Review sheet. |

Trades are chained per account, symbol AND row style, so a swing lot traded
on top of a long-term core position in the same account stays its own trade.

`aar_tags.csv` format:

```
Account,Symbol,OpenedDate,Style
Roth,MSFT,,Long term          <- every MSFT trade in Roth
Joint,NVDA,2026-03-14,Swing   <- one specific trade
```

Hold length does NOT set style automatically. If it did, swing trades that
worked and ran longer would be pulled out of the swing results. Swing trades
held over 40 trading days are listed on the Review sheet instead.

## Step 2b: Decisions across accounts

The same stock or contract entered on the same day in several accounts is one
decision (DecisionID). Setup grading counts decisions, combining the
accounts' recorded Net Profit and TOTAL INVESTMENT, so one entry split
across three family accounts counts once, not three times. Per-account
trades stay available for account filters.

## Step 3: Match trades to scanner flags

- Flags are grouped into episodes: the same ticker flagged on consecutive
  scans (gaps of up to 5 trading days) is one setup episode.
- A trade is a Scanner trade if its entry date is 1 to 5 trading days after a
  ScanDate for that ticker (the option's underlying for options). It links to
  the most recent flag before entry, as the brief says. Entry on the ScanDate
  itself is Discretionary, because ScanDate is the last completed bar and the
  flag did not exist yet.
- Any status counts as a flag (Setup, Near miss, Breakout). The matched
  status is recorded so near-miss trades can be compared to setup trades.
- Recorded on each matched trade: days after the matched flag, days since the
  episode's first flag, and the flag's score, checks, raw values, market
  filter, and risk levels.
- Unmatched trades are Discretionary.

The Review sheet shows the actual spread of entry lags. Once enough trades
exist, that tells us whether 5 days is right.

## Step 4: Grade

**Outcome.** Win, Scratch, or Loss from the trade's recorded net %.

| | Win | Scratch | Loss |
|-|-----|---------|------|
| Stocks | above +2% | -2% to +2% (inclusive) | below -2% |
| Options | above +10% | -10% to +10% (inclusive) | below -10% |

The band is symmetric on purpose: a failed breakout cut near breakeven is
good trade management and is not counted with full stop-outs.

Thresholds live in `.env` (AAR_SCRATCH_STOCK, AAR_SCRATCH_OPTION).

**R-multiple (stocks, flags from 10/12 on).** Risk % = (your actual average
entry price - the scanner's Stop) / your entry price. R = recorded net % /
risk %. Your net % is unchanged; R just expresses it in units of the
scanner's suggested risk. If you actually used a different stop, this
measures against the scanner's stop, which is what we want to grade.

**Entry type.** Pre-breakout if your average entry price was at or below the
matched flag's Pivot, Post-breakout if above.

**Price path during the hold** (from daily bars, entry date to exit date):
maximum favorable and adverse move, whether the scanner's Stop was touched,
and the highest target reached (2R, Measured move, Resistance, 3x ATR, 3R).
This answers "which target method fits my trades" even when you exited early.

**Options.** Each option trade also gets the underlying's close-to-close
move over the same dates, plus its R and targets reached, so an option loss
on a setup that worked is visible as an option problem, not a setup problem.
Daily closes only, so this is approximate.

## Step 5: Flag outcomes (taken vs passed)

To compare setups you took with setups you passed on, every episode is graded
on its own, independent of what you did:

- Triggered: price traded above the Entry level within 5 trading days.
- If triggered: Stop or Target first (a day that touches both counts as
  Stop, to stay conservative), highest target reached, and return 10 and 20
  trading days after the trigger.
- Taken: Yes (with Trade ID) or No.

Flags before 10/12 have no logged Stop or targets, so they get Triggered and
forward returns only.

## Output: aar.xlsx

Each sheet is a named table for Power Query.

| Sheet / table | Contents |
|---------------|----------|
| Summary | Headline tables (below), printed to the console too |
| Trades | One row per trade with everything above |
| Decisions | Closed trades rolled up across accounts (what setup grading counts) |
| TradeRows | Each tracker row with its Trade ID, for checking the grouping |
| Flags | One row per setup episode with its outcome and Taken flag |
| Review | Data issues: untagged long holds, unplaceable rows, entry lag spread, flags with missing levels |
| RunInfo | Run time, tracker file used, scan log range, settings |

## Reports

Every report is split Stocks vs Options, never blended, and filterable by
Account. Swing trades only unless noted. Each line shows N, Win / Scratch /
Loss %, average and median net %, average win, average loss, and expectancy
(and average R once available). Groups under 20 trades are marked "too few to
read"; weight tuning waits for 50+.

1. Scanner vs Discretionary
2. Check combination: all three, squeeze + ATR, squeeze + volume, ATR + volume
3. Each component's raw value in buckets: band width percentile, ATR ratio,
   volume ratio, % below pivot
4. Matched flag status: Setup vs Near miss vs Breakout (tests the 2-of-3 cutoff)
5. Market filter: SPY and QQQ above vs below their 50-day
6. Entry type: Pre-breakout vs Post-breakout
7. Entry lag: days after flag, days since first flag
8. Hold period buckets
9. R-multiples and target hit rates by method
10. Flags taken vs passed (from Flag outcomes)
11. Options: option result vs underlying move
12. Long term and Hedge: simple totals, reported separately

Backtest results (Phase 6b) will carry a Mode column (Live or Backtest) and
are never mixed into live tables.

## Build order

1. Read tracker and scan log, build trades and styles, write Trades and
   TradeRows. Check the grouping against your tracker by hand.
2. Matching and the Review sheet.
3. Grading and Summary for stocks.
4. Price path, R-multiples, target hits, options underlying moves.
5. Flag outcomes and taken vs passed.
6. Power Query dashboard in the tracker.

Steps 1 and 2 are useful right away on your existing 2026 trades (all will
be Discretionary until scanner trades exist).

## Build log

- 2026-10-10 Step 1 built (`aar.py`): trades, styles, decisions, outcomes,
  Review sheet. Tested on tracker 10.9.11: 299 rows, 200 trades, 139 closed
  decisions; trade Net Profit ties to tracker rows to the cent.

## Resolved after review (2026-10-10)

- A. Style comes from the Action column you fill in by hand, with
  `aar_tags.csv` for overrides.
- B. Scratch bands ±2% (stocks) and ±10% (options), symmetric. Confirmed.
- C. Options grouped by contract; a roll starts a new trade. Confirmed.
- D. STC and BTO Action values are hedge legs and map to Hedge.
- E. Same-day entries across accounts count as one decision. Confirmed (and
  becoming rarer going forward).
