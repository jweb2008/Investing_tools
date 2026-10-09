# Investing Tools Roadmap

Swing trade breakout scanner, scan log, and scorecard. Daily bars only.
Holding period: a few days to a few weeks. The scanner builds a shortlist
for chart review on thinkorswim. It is not a buy signal.

## Ground rules

- Read-only Schwab access. No order placement code anywhere.
- Keep the daily-bar swing focus. No intraday data.
- Do not break the trade feed: `sync_trades.py`, `update_tracker.py`,
  `backup.py`, and `auth_setup.py` keep their names and behavior, and share
  the same `.env` and `schwab_token.json`.
- Stagger scanner and trade feed schedules by a few minutes so they never
  refresh the shared token file at the same moment.
- Secrets, the token file, and Excel files stay out of Git (see `.gitignore`).
- Free data preferred. Must run on a Windows PC now and an old laptop
  homelab later.
- Avoid duplicate signals. A new indicator must add information the score
  does not already capture.

## Phase 1: Live data

1. Import existing code into this repo (scanner, scan runner, bot, watchlist,
   trade feed scripts).
2. `py scan.py AAPL` against live Schwab data. Compare last close and
   50/200 SMA with thinkorswim, then with `DATA_SOURCE=yahoo`.
3. Confirm price history requests about 2 years back (200 SMA plus 6-month
   band width lookback needs roughly 330 trading days).
4. A handful of tickers, then the full watchlist. Confirm pacing stays under
   120 requests per minute.
5. Confirm no order-related imports or calls.

## Phase 2: Index universe

- Universe: S&P 500 + S&P 400 + Nasdaq 100 (about 1,000 tickers).
- Rebuilt weekly: apply price $5+, $20M+ average dollar volume, optional
  price ceiling (SCAN_MAX_PRICE), 1 year of history,
  drop ETFs and funds.
- Daily scan runs against the cleaned universe (roughly 9 to 10 minutes).
- `watchlist.txt` stays available as a quick separate check.

## Phase 3: Scan log and market filter (built)

scan_log.csv is the master record; scanner_log.xlsx (table "ScanLog") is
rebuilt from it each run. Logs setups, near misses (uptrend + near pivot with
too few contraction checks), and breakouts. Refine if near misses add noise.

`scanner_log.xlsx` (or CSV), read by the Excel tracker through Power Query.
Built with the scorecard in mind from day one:

- Scan date and ticker as the match key
- Total score and each component score stored separately
- Breakout triggered flag
- SPY and QQQ above or below their 50-day
- Stop, targets, and reward-to-risk (once Phase 4 lands)
- Which bucket the name was moved to, if any (swing, long term, LEAPS)

## Phase 4: Risk levels and price targets

- Suggested stop: base low or an ATR multiple.
- Targets:
  - Measured move: base depth added to the pivot
  - 2R and 3R from entry using the suggested stop
  - Overhead resistance: nearest prior swing high or 52-week high
  - ATR projection over the holding window
- Report a conservative target (lower of measured move and overhead
  resistance) and reward-to-risk. Flag setups under 2:1.
- All levels go into the scan log so the scorecard can measure how often
  each was reached.

## Phase 4b: Position sizing and sector grouping

- Position sizing: fixed risk per trade (for example 1% of the account).
  Stop distance converts to a suggested share count. Account size and risk
  % live in `.env` or a config file, not in code.
- Sector grouping: group the shortlist by sector and flag when several
  setups are concentrated in one group (effectively one bet).

## Phase 4c: Failure alerts

- If the scan fails (expired token, Schwab outage, missing data, empty
  universe), say so clearly instead of producing an empty or partial list.
- Warn ahead of the Saturday token expiry if the token is close to lapsing.
- Notification method TBD (local desktop notice now, Discord later).

## Phase 5: Earnings warning

- Flag setups with earnings in the next 2 to 3 weeks.
- Source: yfinance (Schwab is not a reliable earnings date source).

## Phase 6: Relative strength ranking

- Relative strength vs SPY over 3 to 6 months, used to sort setups.

## Phase 6b: Backtesting

- Replay the scanner over the last several years of daily bars and record
  how each flagged setup played out (target hits, stop hits, return over
  the holding window).
- First read on whether higher scores and each component actually matter,
  before the live scan log has enough history.
- Avoid survivorship bias: today's index members did not all exist or
  qualify in the past. Note this limitation in results if historical
  membership is not available for free.
- Results feed the same scorecard reports, labeled backtest vs live.

## Phase 7: Scorecard (separate component)

Grades results by joining the scan log to the Excel tracker.

- The tracker is the source of truth. Use its net % as recorded. Never
  recalculate P&L.
- Matching rule: a trade is a scanner trade if entered within about 5 trading
  days after a flag, linked to the most recent flag before entry. Unmatched
  trades are labeled discretionary.
- Stocks and options are reported separately, never blended. Options match
  on the underlying ticker.
- For options, also show the underlying stock's move over the same holding
  period, to separate setup quality from option-specific factors.
- Reports:
  - Win rate and average return by score band (60s, 70s, 80+)
  - Which score components predict winners
  - Results with the market filter above vs below the 50-day
  - Pre-breakout entries vs triggered-breakout entries
  - Flagged setups taken vs passed
  - Holding period and R-multiples
  - Target hit rates by method
- Wait for roughly 50+ trades per category before tuning weights.
- Delivered as an Excel dashboard via Power Query or a small Python report.

## Later

- Scoring profiles on one shared engine (one data pull, many profiles).
  Swing first. Long term/LEAPS profile later, weighting trend strength and
  6 to 12 month relative strength over a tight base.
- Discord bot: built but not set up. Staying local for now.
- Move to the homelab laptop by cloning this repo.

## Parked

- TTM Squeeze: overlaps with the Bollinger squeeze and ATR contraction
  criteria. Revisit once the scorecard shows whether momentum direction
  would add value.
- Daily (intraday timing) bucket: needs intraday data, outside the
  daily-bar design.
