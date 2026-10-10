# Trade Feed

Pulls your filled trades from Schwab (read-only) into `trade_log.xlsx` for
your trading tracker.

## Setup
1. Create `.env` in this folder (copy `.env.example`), fill in App Key, Secret,
   and the Callback URL exactly as shown on your Schwab app.
2. `pip install -r requirements.txt`
3. `python auth_setup.py` → open the link, log in, accept, pick your account,
   then paste the full URL from the "can't connect" page back in.
   Repeat every 7 days.
4. `python sync_trades.py`

## What you get in trade_log.xlsx
- **Fills**: every execution (stocks, options, expirations)
- **Closed Trades**: entries matched to exits, FIFO, with P&L, return %, hold days
- **Open Positions**: lots still open
- **Unmatched**: closes with no 2026 entry (positions opened before Jan 1)
- **Sync Info**: when it ran and which date ranges came back

Each sheet is a named Excel table, ready for Power Query.

Run it as often as you like. Nothing duplicates. Close the workbook in Excel
before running, or it can't save.

## Updating the tracker
Put your tracker file in this folder. Then:

- `py update_tracker.py` pulls fills and writes `Review_<version>.xlsx`.
  Nothing in the tracker changes.
- `py update_tracker.py --apply` does the same, then saves a new version
  (`2025-26_Trading_Tracker_M.D.V.xlsx`) through Excel. The previous version is
  never touched, and the Action column is never overwritten.

Rows the script manages get a Sync ID in the last column of the table. Leave it
alone; it's how the script avoids duplicates.

## 2025 lots (starting inventory)
Any row opened before 2026 that is still open (no Closed Date) with its cost
filled in is treated as a lot you held on Jan 1. 2026 sales draw from these:
assignments take the oldest shares first, regular sales use the swing rule.
A partly sold lot is split into a closed row and an open row. These rows get
Sync IDs like `Roth:seed:20251126:1030>`. To fix a 2025 lot's price later,
change it on the lot's open row; the update carries it to the rest.

Each run compares the tracker's open positions to what Schwab says you hold
(Holdings Check sheet in the review). "Tracker has LESS" usually means a 2025
lot is missing.

## Splits and the 2025 carry-forward
- Reverse/forward splits are handled automatically: open lots are converted
  (same total cost, same buy dates). See the Splits sheet in the review.
- The first sync pulls 2025 once (carry_cache.csv) to work out which lots you
  held on Jan 1. Those lots are added or used to correct your hand-entered
  2025 rows; nothing else from 2025 goes in the tracker. Schwab only serves
  about a year back, so older lots stay as you entered them. Re-pull with
  `py sync_trades.py --carry`. Results are on the Carry-forward sheet.

## Schwab history exports (older lots)
For positions older than the API reaches, export the account's history from
schwab.com (History > Export) and save the CSV in the `history` folder, keeping
Schwab's file name (e.g. Roth_XXX946_Transactions_....csv). The update uses
everything before 2026 to work out the lots you held on Jan 1, for the tickers
in that file. You can filter the export to one ticker. Hand-entered 2025 lots
for a ticker an export covers are checked against it.

## Tax-loss harvesting (TLH)
Every run writes TLH_Report.xlsx, and each new tracker version gets a TLH sheet:
realized gains/losses this year (short/long term), open lots at a loss with
their wash-sale status, and every purchase in the last 30 days. Uses your
Schwab lot setting (oldest first), so it lines up with the 1099. Accounts in
NON_TAXABLE_ACCOUNTS (default: Roth) are excluded from gains/losses but still
count for wash sales. Planning tool only, not tax advice.

## Cleanup helpers
- `py highlight_dups.py` marks duplicate hand rows in red.
- `py highlight_dups.py --blanks` marks blank cells to fill (orange) and empty
  placeholder rows to delete (red).

## Back up
`fills_cache.csv` is the master record. Back it up (it's what keeps history
after Schwab stops serving older trades). Keep `.env` and `schwab_token.json`
private.
