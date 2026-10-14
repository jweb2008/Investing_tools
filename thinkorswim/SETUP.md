# thinkorswim setup

Three pieces that put the scanner's levels inside thinkorswim so you can act
during the session, when breakouts actually happen. Each file in this folder
is plain text: open it in Notepad, copy everything, paste into thinkorswim.

All three rebuild the scanner's rules in thinkScript (thinkorswim cannot read
the scanner's files): pivot = highest high of the prior 30 sessions,
stop = tighter of the 30-day base low and 2x ATR(14), never closer than
1 ATR, target = measured move, uptrend = close > 50 SMA > 200 SMA.

Menu names below are from recent thinkorswim versions and may differ slightly
on yours.

## 1. Chart study: BreakoutLevels.ts

Draws pivot (cyan), stop (red dashes), target (green dashes), base low
(gray dots) and the 50/200 SMAs on a daily chart, with labels showing the
values and how far price is below the pivot.

1. Open a chart, set it to a daily timeframe (for example 1Y:1D).
2. Click the Studies (beaker) icon > Edit studies > Create.
3. Name it `BreakoutLevels`, delete the sample code, paste the file, OK.
4. Add it to the chart and Apply. Save the chart style so it loads by default.

It also has a built-in alert when price crosses above the pivot, but chart
alerts only fire while that chart is open. Use part 3 for the whole list.

## 2. Watchlist column: PivotDistance.ts

Live % from the pivot for every stock in the watchlist. Green background =
above the pivot, yellow = within 2%. Sort by it to see who is closest.

1. In your watchlist, right-click a column header > Customize.
2. Find a Custom column (Custom 1, 2, ...) and click its script icon.
3. Paste the file, name it `Pivot %`, and set the aggregation to **D**.
4. OK, then add the column to the watchlist.

## 3. Breakout alert scan: BreakoutTrigger.ts

Notifies you when any stock on the Daily Scanner watchlist moves above its
pivot while in an uptrend.

1. Keep one watchlist named `Daily Scanner`. Each evening, clear it and
   import the new list (Watchlist menu > Import > Paste symbols from
   clipboard; the evening run already copied them).
2. Scan tab > Stock Hacker. Set **Scan in** to your `Daily Scanner` watchlist.
3. Add filter > Study > click the pencil > thinkScript Editor. Paste the
   file and set the aggregation to **D**. OK.
4. Run the scan, then save it (Save scan query) as `Daily Scanner Breakouts`.
5. From the saved scan's menu, choose the alert option that notifies when
   the scan results change, and pick how you want to be notified
   (Setup > Application Settings > Notifications covers pop-up, mobile
   app, email or text).

Scan alerts refresh on thinkorswim's schedule (a few minutes), not tick by
tick, which is fine for a daily-chart swing setup.

## Good to know

- Volume is not part of the live alert. Early in the day volume cannot reach
  1.5x the daily average yet, so confirm volume yourself or wait for the
  evening scan, which only calls it a breakout on 1.5x+ volume.
- The pivot rolls forward each day. If yesterday made a new 30-day high, the
  live pivot is that high, which can be a little above the Entry printed on
  last night's watchlist. The live line is the one that matters for the
  breakout.
- The chart target is the measured move only. The watchlist target can be
  lower when an older swing high sits in the way.
- The market filter (SPY/QQQ vs 50-day) and earnings dates are on the
  watchlist, not in these studies.
