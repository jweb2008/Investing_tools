# BreakoutTrigger
# Custom SCAN FILTER (Stock Hacker). Set the filter's aggregation to D.
# Run it on your Daily Scanner watchlist and save it with an alert so
# thinkorswim notifies you when a stock enters the results.
#
# True when the stock is in an uptrend (close > 50 SMA > 200 SMA) and trading
# above the scanner pivot (highest high of the prior 30 sessions).
#
# Volume is deliberately NOT required here: early in the session, volume
# cannot reach 1.5x the daily average yet. Confirm volume yourself, or wait
# for the evening scan, which only calls it a breakout on 1.5x+ volume.

def pv = Highest(high[1], 30);
def sma50 = SimpleMovingAvg(close, 50);
def sma200 = SimpleMovingAvg(close, 200);
plot scan = close > pv and close > sma50 and sma50 > sma200;
