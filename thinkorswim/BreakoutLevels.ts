# BreakoutLevels
# Chart study for a DAILY chart. Draws the investing_tools scanner's levels
# so the chart matches the Daily Scanner Watchlist.
#
#   Pivot     highest high of the prior 30 sessions (today excluded)
#   Base low  lowest low of the prior 30 sessions
#   Stop      tighter of base low and pivot - 2 x ATR(14, Wilder's),
#             never closer than 1 ATR below the pivot
#   Target    measured move: pivot + (pivot - base low)
#             (the watchlist target may be lower if an older swing high
#             sits in the way; that check needs the scanner)
#   50 / 200  simple moving averages (uptrend = close > 50 > 200)
#
# The lines roll forward each day, so during the session they show the
# live breakout level. Alert fires once per bar when price crosses the pivot
# (only while a chart with this study is open; use the BreakoutTrigger scan
# for alerts across the whole watchlist).

declare upper;

input pivotLength = 30;
input atrLength = 14;
input barsToShow = 40;
input showMovingAverages = yes;

def pv = Highest(high[1], pivotLength);
def bl = Lowest(low[1], pivotLength);
def atrValue = WildersAverage(TrueRange(high, close, low), atrLength);
def atrStop = pv - 2 * atrValue;
def stopRaw = if bl >= atrStop then bl else atrStop;
def stopLevel = if pv - stopRaw < atrValue then pv - atrValue else stopRaw;
def measured = pv + (pv - bl);

def lastBar = HighestAll(if !IsNaN(close) then BarNumber() else 0);
def show = BarNumber() > lastBar - barsToShow;

plot PivotLine = if show then pv else Double.NaN;
PivotLine.SetDefaultColor(Color.CYAN);
PivotLine.SetLineWeight(2);
PivotLine.SetPaintingStrategy(PaintingStrategy.HORIZONTAL);

plot StopLine = if show then stopLevel else Double.NaN;
StopLine.SetDefaultColor(Color.RED);
StopLine.SetStyle(Curve.SHORT_DASH);
StopLine.SetPaintingStrategy(PaintingStrategy.HORIZONTAL);

plot TargetLine = if show then measured else Double.NaN;
TargetLine.SetDefaultColor(Color.GREEN);
TargetLine.SetStyle(Curve.SHORT_DASH);
TargetLine.SetPaintingStrategy(PaintingStrategy.HORIZONTAL);

plot BaseLowLine = if show then bl else Double.NaN;
BaseLowLine.SetDefaultColor(Color.GRAY);
BaseLowLine.SetStyle(Curve.POINTS);

def sma50 = SimpleMovingAvg(close, 50);
def sma200 = SimpleMovingAvg(close, 200);
plot SMA50line = if showMovingAverages then sma50 else Double.NaN;
SMA50line.SetDefaultColor(Color.YELLOW);
plot SMA200line = if showMovingAverages then sma200 else Double.NaN;
SMA200line.SetDefaultColor(Color.MAGENTA);

def uptrend = close > sma50 and sma50 > sma200;
def pctBelow = (pv - close) / pv * 100;

AddLabel(yes, "Pivot " + AsDollars(pv), Color.CYAN);
AddLabel(yes, "Stop " + AsDollars(stopLevel), Color.RED);
AddLabel(yes, "Target " + AsDollars(measured), Color.GREEN);
AddLabel(close > pv, "ABOVE PIVOT", Color.GREEN);
AddLabel(close <= pv, Round(pctBelow, 1) + "% below pivot", if pctBelow <= 5 then Color.YELLOW else Color.GRAY);
AddLabel(!uptrend, "Not in uptrend", Color.ORANGE);

Alert(close crosses above pv, "Breakout above 30-day pivot", Alert.BAR, Sound.Ding);
