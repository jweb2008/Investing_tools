# PivotDistance
# Custom WATCHLIST COLUMN. Set the column's aggregation to D (daily).
# Shows how far the stock is from the scanner pivot (highest high of the
# prior 30 sessions), live during the session:
#   green background  above the pivot (breaking out)
#   yellow            within 2% below the pivot (close to triggering)
#   gray              further below
# Sort the watchlist by this column to see who is closest to breaking out.

def pv = Highest(high[1], 30);
plot PctToPivot = Round((close - pv) / pv * 100, 1);
PctToPivot.AssignValueColor(if close > pv then Color.WHITE
                            else if close >= pv * 0.98 then Color.YELLOW
                            else Color.LIGHT_GRAY);
AssignBackgroundColor(if close > pv then Color.DARK_GREEN else Color.CURRENT);
