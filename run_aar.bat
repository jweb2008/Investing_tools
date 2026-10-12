@echo off
rem After Action Review: weekdays at 6:00 PM, an hour after the 5:00 PM watchlist scan.
rem Read-only. Rebuilds aar.xlsx and aar_report.html; output appends to aar_log.txt.
cd /d C:\Users\jweb2\Desktop\Investing
echo ==== %date% %time% >> aar_log.txt
py -u aar.py >> aar_log.txt 2>&1
