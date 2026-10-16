@echo off
rem After Action Review: weekdays at 6:00 PM, an hour after the 5:00 PM watchlist scan.
rem Read-only. Rebuilds aar.xlsx and aar_report.html; everything shown is also appended to aar_log.txt.
title After Action Review - running, leave this window open
cd /d C:\Users\jweb2\Desktop\Investing
echo After Action Review is running. Leave this window open; it can take a few minutes.
echo.
py -u aar.py --log aar_log.txt
if errorlevel 1 (
  title After Action Review - FAILED
  echo.
  echo AAR stopped with an error ^(shown above, also in aar_log.txt^).
  pause
) else (
  title After Action Review - done
  echo.
  echo Finished. Open aar_report.html to see the dashboard.
  timeout /t 30
)
