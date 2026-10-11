@echo off
rem Evening scan + Daily Scanner Watchlist. Scheduled weekdays at 5 PM.
cd /d %~dp0
if not exist logs mkdir logs
py -u daily_watchlist.py >> logs\daily_runner.log 2>&1
