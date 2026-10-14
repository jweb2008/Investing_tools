@echo off
rem Weekly backtest: refresh cached history, rerun, open the report.
cd /d %~dp0
if not exist logs mkdir logs
py -u backtest.py --download --open >> logs\backtest.log 2>&1
