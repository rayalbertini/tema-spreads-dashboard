@echo off
REM Starts the 5-minute publisher loop on the Bloomberg machine.
set REPO_DATA=C:\spreads\tema-spreads-dashboard-data
cd /d %~dp0
py -3.12 publish.py --repo-dir "%REPO_DATA%" --loop --log-file publisher.log
