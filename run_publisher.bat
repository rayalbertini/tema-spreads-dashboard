@echo off
REM Starts the 5-minute publisher loop on the Bloomberg machine.
REM Edit REPO_DATA to point at your clone of the "data" branch.
set REPO_DATA=C:\spreads\tema-spreads-dashboard-data
cd /d %~dp0
python publish.py --repo-dir "%REPO_DATA%" --loop --log-file publisher.log
