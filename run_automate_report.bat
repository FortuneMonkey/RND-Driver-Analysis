@echo off
REM Runs the gtrack.id monthly export using the correct Python environment.
REM Task Scheduler should point at THIS file, not python.exe directly.

cd /d "C:\Users\saputra\Projects\driver_analysis\code"
"c:\Users\saputra\Projects\global\Scripts\python.exe" automate_report.py

REM Also run the combine step, if you're using combine_exports.py:
REM "c:\Users\saputra\Projects\global\Scripts\python.exe" combine_exports.py
