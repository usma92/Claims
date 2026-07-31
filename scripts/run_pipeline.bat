@echo off
:: run_pipeline.bat — Claims data download
:: Registered in Task Scheduler for daily execution.
:: Self-locating: works regardless of where the project folder lives.

setlocal
cd /d "%~dp0.."

set PYTHONUTF8=1

call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Could not activate venv. Run setup_venv.bat first.
    exit /b 1
)

python scripts\download_planner_browser.py
exit /b %errorlevel%
