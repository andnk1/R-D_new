@echo off
REM Double-click to run the R&D Feasibility app on this computer.
cd /d "%~dp0"
python -m pip install -r requirements.txt --quiet
python -m playwright install chromium
start "" http://localhost:5000
python app.py
