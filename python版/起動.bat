@echo off
cd /d "%~dp0"
py -3 app.py
if not errorlevel 1 goto end
python app.py
if errorlevel 1 pause
:end
