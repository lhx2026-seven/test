@echo off
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 goto fail
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto fail
.venv\Scripts\python.exe app.py
if errorlevel 1 goto fail
exit /b
:fail
 echo Setup failed. Please install Python 3.10 or later and check your Internet connection.
pause
