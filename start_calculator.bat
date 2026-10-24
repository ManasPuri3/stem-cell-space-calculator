@echo off
REM Starts the calculator at http://localhost:8000 . Close this window to stop it.
cd /d "%~dp0"
set PY=%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe
if not exist "%PY%" set PY=py
echo Pitch page:      http://localhost:8000/pitch.html
echo Full calculator: http://localhost:8000  (close this window to stop)
start "" http://localhost:8000/pitch.html
"%PY%" -m http.server 8000 --bind 127.0.0.1 --directory app
