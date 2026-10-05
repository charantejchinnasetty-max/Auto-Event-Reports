@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  python -m venv .venv
  if errorlevel 1 goto failed
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Open http://localhost:8000 in your browser.
".venv\Scripts\python.exe" -m uvicorn main:app --app-dir backend --host 127.0.0.1 --port 8000
pause
exit /b
:failed
echo Setup failed. Check that Python 3.10 or newer is installed and internet is available.
pause
