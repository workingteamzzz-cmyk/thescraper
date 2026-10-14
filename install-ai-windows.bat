@echo off
cd /d "%~dp0"
if not exist .venv (
  echo Run start-windows.bat once first.
  pause
  exit /b 1
)
echo Installing AI analysis packages (this can take several minutes)...
.venv\Scripts\python -m pip install --index-url https://download.pytorch.org/whl/cpu torch || .venv\Scripts\python -m pip install torch
.venv\Scripts\python -m pip install -r requirements-ai.txt
echo.
echo Done. Restart the app and choose "AI model" under Analysis.
pause
