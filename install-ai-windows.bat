@echo off
cd /d "%~dp0"
if not exist .venv (
  echo Run start-windows.bat first.
  pause
  exit /b 1
)
.venv\Scripts\python -m backend.setup_ai
pause
