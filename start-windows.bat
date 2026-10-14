@echo off
cd /d "%~dp0"
if not exist .venv (
  echo Membuat virtual environment (sekali saja)...
  python -m venv .venv || (echo Python belum terpasang. Install dari https://www.python.org/downloads/ lalu centang "Add Python to PATH". & pause & exit /b 1)
  .venv\Scripts\python -m pip install --upgrade pip
  .venv\Scripts\python -m pip install -r requirements.txt
)
.venv\Scripts\python run.py
pause
