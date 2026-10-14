@echo off
cd /d "%~dp0"
title Social Scraper
if not exist .venv (
  echo First run: setting up (one time only, 1-2 minutes)...
  python -m venv .venv || (echo Python is not installed. Get it from https://www.python.org/downloads/ and tick "Add Python to PATH". & pause & exit /b 1)
  .venv\Scripts\python -m pip install --upgrade pip
)
.venv\Scripts\python -m pip install -q -r requirements.txt
.venv\Scripts\python run.py
pause
