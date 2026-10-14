#!/usr/bin/env bash
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "First run: setting up (one time only, 1-2 minutes)..."
  python3 -m venv .venv || { echo "Python 3 is not installed."; exit 1; }
  .venv/bin/python -m pip install --upgrade pip
fi
.venv/bin/python -m pip install -q -r requirements.txt
exec .venv/bin/python run.py
