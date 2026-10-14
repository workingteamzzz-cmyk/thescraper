#!/usr/bin/env bash
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "First run: setting up, one time only..."
  python3 -m venv .venv || { echo "Python 3 is not installed."; exit 1; }
  .venv/bin/python -m pip install --upgrade pip
fi
.venv/bin/python -m pip install -q -r requirements.txt
[ -f .venv/.ai_tried ] || .venv/bin/python -m backend.setup_ai
exec .venv/bin/python run.py
