#!/usr/bin/env bash
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "Membuat virtual environment (sekali saja)..."
  python3 -m venv .venv || { echo "Python 3 belum terpasang."; exit 1; }
  .venv/bin/python -m pip install --upgrade pip
  .venv/bin/python -m pip install -r requirements.txt
fi
exec .venv/bin/python run.py
