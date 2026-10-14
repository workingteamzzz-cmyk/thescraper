#!/usr/bin/env bash
cd "$(dirname "$0")"
if [ ! -d .venv ]; then echo "Run ./start-mac-linux.sh once first."; exit 1; fi
echo "Installing AI analysis packages (this can take several minutes)..."
if [ "$(uname)" = "Linux" ]; then
  .venv/bin/python -m pip install --index-url https://download.pytorch.org/whl/cpu torch || .venv/bin/python -m pip install torch
fi
.venv/bin/python -m pip install -r requirements-ai.txt
echo "Done. Restart the app and choose \"AI model\" under Analysis."
