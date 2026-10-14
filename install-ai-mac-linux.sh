#!/usr/bin/env bash
cd "$(dirname "$0")"
if [ ! -d .venv ]; then echo "Run start-mac.command (Mac) or ./start-mac-linux.sh (Linux) first."; exit 1; fi
.venv/bin/python -m backend.setup_ai
