#!/bin/bash
# ============================================================
#  SALIC Voice -> Slide  |  one-click launcher (macOS)
#  Double-click this file. First run sets everything up
#  (takes a few minutes + a one-time model download).
# ============================================================
cd "$(dirname "$0")" || exit 1
clear
echo "=================================================="
echo "  SALIC Voice -> Slide  (local, private)"
echo "=================================================="
echo

# 1) find python3
if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 is not installed."
  echo "Install it from https://www.python.org/downloads/ (or run: xcode-select --install)"
  echo "Then double-click this file again."
  read -r -p "Press Enter to close."
  exit 1
fi

# 2) create a private virtual environment on first run
if [ ! -d ".venv" ]; then
  echo "[1/3] First-time setup: creating environment…"
  python3 -m venv .venv || { echo "Could not create venv"; read -r -p "Press Enter."; exit 1; }
fi
# shellcheck disable=SC1091
source .venv/bin/activate

# 3) install requirements (quick after the first time)
echo "[2/3] Checking dependencies (first run downloads them, please wait)…"
python -m pip install --upgrade pip -q
python -m pip install -r requirements.txt -q || { echo "Install failed"; read -r -p "Press Enter."; exit 1; }

# 4) launch, then open the browser
echo "[3/3] Starting the app…"
echo
echo "  When it says 'Application startup complete', your browser will open."
echo "  Keep this window open while you use it. Close it (or press Ctrl+C) to stop."
echo
( sleep 4; open "http://localhost:8000" ) &
python -m uvicorn app:app --host 127.0.0.1 --port 8000
