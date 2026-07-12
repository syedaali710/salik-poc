#!/bin/bash
# ============================================================
#  SALIC AI Insights  |  one-click launcher (macOS)
#  Double-click this file. First run sets everything up
#  (takes a few minutes + a one-time model download).
# ============================================================
cd "$(dirname "$0")" || exit 1
clear
echo "=================================================="
echo "  SALIC AI Insights  (uv + FastAPI)"
echo "=================================================="
echo

# 1) find uv (preferred) or install hint
if ! command -v uv >/dev/null 2>&1; then
  echo "uv is not installed."
  echo "Install it from https://docs.astral.sh/uv/getting-started/installation/"
  echo "  curl -LsSf https://astral.sh/uv/install.sh | sh"
  echo "Then double-click this file again."
  read -r -p "Press Enter to close."
  exit 1
fi

# 2) sync dependencies into .venv (idempotent)
echo "[1/2] Syncing dependencies with uv…"
uv sync || { echo "uv sync failed"; read -r -p "Press Enter."; exit 1; }

# 3) launch, then open the browser
echo "[2/2] Starting the app…"
echo
echo "  When it says 'Application startup complete', your browser will open."
echo "  API docs: http://localhost:8000/docs"
echo "  Keep this window open while you use it. Close it (or press Ctrl+C) to stop."
echo
( sleep 4; open "http://localhost:8000" ) &
uv run uvicorn app:app --host 127.0.0.1 --port 8000
