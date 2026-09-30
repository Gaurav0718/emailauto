#!/usr/bin/env bash
# Start both servers: Flask backend (:5001) and Vite frontend (:5173).
# Ctrl-C stops both.
set -e

ROOT="$(cd "$(dirname "$0")" && pwd)"

echo "==> Installing backend deps (first run only)"
python3 -m pip install -q -r "$ROOT/backend/requirements.txt"

echo "==> Installing frontend deps (first run only)"
if [ ! -d "$ROOT/frontend/node_modules" ]; then
  (cd "$ROOT/frontend" && npm install)
fi

echo "==> Starting Flask backend on http://127.0.0.1:5001"
(cd "$ROOT/backend" && python3 app.py) &
BACK=$!

echo "==> Starting Vite frontend on http://localhost:5173"
(cd "$ROOT/frontend" && npm run dev) &
FRONT=$!

trap "echo; echo 'Stopping...'; kill $BACK $FRONT 2>/dev/null" INT TERM
echo
echo "Open http://localhost:5173 in your browser."
wait
