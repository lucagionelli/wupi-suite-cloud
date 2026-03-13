#!/bin/zsh
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

# Create venv locally (project folder)
if [ ! -d ".venv" ]; then
  /usr/bin/python3 -m venv ".venv"
fi

PY="$DIR/.venv/bin/python"

# Ensure pip exists + upgrade
"$PY" -m ensurepip --upgrade >/dev/null 2>&1 || true
"$PY" -m pip install --upgrade pip >/dev/null

# Install deps
"$PY" -m pip install -r "$DIR/src/requirements.txt" >/dev/null

# If port busy, find next
pick_port() {
  /usr/bin/python3 - <<'PY'
import socket
for p in [8501,8502,8503,8504,8505,8506,8507,8508,8509,8510]:
    s=socket.socket()
    try:
        s.bind(("127.0.0.1", p))
        s.close()
        print(p); break
    except OSError:
        try: s.close()
        except: pass
else:
    print(0)
PY
}

PORT="$(pick_port)"
if [ "$PORT" = "0" ]; then
  echo "Nessuna porta libera trovata."
  exit 1
fi

"$PY" -m streamlit run "$DIR/src/app.py" --server.address 127.0.0.1 --server.port "$PORT" --browser.gatherUsageStats false
