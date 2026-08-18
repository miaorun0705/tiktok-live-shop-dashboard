#!/bin/bash
# Double-click to launch the TikTok dashboard in your browser.

cd "$(dirname "$0")"

find_python() {
  bundled_python="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
  if [ -x "$bundled_python" ]; then
    if "$bundled_python" -c "import pandas, openpyxl" 2>/dev/null; then
      echo "$bundled_python"; return
    fi
  fi
  for conda_root in "$HOME/miniconda3" "$HOME/anaconda3" "$HOME/opt/anaconda3" "$HOME/opt/miniconda3"; do
    if [ -x "$conda_root/bin/python" ]; then
      if "$conda_root/bin/python" -c "import pandas, openpyxl" 2>/dev/null; then
        echo "$conda_root/bin/python"; return
      fi
    fi
  done
  for venv in "venv" ".venv"; do
    if [ -x "$venv/bin/python" ]; then
      if "$venv/bin/python" -c "import pandas, openpyxl" 2>/dev/null; then
        echo "$venv/bin/python"; return
      fi
    fi
  done
  for py in "python3" "python"; do
    if command -v "$py" &>/dev/null; then
      if "$py" -c "import pandas, openpyxl" 2>/dev/null; then
        echo "$(command -v $py)"; return
      fi
    fi
  done
  echo ""
}

PYTHON=$(find_python)

if [ -z "$PYTHON" ]; then
  echo "ERROR: Missing dependencies. Run:  pip3 install pandas openpyxl"
  read -p "Press Enter to close..."; exit 1
fi

PORT=5050
echo ""
echo "=== TikTok Dashboard ==="
echo "Opening http://localhost:$PORT ..."
echo ""

# Open browser after 1.5 s so Flask is ready
(sleep 1.5 && open "http://localhost:$PORT") &

"$PYTHON" app.py
