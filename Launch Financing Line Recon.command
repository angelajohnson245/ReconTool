#!/bin/bash
# Double-click: find the app next to this launcher (or one folder down), check deps, then run Streamlit.

# Launcher directory (absolute path, works no matter where the folder lives)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)" || exit 1

RECON_APP=""
if [[ -f "$SCRIPT_DIR/recon_streamlit_app.py" ]]; then
  RECON_APP="$SCRIPT_DIR/recon_streamlit_app.py"
else
  for candidate in "$SCRIPT_DIR"/*/recon_streamlit_app.py; do
    if [[ -f "$candidate" ]]; then
      RECON_APP="$candidate"
      break
    fi
  done
fi

if [[ -z "$RECON_APP" || ! -f "$RECON_APP" ]]; then
  echo "Could not find recon_streamlit_app.py. Please keep all files in the same folder."
  read -r -p "Press Enter to close..."
  exit 1
fi

RECON_DIR="$(dirname "$RECON_APP")"
cd "$RECON_DIR" || exit 1

# Prefer python3 on PATH; fall back to python.org Framework install (common on Mac).
PY=""
FRAMEWORK_PY="/Library/Frameworks/Python.framework/Versions/3.14/bin/python3"
if command -v python3 >/dev/null 2>&1; then
  PY="$(command -v python3)"
elif [[ -x "$FRAMEWORK_PY" ]]; then
  PY="$FRAMEWORK_PY"
  echo "Using Python from: $PY"
else
  echo "Python 3 was not found (not on PATH and not at the usual python.org location)."
  echo "Install Python 3 from https://www.python.org/downloads/ then double-click this launcher again."
  read -r -p "Press Enter to close..."
  exit 1
fi

CHECK_PY='import importlib.util, sys
mods = ("streamlit", "pandas", "openpyxl")
missing = [m for m in mods if importlib.util.find_spec(m) is None]
print(" ".join(missing))
sys.exit(0)'

echo "Checking dependencies..."
MISSING=$("$PY" -c "$CHECK_PY")

if [ -z "$MISSING" ]; then
  echo "Everything is already installed. Opening Finance Recon Tool..."
else
  echo "Installing missing packages..."
  if ! "$PY" -m pip install --upgrade $MISSING; then
    echo ""
    echo "Package install failed. Check the messages above."
    read -r -p "Press Enter to close..."
    exit 1
  fi
fi

echo ""
echo "Launching Finance Recon Tool..."
echo ""

"$PY" -m streamlit run "$(basename "$RECON_APP")"
exit_code=$?

if [ "$exit_code" -ne 0 ]; then
  echo ""
  echo "The app stopped with an error (code $exit_code). See messages above."
  read -r -p "Press Enter to close..."
fi
exit "$exit_code"
