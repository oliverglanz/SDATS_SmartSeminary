#!/usr/bin/env sh
SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
cd "$SCRIPT_DIR" || exit 1
python -m streamlit run app/streamlit_app.py --server.headless=false

