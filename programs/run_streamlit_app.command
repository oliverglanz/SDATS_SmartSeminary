#!/bin/zsh
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
python -m streamlit run app/streamlit_app.py --server.headless=false
