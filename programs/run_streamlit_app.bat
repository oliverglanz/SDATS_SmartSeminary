@echo off
set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"
python -m streamlit run app/streamlit_app.py --server.headless false

