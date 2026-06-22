# Local App Setup

## What works across systems

The Python app code in `programs/app` is written with `pathlib` and relative project paths, so it is suitable for:

- macOS
- Linux
- Windows

The main app command is the same on every system:

```bash
python -m streamlit run app/streamlit_app.py --server.headless=false
```

On Windows, Streamlit expects the equivalent flag form:

```bat
python -m streamlit run app/streamlit_app.py --server.headless false
```

---

## Install dependencies

From the `programs` folder:

```bash
pip install -r requirements_local_app.txt
```

---

## Start the app

### macOS

Use:

- `run_streamlit_app.command`

Or run manually:

```bash
python -m streamlit run app/streamlit_app.py --server.headless=false
```

### Linux

Use:

- `run_streamlit_app.sh`

Make it executable once if needed:

```bash
chmod +x run_streamlit_app.sh
```

Then run:

```bash
./run_streamlit_app.sh
```

### Windows

Use:

- `run_streamlit_app.bat`

Or run manually in Command Prompt:

```bat
python -m streamlit run app/streamlit_app.py --server.headless false
```

---

## Open the app

After startup, open:

- [http://localhost:8501](http://localhost:8501)

---

## Notes

- `python` must be available on the machine PATH.
- If `python` does not work on Windows, try `py` instead.
- The bundled default DonSheet is loaded from `0_source_files/default_DonSheet/DonSheet_default_empty_v20260402.xlsx` relative to the project root, so the same repo layout should be preserved on each machine.
- Packaging into a standalone desktop executable is still a later step. Right now this is cross-platform for Python-based local execution.
- For Linux server deployment with Docker, see `../DEPLOYMENT.md`.
