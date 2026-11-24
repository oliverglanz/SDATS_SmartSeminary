from __future__ import annotations

import os
from pathlib import Path


def _get_env_int(name: str, default: int) -> int:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _get_env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name, "").strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on"}


APP_TITLE = os.getenv("SMART_SEMINARY_APP_TITLE", "Smart Seminary Portal")
DEFAULT_PORT = _get_env_int("SMART_SEMINARY_PORT", 8501)
ENABLE_CLOSE_BUTTON = _get_env_bool("SMART_SEMINARY_ENABLE_CLOSE_BUTTON", True)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEPARTMENT_ALIASES = {"PTDS": "PATH", "PTDI": "PATH", "THCP": "THST", "WMIS": "MSSN", "DSPL": "DSLE"}


def canonical_department(value: object) -> str:
    text = str(value).strip()
    return DEPARTMENT_ALIASES.get(text.upper(), text)


DEFAULT_DONSHEET_PATH = (
    PROJECT_ROOT
    / "0_source_files"
    / "default_DonSheet"
    / "DonSheet_default_empty_v20261020_WithFormulas.xlsx"
)
DEFAULT_SHARISHEET_DIR = Path(
    os.getenv(
        "SMART_SEMINARY_SHARISHEET_DIR",
        str(PROJECT_ROOT.parent / "SmartBudgeting" / "Step1_input_ShariSheets"),
    )
)
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs_app"
SHARED_DONSHEET_LINKS_PATH = PROJECT_ROOT / "0_source_files" / "shared_donsheet_links.json"

HELPER_SHEETS = {"DropDownMenu", "InstructorMap", "_names_lookup", "_courses_lookup"}
ROLE_PASSWORDS = {
    "Karen": os.getenv("SMART_SEMINARY_PASSWORD_KAREN", "neraK"),
    "Shari": os.getenv("SMART_SEMINARY_PASSWORD_SHARI", "irahS"),
}
DEPARTMENT_OPTIONS = [
    "CHIS",
    "DMIN",
    "DSLE",
    "GSEM",
    "MA_Religion",
    "MDivHISP_MAPmENGL_MAPmHISP",
    "MSSN",
    "NTST",
    "OTST",
    "PATH",
    "THST",
]
