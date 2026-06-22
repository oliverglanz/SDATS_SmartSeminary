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


APP_TITLE = os.getenv("SMART_SEMINARY_APP_TITLE", "Smart Seminary Demo")
DEFAULT_PORT = _get_env_int("SMART_SEMINARY_PORT", 8501)
ENABLE_CLOSE_BUTTON = _get_env_bool("SMART_SEMINARY_ENABLE_CLOSE_BUTTON", True)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DONSHEET_PATH = PROJECT_ROOT / "0_source_files" / "default_DonSheet" / "DonSheet_default_empty_v20260402.xlsx"
DEFAULT_SHARISHEET_DIR = Path(
    os.getenv(
        "SMART_SEMINARY_SHARISHEET_DIR",
        str(PROJECT_ROOT.parent / "SmartBudgeting" / "Step1_input_ShariSheets"),
    )
)
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs_app"

HELPER_SHEETS = {"DropDownMenu", "InstructorMap"}
ROLE_PASSWORDS = {
    "Karen": os.getenv("SMART_SEMINARY_PASSWORD_KAREN", "neraK"),
    "Shari": os.getenv("SMART_SEMINARY_PASSWORD_SHARI", "irahS"),
}
DEPARTMENT_OPTIONS = [
    "OTST",
    "NTST",
    "THST",
    "MSSN",
    "CHIS",
    "DSLE",
    "GSEM",
    "PATH",
    "MA_Religion",
    "DMIN",
    "MDivHISP_MAPmENGL_MAPmHISP",
]
