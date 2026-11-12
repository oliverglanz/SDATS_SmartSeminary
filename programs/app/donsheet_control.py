from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from app.config import HELPER_SHEETS
from app.io_donsheet import _red_highlighted_row_indices, normalize_donsheet_headers


REQUIRED_FIELDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Subject", ("Subject",)),
    ("Course Number {Crse Num}", ("Course Number {Crse Num}",)),
    ("Course Section {Seq Crse Num}", ("Course Section {Seq Crse Num}",)),
    ("Program {not in Banner}", ("Program {not in Banner}",)),
    ("Campus {Ssasect Campus}", ("Campus {Ssasect Campus}",)),
    ("Instruction Method {Inst Method}", ("Instruction Method {Inst Method}",)),
    ("Enrollment Cap {Max Enrl}", ("Enrollment Cap {Max Enrl}",)),
    ("Meeting Type", ("Meeting Type", "MeetingType")),
    ("Semester Start Date {Soaterm Start Date}", ("Semester Start Date {Soaterm Start Date}",)),
    ("Semester End Date {Soaterm End Date}", ("Semester End Date {Soaterm End Date}",)),
    ("Weeks", ("Weeks",)),
    ("Instructor Name {Instr Name}", ("Instructor Name {Instr Name}",)),
    ("% Responsibility", ("% Responsibility",)),
    ("load/contract {not in Banner}", ("load/contract {not in Banner}",)),
    ("SEM Department {Scacrse Dept}", ("SEM Department {Scacrse Dept}",)),
)

# These budgeting fields apply only to courses paid by contract.  They remain
# optional for load, continuous, cross-listed, and other course types.
CONTRACT_REQUIRED_FIELDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Reason for Contract {not in Banner}", ("Reason for Contract {not in Banner}",)),
    ("costs per credit {not in Banner}", ("costs per credit {not in Banner}",)),
    ("dept budget {not in Banner}", ("dept budget {not in Banner}",)),
    ("account to be charged {not in Banner}", ("account to be charged {not in Banner}",)),
)

# These details are formula-populated for names selected from the Instructor
# Name dropdown.  Manually entered names that are not in that dropdown are not
# expected to have matching email or ID values yet.
DROPDOWN_INSTRUCTOR_REQUIRED_FIELDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Instructor Email {Instr Email}", ("Instructor Email {Instr Email}",)),
    ("Instructor ID {Instr ID}", ("Instructor ID {Instr ID}",)),
)
INSTRUCTOR_NAME_HEADER = "Instructor Name {Instr Name}"
NOTES_HEADER = "Notes {not in Banner}"
DELETE_CLASS_NOTE = "delete class"
MEETING_TYPE_ALIASES = ("Meeting Type", "MeetingType")
COURSE_PERIOD_DATE_FIELDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Course Period Start Date {Meet Start Date}", ("Course Start Date {Meet Start Date}", "Intensive Period Start Date {Meet Start Date}")),
    ("Course Period End Date {Meet End Date}", ("Course End Date {Meet End Date}", "Intensive Period End Date {Meet End Date}")),
)
PRE_WORK_DATE_FIELDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Pre-work Start Date {not in Banner}", ("Pre-work Start Date {not in Banner}",)),
    ("Pre-work End Date {not in Banner}", ("Pre-work End Date {not in Banner}",)),
)
POST_WORK_DATE_FIELDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Post-work Start Date {not in Banner}", ("Post-work Start Date {not in Banner}",)),
    ("Post-work End Date {not in Banner}", ("Post-work End Date {not in Banner}",)),
)

IDENTITY_HEADERS = ("Subject", "Course Number {Crse Num}", "Course Section {Seq Crse Num}", "CRN")
FACULTY_CREDITS_HEADER = "Faculty Load / Contract Credits {not in Banner}"
LOAD_CONTRACT_HEADER = "load/contract {not in Banner}"
FACULTY_CREDITS_ALLOWED_LOAD_TYPES = {"load", "contract"}
FACULTY_CREDITS_OPTIONAL_LOAD_TYPES = {"na", "n/a", "n / a"}
FACULTY_CREDITS_EMPTY_OR_ZERO_LOAD_TYPES = {"x-list", *FACULTY_CREDITS_OPTIONAL_LOAD_TYPES}
LOAD_CONTRACT_COURSE_KEY_HEADERS = (
    "Subject",
    "Course Number {Crse Num}",
    "Semester Start Date {Soaterm Start Date}",
    INSTRUCTOR_NAME_HEADER,
)
ISSUE_COLUMNS = [
    "Source File",
    "Source Sheet",
    "Source Row",
    "Subject",
    "Course Number",
    "Catalog Title",
    "Teacher",
    "Missing/Wrong Field",
]


@dataclass(frozen=True)
class DonSheetControlResult:
    issues_df: pd.DataFrame
    checked_courses: int
    checked_sheets: int


def _text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none", "<na>"} else text


def _instructor_key(value: object) -> str:
    return " ".join(_text(value).casefold().split())


def _load_contract_course_key(row: pd.Series) -> tuple[str, str, str, str] | None:
    """Return the course identity used to share load/contract assignments."""
    values = tuple(" ".join(_text(row.get(header)).casefold().split()) for header in LOAD_CONTRACT_COURSE_KEY_HEADERS)
    return values if all(values) else None


def _has_load_or_contract(value: object) -> bool:
    normalized = _text(value).casefold()
    return "load" in normalized or "contract" in normalized


def _is_empty_or_zero(value: object) -> bool:
    text = _text(value)
    if not text:
        return True
    try:
        return float(text.replace(",", "")) == 0
    except ValueError:
        return False


def _meeting_type_date_fields(value: object) -> tuple[tuple[str, tuple[str, ...]], ...]:
    """Return the only date fields required for the row's Meeting Type."""
    meeting_type = _text(value).upper()
    if meeting_type == "PRE":
        return PRE_WORK_DATE_FIELDS
    if meeting_type == "POST":
        return POST_WORK_DATE_FIELDS
    # CLAS and ordinary meeting types use the current Course Start/End headers;
    # older DonSheets use the Intensive Period aliases handled above.
    return COURSE_PERIOD_DATE_FIELDS


def _dropdown_instructor_names(workbook: pd.ExcelFile) -> set[str]:
    """Read the names offered by the DonSheet's Instructor Name dropdown."""
    for helper_sheet in ("DropDownMenu", "InstructorMap", "_names_lookup"):
        if helper_sheet not in workbook.sheet_names:
            continue
        frame = pd.read_excel(workbook, sheet_name=helper_sheet, dtype=object, keep_default_na=False)
        candidate_header = INSTRUCTOR_NAME_HEADER if INSTRUCTOR_NAME_HEADER in frame.columns else "Faculty Name"
        if candidate_header in frame.columns:
            return {_instructor_key(value) for value in frame[candidate_header] if _instructor_key(value)}
    return set()


def _issue_base(source_name: str, sheet_name: str, row_number: int | str, row: pd.Series | None = None) -> dict[str, object]:
    row = row if row is not None else pd.Series(dtype="object")
    return {
        "Source File": source_name,
        "Source Sheet": sheet_name,
        "Source Row": row_number,
        "Subject": _text(row.get("Subject")),
        "Course Number": _text(row.get("Course Number {Crse Num}")),
        "Catalog Title": _text(row.get("Catalog Title")),
        "Teacher": _text(row.get("Instructor Name {Instr Name}")),
    }


def analyze_missing_donsheet_fields(payloads: list[tuple[str, bytes]]) -> DonSheetControlResult:
    """Report blank required fields and invalid faculty-credit values on course rows."""
    issues: list[dict[str, object]] = []
    checked_courses = 0
    checked_sheets = 0

    for source_name, source_bytes in payloads:
        workbook = pd.ExcelFile(BytesIO(source_bytes))
        styled_workbook = load_workbook(BytesIO(source_bytes), data_only=True)
        dropdown_instructor_names = _dropdown_instructor_names(workbook)
        for sheet_name in workbook.sheet_names:
            if sheet_name in HELPER_SHEETS or sheet_name.startswith("_"):
                continue
            # Preserve literal values such as NA and N/A.  pandas otherwise
            # converts them to missing values before the control rules can
            # recognize them as valid load/contract categories.
            frame = normalize_donsheet_headers(
                pd.read_excel(workbook, sheet_name=sheet_name, dtype=object, keep_default_na=False)
            )
            if frame.empty:
                continue
            checked_sheets += 1
            headers = set(frame.columns)
            red_rows = (
                _red_highlighted_row_indices(styled_workbook[sheet_name], len(frame.columns))
                if sheet_name in styled_workbook.sheetnames
                else set()
            )
            column_letters = {
                header: get_column_letter(index)
                for index, header in enumerate(frame.columns, start=1)
            }
            for label, aliases in REQUIRED_FIELDS:
                if not any(alias in headers for alias in aliases):
                    issue = _issue_base(source_name, sheet_name, "")
                    issue["Missing/Wrong Field"] = f"Required column is missing: {label}"
                    issues.append(issue)
            if FACULTY_CREDITS_HEADER not in headers:
                issue = _issue_base(source_name, sheet_name, "")
                issue["Missing/Wrong Field"] = f"Required column is missing: {FACULTY_CREDITS_HEADER}"
                issues.append(issue)

            identity_columns = [header for header in IDENTITY_HEADERS if header in headers]
            if not identity_columns:
                continue
            course_rows = frame.loc[
                frame[identity_columns].map(_text).ne("").any(axis=1) & ~frame.index.isin(red_rows)
            ]
            if NOTES_HEADER in headers:
                course_rows = course_rows.loc[
                    course_rows[NOTES_HEADER].map(_text).str.casefold().ne(DELETE_CLASS_NOTE)
                ]
            checked_courses += len(course_rows)

            meeting_type_header = next((header for header in MEETING_TYPE_ALIASES if header in headers), None)
            if meeting_type_header is not None:
                required_date_fields = {
                    field
                    for _, course_row in course_rows.iterrows()
                    for field in _meeting_type_date_fields(course_row.get(meeting_type_header))
                }
                for label, aliases in required_date_fields:
                    if not any(alias in headers for alias in aliases):
                        issue = _issue_base(source_name, sheet_name, "")
                        issue["Missing/Wrong Field"] = f"Required column is missing: {label}"
                        issues.append(issue)

            assigned_load_contract_keys = {
                key
                for _, course_row in course_rows.iterrows()
                if _has_load_or_contract(course_row.get(LOAD_CONTRACT_HEADER))
                if (key := _load_contract_course_key(course_row)) is not None
            }

            has_contract_course = (
                LOAD_CONTRACT_HEADER in headers
                and course_rows[LOAD_CONTRACT_HEADER].map(_text).str.lower().eq("contract").any()
            )
            if has_contract_course:
                for label, aliases in CONTRACT_REQUIRED_FIELDS:
                    if not any(alias in headers for alias in aliases):
                        issue = _issue_base(source_name, sheet_name, "")
                        issue["Missing/Wrong Field"] = f"Required column is missing: {label}"
                        issues.append(issue)

            has_dropdown_instructor = (
                INSTRUCTOR_NAME_HEADER in headers
                and course_rows[INSTRUCTOR_NAME_HEADER].map(_instructor_key).isin(dropdown_instructor_names).any()
            )
            if has_dropdown_instructor:
                for label, aliases in DROPDOWN_INSTRUCTOR_REQUIRED_FIELDS:
                    if not any(alias in headers for alias in aliases):
                        issue = _issue_base(source_name, sheet_name, "")
                        issue["Missing/Wrong Field"] = f"Required column is missing: {label}"
                        issues.append(issue)

            for index, row in course_rows.iterrows():
                load_contract = _text(row.get(LOAD_CONTRACT_HEADER)).lower()
                for label, aliases in REQUIRED_FIELDS:
                    source_header = next((alias for alias in aliases if alias in headers), None)
                    if source_header is not None and _text(row.get(source_header)) == "":
                        if (
                            source_header == LOAD_CONTRACT_HEADER
                            and _load_contract_course_key(row) in assigned_load_contract_keys
                        ):
                            continue
                        if (
                            source_header == "Instruction Method {Inst Method}"
                            and load_contract in FACULTY_CREDITS_OPTIONAL_LOAD_TYPES
                        ):
                            continue
                        issue = _issue_base(source_name, sheet_name, index + 2, row)
                        issue["Missing/Wrong Field"] = f"{column_letters[source_header]}_{source_header}"
                        issues.append(issue)

                if meeting_type_header is not None:
                    for _, aliases in _meeting_type_date_fields(row.get(meeting_type_header)):
                        source_header = next((alias for alias in aliases if alias in headers), None)
                        if source_header is not None and _text(row.get(source_header)) == "":
                            issue = _issue_base(source_name, sheet_name, index + 2, row)
                            issue["Missing/Wrong Field"] = f"{column_letters[source_header]}_{source_header}"
                            issues.append(issue)

                selected_instructor = _instructor_key(row.get(INSTRUCTOR_NAME_HEADER)) in dropdown_instructor_names
                if selected_instructor:
                    for _, aliases in DROPDOWN_INSTRUCTOR_REQUIRED_FIELDS:
                        source_header = next((alias for alias in aliases if alias in headers), None)
                        if source_header is not None and _text(row.get(source_header)) == "":
                            issue = _issue_base(source_name, sheet_name, index + 2, row)
                            issue["Missing/Wrong Field"] = f"{column_letters[source_header]}_{source_header}"
                            issues.append(issue)

                if load_contract == "contract":
                    for _, aliases in CONTRACT_REQUIRED_FIELDS:
                        source_header = next((alias for alias in aliases if alias in headers), None)
                        if source_header is not None and _text(row.get(source_header)) == "":
                            issue = _issue_base(source_name, sheet_name, index + 2, row)
                            issue["Missing/Wrong Field"] = f"{column_letters[source_header]}_{source_header}"
                            issues.append(issue)

                faculty_credits = row.get(FACULTY_CREDITS_HEADER)
                if FACULTY_CREDITS_HEADER not in headers:
                    continue
                faculty_field = f"{column_letters[FACULTY_CREDITS_HEADER]}_{FACULTY_CREDITS_HEADER}"
                if load_contract in FACULTY_CREDITS_EMPTY_OR_ZERO_LOAD_TYPES:
                    if not _is_empty_or_zero(faculty_credits):
                        issue = _issue_base(source_name, sheet_name, index + 2, row)
                        issue["Missing/Wrong Field"] = (
                            f"{faculty_field} must be empty or 0 when "
                            f"{column_letters.get(LOAD_CONTRACT_HEADER, '?')}_{LOAD_CONTRACT_HEADER} is '{load_contract}'"
                        )
                        issues.append(issue)
                    continue
                if load_contract in FACULTY_CREDITS_ALLOWED_LOAD_TYPES:
                    if _text(faculty_credits) == "":
                        issue = _issue_base(source_name, sheet_name, index + 2, row)
                        issue["Missing/Wrong Field"] = faculty_field
                        issues.append(issue)
                elif _text(faculty_credits) != "":
                    issue = _issue_base(source_name, sheet_name, index + 2, row)
                    issue["Missing/Wrong Field"] = (
                        f"{faculty_field} must be empty when "
                        f"{column_letters.get(LOAD_CONTRACT_HEADER, '?')}_{LOAD_CONTRACT_HEADER} is '{load_contract or 'blank'}'"
                    )
                    issues.append(issue)

    return DonSheetControlResult(
        issues_df=pd.DataFrame(issues, columns=ISSUE_COLUMNS),
        checked_courses=checked_courses,
        checked_sheets=checked_sheets,
    )
