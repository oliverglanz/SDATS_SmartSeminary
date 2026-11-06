from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO

import pandas as pd
from openpyxl.utils import get_column_letter

from app.config import HELPER_SHEETS


REQUIRED_FIELDS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Subject", ("Subject",)),
    ("Course Number {Crse Num}", ("Course Number {Crse Num}",)),
    ("Course Section {Seq Crse Num}", ("Course Section {Seq Crse Num}",)),
    ("Program {not in Banner}", ("Program {not in Banner}",)),
    ("Campus {Ssasect Campus}", ("Campus {Ssasect Campus}",)),
    ("Instruction Method {Inst Method}", ("Instruction Method {Inst Method}",)),
    ("Enrollment Cap {Max Enrl}", ("Enrollment Cap {Max Enrl}",)),
    ("Meeting Type", ("Meeting Type",)),
    ("Semester Start Date {Soaterm Start Date}", ("Semester Start Date {Soaterm Start Date}",)),
    ("Course Period Start Date {Meet Start Date}", ("Course Start Date {Meet Start Date}", "Intensive Period Start Date {Meet Start Date}")),
    ("Course Period End Date {Meet End Date}", ("Course End Date {Meet End Date}", "Intensive Period End Date {Meet End Date}")),
    ("Semester End Date {Soaterm End Date}", ("Semester End Date {Soaterm End Date}",)),
    ("Weeks", ("Weeks",)),
    ("Instructor Name {Instr Name}", ("Instructor Name {Instr Name}",)),
    ("Instructor Email {Instr Email}", ("Instructor Email {Instr Email}",)),
    ("Instructor ID {Instr ID}", ("Instructor ID {Instr ID}",)),
    ("% Responsibility", ("% Responsibility",)),
    ("load/contract {not in Banner}", ("load/contract {not in Banner}",)),
    ("Reason for Contract {not in Banner}", ("Reason for Contract {not in Banner}",)),
    ("costs per credit {not in Banner}", ("costs per credit {not in Banner}",)),
    ("dept budget {not in Banner}", ("dept budget {not in Banner}",)),
    ("account to be charged {not in Banner}", ("account to be charged {not in Banner}",)),
    ("SEM Department {Scacrse Dept}", ("SEM Department {Scacrse Dept}",)),
)

IDENTITY_HEADERS = ("Subject", "Course Number {Crse Num}", "Course Section {Seq Crse Num}", "CRN")
FACULTY_CREDITS_HEADER = "Faculty Load / Contract Credits {not in Banner}"
LOAD_CONTRACT_HEADER = "load/contract {not in Banner}"
FACULTY_CREDITS_ALLOWED_LOAD_TYPES = {"load", "contract"}
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
        for sheet_name in workbook.sheet_names:
            if sheet_name in HELPER_SHEETS or sheet_name.startswith("_"):
                continue
            frame = pd.read_excel(workbook, sheet_name=sheet_name, dtype=object)
            if frame.empty:
                continue
            checked_sheets += 1
            headers = set(frame.columns)
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
            course_rows = frame.loc[frame[identity_columns].map(_text).ne("").any(axis=1)]
            checked_courses += len(course_rows)
            for index, row in course_rows.iterrows():
                for label, aliases in REQUIRED_FIELDS:
                    source_header = next((alias for alias in aliases if alias in headers), None)
                    if source_header is not None and _text(row.get(source_header)) == "":
                        issue = _issue_base(source_name, sheet_name, index + 2, row)
                        issue["Missing/Wrong Field"] = f"{column_letters[source_header]}_{source_header}"
                        issues.append(issue)

                faculty_credits = row.get(FACULTY_CREDITS_HEADER)
                load_contract = _text(row.get(LOAD_CONTRACT_HEADER)).lower()
                if FACULTY_CREDITS_HEADER not in headers:
                    continue
                faculty_field = f"{column_letters[FACULTY_CREDITS_HEADER]}_{FACULTY_CREDITS_HEADER}"
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
