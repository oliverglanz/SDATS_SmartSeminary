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
    ("Faculty Load / Contract Credits {not in Banner}", ("Faculty Load / Contract Credits {not in Banner}",)),
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
ISSUE_COLUMNS = [
    "Source File",
    "Source Sheet",
    "Source Row",
    "Subject",
    "Course Number",
    "Catalog Title",
    "Teacher",
    "Missing Field",
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
    """Report one record for every required field that is blank on a course row."""
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
                    issue["Missing Field"] = f"Required column is missing: {label}"
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
                        issue["Missing Field"] = f"{column_letters[source_header]}_{source_header}"
                        issues.append(issue)

    return DonSheetControlResult(
        issues_df=pd.DataFrame(issues, columns=ISSUE_COLUMNS),
        checked_courses=checked_courses,
        checked_sheets=checked_sheets,
    )
