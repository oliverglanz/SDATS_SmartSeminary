from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


LOAD_CONTRACT_COLUMN = "load/contract"
INSTRUCTOR_COLUMN = "Faculty Name"
CLASS_KEY_COLUMNS = ["Subject", "Course Number", INSTRUCTOR_COLUMN]
SHARED_CLASS_FIELDS = [
    "Faculty Load / Contract Credits {not in Banner}",
    "total contract hours {not in Banner}",
    "costs per credit {not in Banner}",
]


@dataclass(frozen=True)
class BudgetReadinessResult:
    issues_df: pd.DataFrame
    checked_rows: int
    issue_counts: dict[str, int]


def _text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "none", "<na>"} else text


def _issue_row(row: pd.Series, issue: str, field: str, value: object = "") -> dict[str, object]:
    return {
        "Issue": issue,
        "Department": _text(row.get("Department")),
        "Source Sheet": _text(row.get("Source Sheet")),
        "Subject": _text(row.get("Subject")),
        "Course Number": _text(row.get("Course Number")),
        "Course Section": _text(row.get("Course Section")),
        "Instructor": _text(row.get(INSTRUCTOR_COLUMN)),
        "Field": field,
        "Value": _text(value),
    }


def analyze_budget_readiness(canonical_df: pd.DataFrame) -> BudgetReadinessResult:
    """Return row-level issues that should be corrected before budget reporting."""
    issues: list[dict[str, object]] = []
    if canonical_df.empty:
        return BudgetReadinessResult(
            issues_df=pd.DataFrame(columns=["Issue", "Department", "Source Sheet", "Subject", "Course Number", "Course Section", "Instructor", "Field", "Value"]),
            checked_rows=0,
            issue_counts={},
        )

    working = canonical_df.copy()
    for column in [*CLASS_KEY_COLUMNS, "Course Section", "Department", "Source Sheet", LOAD_CONTRACT_COLUMN]:
        if column not in working.columns:
            working[column] = pd.NA

    load_contract_missing = working[LOAD_CONTRACT_COLUMN].map(_text).eq("")
    for _, row in working.loc[load_contract_missing].iterrows():
        issues.append(_issue_row(row, "Missing load/contract information", LOAD_CONTRACT_COLUMN))

    instructor_missing = working[INSTRUCTOR_COLUMN].map(_text).eq("")
    for _, row in working.loc[instructor_missing].iterrows():
        issues.append(_issue_row(row, "Missing instructor", INSTRUCTOR_COLUMN))

    available_shared_fields = [field for field in SHARED_CLASS_FIELDS if field in working.columns]
    for field in SHARED_CLASS_FIELDS:
        if field not in working.columns:
            issues.append(
                {
                    "Issue": "Required field is missing from DonSheet",
                    "Department": "",
                    "Source Sheet": "",
                    "Subject": "",
                    "Course Number": "",
                    "Course Section": "",
                    "Instructor": "",
                    "Field": field,
                    "Value": "",
                }
            )

    class_rows = working.loc[
        working["Subject"].map(_text).ne("")
        & working["Course Number"].map(_text).ne("")
        & working[INSTRUCTOR_COLUMN].map(_text).ne("")
    ].copy()
    if available_shared_fields and not class_rows.empty:
        for _, group in class_rows.groupby(CLASS_KEY_COLUMNS, dropna=False, sort=False):
            if len(group) < 2:
                continue
            for field in available_shared_fields:
                filled = group[field].map(_text).ne("")
                if int(filled.sum()) <= 1:
                    continue
                for _, row in group.loc[filled].iterrows():
                    issues.append(
                        _issue_row(
                            row,
                            "Shared class field is filled more than once",
                            field,
                            row[field],
                        )
                    )

    issue_columns = ["Issue", "Department", "Source Sheet", "Subject", "Course Number", "Course Section", "Instructor", "Field", "Value"]
    issues_df = pd.DataFrame(issues, columns=issue_columns)
    issue_counts = issues_df["Issue"].value_counts().to_dict() if not issues_df.empty else {}
    return BudgetReadinessResult(
        issues_df=issues_df,
        checked_rows=len(working),
        issue_counts=issue_counts,
    )
