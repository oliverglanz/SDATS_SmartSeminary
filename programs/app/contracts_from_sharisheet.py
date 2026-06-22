from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd

from .contracts_from_donsheet import (
    ContractGenerationResult,
    create_zip_from_contracts,
    generate_contract_pdfs,
    resolve_app_paths,
    sanitize_filename,
)
from .sharisheet import normalize_sharisheet_uploads


def generate_contracts_from_sharisheet_uploads(
    uploads: list[tuple[str, bytes]],
    base_dir: Path | None = None,
) -> ContractGenerationResult:
    paths = resolve_app_paths(base_dir)
    work_root = Path(tempfile.mkdtemp(prefix="smartseminary_sharisheet_contracts_"))
    output_dir = work_root / "generated_contracts"
    output_dir.mkdir(parents=True, exist_ok=True)

    result = normalize_sharisheet_uploads(uploads)
    extracted_rows = _contract_rows_for_pdf(result.canonical_df)
    contracts = []

    if not extracted_rows.empty:
        for source_file, source_df in extracted_rows.groupby("Source File", dropna=False):
            source_text = str(source_file or "ShariSheet")
            contracts.extend(
                generate_contract_pdfs(
                    source_df,
                    paths,
                    output_dir,
                    filename_prefix=f"{sanitize_filename(Path(source_text).stem)}__",
                    source_file=source_text,
                )
            )

    zip_path = create_zip_from_contracts(contracts, output_dir / "contracts.zip")
    return ContractGenerationResult(
        output_dir=output_dir,
        zip_path=zip_path,
        contracts=contracts,
        extracted_rows=extracted_rows,
        records_processed=len(extracted_rows),
    )


def _contract_rows_for_pdf(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "load/contract" not in df.columns:
        return pd.DataFrame()

    working = df.copy()
    load_type = working["load/contract"].astype("string").str.strip().str.lower()
    working = working.loc[load_type.eq("contract")].copy()
    if working.empty:
        return pd.DataFrame()

    working["Instructor"] = _text_column(working, "Faculty Name")
    working["ID"] = _text_column(working, "ID#")
    working["Email"] = _text_column(working, "email")
    working["Telephone"] = _text_column(working, "telephone")
    working["Account"] = _text_column(working, "account to be charged")
    working["Remote"] = _text_column(working, "remote employee")
    working["International"] = _text_column(working, "international")
    working["DeptContact"] = _text_column(working, "Dept contact name")
    working["DeptContactID"] = _text_column(working, "Dept contact ID#")
    working["Reason"] = _combined_reason(working)
    working["Year"] = _text_column(working, "Semester").str.extract(r"(\d{4})$", expand=False).fillna("")
    working["CourseRaw"] = _text_column(working, "Course")
    working["Dept"] = _text_column(working, "Subject")
    missing_dept = working["Dept"].eq("")
    if missing_dept.any():
        working.loc[missing_dept, "Dept"] = _text_column(working, "Course").str.extract(r"^([A-Za-z]+)", expand=False).fillna("")
    working["CourseNum"] = _text_column(working, "Course Number")
    missing_num = working["CourseNum"].eq("")
    if missing_num.any():
        working.loc[missing_num, "CourseNum"] = _text_column(working, "Course").str.extract(r"^[A-Za-z]+(\d+)", expand=False).fillna("")
    working["CourseSection"] = _text_column(working, "Course Section")
    working["CourseTitle"] = _text_column(working, "Catalog Title")
    working["Credits"] = _numeric_column(working, "Cr.")
    working["Rate"] = _numeric_column(working, "rate per credit")
    working["DeptBudget"] = _text_column(working, "dept budget")
    working["BeginDate"] = working.get("Begin Date")
    working["EndDate"] = working.get("End Date")
    working["TotalContractHours"] = _numeric_column(working, "total contract hours")
    working["Department"] = _text_column(working, "Department")
    working["Source File"] = _text_column(working, "SourceFile")
    working["Source Sheet"] = "contract-load_courses"
    working["Source Row"] = working.index + 2

    rename_map = {
        "pre-work period start": "prework_start",
        "pre-work period end": "prework_end",
        "pre-work # of weeks": "prework_weeks",
        "pre-work hours/week": "prework_hours_week",
        "pre-work hours/period": "prework_hours_period",
        "intensive period start": "intensive_start",
        "intensive period end": "intensive_end",
        "intensive # of weeks": "intensive_weeks",
        "intensive hours/week": "intensive_hours_week",
        "intensive hours/period": "intensive_hours_period",
        "post-work period start": "postwork_start",
        "post-work period end": "postwork_end",
        "post-work # of weeks": "postwork_weeks",
        "post-work hours/week": "postwork_hours_week",
        "post-work hours/period": "postwork_hours_period",
        "total contract amount": "total contract amount",
    }
    for source, target in rename_map.items():
        working[target] = working[source] if source in working.columns else ""

    return working


def _text_column(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series("", index=df.index, dtype="string")
    return df[column].astype("string").fillna("").str.strip()


def _numeric_column(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(0, index=df.index, dtype="float64")
    return pd.to_numeric(df[column], errors="coerce").fillna(0)


def _combined_reason(df: pd.DataFrame) -> pd.Series:
    parts = []
    for column in ["Program", "Location", "Reason for Contract"]:
        parts.append(_text_column(df, column))
    out = []
    for values in zip(*parts):
        out.append("/".join(value for value in values if value))
    return pd.Series(out, index=df.index, dtype="string")
