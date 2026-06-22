from __future__ import annotations

import re
import warnings
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd

from .models import NormalizationResult, WorkbookSummary


SHEET_NAME = "contract-load_courses"
IDENTITY_COLS = ["Faculty Name", "ID#", "email"]
OPENPYXL_WARNING = "Data Validation extension is not supported and will be removed"

ORG_CODE_MAP = {
    "CHIS": 1100,
    "PATH": 1150,
    "DMIN": 1170,
    "MSSN": 1200,
    "NTST": 1250,
    "OTST": 1300,
    "DSLE": 1320,
    "THST": 1350,
}

TERM_SUFFIX = {
    "Spring": "21",
    "Summer": "31",
    "Fall": "41",
}


@dataclass
class ShariSheetPayload:
    name: str
    data: bytes


def normalize_sharisheet_uploads(uploads: list[tuple[str, bytes]]) -> NormalizationResult:
    frames: list[pd.DataFrame] = []
    row_counts: dict[str, int] = {}
    warnings_out: list[str] = []
    source_names: list[str] = []

    for source_name, source_bytes in uploads:
        source_names.append(source_name)
        try:
            frame = normalize_single_sharisheet(source_name, source_bytes)
            row_counts[source_name] = len(frame)
            frames.append(frame)
        except Exception as exc:
            warnings_out.append(f"{source_name}: {exc}")

    canonical_df = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    summary = WorkbookSummary(
        source_name=", ".join(source_names) if source_names else "ShariSheets",
        sheet_names=[SHEET_NAME],
        data_sheet_names=source_names,
        helper_sheet_names=[],
        row_counts=row_counts,
        warnings=warnings_out,
    )
    return NormalizationResult(canonical_df=canonical_df, summary=summary, warnings=warnings_out)


def normalize_single_sharisheet(source_name: str, source_bytes: bytes) -> pd.DataFrame:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=OPENPYXL_WARNING)
        excel_file = pd.ExcelFile(BytesIO(source_bytes))
        sheet_name = SHEET_NAME if SHEET_NAME in excel_file.sheet_names else excel_file.sheet_names[0]
        df = pd.read_excel(excel_file, sheet_name=sheet_name)

    for col in IDENTITY_COLS:
        if col not in df.columns:
            df[col] = pd.NA

    df = normalize_object_cols(df)
    bad_rows = is_header_or_totals_row(df)
    df.loc[bad_rows, IDENTITY_COLS] = pd.NA
    for col in IDENTITY_COLS:
        df[col] = df[col].ffill()

    filtered = df.loc[~bad_rows].copy()
    filtered = filtered.loc[filtered["Faculty Name"].notna()].copy()
    filtered["Department"] = department_from_filename(source_name)
    filtered["SourceFile"] = source_name
    filtered["Source Workbook"] = source_name

    reshaped = reshape_single_sharisheet(filtered)
    enrich_sharisheet_frame(reshaped)
    return reorder_sharisheet_columns(reshaped)


def normalize_object_cols(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.select_dtypes(include="object").columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
        )
    return df


def is_header_or_totals_row(df: pd.DataFrame) -> pd.Series:
    fac = df["Faculty Name"].astype(str).str.strip()
    is_totals = fac.str.startswith("Totals:", na=False)
    is_section_header = fac.isin(["Contract Teachers", "Load Teachers"])
    is_repeated_header = pd.Series(False, index=df.index)
    if "ID#" in df.columns:
        is_repeated_header = is_repeated_header | df["ID#"].astype(str).str.strip().eq("ID#")
    if "email" in df.columns:
        is_repeated_header = is_repeated_header | df["email"].astype(str).str.strip().eq("email")
    return is_totals | is_section_header | is_repeated_header


def _norm_key(value: object) -> str:
    text = str(value).strip()
    text = re.sub(r"\.\d+$", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.lower()


CANONICAL = {
    _norm_key("Course Section"): "Course Section",
    _norm_key("Catalog Title"): "Catalog Title",
    _norm_key("Mode"): "Mode",
    _norm_key("Program"): "Program",
    _norm_key("Location"): "Location",
    _norm_key("dept budget"): "dept budget",
    _norm_key("rate per credit"): "rate per credit",
    _norm_key("total contract amount"): "total contract amount",
    _norm_key("Begin Date"): "Begin Date",
    _norm_key("End Date"): "End Date",
    _norm_key("pre-work period start"): "pre-work period start",
    _norm_key("pre-work period end"): "pre-work period end",
    _norm_key("pre-work # of weeks"): "pre-work # of weeks",
    _norm_key("pre-work hours/week"): "pre-work hours/week",
    _norm_key("pre-work hours/period"): "pre-work hours/period",
    _norm_key("intensive period start"): "intensive period start",
    _norm_key("intensive period end"): "intensive period end",
    _norm_key("intensive # of weeks"): "intensive # of weeks",
    _norm_key("intensive hours/week"): "intensive hours/week",
    _norm_key("intensive hours/period"): "intensive hours/period",
    _norm_key("post-work period start"): "post-work period start",
    _norm_key("post-work period end"): "post-work period end",
    _norm_key("post-work # of weeks"): "post-work # of weeks",
    _norm_key("post-work hours/week"): "post-work hours/week",
    _norm_key("post-work hours/period"): "post-work hours/period",
    _norm_key("sum of weeks (for check)"): "sum of weeks (for check)",
    _norm_key("total contract hours"): "total contract hours",
    _norm_key("Cr."): "Cr.",
    _norm_key("Total Cr."): "Total Cr.",
    _norm_key("order"): "order",
    _norm_key("load/contract"): "load/contract",
}

TERM_START_RX = re.compile(r"^(Summer|Fall|Spring)\s+(\d{4})\b")


def canonicalize_columns_preserve_suffix(cols) -> dict[object, str]:
    rename = {}
    for col in cols:
        text = str(col).strip()
        suffix = re.search(r"(\.\d+)$", text)
        base_key = _norm_key(text)
        if base_key in CANONICAL:
            rename[col] = CANONICAL[base_key] + (suffix.group(1) if suffix else "")
    return rename


def build_canonical_occurrences(df: pd.DataFrame) -> dict[str, list[str]]:
    cols = list(df.columns)
    pos = {col: idx for idx, col in enumerate(cols)}
    occurrences = {base_key: [] for base_key in CANONICAL}
    for col in cols:
        key = _norm_key(col)
        if key in CANONICAL:
            occurrences[key].append(col)
    for key, values in occurrences.items():
        occurrences[key] = sorted(values, key=lambda col: pos[col])
    return occurrences


def get_all_term_blocks(df: pd.DataFrame) -> list[tuple[int, str, list[str]]]:
    cols = list(df.columns)
    starts = []
    for idx, col in enumerate(cols):
        match = TERM_START_RX.match(str(col))
        if match:
            starts.append((idx, f"{match.group(1)} {match.group(2)}"))
    if not starts:
        raise ValueError("No Summer/Fall/Spring YYYY term columns were found.")
    blocks = []
    for block_idx, (start_idx, label) in enumerate(starts):
        end_idx = starts[block_idx + 1][0] if block_idx + 1 < len(starts) else len(cols)
        blocks.append((block_idx, label, cols[start_idx:end_idx]))
    return blocks


def coalesce_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame(index=df.index)
    for col in pd.unique(df.columns):
        sub = df.loc[:, df.columns == col]
        if sub.shape[1] == 1:
            out[col] = sub.iloc[:, 0]
            continue
        sub2 = sub.copy()
        sub2 = sub2.map(lambda value: pd.NA if isinstance(value, str) and value.strip() == "" else value)
        out[col] = sub2.bfill(axis=1).iloc[:, 0]
    return out


def reshape_single_sharisheet(df: pd.DataFrame) -> pd.DataFrame:
    df = df.drop(columns=[c for c in ["Term 1", "Term 2", "Term 3"] if c in df.columns], errors="ignore")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed:")].copy()
    canon_occ = build_canonical_occurrences(df)

    try:
        identity_cols = list(df.loc[:, "Faculty Name":"Dept contact ID#"].columns)
    except Exception:
        identity_cols = [c for c in df.columns if c in IDENTITY_COLS]
    for col in ["Department", "SourceFile", "Source Workbook"]:
        if col in df.columns and col not in identity_cols:
            identity_cols.append(col)

    def extract_block(cols: list[str], semester_label: str, block_idx: int) -> pd.DataFrame:
        extra_cols = []
        for col_list in canon_occ.values():
            if not col_list:
                continue
            extra_cols.append(col_list[block_idx] if block_idx < len(col_list) else col_list[-1])
        seen = set()
        extra_cols = [col for col in extra_cols if not (col in seen or seen.add(col))]

        block = df[identity_cols + cols + extra_cols].copy()
        cleaned_cols = {c: re.sub(r"^(Summer|Fall|Spring)\s+\d{4}\s*", "", str(c)) for c in cols}
        block = block.rename(columns=cleaned_cols)
        block.columns = [str(c).strip() for c in block.columns]
        block = block.rename(columns=canonicalize_columns_preserve_suffix(block.columns))
        block.columns = [re.sub(r"\.\d+$", "", str(c)) for c in block.columns]
        block = coalesce_duplicate_columns(block)
        block["Semester"] = semester_label
        if "Course" not in block.columns:
            raise KeyError(f"No Course column found after reshaping {semester_label}.")
        return block.loc[block["Course"].notna() & block["Course"].astype(str).str.strip().ne("")].copy()

    blocks = [extract_block(cols, label, block_idx) for block_idx, label, cols in get_all_term_blocks(df)]
    return pd.concat(blocks, ignore_index=True, sort=False)


def enrich_sharisheet_frame(df: pd.DataFrame) -> None:
    df["ORG code"] = df["Department"].astype(str).str.strip().map(ORG_CODE_MAP)
    df["Term Code"] = df["Semester"].apply(semester_to_term_code)
    df["Term Code"] = pd.to_numeric(df["Term Code"], errors="coerce").astype("Int64")
    year = df["Term Code"] // 100
    semester = df["Term Code"] % 100
    df["Fiscal Year"] = pd.NA
    df.loc[semester.isin([31, 41]), "Fiscal Year"] = year[semester.isin([31, 41])] + 1
    df.loc[semester.eq(21), "Fiscal Year"] = year[semester.eq(21)]
    df["Fiscal Year"] = df["Fiscal Year"].astype("Int64")
    df["sourceDept"] = df["SourceFile"].apply(department_from_filename)

    parsed = df["Course"].astype(str).str.extract(r"^([A-Za-z]+)(\d+)")
    df["Subject"] = parsed[0].str.upper()
    df["Course Number"] = parsed[1]
    df["Course Number {Crse Num}"] = parsed[1]


def semester_to_term_code(value: object) -> object:
    if pd.isna(value):
        return pd.NA
    parts = str(value).strip().split()
    if len(parts) != 2:
        return pd.NA
    season, year = parts
    suffix = TERM_SUFFIX.get(season)
    return f"{year}{suffix}" if suffix and year.isdigit() else pd.NA


def department_from_filename(filename: object) -> str:
    if pd.isna(filename):
        return ""
    parts = Path(str(filename)).name.split("_")
    if len(parts) >= 3:
        return parts[1]
    for part in parts:
        if part.isalpha() and part.isupper() and 3 <= len(part) <= 16:
            return part
    return Path(str(filename)).stem


def reorder_sharisheet_columns(df: pd.DataFrame) -> pd.DataFrame:
    desired_order = [
        "Semester",
        "Faculty Name",
        "ID#",
        "email",
        "telephone",
        "status",
        "account to be charged",
        "Reason for Contract",
        "remote employee",
        "international",
        "Dept contact name",
        "Dept contact ID#",
        "Department",
        "ORG code",
        "Course",
        "Course Section",
        "Catalog Title",
        "Mode",
        "Program",
        "Location",
        "load/contract",
        "dept budget",
        "rate per credit",
        "total contract amount",
        "Begin Date",
        "End Date",
        "pre-work period start",
        "pre-work period end",
        "pre-work # of weeks",
        "pre-work hours/week",
        "pre-work hours/period",
        "intensive period start",
        "intensive period end",
        "intensive # of weeks",
        "intensive hours/week",
        "intensive hours/period",
        "post-work period start",
        "post-work period end",
        "post-work # of weeks",
        "post-work hours/week",
        "post-work hours/period",
        "sum of weeks (for check)",
        "total contract hours",
        "SourceFile",
        "Source Workbook",
        "CRN",
        "Cr.",
        "Total Cr.",
        "order",
        "Term Code",
        "Fiscal Year",
        "sourceDept",
        "Subject",
        "Course Number",
    ]
    return df[[col for col in desired_order if col in df.columns] + [col for col in df.columns if col not in desired_order]]


def sharisheet_payloads_from_directory(directory: Path) -> list[tuple[str, bytes]]:
    if not directory.exists():
        return []
    payloads = []
    for path in sorted(directory.glob("*.xlsx")):
        if path.name.startswith("._") or path.name.startswith("~$"):
            continue
        payloads.append((path.name, path.read_bytes()))
    return payloads
