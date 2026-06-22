from __future__ import annotations

import re
from collections.abc import Sequence

import pandas as pd

from .io_donsheet import load_workbook, read_department_sheets, read_helper_sheet
from .models import NormalizationResult, WorkbookSummary


COLUMN_MAP = {
    "Subject": "Subject",
    "Course Number {Crse Num}": "Course Number",
    "Course Section {Seq Crse Num}": "Course Section",
    "Program {not in Banner}": "Program",
    "Catalog Title": "Catalog Title",
    "CRN": "CRN",
    "Instruction Method {Inst Method}": "Mode Raw",
    "Campus {Ssasect Campus}": "Campus Code",
    "Instructor Name {Instr Name}": "Faculty Name",
    "Instructor Email {Instr Email}": "email",
    "Instructor ID {Instr ID}": "ID#",
    "load/contract {not in Banner}": "load/contract",
    "costs per credit {not in Banner}": "rate per credit",
    "total costs {not in Banner}": "total contract amount",
    "account to be charged {not in Banner}": "account to be charged",
    "Credits {Sect Crs}": "Cr.",
    "Pre-work Start Date {not in Banner}": "pre-work period start",
    "Pre-work End Date {not in Banner}": "pre-work period end",
    "Course Start Date {Meet Start Date}": "Begin Date",
    "Course End Date {Meet End Date}": "End Date",
    "Post-work Start Date {not in Banner}": "post-work period start",
    "Post-work End Date {not in Banner}": "post-work period end",
    "Semester {not in Banner}": "Semester",
    "Year {not in Banner}": "Calendar Year",
    "SEM Department {Scacrse Dept}": "SEM Department",
    "Campus Restriction {Camp Restr CC}": "Location",
}


MODE_MAP = {
    "ASYNC": "ASYNC",
    "SYNC": "SYNC",
    "INPERSON": "In-Person",
    "INTEN": "Intensive",
    "BLENDED": "Blended",
    "WEB": "Web",
}

TERM_SEASON_MAP = {
    "2": "Spring",
    "3": "Summer",
    "4": "Fall",
}


def _clean_object_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for column in out.columns:
        if pd.api.types.is_object_dtype(out[column]) or pd.api.types.is_string_dtype(out[column]):
            out[column] = (
                out[column]
                .replace({pd.NA: None})
                .astype("string")
                .str.strip()
                .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
            )
    return out


def _normalize_mode(value: object) -> str | pd.NA:
    if pd.isna(value):
        return pd.NA
    text = re.sub(r"[^A-Za-z]", "", str(value)).upper()
    return MODE_MAP.get(text, str(value).strip())


def _parse_donsheet_dates(values: pd.Series) -> pd.Series:
    return pd.to_datetime(values, format="%d-%b-%y", errors="coerce")


def _parse_term_code(values: pd.Series) -> tuple[pd.Series, pd.Series, pd.Series]:
    text = values.astype("string").str.strip()
    year = pd.to_numeric(text.str.slice(0, 4), errors="coerce").astype("Int64")
    season_code = text.str.slice(4, 5)
    season = season_code.map(TERM_SEASON_MAP).astype("string")
    label = season.where(year.isna(), season + " " + year.astype("string"))
    return year, season, label


def _numeric_string(values: pd.Series) -> pd.Series:
    return values.astype("string").str.strip().str.replace(r"\.0$", "", regex=True)


def _looks_like_shifted_contract_layout(df: pd.DataFrame) -> bool:
    if df.empty:
        return False
    required = [
        "Instructor Name {Instr Name}",
        "Instructor Email {Instr Email}",
        "Instructor ID {Instr ID}",
        "Remote Employee {not in Banner}",
        "Catalog Title",
        "Section Title",
        "Course Number {Crse Num}",
        "Course Section {Seq Crse Num}",
        "Program {not in Banner}",
        "Reason for Contract {not in Banner}",
        "dept budget {not in Banner}",
        "total costs {not in Banner}",
        "Room {Meet Room}",
        "Semester {not in Banner}",
    ]
    if any(column not in df.columns for column in required):
        return False

    instructor_name = _numeric_string(df["Instructor Name {Instr Name}"].head(5))
    instructor_email = _numeric_string(df["Instructor Email {Instr Email}"].head(5))
    instructor_id = _numeric_string(df["Instructor ID {Instr ID}"].head(5))
    remote_id = _numeric_string(df["Remote Employee {not in Banner}"].head(5))
    room = _numeric_string(df["Room {Meet Room}"].head(5))
    reason = _numeric_string(df["Reason for Contract {not in Banner}"].head(5)).str.lower()

    return bool(
        instructor_name.str.contains("Seminary Building", na=False).any()
        and instructor_email.str.contains(",", na=False).any()
        and instructor_id.str.contains("@", na=False).any()
        and remote_id.str.fullmatch(r"\d+", na=False).any()
        and room.isin(["Spring", "Summer", "Fall"]).any()
        and reason.isin(["load", "contract"]).any()
    )


def _apply_shifted_contract_layout(df: pd.DataFrame) -> pd.DataFrame:
    working = df.copy()
    subject_fallback = _string_series(working.get("Course Number {Crse Num}"), working.index)
    course_number_fallback = _numeric_string(_string_series(working.get("Course Section {Seq Crse Num}"), working.index))
    section_fallback = _numeric_string(_string_series(working.get("Program {not in Banner}"), working.index)).str.zfill(3)

    existing_subject = _string_series(working.get("Subject"), working.index)
    working["Subject"] = existing_subject.where(existing_subject.notna(), subject_fallback)
    working["Course Number {Crse Num}"] = course_number_fallback.where(course_number_fallback.notna(), _string_series(working.get("Course Number {Crse Num}"), working.index))
    working["Course Section {Seq Crse Num}"] = section_fallback.where(section_fallback.notna(), _string_series(working.get("Course Section {Seq Crse Num}"), working.index))
    working["Program {not in Banner}"] = _string_series(working.get("Catalog Title"), working.index)
    working["Catalog Title"] = _string_series(working.get("Section Title"), working.index)

    fallback_name = _string_series(working.get("Instructor Email {Instr Email}"), working.index)
    fallback_email = _string_series(working.get("Instructor ID {Instr ID}"), working.index)
    fallback_id = _numeric_string(_string_series(working.get("Remote Employee {not in Banner}"), working.index))
    working["Instructor Name {Instr Name}"] = fallback_name.where(fallback_name.str.contains(",", na=False), _string_series(working.get("Instructor Name {Instr Name}"), working.index))
    working["Instructor Email {Instr Email}"] = fallback_email.where(fallback_email.str.contains("@", na=False), _string_series(working.get("Instructor Email {Instr Email}"), working.index))
    working["Instructor ID {Instr ID}"] = fallback_id.where(fallback_id.notna(), _string_series(working.get("Instructor ID {Instr ID}"), working.index))

    load_contract = _string_series(working.get("Reason for Contract {not in Banner}"), working.index).str.lower()
    load_contract = load_contract.where(load_contract.isin(["load", "contract"]), _string_series(working.get("Enrollment Cap {Max Enrl}"), working.index).str.lower())
    working["load/contract {not in Banner}"] = load_contract
    working["costs per credit {not in Banner}"] = pd.to_numeric(working.get("total costs {not in Banner}"), errors="coerce")
    working["total costs {not in Banner}"] = pd.to_numeric(working.get("dept budget {not in Banner}"), errors="coerce")

    account_raw = _string_series(working.get("SEM Department {Scacrse Dept}"), working.index)
    working["account to be charged {not in Banner}"] = account_raw.where(account_raw.str.contains(":", na=False), _string_series(working.get("account to be charged {not in Banner}"), working.index))
    working["Campus Restriction {Camp Restr CC}"] = _string_series(working.get("Fee Amount {Fees Amt}"), working.index)
    working["SEM Department {Scacrse Dept}"] = _string_series(working.get("Fee Amount {Fees Amt}"), working.index)

    season = _string_series(working.get("Room {Meet Room}"), working.index)
    year = _numeric_string(_string_series(working.get("Semester {not in Banner}"), working.index))
    working["Semester {not in Banner}"] = season
    working["Year {not in Banner}"] = year

    mode_candidate = _string_series(working.get("Level {Level SC}"), working.index)
    working["Instruction Method {Inst Method}"] = mode_candidate.where(mode_candidate.notna(), _string_series(working.get("Instruction Method {Inst Method}"), working.index))
    working["Course Start Date {Meet Start Date}"] = working.get("Pre-work Start Date {not in Banner}")
    working["Course End Date {Meet End Date}"] = working.get("Course Beginning Time {Meet Beg Time}")
    return working


def _parse_donsheet_dates_flexible(values: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(values, format="%d-%b-%y", errors="coerce")
    return parsed.fillna(pd.to_datetime(values, errors="coerce"))


def _derive_semester_label(raw_semester: object, begin_date: pd.Series, calendar_year: pd.Series | None = None) -> pd.Series:
    cleaned = pd.Series(raw_semester, copy=True)
    cleaned = cleaned.astype("string").str.strip()
    missing = cleaned.isna() | cleaned.eq("")

    if missing.any():
        months = pd.to_datetime(begin_date, errors="coerce").dt.month
        derived = pd.Series(pd.NA, index=cleaned.index, dtype="string")
        derived = derived.mask(months.isin([1, 2, 3, 4]), "Spring")
        derived = derived.mask(months.isin([5, 6, 7]), "Summer")
        derived = derived.mask(months.isin([8, 9, 10, 11, 12]), "Fall")
        cleaned = cleaned.mask(missing, derived)

    year = pd.to_datetime(begin_date, errors="coerce").dt.year.astype("Int64")
    if calendar_year is not None:
        fallback_year = pd.to_numeric(calendar_year, errors="coerce").astype("Int64")
        year = fallback_year.where(cleaned.notna() & cleaned.ne(""), year.fillna(fallback_year))
    with_year = cleaned.where(year.isna(), cleaned + " " + year.astype("Int64").astype("string"))
    return with_year


def _derive_fiscal_year(begin_date: pd.Series, calendar_year: pd.Series | None = None) -> pd.Series:
    dates = pd.to_datetime(begin_date, errors="coerce")
    years = dates.dt.year
    fiscal_year = years.where(dates.dt.month.lt(7), years + 1)
    if calendar_year is not None:
        fallback = pd.to_numeric(calendar_year, errors="coerce").astype("Int64")
        fiscal_year = fallback.fillna(fiscal_year)
    return fiscal_year.astype("Int64")


def _build_course_code(df: pd.DataFrame) -> pd.Series:
    subject = df.get("Subject", pd.Series(index=df.index, dtype="string")).astype("string").fillna("")
    number = df.get("Course Number", pd.Series(index=df.index, dtype="string")).astype("string").fillna("")
    section = df.get("Course Section", pd.Series(index=df.index, dtype="string")).astype("string").fillna("")
    code = (subject.str.strip() + number.str.strip()).str.strip()
    return code.where(section.eq("") | section.isna(), code + "-" + section.str.strip())


def _normalize_name_key(values: pd.Series) -> pd.Series:
    return (
        values.astype("string")
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
    )


def _string_series(values: pd.Series | object, index: pd.Index) -> pd.Series:
    return pd.Series(values, index=index, copy=False).astype("string").str.strip()


def _fill_instructor_details(df: pd.DataFrame, instructor_map: pd.DataFrame) -> pd.DataFrame:
    if instructor_map.empty:
        return df

    if "Faculty Name" not in df.columns:
        df = df.copy()
        df["Faculty Name"] = pd.Series(index=df.index, dtype="string")

    working = df.copy()
    working["Faculty Name"] = _normalize_name_key(working["Faculty Name"])

    lookup = instructor_map.rename(
        columns={
            "Instructor Name {Instr Name}": "Faculty Name",
            "Instructor Email {Instr Email}": "email_lookup",
            "Instructor ID {Instr ID}": "id_lookup",
        }
    )
    lookup = lookup[["Faculty Name", "email_lookup", "id_lookup"]].dropna(how="all")
    lookup["Faculty Name"] = _normalize_name_key(lookup["Faculty Name"])
    lookup["email_lookup"] = _string_series(lookup["email_lookup"], lookup.index)
    lookup["id_lookup"] = _string_series(lookup["id_lookup"], lookup.index)
    lookup = lookup.dropna(subset=["Faculty Name"])
    lookup = lookup.drop_duplicates(subset=["Faculty Name"], keep="first")

    merged = working.merge(lookup, on="Faculty Name", how="left")
    merged["email"] = _string_series(merged.get("email"), merged.index).fillna(
        _string_series(merged["email_lookup"], merged.index)
    )
    merged["ID#"] = _string_series(merged.get("ID#"), merged.index).fillna(
        _string_series(merged["id_lookup"], merged.index)
    )
    return merged.drop(columns=["email_lookup", "id_lookup"])


def build_canonical_dataframe(source: str | bytes | object) -> NormalizationResult:
    excel_file, summary = load_workbook(source)
    department_frames = read_department_sheets(excel_file, summary.data_sheet_names, source=source)
    instructor_map = read_helper_sheet(excel_file, "InstructorMap")

    normalized_frames: list[pd.DataFrame] = []
    warnings = list(summary.warnings)

    for sheet_name, raw_df in department_frames.items():
        df = raw_df.copy()
        if _looks_like_shifted_contract_layout(df):
            df = _apply_shifted_contract_layout(df)
        df["Source Workbook"] = summary.source_name
        df["Source Sheet"] = sheet_name
        df["Department"] = sheet_name
        df = df.rename(columns=COLUMN_MAP)
        df = _clean_object_columns(df)

        available_columns = [column for column in COLUMN_MAP.values() if column in df.columns]
        passthrough_columns = [column for column in raw_df.columns if column in df.columns]
        selected_columns = [
            *dict.fromkeys(["Source Workbook", "Department", "Source Sheet", *available_columns, *passthrough_columns])
        ]
        df = df[selected_columns].copy()

        df["Faculty Name"] = df.get("Faculty Name", pd.Series(index=df.index, dtype="string"))
        df["load/contract"] = (
            df.get("load/contract", pd.Series(index=df.index, dtype="string"))
            .astype("string")
            .str.lower()
            .str.strip()
        )
        df["Mode"] = df.get("Mode Raw", pd.Series(index=df.index, dtype="object")).map(_normalize_mode)
        df["Cr."] = pd.to_numeric(df.get("Cr.", pd.Series(index=df.index)), errors="coerce")
        df["rate per credit"] = pd.to_numeric(df.get("rate per credit", pd.Series(index=df.index)), errors="coerce")
        df["total contract amount"] = pd.to_numeric(
            df.get("total contract amount", pd.Series(index=df.index)),
            errors="coerce",
        )
        df["Begin Date"] = _parse_donsheet_dates_flexible(df.get("Begin Date", pd.Series(index=df.index)))
        df["End Date"] = _parse_donsheet_dates_flexible(df.get("End Date", pd.Series(index=df.index)))
        df["Semester"] = _derive_semester_label(
            df.get("Semester", pd.Series(index=df.index)),
            df["Begin Date"],
            df.get("Calendar Year"),
        )
        df["Fiscal Year"] = _derive_fiscal_year(df["Begin Date"], df.get("Calendar Year"))
        df["Course"] = _build_course_code(df)
        df["Location"] = _string_series(df.get("Location"), df.index).fillna(
            _string_series(df.get("Campus Code"), df.index)
        )
        df["Location"] = df["Location"].fillna(_string_series(df.get("Campus Restriction {Camp Restr CC}"), df.index))
        if "Term Code" in df.columns:
            term_year, term_season, term_label = _parse_term_code(df["Term Code"])
            df["Term Year"] = term_year
            df["Term Season"] = term_season
            df["Term Label"] = term_label
        else:
            df["Term Year"] = pd.to_numeric(df.get("Calendar Year"), errors="coerce").astype("Int64")
            df["Term Season"] = _string_series(df.get("Semester"), df.index)
            df["Term Label"] = _string_series(df.get("Semester"), df.index)
        missing_term_label = _string_series(df.get("Term Label"), df.index).isna()
        df.loc[missing_term_label, "Term Label"] = _string_series(df.get("Semester"), df.index)

        df = _fill_instructor_details(df, instructor_map)

        useful_rows = df["Faculty Name"].notna() | df["Course"].notna()
        df = df.loc[useful_rows].copy()
        normalized_frames.append(df)

        if df.empty:
            warnings.append(f"Sheet '{sheet_name}' produced no usable rows after normalization.")
        elif df["load/contract"].dropna().empty:
            warnings.append(f"Sheet '{sheet_name}' has no load/contract values filled yet.")

    non_empty_frames = [frame for frame in normalized_frames if not frame.empty]
    canonical_df = pd.concat(non_empty_frames, ignore_index=True, sort=False) if non_empty_frames else pd.DataFrame()

    if not canonical_df.empty:
        if "ID#" in canonical_df.columns:
            canonical_df["ID#"] = canonical_df["ID#"].astype("string")
        if "email" in canonical_df.columns:
            canonical_df["email"] = canonical_df["email"].astype("string")
        if "total contract amount" in canonical_df.columns and canonical_df["total contract amount"].dropna().empty:
            warnings.append("No total contract amount values were found in the normalized workbook.")

    return NormalizationResult(
        canonical_df=canonical_df,
        summary=summary,
        warnings=warnings,
    )


def combine_normalization_results(results: Sequence[NormalizationResult]) -> NormalizationResult:
    valid_results = [result for result in results if result is not None]
    if not valid_results:
        empty_summary = WorkbookSummary(
            source_name="uploaded workbooks",
            sheet_names=[],
            data_sheet_names=[],
            helper_sheet_names=[],
            row_counts={},
            warnings=[],
        )
        return NormalizationResult(canonical_df=pd.DataFrame(), summary=empty_summary, warnings=[])

    canonical_frames = [result.canonical_df for result in valid_results if not result.canonical_df.empty]
    canonical_df = pd.concat(canonical_frames, ignore_index=True, sort=False) if canonical_frames else pd.DataFrame()

    if not canonical_df.empty:
        if "ID#" in canonical_df.columns:
            canonical_df["ID#"] = canonical_df["ID#"].astype("string")
        if "email" in canonical_df.columns:
            canonical_df["email"] = canonical_df["email"].astype("string")

    merged_sheet_names: list[str] = []
    merged_data_sheet_names: list[str] = []
    merged_helper_sheet_names: list[str] = []
    merged_row_counts: dict[str, int] = {}
    warnings: list[str] = []
    source_names: list[str] = []

    for result in valid_results:
        summary = result.summary
        source_names.append(str(summary.source_name))
        merged_sheet_names.extend(f"{summary.source_name}: {sheet}" for sheet in summary.sheet_names)
        merged_data_sheet_names.extend(f"{summary.source_name}: {sheet}" for sheet in summary.data_sheet_names)
        merged_helper_sheet_names.extend(f"{summary.source_name}: {sheet}" for sheet in summary.helper_sheet_names)
        merged_row_counts.update(
            {
                f"{summary.source_name}: {sheet_name}": row_count
                for sheet_name, row_count in summary.row_counts.items()
            }
        )
        warnings.extend(result.warnings)

    combined_summary = WorkbookSummary(
        source_name=", ".join(source_names),
        sheet_names=merged_sheet_names,
        data_sheet_names=merged_data_sheet_names,
        helper_sheet_names=merged_helper_sheet_names,
        row_counts=merged_row_counts,
        warnings=warnings,
    )
    return NormalizationResult(
        canonical_df=canonical_df,
        summary=combined_summary,
        warnings=warnings,
    )
