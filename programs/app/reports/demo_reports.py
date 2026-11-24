from __future__ import annotations

import re
from copy import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from ..excel_utils import workbook_to_excel_bytes
from ..config import DEPARTMENT_ALIASES
from ..models import ReportBundle, ReportResult
from ..schedule_from_donsheet import (
    DEPARTMENT_ADMIN_SCHEDULE_COLUMN_WIDTH,
    DEPARTMENT_ADMIN_WEEKLY_ROOMS,
    generate_schedule_outputs_from_payloads,
)


DEFAULT_PROGRAM_ORDER = [
    "DIS",
    "DMIN-INT",
    "DMIN-NAD",
    "EGW-Cert",
    "MAPM-ENGL",
    "MAPM-HISP",
    "MAR-INT",
    "MAR-NAD",
    "MAYYM",
    "MDIV",
    "MDIV-FL",
    "PHD/THD",
    "SOLC",
    "Urban",
]

DEFAULT_DEPARTMENT_ORDER = [
    "CHIS",
    "DMIN",
    "DSLE",
    "GSEM",
    "MA_Religion",
    "MSSN",
    "NTST",
    "OTST",
    "PATH",
    "THST",
]

# Legacy Banner department codes that should roll up under the current SEM
# department in all budget reports.
SEM_DEPARTMENT_REPORT_ALIASES = DEPARTMENT_ALIASES

FACULTY_STATUS_SOURCE = (
    Path(__file__).resolve().parents[3]
    / "0_source_files"
    / "0000_BuildingFiles"
    / "source_AutoFill_names_v20261124.xlsx"
)


def _copy_frame(df: pd.DataFrame) -> pd.DataFrame:
    return df.copy() if df is not None else pd.DataFrame()


def _string_column(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(index=df.index, dtype="string")
    return df[column].astype("string").str.strip()


def _numeric_column(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(index=df.index, dtype="float64")
    return pd.to_numeric(df[column], errors="coerce")


def _safe_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _faculty_match_key(value: object) -> str:
    """Normalize a faculty name for case- and whitespace-insensitive lookup."""
    return " ".join(_safe_text(value).casefold().split())


def _faculty_id_key(value: object) -> str:
    text = _safe_text(value)
    return re.sub(r"\.0$", "", text)


def _salaried_faculty_lookup() -> tuple[set[str], set[str]]:
    """Return salaried faculty names and IDs from the maintained source workbook."""
    if not FACULTY_STATUS_SOURCE.exists():
        return set(), set()
    source = pd.read_excel(FACULTY_STATUS_SOURCE)
    required_columns = {"Faculty Name", "salaried"}
    if not required_columns.issubset(source.columns):
        return set(), set()
    salaried = source.loc[
        source["salaried"].astype("string").str.strip().str.casefold().eq("salaried")
    ].copy()
    names = {_faculty_match_key(value) for value in salaried["Faculty Name"]}
    ids = (
        {_faculty_id_key(value) for value in salaried["ID#"]}
        if "ID#" in salaried.columns
        else set()
    )
    return names - {""}, ids - {""}


def _is_salaried_faculty(name: object, faculty_ids: pd.Series, salaried_names: set[str], salaried_ids: set[str]) -> bool:
    if _faculty_match_key(name) in salaried_names:
        return True
    return any(_faculty_id_key(value) in salaried_ids for value in faculty_ids.dropna())


def _program_equals(series: pd.Series, program: str) -> pd.Series:
    return series.astype("string").fillna("").eq(str(program))


def _prep_working_df(df: pd.DataFrame, *, department_from_sem: bool = False) -> pd.DataFrame:
    working = _copy_frame(df)
    if working.empty:
        return working
    text_columns = [
        "Department",
        "SEM Department",
        "Program",
        "Faculty Name",
        "Course",
        "Catalog Title",
        "Semester",
        "Mode",
        "Location",
        "account to be charged",
        "Source Workbook",
        "Subject",
        "Course Number",
        "Course Section",
        "Term Season",
        "Term Label",
    ]
    for column in text_columns:
        working[column] = _string_column(working, column)
    # Do not silently drop otherwise usable rows from term-grouped reports
    # when their DonSheet start date is not set.  A visible label lets the
    # department correct the date while preserving the reported credits/costs.
    working["Semester"] = working["Semester"].fillna("Term not set")
    for column in ["Department", "SEM Department"]:
        aliases = working[column].str.upper().map(SEM_DEPARTMENT_REPORT_ALIASES)
        working[column] = aliases.fillna(working[column])
    working["load/contract"] = _string_column(working, "load/contract").str.lower().fillna("")
    working["Cr."] = _numeric_column(working, "Cr.").fillna(0)
    working["rate per credit"] = _numeric_column(working, "rate per credit").fillna(0)
    working["total contract amount"] = _numeric_column(working, "total contract amount").fillna(0)
    working["Fiscal Year"] = pd.to_numeric(working.get("Fiscal Year"), errors="coerce")
    working["Term Year"] = pd.to_numeric(working.get("Term Year"), errors="coerce")
    if department_from_sem:
        # Budget reports are organized by the DonSheet's designated SEM department,
        # rather than by the worksheet that happened to contain the course.
        working["Department"] = working["SEM Department"]
    else:
        working["Department"] = working["Department"].fillna(working["SEM Department"])
    working["ORG code"] = working["account to be charged"].map(_org_code_from_account)
    working["Term Code Derived"] = working["Semester"].map(_term_code_from_semester)
    working["Course Number Text"] = _course_number_text(working)
    working["Course Text"] = _course_text(working)
    working["Credits from Load"] = working["Cr."].where(working["load/contract"].eq("load"), 0)
    working["Credits from Contract"] = working["Cr."].where(working["load/contract"].eq("contract"), 0)
    return working


def _filter_department(working: pd.DataFrame, department: str) -> pd.DataFrame:
    if working.empty:
        return working
    department_value = str(department).strip()
    mask = working["Department"].eq(department_value) | working["SEM Department"].eq(department_value)
    return working.loc[mask].copy()


def _program_columns(working: pd.DataFrame) -> list[str]:
    seen = _string_column(working, "Program").dropna().unique().tolist() if not working.empty else []
    ordered = [program for program in DEFAULT_PROGRAM_ORDER if program in seen]
    extras = sorted(program for program in seen if program not in ordered)
    return ordered + extras


def _department_order(working: pd.DataFrame) -> list[str]:
    seen = _string_column(working, "Department").dropna().unique().tolist() if not working.empty else []
    ordered = [department for department in DEFAULT_DEPARTMENT_ORDER if department in seen]
    extras = sorted(department for department in seen if department not in ordered)
    return ordered + extras


def _split_name(name: object) -> tuple[str, str]:
    text = _safe_text(name)
    if "," in text:
        last, first = text.split(",", 1)
        return last.strip(), first.strip()
    parts = text.split()
    if len(parts) >= 2:
        return parts[-1], " ".join(parts[:-1])
    return text, ""


def _org_code_from_account(value: object) -> str:
    text = _safe_text(value)
    if ":" in text:
        text = text.split(":", 1)[1].strip()
    match = re.search(r"\b(\d{4})\b", text)
    return match.group(1) if match else ""


def _term_code_from_semester(semester: object) -> str:
    text = _safe_text(semester)
    if not text:
        return ""
    parts = text.split()
    if len(parts) >= 2 and parts[-1].isdigit():
        season = parts[0].lower()
        season_code = {"spring": "1", "summer": "2", "fall": "3"}.get(season, "0")
        return f"{parts[-1]}{season_code}"
    return text


def _course_number_text(working: pd.DataFrame) -> pd.Series:
    if "Course Number" not in working.columns:
        return _string_column(working, "Course")
    values = working["Course Number"]
    return values.apply(
        lambda value: str(int(value)) if isinstance(value, float) and float(value).is_integer() else _safe_text(value)
    ).astype("string")


def _course_text(working: pd.DataFrame) -> pd.Series:
    if "Course" in working.columns:
        course = _string_column(working, "Course")
        if course.fillna("").str.len().gt(0).any():
            return course
    subject = _string_column(working, "Subject").fillna("")
    number = _course_number_text(working).fillna("")
    section = _string_column(working, "Course Section").fillna("")
    return (
        subject
        + number.where(number.eq(""), number)
        + section.map(lambda value: f"-{value}" if pd.notna(value) and value != "" else "")
    ).astype("string")


def _semester_sort_key(value: object) -> str:
    """Return a scalar key so Arrow-backed pandas columns can be sorted safely."""
    text = _safe_text(value)
    if not text:
        return "9999999"
    parts = text.split()
    if len(parts) < 2 or not parts[-1].isdigit():
        return "9999999"
    season_order = {"summer": 0, "fall": 1, "spring": 2}
    return f"{int(parts[-1]):04d}{season_order.get(parts[0].lower(), 999):03d}"


def _ordered_semesters(series: pd.Series) -> list[str]:
    values = [str(value).strip() for value in series.dropna().tolist() if str(value).strip()]
    return sorted(dict.fromkeys(values), key=_semester_sort_key)


def _academic_year_sort_key(value: object) -> tuple[int, int]:
    text = _safe_text(value)
    if not text:
        return (9999, 999)
    parts = text.split()
    if len(parts) < 2 or not parts[-1].isdigit():
        return (9999, 999)
    season = parts[0].lower()
    year = int(parts[-1])
    if season == "spring":
        return (year - 1, 2)
    if season == "summer":
        return (year, 0)
    if season == "fall":
        return (year, 1)
    return (year, 999)


def _academic_year_semesters(series: pd.Series) -> list[str]:
    values = [str(value).strip() for value in series.dropna().tolist() if str(value).strip()]
    ordered = sorted(dict.fromkeys(values), key=_academic_year_sort_key)
    if not ordered:
        return []
    first_year = _academic_year_sort_key(ordered[0])[0]
    return [value for value in ordered if _academic_year_sort_key(value)[0] == first_year]


def _department_cycle_semesters(series: pd.Series) -> list[str]:
    values = [str(value).strip() for value in series.dropna().tolist() if str(value).strip()]
    values = [value for value in values if _academic_year_sort_key(value)[0] != 9999]
    if not values:
        return []
    starts = [_academic_year_sort_key(value)[0] for value in values]
    start_year = min(starts)
    target_labels = [f"Summer {start_year}", f"Fall {start_year}", f"Spring {start_year + 1}"]
    seen = set(values)
    return [label for label in target_labels if label in seen]


def _latest_department_cycle_semesters(series: pd.Series) -> list[str]:
    """Return the most recent Summer/Fall/Spring cycle represented in the data."""
    values = [str(value).strip() for value in series.dropna().tolist() if str(value).strip()]
    # Labels such as "Term not set" belong in the detail reports, but they
    # must not be considered an academic year when choosing the schedule
    # cycle.  Their sort key intentionally uses year 9999.
    values = [value for value in values if _academic_year_sort_key(value)[0] != 9999]
    if not values:
        return []
    start_year = max(_academic_year_sort_key(value)[0] for value in values)
    target_labels = [f"Summer {start_year}", f"Fall {start_year}", f"Spring {start_year + 1}"]
    seen = set(values)
    return [label for label in target_labels if label in seen]


def _clean_numeric_display(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = frame.copy()
    for column in columns:
        if column in out.columns:
            out[column] = out[column].apply(
                lambda x: int(x) if isinstance(x, float) and float(x).is_integer() else x
            )
    return out


def _blank_zeroes(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = frame.copy()
    for column in columns:
        if column in out.columns:
            out[column] = out[column].apply(lambda value: pd.NA if pd.notna(value) and value == 0 else value)
    return out


def _blank_row(columns: list[str]) -> dict[str, object]:
    return {column: pd.NA for column in columns}


def _clone_sheet_into_workbook(source_ws, target_wb, title: str) -> None:
    target_ws = target_wb.create_sheet(title=title[:31] or "Sheet")
    for row in source_ws.iter_rows():
        for cell in row:
            target = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                target.font = copy(cell.font)
                target.fill = copy(cell.fill)
                target.border = copy(cell.border)
                target.alignment = copy(cell.alignment)
                target.number_format = cell.number_format
                target.protection = copy(cell.protection)
            if cell.hyperlink:
                target._hyperlink = copy(cell.hyperlink)
            if cell.comment:
                target.comment = copy(cell.comment)
    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))
    for key, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[key].width = dim.width
        target_ws.column_dimensions[key].hidden = dim.hidden
    for key, dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[key].height = dim.height
        target_ws.row_dimensions[key].hidden = dim.hidden
    target_ws.freeze_panes = source_ws.freeze_panes
    target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines


def _set_schedule_class_column_widths(ws, width: float) -> None:
    if ws.max_column < 2:
        return
    for column_index in range(2, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(column_index)].width = width


def _build_contracts_overview_template(start_year: int | None) -> Workbook:
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Contracts"
    ws.sheet_view.showGridLines = False

    widths = {
        "A": 3.5,
        "B": 66.83,
        "C": 25.16,
        "D": 15.0,
        "E": 13.5,
        "F": 15.5,
        "G": 11.66,
        "H": 12.16,
        "I": 9.33,
        "J": 8.83,
        "K": 13.0,
        "L": 13.0,
        "M": 13.0,
    }
    for key, value in widths.items():
        ws.column_dimensions[key].width = value

    thin = Side(style="thin", color="000000")
    medium = Side(style="medium", color="000000")
    header_font = Font(name="Arial", size=11, bold=True)
    body_font = Font(name="Arial", size=11)
    title_font = Font(name="Arial", size=11, bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    currency_fill = PatternFill(fill_type="solid", fgColor="CCFFFF")
    light_blue_fill = PatternFill(fill_type="solid", fgColor="DDEBF7")
    olive_fill = PatternFill(fill_type="solid", fgColor="D8E4BC")

    def merge(range_ref: str) -> None:
        ws.merge_cells(range_ref)

    def set_text(cell_ref: str, value, *, font=body_font, alignment=None):
        cell = ws[cell_ref]
        cell.value = value
        cell.font = copy(font)
        if alignment is not None:
            cell.alignment = alignment

    def fill_range(row: int, col_start: int, col_end: int, fill) -> None:
        for col in range(col_start, col_end + 1):
            ws.cell(row, col).fill = copy(fill)

    def fill_table_columns(row_start: int, row_end: int) -> None:
        white_fill = PatternFill(fill_type="solid", fgColor="FFFFFF")
        for row in range(row_start, row_end + 1):
            for col in range(2, 6):
                ws.cell(row, col).fill = copy(white_fill)
            ws.cell(row, 6).fill = copy(light_blue_fill)
            ws.cell(row, 7).fill = copy(olive_fill)
            ws.cell(row, 8).fill = copy(olive_fill)

    def box(row: int, col_start: int, col_end: int, *, left_side=thin, right_side=thin, top_side=thin, bottom_side=thin):
        for col in range(col_start, col_end + 1):
            cell = ws.cell(row, col)
            cell.border = Border(
                left=left_side if col == col_start else thin,
                right=right_side if col == col_end else thin,
                top=top_side,
                bottom=bottom_side,
            )

    budget_end_year = start_year + 1 if start_year is not None else None
    budget_label = f"May {start_year}-April {budget_end_year}" if start_year is not None else "May - April"
    budget_label_spaced = f"May {start_year} - April {budget_end_year}" if start_year is not None else "May - April"
    summer_label = f"Summer {start_year}" if start_year is not None else "Summer"
    fall_label = f"Fall {start_year}" if start_year is not None else "Fall"
    spring_label = f"Spring {budget_end_year}" if budget_end_year is not None else "Spring"

    for rng in [
        "B2:F2", "B3:F3", "B4:F4", "B5:F5",
        "B7:E7", "B8:E8", "B9:E9", "B10:E10",
        "B12:B13", "C12:C13", "D12:E13", "F12:F13",
        "D14:E14", "D15:E15", "D16:E16", "D17:E17", "D18:E18",
        "B19:E19", "B20:E20", "B21:E21", "B22:E22",
        "B24:H24", "B25:H25", "B26:H26", "B27:H27",
        "B55:C55", "B93:C93", "B94:C94",
    ]:
        merge(rng)

    for rng in [
        "J57:L57", "J58:L58", "J59:L59", "J60:L60", "J61:L61", "J62:L62",
        "J63:L63", "J64:L64", "J65:L65", "J66:L66", "J67:L67", "J68:L68",
        "J69:L69", "J70:L70", "J71:L71",
    ]:
        merge(rng)

    set_text("B2", "ANDREWS UNIVERSITY", font=title_font, alignment=center)
    set_text("B3", "Budget Request", font=title_font, alignment=center)
    set_text("B4", budget_label, font=title_font, alignment=center)
    set_text("B5", "SCHEDULE OF DETAILS", font=title_font, alignment=center)
    set_text("B7", f"Credits on {summer_label} schedule", font=header_font, alignment=left)
    set_text("B8", f"Credits on {fall_label} schedule", font=header_font, alignment=left)
    set_text("B9", f"Credits on {spring_label} schedule", font=header_font, alignment=left)
    set_text("B10", "Total proposed credits offered for the year", font=header_font, alignment=left)
    ws["F10"] = "=SUM(F7:F9)"

    for row in [7, 8, 9]:
        box(row, 2, 6)
        ws.cell(row, 6).font = copy(header_font)
        ws.cell(row, 6).alignment = center
        ws.cell(row, 6).number_format = "0"
    box(10, 2, 6, left_side=thin, right_side=medium, top_side=thin, bottom_side=medium)
    ws["F10"].font = copy(header_font)
    ws["F10"].alignment = center

    set_text("B12", "Salaried Faculty Name", font=header_font, alignment=center)
    set_text("C12", "Load Reduction Reason", font=header_font, alignment=center)
    set_text("D12", "Load by Policy", font=header_font, alignment=center)
    set_text("F12", "Teaching Load", font=header_font, alignment=center)
    fill_range(12, 2, 6, light_blue_fill)
    fill_range(13, 2, 6, light_blue_fill)
    for row in range(12, 19):
        box(row, 2, 8, left_side=thin, right_side=medium, top_side=thin, bottom_side=thin)
        for col in range(2, 9):
            ws.cell(row, col).font = copy(body_font)
        for col in [7, 8]:
            ws.cell(row, col).border = Border()
    for ref in ["B12", "C12", "D12", "F12"]:
        ws[ref].font = copy(header_font)
    set_text("B14", "Instructor1")
    set_text("C14", "Dept chair")
    set_text("B15", "Instructor2")
    set_text("B16", "Instructor3")
    set_text("B17", "Instructor4")
    set_text("B18", "Instructor5")
    set_text("B20", "Proposed credits to be taught by Seminary faculty", font=header_font, alignment=left)
    set_text("B21", "Proposed credits to be taught by contract", font=header_font, alignment=left)
    set_text("B22", "TOTAL", font=header_font, alignment=left)
    ws["F20"] = "=SUM(F14:F17)"
    ws["F21"] = "=+D94"
    ws["F22"] = "=SUM(F20:F21)"
    for row in [20, 21]:
        box(row, 2, 6)
    box(22, 2, 6, left_side=thin, right_side=medium, top_side=thin, bottom_side=medium)

    set_text("B24", "ANDREWS UNIVERSITY", font=title_font, alignment=center)
    set_text("B25", "OT Budget Request", font=title_font, alignment=center)
    set_text("B26", budget_label_spaced, font=title_font, alignment=center)
    set_text("B27", "SCHEDULE OF DETAILS", font=title_font, alignment=center)
    set_text("B29", "Schedule 1(b) -- Contract Salaries", font=title_font)
    set_text("B30", "Persons working less than half time for ", font=header_font)
    set_text("B31", "university and/or employed on a ", font=header_font)
    set_text("B32", "contract basis, including retirees", font=header_font)
    set_text("C30", "Academic Rank", font=header_font, alignment=center)
    set_text("C31", "or Equivalent", font=header_font, alignment=center)
    set_text("D30", "Number of", font=header_font, alignment=center)
    set_text("D31", "Credits Taught", font=header_font, alignment=center)
    set_text("E30", "Rate per", font=header_font, alignment=center)
    set_text("E31", "Credit", font=header_font, alignment=center)
    set_text("F30", "Contract", font=header_font, alignment=center)
    set_text("F31", "Amount", font=header_font, alignment=center)
    set_text("G30", "Salary", font=header_font, alignment=center)
    set_text("G31", "Contracts", font=header_font, alignment=center)
    set_text("H30", "PhD", font=header_font, alignment=center)
    set_text("H31", "Contracts", font=header_font, alignment=center)
    ws["G32"] = 9120
    ws["H32"] = 9260
    fill_table_columns(30, 94)
    for row in [30, 31, 32]:
        box(row, 2, 8, left_side=medium if row == 32 else thin, right_side=medium, top_side=thin, bottom_side=thin)

    def contract_section_header(row: int, title: str) -> None:
        set_text(f"B{row}", title, font=header_font)
        box(row, 2, 8, left_side=medium, right_side=medium, top_side=medium if row in {33, 57, 67, 78, 85} else thin, bottom_side=medium if row in {57, 67, 78, 85} else thin)

    def semester_row(row: int, label: str) -> None:
        set_text(f"B{row}", label, font=header_font)
        box(row, 2, 8, left_side=medium, right_side=medium, top_side=thin, bottom_side=thin)

    def detail_row(row: int) -> None:
        box(row, 2, 8, left_side=thin, right_side=medium, top_side=thin, bottom_side=thin)
        ws[f"F{row}"] = f"=E{row}*D{row}"
        ws[f"D{row}"].number_format = '0.##'
        ws[f"E{row}"].number_format = '"$"#,##0.00'
        ws[f"F{row}"].number_format = '"$"#,##0.00'
        ws[f"G{row}"].number_format = '"$"#,##0.00'
        ws[f"H{row}"].number_format = '"$"#,##0.00'

    contract_section_header(33, "ON-CAMPUS CONTRACTS")
    semester_row(34, summer_label)
    for row in [35, 36, 37]:
        detail_row(row)
    semester_row(39, fall_label)
    for row in [40, 41, 42, 43, 44, 45]:
        detail_row(row)
    semester_row(47, spring_label)
    for row in [48, 49, 50, 51, 52, 53]:
        detail_row(row)
    set_text("B55", "Total On Campus Contracts", font=header_font, alignment=right)
    ws["D55"] = "=SUM(D35:D54)"
    ws["F55"] = "=SUM(F35:F54)"
    ws["G55"] = "=SUM(G35:G54)"
    ws["H55"] = "=SUM(H35:H54)"
    box(55, 2, 8, left_side=medium, right_side=medium, top_side=medium, bottom_side=medium)

    contract_section_header(57, "MAPM - ENG CONTRACT")
    semester_row(59, fall_label)
    for row in [60]:
        detail_row(row)
    semester_row(62, spring_label)
    for row in [63]:
        detail_row(row)

    contract_section_header(67, "MAPM - HISP CONTRACT")
    semester_row(68, summer_label)
    for row in [69]:
        detail_row(row)
    semester_row(71, fall_label)
    for row in [72, 73]:
        detail_row(row)
    semester_row(75, spring_label)
    for row in [76]:
        detail_row(row)

    contract_section_header(78, "MA RELIGION CONTRACT")
    semester_row(80, fall_label)
    for row in [81, 82]:
        detail_row(row)

    contract_section_header(85, "MDIV Spanish FL Cohort")
    semester_row(87, fall_label)
    for row in [88]:
        detail_row(row)

    set_text("B93", "Total Extension School Classes", font=header_font, alignment=right)
    ws["D93"] = "=SUM(D58:D92)"
    ws["F93"] = "=SUM(F58:F92)"
    ws["G93"] = "=SUM(G58:G92)"
    ws["H93"] = "=SUM(H58:H92)"
    box(93, 2, 8, left_side=medium, right_side=medium, top_side=medium, bottom_side=medium)

    set_text("B94", "Total Contracts", font=header_font, alignment=right)
    ws["D94"] = "=+D55+D93"
    ws["F94"] = "=+F55+F93"
    ws["G94"] = "=+G55+G93"
    ws["H94"] = "=+H55+H93"
    ws["F94"].number_format = '"$"#,##0.00'
    ws["G94"].number_format = '"$"#,##0.00'
    ws["H94"].number_format = '"$"#,##0.00'
    box(94, 2, 8, left_side=medium, right_side=medium, top_side=medium, bottom_side=medium)

    for row in range(33, 95):
        if row not in {33, 55, 57, 67, 78, 85, 93, 94}:
            box(row, 2, 8, left_side=thin, right_side=medium, top_side=thin, bottom_side=thin)

    location_labels = [
        "LOCATION ABBREVIATIONS",
        "Columbia Union - CU",
        "Lake Union - LU",
        "Mid-America Union - MAU",
        "Pacific Union - PU",
        "North Pacific Union - NPU",
        "Poland/Baltic States - PolBal",
        "Poland/Baltic States - PolBal",
        "Poland/Baltic States - PolBal",
        "Romania - Rom",
        "Ukraine - Ukr",
        "Main Campus - MC",
        "Atlantic Union - AtlU",
        "Hong Kong - HK",
        "Lake Union - LU",
    ]
    location_rows = [57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 68, 69, 74, 83, 71]
    for row, label in zip(location_rows, location_labels):
        merge(f"J{row}:L{row}")
        set_text(f"J{row}", label, font=body_font if row != 57 else header_font, alignment=center if row != 57 else left)

    return workbook


def _safe_sum(series: pd.Series) -> float:
    return float(pd.to_numeric(series, errors="coerce").fillna(0).sum())


def _department_org_code(df: pd.DataFrame) -> str:
    if df.empty or "ORG code" not in df.columns:
        return ""
    non_blank = df["ORG code"].dropna().astype("string").str.strip()
    non_blank = non_blank[non_blank.ne("")]
    return str(non_blank.iloc[0]) if not non_blank.empty else ""


def _faculty_program_matrix(
    df: pd.DataFrame,
    value_column: str,
    programs: list[str],
    include_total: bool = True,
) -> pd.DataFrame:
    if df.empty:
        columns = ["Faculty Name", *programs]
        if include_total:
            columns.append("Total")
        return pd.DataFrame(columns=columns)
    pivot = (
        pd.pivot_table(
            df,
            index="Faculty Name",
            columns="Program",
            values=value_column,
            aggfunc="sum",
            fill_value=0,
        )
        .reindex(columns=programs, fill_value=0)
        .reset_index()
    )
    if include_total:
        pivot["Total"] = pivot[programs].sum(axis=1)
    return pivot


def _report_001a_columns(program_cols: list[str]) -> list[str]:
    return [
        "Department",
        "ORG code",
        "Mode",
        "Semester",
        "LastName",
        "FirstName",
        "CourseNumber",
        "Cr.",
        "Cost",
        *program_cols,
        "Check Total",
        "Notes",
    ]


def _overview_report_001a(working: pd.DataFrame) -> pd.DataFrame:
    program_cols = _program_columns(working)
    columns = _report_001a_columns(program_cols)
    if working.empty:
        return pd.DataFrame(columns=columns)

    report = working.copy()
    names = report["Faculty Name"].map(_split_name)
    report["LastName"] = names.map(lambda value: value[0])
    report["FirstName"] = names.map(lambda value: value[1])
    report["CourseNumber"] = report["Course Text"]
    report["Cost"] = report["total contract amount"]
    for program in program_cols:
        report[program] = report["Cost"].where(_program_equals(report["Program"], program), 0)
    report["Check Total"] = report[program_cols].sum(axis=1) if program_cols else report["Cost"]
    report["Notes"] = pd.NA
    report = report.sort_values(
        ["Department", "Semester", "LastName", "FirstName", "CourseNumber"],
        key=lambda col: col.map(_semester_sort_key) if col.name == "Semester" else col,
        na_position="last",
    )

    lines: list[dict[str, object]] = []
    for department in _department_order(report):
        dept_df = report.loc[report["Department"].eq(department), columns].copy()
        if dept_df.empty:
            continue
        lines.extend(dept_df.to_dict("records"))
        summary = {column: pd.NA for column in columns}
        summary["Department"] = department
        summary["ORG code"] = _department_org_code(dept_df)
        summary["LastName"] = "DEPARTMENT TOTAL"
        for column in ["Cr.", "Cost", "Check Total", *program_cols]:
            summary[column] = _safe_sum(dept_df[column])
        lines.append(summary)
        lines.append(_blank_row(columns))
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), ["Cr.", "Cost", "Check Total", *program_cols])


def _overview_report_001b(working: pd.DataFrame) -> pd.DataFrame:
    columns = ["Faculty", "Course", "Mode", "Location", "Credits", "Rate", "Total"]
    if working.empty:
        return pd.DataFrame(columns=columns)

    contract_rows = working.loc[working["load/contract"].eq("contract")].copy()
    if contract_rows.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for department in _department_order(contract_rows):
        dept_df = contract_rows.loc[contract_rows["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        lines.append({"Faculty": department})
        for program in _program_columns(dept_df):
            program_df = dept_df.loc[dept_df["Program"].eq(program)].copy()
            if program_df.empty:
                continue
            lines.append({"Faculty": f"{program} contracts"})
            for semester in _ordered_semesters(program_df["Semester"]):
                semester_df = program_df.loc[program_df["Semester"].eq(semester)].copy()
                lines.append({"Faculty": semester})
                semester_df = semester_df.sort_values(["Faculty Name", "Course Text"], na_position="last")
                for _, row in semester_df.iterrows():
                    lines.append(
                        {
                            "Faculty": row["Faculty Name"],
                            "Course": row["Course Text"],
                            "Mode": row["Mode"],
                            "Location": row["Location"],
                            "Credits": row["Cr."],
                            "Rate": row["rate per credit"],
                            "Total": row["total contract amount"],
                        }
                    )
                lines.append(_blank_row(columns))
            lines.append(
                {
                    "Faculty": f"Total {program}",
                    "Credits": _safe_sum(program_df["Cr."]),
                    "Total": _safe_sum(program_df["total contract amount"]),
                }
            )
            lines.append(_blank_row(columns))
        lines.append(
            {
                "Faculty": "Total contracts",
                "Credits": _safe_sum(dept_df["Cr."]),
                "Total": _safe_sum(dept_df["total contract amount"]),
            }
        )
        lines.append(_blank_row(columns))
        lines.append(_blank_row(columns))
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), ["Credits", "Rate", "Total"])


def _overview_report_001c(working: pd.DataFrame) -> pd.DataFrame:
    columns = ["Faculty", "Course", "Program", "LoadType", "Mode", "Location", "Credits"]
    if working.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for department in _department_order(working):
        dept_df = working.loc[working["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        lines.append({"Faculty": department})
        for semester in _ordered_semesters(dept_df["Semester"]):
            semester_df = dept_df.loc[dept_df["Semester"].eq(semester)].copy()
            if semester_df.empty:
                continue
            lines.append({"Faculty": semester})
            semester_df = semester_df.sort_values(["Faculty Name", "Course Text"], na_position="last")
            for _, row in semester_df.iterrows():
                lines.append(
                    {
                        "Faculty": row["Faculty Name"],
                        "Course": row["Course Text"],
                        "Program": row["Program"],
                        "LoadType": _report_001c_load_type(row["load/contract"]),
                        "Mode": row["Mode"],
                        "Location": row["Location"],
                        "Credits": row["Cr."],
                    }
                )
            lines.append(
                {
                    "Faculty": f"Total {semester}",
                    "Credits": _safe_sum(semester_df["Cr."]),
                }
            )
            lines.append(_blank_row(columns))
        lines.append(
            {
                "Faculty": f"{department} Total Credits",
                "Credits": _safe_sum(dept_df["Cr."]),
            }
        )
        lines.append(_blank_row(columns))
        lines.append(_blank_row(columns))
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), ["Credits"])


def _report_001c_load_type(value: object) -> str:
    """Keep the two budget categories visible and consolidate all other labels."""
    normalized = _safe_text(value).lower()
    if normalized in {"load", "contract"}:
        return normalized
    return "Other (continuous/x-list/NA/etc.)"


def _overview_report_001f(working: pd.DataFrame) -> pd.DataFrame:
    columns = ["Department", "Faculty", "Course", "Program", "LoadType", "Mode", "Location", "Credits", "Rate", "Total"]
    if working.empty:
        return pd.DataFrame(columns=columns)

    department = str(_string_column(working, "Department").dropna().iloc[0]) if not working.empty else ""
    lines: list[dict[str, object]] = []
    faculty_names = sorted(_string_column(working, "Faculty Name").dropna().unique().tolist())

    first_faculty = True
    for faculty in faculty_names:
        faculty_df = working.loc[working["Faculty Name"].eq(faculty)].copy()
        if faculty_df.empty:
            continue
        faculty_total_credits = 0.0
        faculty_total_amount = 0.0
        first_semester_row = True
        for semester in _department_cycle_semesters(faculty_df["Semester"]):
            semester_df = faculty_df.loc[faculty_df["Semester"].eq(semester)].copy()
            if semester_df.empty:
                continue
            lines.append(
                {
                    "Department": department if first_faculty and first_semester_row else pd.NA,
                    "Faculty": faculty if first_semester_row else pd.NA,
                    "Course": semester,
                }
            )
            semester_total_credits = 0.0
            semester_total_amount = 0.0
            semester_df = semester_df.sort_values(["Course Text", "Mode", "Location"], na_position="last")
            for _, row in semester_df.iterrows():
                semester_total_credits += float(row["Cr."])
                semester_total_amount += float(row["total contract amount"])
                lines.append(
                    {
                        "Course": row["Course Text"],
                        "Program": row["Program"],
                        "LoadType": _report_001c_load_type(row["load/contract"]),
                        "Mode": row["Mode"],
                        "Location": row["Location"],
                        "Credits": row["Cr."],
                        "Rate": row["rate per credit"] if _is_contract_value(row["load/contract"]) else 0,
                        "Total": row["total contract amount"] if _is_contract_value(row["load/contract"]) else 0,
                    }
                )
            lines.append({"Course": f"Total {semester}", "Credits": semester_total_credits, "Total": semester_total_amount})
            faculty_total_credits += semester_total_credits
            faculty_total_amount += semester_total_amount
            first_semester_row = False
        lines.append({"Course": "TOTAL", "Credits": faculty_total_credits, "Total": faculty_total_amount})
        first_faculty = False
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), ["Credits", "Rate", "Total"])


def _overview_report_001f_all_departments(working: pd.DataFrame) -> pd.DataFrame:
    columns = ["Department", "Faculty", "Course", "Program", "LoadType", "Mode", "Location", "Credits", "Rate", "Total"]
    if working.empty:
        return pd.DataFrame(columns=columns)
    frames = []
    for department in _department_order(working):
        dept_df = working.loc[working["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        frames.append(_overview_report_001f(dept_df))
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=columns)


def _report_002a_columns(program_cols: list[str]) -> list[str]:
    base_columns = [
        "Department",
        "ORG code",
        "Account Charged",
        "Load/Contract",
        "Term Code",
        "Mode/Cohort",
        "Semester",
        "Faculty Name",
        "Subject",
        "Course Number",
        "Course Section",
        "Credits from Load",
        "Credits from Contract",
        "Cr.",
        "Cost",
    ]
    pairs = []
    for program in program_cols:
        pairs.extend([f"{program} Credits", f"{program} Cost"])
    return base_columns + pairs


def _overview_report_002a(working: pd.DataFrame) -> pd.DataFrame:
    program_cols = _program_columns(working)
    columns = _report_002a_columns(program_cols)
    if working.empty:
        return pd.DataFrame(columns=columns)

    report = working.copy()
    report["Account Charged"] = report["account to be charged"]
    report["Load/Contract"] = report["load/contract"]
    report["Term Code"] = report["Term Code"] if "Term Code" in report.columns else report["Term Code Derived"]
    report["Mode/Cohort"] = report["Mode"]
    report["Cost"] = report["total contract amount"].where(report["load/contract"].eq("contract"), 0)
    for program in program_cols:
        program_mask = _program_equals(report["Program"], program)
        report[f"{program} Credits"] = report["Cr."].where(program_mask, 0)
        report[f"{program} Cost"] = report["Cost"].where(program_mask, 0)
    report = report[columns].sort_values(
        ["Department", "Semester", "Faculty Name", "Subject", "Course Number"],
        key=lambda col: col.map(_semester_sort_key) if col.name == "Semester" else col,
        na_position="last",
    )

    lines: list[dict[str, object]] = []
    numeric_columns = ["Credits from Load", "Credits from Contract", "Cr.", "Cost", *[f"{program} Credits" for program in program_cols], *[f"{program} Cost" for program in program_cols]]
    for department in _department_order(report):
        dept_df = report.loc[report["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        lines.extend(dept_df.to_dict("records"))
        summary = {column: pd.NA for column in columns}
        summary["Department"] = department
        summary["ORG code"] = _department_org_code(dept_df)
        summary["Faculty Name"] = "DEPARTMENT TOTAL"
        for column in numeric_columns:
            summary[column] = _safe_sum(dept_df[column])
        lines.append(summary)
        lines.append(_blank_row(columns))
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), numeric_columns)


def _overview_report_003a(working: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Department",
        "Contract Cost",
        "Load_credits",
        "Contract_credits",
        "Load/Contract Ratio",
        "Sum of all Credits",
        "Sum of all Credits wo Lilly",
        "Sum of all Credits wo NADEI's MDIV courses",
        "Load_credits wo NADEI's MDIV courses",
        "Load/Contract Ratio wo NADEI's MDIV courses",
        "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)",
        "Contract Credits wo Lilly funded Courses",
    ]
    if working.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for department in _department_order(working):
        dept_df = working.loc[working["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        load_credits = _safe_sum(dept_df.loc[dept_df["load/contract"].eq("load"), "Cr."])
        contract_credits = _safe_sum(dept_df.loc[dept_df["load/contract"].eq("contract"), "Cr."])
        contract_cost = _safe_sum(dept_df.loc[dept_df["load/contract"].eq("contract"), "total contract amount"])
        total_credits = load_credits + contract_credits
        # The current DonSheet does not expose Lilly/NADEI flags in a stable way, so these
        # legacy columns currently mirror the full totals until those markers are formalized.
        lines.append(
            {
                "Department": department,
                "Contract Cost": contract_cost,
                "Load_credits": load_credits,
                "Contract_credits": contract_credits,
                "Load/Contract Ratio": (contract_credits / load_credits) if load_credits else pd.NA,
                "Sum of all Credits": total_credits,
                "Sum of all Credits wo Lilly": total_credits,
                "Sum of all Credits wo NADEI's MDIV courses": total_credits,
                "Load_credits wo NADEI's MDIV courses": load_credits,
                "Load/Contract Ratio wo NADEI's MDIV courses": (contract_credits / load_credits) if load_credits else pd.NA,
                "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)": total_credits,
                "Contract Credits wo Lilly funded Courses": contract_credits,
            }
        )
    report = pd.DataFrame(lines, columns=columns)
    totals = {
        "Department": "TOTAL",
        "Contract Cost": _safe_sum(report["Contract Cost"]),
        "Load_credits": _safe_sum(report["Load_credits"]),
        "Contract_credits": _safe_sum(report["Contract_credits"]),
        "Sum of all Credits": _safe_sum(report["Sum of all Credits"]),
        "Sum of all Credits wo Lilly": _safe_sum(report["Sum of all Credits wo Lilly"]),
        "Sum of all Credits wo NADEI's MDIV courses": _safe_sum(report["Sum of all Credits wo NADEI's MDIV courses"]),
        "Load_credits wo NADEI's MDIV courses": _safe_sum(report["Load_credits wo NADEI's MDIV courses"]),
        "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)": _safe_sum(report["Sum of all Credits wo (Lilly + NADEI's Mdiv courses)"]),
        "Contract Credits wo Lilly funded Courses": _safe_sum(report["Contract Credits wo Lilly funded Courses"]),
    }
    totals["Load/Contract Ratio"] = totals["Contract_credits"] / totals["Load_credits"] if totals["Load_credits"] else pd.NA
    totals["Load/Contract Ratio wo NADEI's MDIV courses"] = (
        totals["Contract Credits wo Lilly funded Courses"] / totals["Load_credits wo NADEI's MDIV courses"]
        if totals["Load_credits wo NADEI's MDIV courses"]
        else pd.NA
    )
    report = pd.concat([report, pd.DataFrame([totals])], ignore_index=True)
    return _clean_numeric_display(
        report,
        [
            "Contract Cost",
            "Load_credits",
            "Contract_credits",
            "Sum of all Credits",
            "Sum of all Credits wo Lilly",
            "Sum of all Credits wo NADEI's MDIV courses",
            "Load_credits wo NADEI's MDIV courses",
            "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)",
            "Contract Credits wo Lilly funded Courses",
        ],
    )


def _budget_report_002b(working: pd.DataFrame) -> pd.DataFrame:
    program_cols = _program_columns(working)
    credit_columns: list[str] = []
    for program in program_cols:
        credit_columns.append(f"{program} Credits")
    columns = [
        "Department",
        "ORG code",
        "Course",
        "Faculty Name",
        "Semester",
        "load/contract",
        "Program",
        "Cr.",
        *credit_columns,
        "Total Credits",
    ]
    if working.empty:
        return pd.DataFrame(columns=columns)

    report = working.copy()
    for program in program_cols:
        report[f"{program} Credits"] = report["Cr."].where(_program_equals(report["Program"], program), 0)
    report["Total Credits"] = report["Cr."]
    report = report[columns].sort_values(
        ["Department", "Program", "Semester", "Course", "Faculty Name"],
        key=lambda col: col.map(_semester_sort_key) if col.name == "Semester" else col,
        na_position="last",
    )

    lines: list[dict[str, object]] = []
    numeric_columns = ["Cr.", *credit_columns, "Total Credits"]
    for department in _department_order(report):
        dept_df = report.loc[report["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        lines.extend(dept_df.to_dict("records"))
        summary = {column: pd.NA for column in columns}
        summary["Department"] = department
        summary["ORG code"] = _department_org_code(dept_df)
        summary["Course"] = "DEPARTMENT TOTAL"
        for column in numeric_columns:
            summary[column] = _safe_sum(dept_df[column])
        lines.append(summary)
        lines.append(_blank_row(columns))
    out = _clean_numeric_display(pd.DataFrame(lines, columns=columns), numeric_columns)
    return _blank_zeroes(out, ["Cr.", *credit_columns, "Total Credits"])


def _budget_report_003b(working: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Subject",
        "Contract Cost",
        "Load_credits",
        "Contract_credits",
        "Load/Contract Ratio",
        "Sum of all Credits",
        "Sum of all Credits wo Lilly",
        "Sum of all Credits wo NADEI's MDIV courses",
        "Load_credits wo NADEI's MDIV courses",
        "Load/Contract Ratio wo NADEI's MDIV courses",
        "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)",
        "Contract Credits wo Lilly funded Courses",
    ]
    if working.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    subjects = sorted(_string_column(working, "Subject").dropna().unique().tolist())
    for subject in subjects:
        subject_df = working.loc[working["Subject"].eq(subject)].copy()
        load_credits = _safe_sum(subject_df.loc[subject_df["load/contract"].eq("load"), "Cr."])
        contract_credits = _safe_sum(subject_df.loc[subject_df["load/contract"].eq("contract"), "Cr."])
        contract_cost = _safe_sum(subject_df.loc[subject_df["load/contract"].eq("contract"), "total contract amount"])
        total_credits = load_credits + contract_credits
        lines.append(
            {
                "Subject": subject,
                "Contract Cost": contract_cost,
                "Load_credits": load_credits,
                "Contract_credits": contract_credits,
                "Load/Contract Ratio": (contract_credits / load_credits) if load_credits else pd.NA,
                "Sum of all Credits": total_credits,
                "Sum of all Credits wo Lilly": total_credits,
                "Sum of all Credits wo NADEI's MDIV courses": total_credits,
                "Load_credits wo NADEI's MDIV courses": load_credits,
                "Load/Contract Ratio wo NADEI's MDIV courses": (contract_credits / load_credits) if load_credits else pd.NA,
                "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)": total_credits,
                "Contract Credits wo Lilly funded Courses": contract_credits,
            }
        )
    report = pd.DataFrame(lines, columns=columns)
    totals = {
        "Subject": "TOTAL",
        "Contract Cost": _safe_sum(report["Contract Cost"]),
        "Load_credits": _safe_sum(report["Load_credits"]),
        "Contract_credits": _safe_sum(report["Contract_credits"]),
        "Sum of all Credits": _safe_sum(report["Sum of all Credits"]),
        "Sum of all Credits wo Lilly": _safe_sum(report["Sum of all Credits wo Lilly"]),
        "Sum of all Credits wo NADEI's MDIV courses": _safe_sum(report["Sum of all Credits wo NADEI's MDIV courses"]),
        "Load_credits wo NADEI's MDIV courses": _safe_sum(report["Load_credits wo NADEI's MDIV courses"]),
        "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)": _safe_sum(report["Sum of all Credits wo (Lilly + NADEI's Mdiv courses)"]),
        "Contract Credits wo Lilly funded Courses": _safe_sum(report["Contract Credits wo Lilly funded Courses"]),
    }
    totals["Load/Contract Ratio"] = (
        totals["Contract_credits"] / totals["Load_credits"] if totals["Load_credits"] else pd.NA
    )
    totals["Load/Contract Ratio wo NADEI's MDIV courses"] = (
        totals["Contract Credits wo Lilly funded Courses"] / totals["Load_credits wo NADEI's MDIV courses"]
        if totals["Load_credits wo NADEI's MDIV courses"]
        else pd.NA
    )
    report = pd.concat([report, pd.DataFrame([totals])], ignore_index=True)
    return _clean_numeric_display(
        report,
        [
            "Contract Cost",
            "Load_credits",
            "Contract_credits",
            "Sum of all Credits",
            "Sum of all Credits wo Lilly",
            "Sum of all Credits wo NADEI's MDIV courses",
            "Load_credits wo NADEI's MDIV courses",
            "Sum of all Credits wo (Lilly + NADEI's Mdiv courses)",
            "Contract Credits wo Lilly funded Courses",
        ],
    )


def _budget_report_004a(working: pd.DataFrame) -> pd.DataFrame:
    programs = _program_columns(working)
    columns = ["Department", "Faculty Name", *programs, "Total"]
    if working.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for department in _department_order(working):
        dept_df = working.loc[working["Department"].eq(department)].copy()
        faculty_matrix = _faculty_program_matrix(dept_df, "Cr.", programs)
        for _, row in faculty_matrix.sort_values("Faculty Name", na_position="last").iterrows():
            line = {"Department": department, "Faculty Name": row["Faculty Name"]}
            for program in programs:
                line[program] = row[program]
            line["Total"] = row["Total"]
            lines.append(line)
        total_line = {"Department": department, "Faculty Name": "DEPARTMENT TOTAL"}
        for program in programs:
            total_line[program] = _safe_sum(dept_df.loc[dept_df["Program"].eq(program), "Cr."])
        total_line["Total"] = sum(total_line[program] for program in programs)
        lines.append(total_line)
        lines.append(_blank_row(columns))
        lines.append(_blank_row(columns))
    grand_total = {"Department": "GRAND TOTAL", "Faculty Name": pd.NA}
    for program in programs:
        grand_total[program] = _safe_sum(working.loc[working["Program"].eq(program), "Cr."])
    grand_total["Total"] = sum(grand_total[program] for program in programs)
    lines.append(grand_total)
    out = _clean_numeric_display(pd.DataFrame(lines, columns=columns), [*programs, "Total"])
    return _blank_zeroes(out, [*programs, "Total"])


def _budget_report_004b(working: pd.DataFrame) -> pd.DataFrame:
    programs = _program_columns(working)
    total_credit_columns = [
        "Total credits per person",
        "Total Load-Credits per person",
        "Total Contract-Credits per person",
    ]
    columns = ["Department", "Faculty Name", *total_credit_columns, *programs]
    if working.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for department in _department_order(working):
        dept_df = working.loc[working["Department"].eq(department)].copy()
        load_df = dept_df.loc[dept_df["load/contract"].eq("load")].copy()
        contract_df = dept_df.loc[dept_df["load/contract"].eq("contract")].copy()
        credit_df = pd.concat([load_df, contract_df], ignore_index=True)
        faculty_names = sorted(_string_column(credit_df, "Faculty Name").dropna().unique().tolist())

        for faculty_name in faculty_names:
            faculty_load_df = load_df.loc[load_df["Faculty Name"].eq(faculty_name)]
            faculty_contract_df = contract_df.loc[contract_df["Faculty Name"].eq(faculty_name)]
            load_credits = _safe_sum(faculty_load_df["Cr."])
            contract_credits = _safe_sum(faculty_contract_df["Cr."])
            line = {
                "Department": department,
                "Faculty Name": faculty_name,
                "Total credits per person": load_credits + contract_credits,
                "Total Load-Credits per person": load_credits,
                "Total Contract-Credits per person": contract_credits,
            }
            for program in programs:
                line[program] = _safe_sum(faculty_load_df.loc[faculty_load_df["Program"].eq(program), "Cr."])
            lines.append(line)

        def _summary_row(label: str, subset: pd.DataFrame, load_credits: float, contract_credits: float) -> dict[str, object]:
            line = {
                "Department": department,
                "Faculty Name": label,
                "Total credits per person": load_credits + contract_credits,
                "Total Load-Credits per person": load_credits,
                "Total Contract-Credits per person": contract_credits,
            }
            for program in programs:
                line[program] = _safe_sum(subset.loc[subset["Program"].eq(program), "Cr."])
            return line

        load_total = _safe_sum(load_df["Cr."])
        contract_total = _safe_sum(contract_df["Cr."])
        lines.append(_summary_row("DEPARTMENT TOTAL Load Credits", load_df, load_total, 0))
        lines.append(_summary_row("DEPARTMENT TOTAL Contract Credits", contract_df, 0, contract_total))
        lines.append(_summary_row("DEPARTMENT TOTAL Contract+Load Credits", credit_df, load_total, contract_total))
        lines.append(_blank_row(columns))
        lines.append(_blank_row(columns))
    all_load_df = working.loc[working["load/contract"].eq("load")]
    all_contract_df = working.loc[working["load/contract"].eq("contract")]
    all_credit_df = pd.concat([all_load_df, all_contract_df], ignore_index=True)
    grand_total = {
        "Department": "GRAND TOTAL",
        "Faculty Name": pd.NA,
        "Total credits per person": _safe_sum(all_credit_df["Cr."]),
        "Total Load-Credits per person": _safe_sum(all_load_df["Cr."]),
        "Total Contract-Credits per person": _safe_sum(all_contract_df["Cr."]),
    }
    for program in programs:
        grand_total[program] = _safe_sum(all_credit_df.loc[all_credit_df["Program"].eq(program), "Cr."])
    lines.append(grand_total)
    out = _clean_numeric_display(pd.DataFrame(lines, columns=columns), [*total_credit_columns, *programs])
    return _blank_zeroes(out, [*total_credit_columns, *programs])


def _budget_report_005(working: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Program",
        "credits from Load",
        "credits from Contract",
        "credits from neither Load nor Contract",
        "Total credits",
        "Contract/Load Ratio",
    ]
    if working.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    programs = _program_columns(working)
    for program in programs:
        program_df = working.loc[working["Program"].eq(program)].copy()
        load_credits = _safe_sum(program_df.loc[program_df["load/contract"].eq("load"), "Cr."])
        contract_credits = _safe_sum(program_df.loc[program_df["load/contract"].eq("contract"), "Cr."])
        neither_credits = _safe_sum(program_df.loc[~program_df["load/contract"].isin(["load", "contract"]), "Cr."])
        total_credits = load_credits + contract_credits + neither_credits
        lines.append(
            {
                "Program": program,
                "credits from Load": load_credits,
                "credits from Contract": contract_credits,
                "credits from neither Load nor Contract": neither_credits,
                "Total credits": total_credits,
                "Contract/Load Ratio": (contract_credits / load_credits) if load_credits else pd.NA,
            }
        )
    report = pd.DataFrame(lines, columns=columns)
    totals = {
        "Program": "TOTAL",
        "credits from Load": _safe_sum(report["credits from Load"]),
        "credits from Contract": _safe_sum(report["credits from Contract"]),
        "credits from neither Load nor Contract": _safe_sum(report["credits from neither Load nor Contract"]),
        "Total credits": _safe_sum(report["Total credits"]),
    }
    totals["Contract/Load Ratio"] = (
        totals["credits from Contract"] / totals["credits from Load"] if totals["credits from Load"] else pd.NA
    )
    report = pd.concat([report, pd.DataFrame([totals])], ignore_index=True)
    return _clean_numeric_display(
        report,
        ["credits from Load", "credits from Contract", "credits from neither Load nor Contract", "Total credits"],
    )


def _budget_report_006(working: pd.DataFrame) -> pd.DataFrame:
    columns = ["Department", "Program", "contract_count", "contract_credits", "contract_dollars"]
    if working.empty:
        return pd.DataFrame(columns=columns)

    contract_df = working.loc[working["load/contract"].eq("contract")].copy()
    if contract_df.empty:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for department in _department_order(contract_df):
        dept_df = contract_df.loc[contract_df["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        for program in _program_columns(dept_df):
            program_df = dept_df.loc[dept_df["Program"].eq(program)].copy()
            if program_df.empty:
                continue
            lines.append(
                {
                    "Department": department,
                    "Program": program,
                    "contract_count": int(len(program_df)),
                    "contract_credits": _safe_sum(program_df["Cr."]),
                    "contract_dollars": _safe_sum(program_df["total contract amount"]),
                }
            )
        lines.append(
            {
                "Department": department,
                "Program": "DEPARTMENT TOTAL",
                "contract_count": int(len(dept_df)),
                "contract_credits": _safe_sum(dept_df["Cr."]),
                "contract_dollars": _safe_sum(dept_df["total contract amount"]),
            }
        )
        lines.append(_blank_row(columns))
        lines.append(_blank_row(columns))
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), ["contract_count", "contract_credits", "contract_dollars"])


def _fiscal_years(working: pd.DataFrame) -> list[int]:
    years = [int(value) for value in working["Fiscal Year"].dropna().tolist()] if not working.empty else []
    return sorted(dict.fromkeys(years))


def _year_columns(years: list[int], prefix: str = "FY") -> list[str]:
    return [f"{prefix}{year}" for year in years]


def _budget_report_007(working: pd.DataFrame) -> pd.DataFrame:
    years = _fiscal_years(working)
    columns = ["Type", "Department", *_year_columns(years)]
    if working.empty or not years:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for load_type, label in [("load", "load"), ("contract", "contract"), ("all", "total")]:
        for department in _department_order(working):
            dept_df = working.loc[working["Department"].eq(department)].copy()
            if load_type != "all":
                dept_df = dept_df.loc[dept_df["load/contract"].eq(load_type)].copy()
            line = {"Type": label, "Department": department}
            for year in years:
                line[f"FY{year}"] = _safe_sum(dept_df.loc[dept_df["Fiscal Year"].eq(year), "Cr."])
            lines.append(line)
        total_line = {"Type": label, "Department": "total"}
        for year in years:
            subset = working.loc[working["Fiscal Year"].eq(year)].copy()
            if load_type != "all":
                subset = subset.loc[subset["load/contract"].eq(load_type)].copy()
            total_line[f"FY{year}"] = _safe_sum(subset["Cr."])
        lines.append(total_line)
        lines.append(_blank_row(columns))
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), _year_columns(years))


def _budget_report_008a(working: pd.DataFrame) -> pd.DataFrame:
    years = _fiscal_years(working)
    columns = ["Program"]
    for year in years:
        columns.extend([f"FY{year} Credits", f"FY{year} Δ vs prev"])
    if working.empty or not years:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for program in _program_columns(working):
        program_df = working.loc[working["Program"].eq(program)].copy()
        line = {"Program": program}
        previous = 0.0
        for year in years:
            credits = _safe_sum(program_df.loc[program_df["Fiscal Year"].eq(year), "Cr."])
            line[f"FY{year} Credits"] = credits or pd.NA
            line[f"FY{year} Δ vs prev"] = (credits - previous) if years.index(year) > 0 or previous else pd.NA
            previous = credits
        lines.append(line)
    report = pd.DataFrame(lines, columns=columns)
    total_line = {"Program": "TOTAL"}
    for year in years:
        total_line[f"FY{year} Credits"] = _safe_sum(report[f"FY{year} Credits"])
    for idx, year in enumerate(years):
        total_line[f"FY{year} Δ vs prev"] = (
            total_line[f"FY{year} Credits"] - total_line[f"FY{years[idx - 1]} Credits"]
            if idx > 0
            else pd.NA
        )
    report = pd.concat([report, pd.DataFrame([total_line])], ignore_index=True)
    return _clean_numeric_display(report, [column for column in columns if column != "Program"])


def _budget_report_008b(working: pd.DataFrame) -> pd.DataFrame:
    years = _fiscal_years(working)
    columns = ["Department", "Program"]
    for year in years:
        columns.extend([f"FY{year} Credits", f"FY{year} Δ vs prev"])
    if working.empty or not years:
        return pd.DataFrame(columns=columns)

    lines: list[dict[str, object]] = []
    for department in _department_order(working):
        dept_df = working.loc[working["Department"].eq(department)].copy()
        if dept_df.empty:
            continue
        for program in _program_columns(dept_df):
            program_df = dept_df.loc[dept_df["Program"].eq(program)].copy()
            if program_df.empty:
                continue
            line = {"Department": department, "Program": program}
            previous = 0.0
            for idx, year in enumerate(years):
                credits = _safe_sum(program_df.loc[program_df["Fiscal Year"].eq(year), "Cr."])
                line[f"FY{year} Credits"] = credits or pd.NA
                line[f"FY{year} Δ vs prev"] = (credits - previous) if idx > 0 else pd.NA
                previous = credits
            lines.append(line)
        total_line = {"Department": department, "Program": "DEPARTMENT TOTAL"}
        previous = 0.0
        for idx, year in enumerate(years):
            credits = _safe_sum(dept_df.loc[dept_df["Fiscal Year"].eq(year), "Cr."])
            total_line[f"FY{year} Credits"] = credits or pd.NA
            total_line[f"FY{year} Δ vs prev"] = (credits - previous) if idx > 0 else pd.NA
            previous = credits
        lines.append(total_line)
        lines.append(_blank_row(columns))
        lines.append(_blank_row(columns))
    report = pd.DataFrame(lines, columns=columns)
    return _clean_numeric_display(report, [column for column in columns if column not in {"Department", "Program"}])


def _budget_report_009(working: pd.DataFrame) -> pd.DataFrame:
    years = _fiscal_years(working)
    max_cols = max(1 + len(years) * 2, 1 + len(years))
    columns = [f"col_{idx}" for idx in range(1, max_cols + 1)]
    if working.empty or not years:
        return pd.DataFrame(columns=columns)

    lines: list[list[object]] = []
    load_df = working.loc[working["load/contract"].eq("load")].copy()
    contract_df = working.loc[working["load/contract"].eq("contract")].copy()

    def _pad(row: list[object]) -> list[object]:
        return row + [pd.NA] * (len(columns) - len(row))

    def _simple_header(label: str) -> list[object]:
        return _pad([label, *[f"FY {year}" for year in years]])

    def _contract_header(label: str) -> list[object]:
        row = [label]
        for year in years:
            row.extend([f"FY {year}", f"% contracts to load FY {year}"])
        return _pad(row)

    def _simple_data_row(program: str, df: pd.DataFrame) -> list[object]:
        return _pad([program, *[_safe_sum(df.loc[df["Fiscal Year"].eq(year), "Cr."]) for year in years]])

    def _contract_data_row(program: str, df: pd.DataFrame) -> list[object]:
        row = [program]
        for year in years:
            contract_value = _safe_sum(df.loc[df["Fiscal Year"].eq(year), "Cr."])
            load_value = _safe_sum(
                load_df.loc[
                    load_df["Program"].eq(program) & load_df["Fiscal Year"].eq(year),
                    "Cr.",
                ]
            )
            ratio = contract_value / load_value if load_value else 0
            row.extend([contract_value, ratio])
        return _pad(row)

    programs = _program_columns(working)
    lines.append(_pad(["COMPARISON OF CREDITS BUDGETED per PROGRAM"]))
    lines.append(_pad([]))
    lines.append(_simple_header("TOTAL LOAD"))
    for program in programs:
        program_df = load_df.loc[load_df["Program"].eq(program)].copy()
        lines.append(_simple_data_row(program, program_df))
    lines.append(_simple_data_row("TOTAL", load_df))
    lines.extend([_pad([]), _pad([]), _pad([])])

    lines.append(_contract_header("TOTAL CONTRACT"))
    for program in programs:
        program_df = contract_df.loc[contract_df["Program"].eq(program)].copy()
        lines.append(_contract_data_row(program, program_df))
    total_contract_row = ["TOTAL"]
    for year in years:
        contract_value = _safe_sum(contract_df.loc[contract_df["Fiscal Year"].eq(year), "Cr."])
        load_value = _safe_sum(load_df.loc[load_df["Fiscal Year"].eq(year), "Cr."])
        total_contract_row.extend([contract_value, (contract_value / load_value) if load_value else 0])
    lines.append(_pad(total_contract_row))
    grand_total_row = ["GRAND TOTAL"]
    for year in years:
        grand_total_row.extend([_safe_sum(working.loc[working["Fiscal Year"].eq(year), "Cr."]), pd.NA])
    lines.append(_pad(grand_total_row))
    lines.extend([_pad([]), _pad([]), _pad([])])

    lines.append(_simple_header("TOTAL LOAD AND CONTRACT"))
    for program in programs:
        program_df = working.loc[working["Program"].eq(program)].copy()
        lines.append(_simple_data_row(program, program_df))
    lines.append(_simple_data_row("TOTAL", working))
    return _clean_numeric_display(pd.DataFrame(lines, columns=columns), columns)


def _budget_report_010(working: pd.DataFrame) -> pd.DataFrame:
    columns = ["Program", "NADEI_courses", "total_credits"]
    if working.empty:
        return pd.DataFrame(columns=columns)
    nadei = working.loc[_string_column(working, "Faculty Name").str.startswith("NADEI:", na=False)].copy()
    if nadei.empty:
        return pd.DataFrame(columns=columns)
    report = (
        nadei.groupby("Program", dropna=False)
        .agg(
            NADEI_courses=("Program", "size"),
            total_credits=("Cr.", "sum"),
        )
        .reset_index()
        .sort_values("Program", na_position="last")
    )
    return _clean_numeric_display(report, ["NADEI_courses", "total_credits"])


def _location_abbreviation(value: object) -> str:
    text = _safe_text(value)
    if not text:
        return ""
    if " - " in text:
        return text.rsplit(" - ", 1)[-1].strip()
    if text.upper() in {"MC", "ASYNC", "SYNC", "AU", "MC/DUAL", "DUAL"}:
        return text
    return text


def _contract_mode_location(row: pd.Series) -> str:
    """Use the DonSheet instruction-method and building fields for contracts."""
    instruction_method = (
        _safe_text(row.get("Mode Raw"))
        or _safe_text(row.get("Instruction Method {Inst Method}"))
        or _safe_text(row.get("Mode"))
    )
    building = _safe_text(row.get("Building {Meet Bldg}")) or _safe_text(row.get("Building"))
    if not building:
        building = _location_abbreviation(row.get("Location"))
    return " ".join(part for part in [instruction_method, building] if part)


def _course_display_for_contract(row: pd.Series) -> str:
    pieces = [_safe_text(row.get("Faculty Name"))]
    course = _safe_text(row.get("Course Text"))
    title = _safe_text(row.get("Catalog Title"))
    if course:
        pieces.append(course)
    if title:
        pieces.append(title)
    return " ".join(piece for piece in pieces if piece)


def _is_contract_value(value: object) -> bool:
    return _safe_text(value).lower() == "contract"


def _write_contract_block(ws, rows_df: pd.DataFrame, slot_rows: dict[str, list[int]], semesters: list[str]) -> None:
    sem_to_label = {semester: semester for semester in semesters}
    for semester, rows in slot_rows.items():
        label_row = rows[0] - 1
        ws.cell(label_row, 2).value = sem_to_label.get(semester, semester)
        semester_df = rows_df.loc[rows_df["Semester"].eq(semester)].copy()
        semester_df = semester_df.sort_values(["Faculty Name", "Course Text"], na_position="last")
        for target_row, (_, row) in zip(rows, semester_df.iterrows()):
            ws.cell(target_row, 2).value = _course_display_for_contract(row)
            ws.cell(target_row, 3).value = _location_abbreviation(row.get("Location"))
            ws.cell(target_row, 4).value = row.get("Cr.")
            ws.cell(target_row, 5).value = row.get("rate per credit")
            ws.cell(target_row, 7).value = f"=+F{target_row}"
            ws.cell(target_row, 8).value = None


def generate_department_contracts_overview_report(df: pd.DataFrame, department: str) -> ReportResult:
    working = _filter_department(_prep_working_df(df), department)
    top_semesters = _department_cycle_semesters(working["Semester"]) if not working.empty else []
    start_year = _academic_year_sort_key(top_semesters[0])[0] if top_semesters else None
    workbook = _build_contracts_overview_template(start_year)
    ws = workbook.active

    if top_semesters:
        ws["B4"] = f"May {start_year}-April {start_year + 1}"
        ws["B26"] = f"May {start_year} - April {start_year + 1}"
    semester_rows = {7: f"Summer {start_year}" if start_year is not None else "", 8: f"Fall {start_year}" if start_year is not None else "", 9: f"Spring {start_year + 1}" if start_year is not None else ""}
    for row_idx, semester in semester_rows.items():
        ws.cell(row_idx, 2).value = f"Credits on {semester} schedule" if semester else None
        ws.cell(row_idx, 6).value = _safe_sum(working.loc[working["Semester"].eq(semester), "Cr."]) if semester else None

    load_df = working.loc[working["load/contract"].eq("load")].copy()
    salaried = (
        load_df.groupby("Faculty Name", dropna=False)["Cr."].sum().reset_index().sort_values(["Cr.", "Faculty Name"], ascending=[False, True])
        if not load_df.empty
        else pd.DataFrame(columns=["Faculty Name", "Cr."])
    )
    for row_idx in range(14, 19):
        ws.cell(row_idx, 2).value = None
        ws.cell(row_idx, 6).value = None
    for row_idx, (_, row) in zip(range(14, 19), salaried.iterrows()):
        ws.cell(row_idx, 2).value = row["Faculty Name"]
        ws.cell(row_idx, 6).value = row["Cr."]

    contract_df = working.loc[working["load/contract"].eq("contract")].copy()
    extension_programs = {
        "MAPM-ENGL": "mapm_engl",
        "MAPM-HISP": "mapm_hisp",
        "MAR-INT": "ma_religion",
        "MDivHISP_MAPmENGL_MAPmHISP": "mdiv_spanish",
    }
    extension_mask = contract_df["Program"].isin(extension_programs)
    on_campus_df = contract_df.loc[~extension_mask].copy()
    start_year = start_year or 0
    on_campus_slots = {
        f"Summer {start_year}": [35, 36, 37],
        f"Fall {start_year}": [40, 41, 42, 43, 44, 45],
        f"Spring {start_year + 1}": [48, 49, 50, 51, 52, 53],
    }
    _write_contract_block(ws, on_campus_df, on_campus_slots, top_semesters)

    extension_sections = {
        "MAPM-ENGL": {f"Fall {start_year}": [60, 61], f"Spring {start_year + 1}": [63, 64, 65, 66]},
        "MAPM-HISP": {f"Summer {start_year}": [69, 70], f"Fall {start_year}": [72, 73, 74], f"Spring {start_year + 1}": [76, 77]},
        "MAR-INT": {f"Fall {start_year}": [81, 82, 83, 84]},
        "MDivHISP_MAPmENGL_MAPmHISP": {f"Fall {start_year}": [88, 89, 90, 91, 92]},
    }
    for program, slot_rows in extension_sections.items():
        program_df = contract_df.loc[contract_df["Program"].eq(program)].copy()
        _write_contract_block(ws, program_df, slot_rows, top_semesters)

    output = BytesIO()
    workbook.save(output)
    preview_df = contract_df.loc[
        :,
        [column for column in ["Semester", "Faculty Name", "Course Text", "Program", "Location", "Cr.", "rate per credit", "total contract amount"] if column in contract_df.columns],
    ].copy()
    preview_df = preview_df.rename(columns={"Course Text": "Course"})
    return ReportResult(
        report_name=f"{department} Contracts Overview",
        preview_df=preview_df,
        excel_bytes=output.getvalue(),
        output_filename=f"{department}_ContractsOverview.xlsx",
        worksheet_name="Contracts",
        metadata={"department": department},
    )


def _course_schedule_label(row: pd.Series) -> str:
    """Build the course label shown in the department teaching schedule."""
    course = _safe_text(row.get("Course Text"))
    title = _safe_text(row.get("Catalog Title"))
    return " ".join(part for part in [course, title] if part)


def _is_off_campus_teaching_appointment(row: pd.Series) -> bool:
    """Match the reference schedule's red-font treatment for non-main-campus courses."""
    campus = _safe_text(row.get("Campus Code")).upper()
    if not campus:
        campus = _safe_text(row.get("Location")).upper()
    return campus not in {"", "0", "0.0", "7", "7.0", "MC", "MAIN CAMPUS"}


def _summer_part_of_term_marker(row: pd.Series, term_number: int) -> str:
    """Return an x when a Summer course is assigned to the requested term."""
    value = _safe_text(row.get("Part-Of-Term")) or _safe_text(row.get("Part-of-Term"))
    term_values = set(re.findall(r"\d+", value))
    return "x" if str(term_number) in term_values else ""


def _add_department_budget_course_schedule(workbook, working: pd.DataFrame, department: str) -> None:
    """Add a faculty-by-semester teaching-load schedule to an overview workbook."""
    sheet_name = "Budget Course Schedule A-B"
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    ws = workbook.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    schedule = working.copy()
    if not schedule.empty:
        schedule = schedule.loc[
            _string_column(schedule, "Faculty Name").notna()
            & _string_column(schedule, "Faculty Name").ne("")
            & _string_column(schedule, "Semester").notna()
            & _string_column(schedule, "Semester").ne("")
        ].copy()

    semesters = _latest_department_cycle_semesters(schedule["Semester"]) if not schedule.empty else []
    semester_layouts: list[dict[str, object]] = []
    next_column = 2
    for semester in semesters:
        is_summer = semester.lower().startswith("summer ")
        labels = ["Course", "Term 1", "Term 2", "Term 3", "Mode", "Location", "Program", "Cr."] if is_summer else [
            "Course", "Mode", "Location", "Program", "Cr."
        ]
        semester_layouts.append(
            {
                "semester": semester,
                "start_column": next_column,
                "labels": labels,
                "credit_column": next_column + len(labels) - 1,
                "is_summer": is_summer,
            }
        )
        next_column += len(labels)
    total_column = next_column
    last_column_letter = get_column_letter(total_column)

    ws.merge_cells(f"A1:{last_column_letter}1")
    ws["A1"] = f"{department} Budget Course Schedule A-B"
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells(f"A2:{last_column_letter}2")
    ws["A2"] = ", ".join(semesters) if semesters else "No scheduled courses found"
    ws["A2"].font = Font(name="Arial", size=10, italic=True)
    ws["A2"].alignment = Alignment(horizontal="center")
    ws["A3"] = "Red = Off-campus Teaching Appointments"
    ws["A3"].font = Font(name="Arial", size=10, bold=True, color="FF0000")

    if not semesters:
        ws["A4"] = "No course rows with both an instructor and semester were found for this department."
        ws.column_dimensions["A"].width = 78
        return

    header_fill = PatternFill("solid", fgColor="1F4E78")
    semester_fills = {
        "summer": PatternFill("solid", fgColor="FFF475"),
        "fall": PatternFill("solid", fgColor="FF99CC"),
        "spring": PatternFill("solid", fgColor="CCFF90"),
    }
    faculty_fill = PatternFill("solid", fgColor="92D050")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    semester_header_font = Font(name="Arial", size=10, bold=True, color="000000")
    faculty_font = Font(name="Arial", size=10, bold=True)
    off_campus_font = Font(name="Arial", size=10, color="FF0000")
    thin_border = Border(bottom=Side(style="thin", color="B7B7B7"))

    # The load and contract schedules are separate sections.  Use the same
    # visual treatment as the Faculty Name header to make each section clear.
    for column in range(1, total_column + 1):
        ws.cell(4, column).fill = header_fill
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=total_column)
    ws["A4"] = "Load Teachers:"
    ws["A4"].font = header_font
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A5:A6")
    ws["A5"] = "Faculty Name"
    ws["A5"].fill = header_fill
    ws["A5"].font = header_font
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")

    for layout in semester_layouts:
        semester = layout["semester"]
        start_column = layout["start_column"]
        labels = layout["labels"]
        end_column = start_column + len(labels) - 1
        ws.merge_cells(start_row=5, start_column=start_column, end_row=5, end_column=end_column)
        semester_header = ws.cell(5, start_column, semester)
        season = semester.split(" ", 1)[0].lower()
        semester_fill = semester_fills.get(season, header_fill)
        semester_header.fill = semester_fill
        semester_header.font = semester_header_font
        semester_header.alignment = Alignment(horizontal="center", vertical="center")
        for column, label in enumerate(labels, start=start_column):
            cell = ws.cell(6, column, label)
            cell.fill = semester_fill
            cell.font = semester_header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells(start_row=5, start_column=total_column, end_row=6, end_column=total_column)
    total_header = ws.cell(5, total_column, "Total Cr.")
    total_header.fill = header_fill
    total_header.font = header_font
    total_header.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 7
    def write_faculty_section(section: pd.DataFrame, start_row: int) -> int:
        faculty_names = sorted(
            _string_column(section, "Faculty Name").dropna().unique().tolist(),
            key=lambda value: str(value).casefold(),
        )
        current = start_row
        for faculty_name in faculty_names:
            faculty_rows = section.loc[_string_column(section, "Faculty Name").eq(faculty_name)].copy()
            semester_rows = {
                semester: faculty_rows.loc[faculty_rows["Semester"].eq(semester)].sort_values(
                    ["Course Text", "Catalog Title"], na_position="last"
                )
                for semester in semesters
            }
            course_row_count = max(len(rows) for rows in semester_rows.values())
            course_start_row = current + 1
            course_end_row = current + course_row_count

            for column in range(1, total_column + 1):
                cell = ws.cell(current, column)
                cell.fill = faculty_fill
                cell.font = faculty_font
                cell.border = thin_border
            ws.cell(current, 1, faculty_name)

            total_cells: list[str] = []
            for layout in semester_layouts:
                semester = layout["semester"]
                credit_column = layout["credit_column"]
                credit_cell = ws.cell(current, credit_column)
                if not semester_rows[semester].empty:
                    credit_cell.value = f"=SUM({get_column_letter(credit_column)}{course_start_row}:{get_column_letter(credit_column)}{course_end_row})"
                    total_cells.append(credit_cell.coordinate)
                credit_cell.number_format = "0.##"
                credit_cell.alignment = Alignment(horizontal="center")
            ws.cell(current, total_column).value = f"=SUM({','.join(total_cells)})" if total_cells else 0
            ws.cell(current, total_column).number_format = "0.##"
            ws.cell(current, total_column).alignment = Alignment(horizontal="center")

            for offset in range(course_row_count):
                target_row = course_start_row + offset
                for layout in semester_layouts:
                    semester = layout["semester"]
                    entries = semester_rows[semester]
                    if offset >= len(entries):
                        continue
                    entry = entries.iloc[offset]
                    start_column = layout["start_column"]
                    values = [_course_schedule_label(entry)]
                    if layout["is_summer"]:
                        values.extend(_summer_part_of_term_marker(entry, term_number) for term_number in range(1, 4))
                    values.extend(
                        [
                            _safe_text(entry.get("Mode")),
                            _location_abbreviation(entry.get("Location")),
                            _safe_text(entry.get("Program")),
                            entry.get("Cr."),
                        ]
                    )
                    for column, value in enumerate(values, start=start_column):
                        cell = ws.cell(target_row, column, value)
                        cell.alignment = Alignment(vertical="top", wrap_text=column == start_column)
                        if _is_off_campus_teaching_appointment(entry):
                            cell.font = off_campus_font
                        if column == layout["credit_column"]:
                            cell.number_format = "0.##"
                            cell.alignment = Alignment(horizontal="center", vertical="top")
            current = course_end_row + 2
        return current

    load_schedule = schedule.loc[schedule["load/contract"].eq("load")].copy()
    contract_schedule = schedule.loc[schedule["load/contract"].eq("contract")].copy()
    current_row = write_faculty_section(load_schedule, current_row)
    if not contract_schedule.empty:
        # The intentionally unused row above the heading stays white as a
        # visual break between load and contract teachers.
        for column in range(1, total_column + 1):
            ws.cell(current_row - 1, column).fill = PatternFill(fill_type=None)
        for column in range(1, total_column + 1):
            ws.cell(current_row, column).fill = header_fill
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=total_column)
        contract_heading = ws.cell(current_row, 1, "Contract Teachers:")
        contract_heading.font = Font(name="Arial", size=11, bold=True, color="FF0000")
        contract_heading.alignment = Alignment(horizontal="left", vertical="center")

        # Repeat the two schedule-header rows for the contract section.
        contract_header_row = current_row + 1
        for source_row, target_row in [(5, contract_header_row), (6, contract_header_row + 1)]:
            for column in range(1, total_column + 1):
                source_cell = ws.cell(source_row, column)
                target_cell = ws.cell(target_row, column)
                target_cell._style = copy(source_cell._style)
                if source_cell.value is not None:
                    target_cell.value = source_cell.value
        ws.merge_cells(start_row=contract_header_row, start_column=1, end_row=contract_header_row + 1, end_column=1)
        for layout in semester_layouts:
            start_column = layout["start_column"]
            end_column = start_column + len(layout["labels"]) - 1
            ws.merge_cells(
                start_row=contract_header_row,
                start_column=start_column,
                end_row=contract_header_row,
                end_column=end_column,
            )
        ws.merge_cells(
            start_row=contract_header_row,
            start_column=total_column,
            end_row=contract_header_row + 1,
            end_column=total_column,
        )
        current_row = write_faculty_section(contract_schedule, contract_header_row + 2)

    last_filled_row = current_row - 2
    for layout in semester_layouts:
        semester = layout["semester"]
        season = semester.split(" ", 1)[0].lower()
        credit_fill = semester_fills.get(season, header_fill)
        credit_column = layout["credit_column"]
        for row_number in range(5, last_filled_row + 1):
            ws.cell(row_number, credit_column).fill = credit_fill

    ws.freeze_panes = "B7"
    ws.column_dimensions["A"].width = 24
    for layout in semester_layouts:
        start_column = layout["start_column"]
        widths = [34, 10, 10, 10, 14, 14, 18, 8] if layout["is_summer"] else [34, 14, 14, 18, 8]
        for column, width in enumerate(widths, start=start_column):
            ws.column_dimensions[get_column_letter(column)].width = width
    ws.column_dimensions[last_column_letter].width = 11


def _add_department_contracts_sheet(workbook, working: pd.DataFrame, department: str) -> None:
    """Add the department's budget-request style contract schedule."""
    sheet_name = f"{department} CONTRACTS"[:31]
    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]
    ws = workbook.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    source = working.copy()
    for column in ["Semester", "load/contract", "Faculty Name", "Cr.", "rate per credit"]:
        if column not in source.columns:
            source[column] = pd.NA
    semesters = _latest_department_cycle_semesters(source["Semester"]) if not source.empty else []
    start_year = _academic_year_sort_key(semesters[0])[0] if semesters else None
    period_label = f"May {start_year} - April {start_year + 1}" if start_year is not None else ""
    all_cycle_rows = source.loc[source["Semester"].isin(semesters)].copy() if semesters else source.iloc[0:0].copy()
    contract_rows = all_cycle_rows.loc[all_cycle_rows["load/contract"].eq("contract")].copy()

    for row, text, italic in [
        (2, "ANDREWS UNIVERSITY", False),
        (3, "Budget Request", False),
        (4, period_label, False),
        (5, "SCHEDULE OF DETAILS", True),
    ]:
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
        cell = ws.cell(row, 2, text)
        cell.font = Font(name="Arial", size=11, bold=True, italic=italic)
        cell.alignment = Alignment(horizontal="center")
    ws["B7"] = f"{department} Department"
    ws["B7"].font = Font(name="Arial", size=10, bold=True)

    currency_format = '"$"#,##0.00;[Red]("$"#,##0.00);-'
    credit_format = '#,##0.0;(#,##0.0);-'
    signed_credit_format = '#,##0.0;-#,##0.0;-'
    aqua_blue_fill = PatternFill("solid", fgColor="00FFFF")
    light_green_fill = PatternFill("solid", fgColor="CCFF90")
    summary_fill = PatternFill("solid", fgColor="D9EAF7")
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin_bottom = Side(style="thin", color="7F7F7F")
    medium_bottom = Side(style="medium", color="000000")
    medium_top = Side(style="medium", color="000000")

    for offset, semester in enumerate(semesters, start=9):
        ws.cell(offset, 2, f"Credits on {semester} schedule").font = Font(name="Arial", size=10, bold=True)
        ws.cell(offset, 3, _safe_sum(all_cycle_rows.loc[all_cycle_rows["Semester"].eq(semester), "Cr."]))
        ws.cell(offset, 3).number_format = credit_format
        ws.cell(offset, 3).fill = aqua_blue_fill
        ws.cell(offset, 3).alignment = Alignment(horizontal="right")
        ws.cell(offset, 3).border = Border(bottom=thin_bottom)
    total_summary_row = 9 + len(semesters)
    ws.cell(total_summary_row, 2, "Total proposed credits offered for the year").font = Font(name="Arial", size=10, bold=True)
    ws.cell(total_summary_row, 3, f"=SUM(C9:C{total_summary_row - 1})" if semesters else 0)
    ws.cell(total_summary_row, 3).number_format = credit_format
    ws.cell(total_summary_row, 3).fill = aqua_blue_fill
    ws.cell(total_summary_row, 3).alignment = Alignment(horizontal="right")
    ws.cell(total_summary_row, 3).font = Font(name="Arial", size=10, bold=True)

    faculty_header_row = total_summary_row + 3
    faculty_headers = [
        "Faculty Name",
        "salaried/non-salaried",
        "Load Reduction Reason",
        "Load/Policy",
        "Load Credits",
        "Contract Credits",
        "Total Credits",
        "Long/short of Load",
    ]
    for column, label in enumerate(faculty_headers, start=2):
        cell = ws.cell(faculty_header_row, column, label)
        cell.fill = header_fill
        cell.font = Font(name="Arial", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(bottom=thin_bottom)

    faculty_rows = all_cycle_rows.loc[all_cycle_rows["Faculty Name"].notna() & all_cycle_rows["Faculty Name"].ne("")].copy()
    salaried_names, salaried_ids = _salaried_faculty_lookup()
    faculty_entries: list[tuple[str, bool]] = []
    for faculty_name in _string_column(faculty_rows, "Faculty Name").dropna().unique().tolist():
        person_rows = faculty_rows.loc[faculty_rows["Faculty Name"].eq(faculty_name)]
        faculty_ids = person_rows["ID#"] if "ID#" in person_rows.columns else pd.Series(dtype="string")
        faculty_entries.append(
            (
                faculty_name,
                _is_salaried_faculty(faculty_name, faculty_ids, salaried_names, salaried_ids),
            )
        )
    # Salaried faculty must appear first; each group is alphabetized by name.
    faculty_entries.sort(key=lambda entry: (not entry[1], str(entry[0]).casefold()))
    row_number = faculty_header_row + 1
    for faculty_name, is_salaried in faculty_entries:
        person_rows = faculty_rows.loc[faculty_rows["Faculty Name"].eq(faculty_name)]
        load_credits = _safe_sum(person_rows.loc[person_rows["load/contract"].eq("load"), "Cr."])
        contract_credits = _safe_sum(person_rows.loc[person_rows["load/contract"].eq("contract"), "Cr."])
        if not load_credits and not contract_credits:
            continue
        ws.cell(row_number, 2, faculty_name)
        ws.cell(row_number, 3, "salaried" if is_salaried else "non-salaried")
        # Columns D and E intentionally remain blank for the department
        # secretary to enter the reduction reason and approved load/policy.
        ws.cell(row_number, 6, load_credits)
        ws.cell(row_number, 7, contract_credits)
        ws.cell(row_number, 8, f"=SUM(F{row_number}:G{row_number})")
        ws.cell(row_number, 9, f"=E{row_number}-H{row_number}")
        for column in [6, 7, 8, 9]:
            ws.cell(row_number, column).number_format = credit_format
            ws.cell(row_number, column).alignment = Alignment(horizontal="center")
        ws.cell(row_number, 9).number_format = signed_credit_format
        row_number += 1

    first_faculty_row = faculty_header_row + 1
    last_faculty_row = row_number - 1
    if last_faculty_row >= first_faculty_row:
        status_validation = DataValidation(type="list", formula1='"salaried,non-salaried"', allow_blank=True)
        status_validation.error = "Choose salaried or non-salaried."
        status_validation.errorTitle = "Invalid employment status"
        ws.add_data_validation(status_validation)
        status_validation.add(f"C{first_faculty_row}:C{last_faculty_row}")

        reduction_validation = DataValidation(type="list", formula1='"AUSS,Chair,25%,33%,50%,66%"', allow_blank=True)
        reduction_validation.error = "Choose a listed load reduction reason."
        reduction_validation.errorTitle = "Invalid load reduction reason"
        ws.add_data_validation(reduction_validation)
        reduction_validation.add(f"D{first_faculty_row}:D{last_faculty_row}")

    salaried_total_row = row_number
    non_salaried_total_row = row_number + 3
    faculty_total_row = row_number + 4
    total_rows = [
        (salaried_total_row, "Proposed credits to be taught by salaried Seminary faculty", "salaried", "H"),
        (row_number + 1, "Proposed load credits to be taught by salaried professors", "salaried", "F"),
        (row_number + 2, "Proposed contract credits to be taught by salaried professors", "salaried", "G"),
        (non_salaried_total_row, "Proposed credits to be taught by non-salaried professors", "non-salaried", "H"),
    ]
    for total_row, label, status, credit_column in total_rows:
        ws.cell(total_row, 2, label).font = Font(name="Arial", size=10, bold=True)
        formula = f'=SUMIF(C{first_faculty_row}:C{last_faculty_row},"{status}",{credit_column}{first_faculty_row}:{credit_column}{last_faculty_row})'
        ws.cell(total_row, 3, formula if last_faculty_row >= first_faculty_row else 0)
        ws.cell(total_row, 3).number_format = credit_format
        ws.cell(total_row, 3).font = Font(name="Arial", size=10, bold=True)
        fill = PatternFill("solid", fgColor="99D6D0" if status == "salaried" else "FFC7CE")
        for column in (2, 3):
            ws.cell(total_row, column).fill = fill

    ws.cell(faculty_total_row, 2, "Proposed total credits").font = Font(name="Arial", size=10, bold=True)
    ws.cell(faculty_total_row, 3, f"=C{salaried_total_row}+C{non_salaried_total_row}")
    ws.cell(faculty_total_row, 3).number_format = credit_format
    ws.cell(faculty_total_row, 3).font = Font(name="Arial", size=10, bold=True)
    for column in range(2, 10):
        ws.cell(salaried_total_row, column).border = Border(top=medium_top)

    row_number = faculty_total_row + 4
    ws.merge_cells(start_row=row_number, start_column=2, end_row=row_number, end_column=9)
    ws.cell(row_number, 2, "Schedule 1(b) -- Contract Salaries").font = Font(name="Arial", size=11, bold=True)
    row_number += 2
    contract_header_row = row_number
    contract_headers = [
        "Instructor and Course",
        "salaried/non-salaried",
        "Mode / Location",
        "Credits Taught",
        "Rate per Credit",
        "Contract Amount",
        "Contracts",
    ]
    for column, label in enumerate(contract_headers, start=2):
        cell = ws.cell(row_number, column, label)
        cell.fill = header_fill
        cell.font = Font(name="Arial", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = Border(bottom=thin_bottom)
    row_number += 1

    def write_contract_section(section_title: str, rows: pd.DataFrame, start_row: int) -> tuple[int, int, int]:
        section_start = start_row
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=8)
        section_cell = ws.cell(start_row, 2, section_title)
        section_cell.font = Font(name="Arial", size=10, bold=True)
        section_cell.fill = summary_fill
        start_row += 1
        first_detail_row = start_row
        for semester in semesters:
            semester_rows = rows.loc[rows["Semester"].eq(semester)].sort_values(["Faculty Name", "Course Text"], na_position="last")
            if semester_rows.empty:
                continue
            ws.cell(start_row, 2, semester)
            ws.cell(start_row, 2).font = Font(name="Arial", size=10, bold=True)
            ws.cell(start_row, 2).border = Border(bottom=thin_bottom)
            start_row += 1
            for _, contract_row in semester_rows.iterrows():
                ws.cell(start_row, 2, " ".join(part for part in [_safe_text(contract_row.get("Faculty Name")), _course_schedule_label(contract_row)] if part))
                contract_ids = pd.Series([contract_row.get("ID#")], dtype="object")
                is_salaried = _is_salaried_faculty(
                    contract_row.get("Faculty Name"), contract_ids, salaried_names, salaried_ids
                )
                ws.cell(start_row, 3, "salaried" if is_salaried else "non-salaried")
                ws.cell(start_row, 4, _contract_mode_location(contract_row))
                ws.cell(start_row, 5, contract_row.get("Cr."))
                ws.cell(start_row, 6, contract_row.get("rate per credit"))
                ws.cell(start_row, 7, f"=E{start_row}*F{start_row}")
                ws.cell(start_row, 8, f"=G{start_row}")
                for column in range(2, 9):
                    ws.cell(start_row, column).border = Border(bottom=thin_bottom)
                ws.cell(start_row, 4).alignment = Alignment(horizontal="left")
                ws.cell(start_row, 5).number_format = credit_format
                for column in [6, 7, 8]:
                    ws.cell(start_row, column).number_format = currency_format
                start_row += 1
        subtotal_row = start_row
        ws.cell(subtotal_row, 4, f"Total {section_title}")
        ws.cell(subtotal_row, 4).font = Font(name="Arial", size=10, bold=True)
        for column in [5, 7, 8]:
            column_letter = get_column_letter(column)
            ws.cell(subtotal_row, column, f"=SUM({column_letter}{first_detail_row}:{column_letter}{subtotal_row - 1})" if subtotal_row > first_detail_row else 0)
            ws.cell(subtotal_row, column).font = Font(name="Arial", size=10, bold=True)
            ws.cell(subtotal_row, column).border = Border(bottom=medium_bottom)
        ws.cell(subtotal_row, 5).number_format = credit_format
        for column in [7, 8]:
            ws.cell(subtotal_row, column).number_format = currency_format
        return subtotal_row + 2, subtotal_row, section_start

    if contract_rows.empty:
        on_campus = contract_rows.copy()
        extension = contract_rows.copy()
    else:
        off_campus = contract_rows.apply(_is_off_campus_teaching_appointment, axis=1)
        on_campus = contract_rows.loc[~off_campus].copy()
        extension = contract_rows.loc[off_campus].copy()
    row_number, on_campus_total_row, _ = write_contract_section("On Campus Contracts", on_campus, row_number)
    row_number, extension_total_row, _ = write_contract_section("Extension School Classes", extension, row_number)

    ws.cell(row_number, 4, "Total Contracts")
    ws.cell(row_number, 4).font = Font(name="Arial", size=11, bold=True)
    for column in [5, 7, 8]:
        column_letter = get_column_letter(column)
        ws.cell(row_number, column, f"={column_letter}{on_campus_total_row}+{column_letter}{extension_total_row}")
        ws.cell(row_number, column).font = Font(name="Arial", size=11, bold=True)
        ws.cell(row_number, column).border = Border(bottom=medium_bottom)
    ws.cell(row_number, 5).number_format = credit_format
    for column in [7, 8]:
        ws.cell(row_number, column).number_format = currency_format

    for detail_row in range(contract_header_row, row_number + 1):
        ws.cell(detail_row, 7).fill = aqua_blue_fill
        ws.cell(detail_row, 8).fill = light_green_fill

    for column, width in {
        "A": 3,
        "B": 40,
        "C": 22,
        "D": 24,
        "E": 14,
        "F": 14,
        "G": 16,
        "H": 18,
        "I": 18,
    }.items():
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "B15"


def generate_department_overview_bundle(
    df: pd.DataFrame,
    department: str,
    donsheet_payloads: list[tuple[str, bytes]] | None = None,
) -> ReportResult:
    working = _filter_department(_prep_working_df(df), department)
    sheets = {
        "report_001a": _overview_report_001a(working),
        "report_001b": _overview_report_001b(working),
        "report_001c": _overview_report_001c(working),
        "report_001f": _overview_report_001f(working),
        "report_002a": _overview_report_002a(working),
        "report_003a": _overview_report_003a(working),
    }
    workbook_bytes = workbook_to_excel_bytes(sheets, style_plan={key: key for key in sheets})
    contracts_report = generate_department_contracts_overview_report(df, department)
    workbook = load_workbook(BytesIO(workbook_bytes))
    contracts_workbook = load_workbook(BytesIO(contracts_report.excel_bytes))
    _clone_sheet_into_workbook(contracts_workbook.active, workbook, "ContractsOverview")
    if donsheet_payloads:
        schedule_report = generate_schedule_outputs_from_payloads(
            donsheet_payloads,
            source_name=f"{department}_department_admin",
            department=department,
            weekly_room_order=DEPARTMENT_ADMIN_WEEKLY_ROOMS,
            weekly_class_column_width=DEPARTMENT_ADMIN_SCHEDULE_COLUMN_WIDTH,
            sync_class_column_width=28,
        )
        schedule_workbook = load_workbook(BytesIO(schedule_report.excel_bytes))
        for sheet_name in ["Weekly Schedule", "SYNC Schedule", "ASYNC Classes"]:
            if sheet_name in schedule_workbook.sheetnames:
                _clone_sheet_into_workbook(schedule_workbook[sheet_name], workbook, sheet_name)
        if "Weekly Schedule" in workbook.sheetnames:
            _set_schedule_class_column_widths(workbook["Weekly Schedule"], DEPARTMENT_ADMIN_SCHEDULE_COLUMN_WIDTH)
        if "SYNC Schedule" in workbook.sheetnames:
            _set_schedule_class_column_widths(workbook["SYNC Schedule"], 28)
    _add_department_budget_course_schedule(workbook, working, department)
    _add_department_contracts_sheet(workbook, working, department)
    # Put the two department planning sheets first so recipients see the
    # actionable contracts and teaching schedule before the report tabs.
    first_sheets = [f"{department} CONTRACTS"[:31], "Budget Course Schedule A-B"]
    for target_index, sheet_name in enumerate(first_sheets):
        if sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            workbook.move_sheet(worksheet, offset=target_index - workbook.index(worksheet))
    combined_buffer = BytesIO()
    workbook.save(combined_buffer)
    preview_df = sheets["report_001a"]
    return ReportResult(
        report_name=f"{department} Department Overview Reports",
        preview_df=preview_df,
        excel_bytes=combined_buffer.getvalue(),
        output_filename=f"{department}_department_overview_reports.xlsx",
        worksheet_name="report_001a",
        metadata={"department": department, "sheet_count": str(len(workbook.sheetnames))},
    )


def generate_contract_forms_demo(df: pd.DataFrame) -> ReportResult:
    working = _prep_working_df(df)
    if not working.empty:
        working = working.loc[working["load/contract"].eq("contract")].copy()
    selected_columns = [
        "Department",
        "Semester",
        "Faculty Name",
        "ID#",
        "email",
        "Course",
        "Catalog Title",
        "Cr.",
        "rate per credit",
        "total contract amount",
        "Begin Date",
        "End Date",
        "account to be charged",
        "Location",
        "Mode",
        "Source Workbook",
        "Source Sheet",
    ]
    preview_df = working.loc[:, [c for c in selected_columns if c in working.columns]].copy()
    preview_df = preview_df.rename(columns={"Cr.": "Credits"})
    sheets = {
        "contract_rows": preview_df,
        "summary_by_faculty": (
            preview_df.groupby(["Faculty Name", "Semester"], dropna=False)
            .agg(
                Credits=("Credits", "sum"),
                Contract_Cost=("total contract amount", "sum"),
                Course_Count=("Course", lambda s: s.dropna().nunique()),
            )
            .reset_index()
            .rename(columns={"Contract_Cost": "Contract Cost", "Course_Count": "Course Count"})
            .sort_values(["Faculty Name", "Semester"], na_position="last")
            if not preview_df.empty
            else pd.DataFrame(columns=["Faculty Name", "Semester", "Credits", "Contract Cost", "Course Count"])
        ),
    }
    return ReportResult(
        report_name="Contract Forms Demo Export",
        preview_df=preview_df,
        excel_bytes=workbook_to_excel_bytes(sheets),
        output_filename="contract_forms_demo_from_donsheet.xlsx",
        worksheet_name="contract_rows",
        metadata={"rows": str(len(preview_df))},
    )


def generate_schedule_demo(df: pd.DataFrame) -> ReportResult:
    working = _prep_working_df(df)
    day_columns = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]
    for column in day_columns:
        working[column] = _string_column(working, column)
    schedule_columns = [
        "Department",
        "Semester",
        "Course",
        "Catalog Title",
        "Faculty Name",
        "Mode",
        "Location",
        "Course Beginning Time {Meet Beg Time}",
        "Course Ending Time {Meet End Time}",
        *day_columns,
        "Room {Meet Room}",
        "Building {Meet Bldg}",
        "Source Sheet",
    ]
    preview_df = working.loc[:, [c for c in schedule_columns if c in working.columns]].copy()
    preview_df = preview_df.rename(
        columns={
            "Course Beginning Time {Meet Beg Time}": "Begin Time",
            "Course Ending Time {Meet End Time}": "End Time",
            "Room {Meet Room}": "Room",
            "Building {Meet Bldg}": "Building",
        }
    )
    preview_df = preview_df.sort_values(["Semester", "Department", "Course"], na_position="last")
    summary_df = (
        preview_df.groupby(["Semester", "Department"], dropna=False)
        .agg(
            Course_Count=("Course", lambda s: s.dropna().nunique()),
            Faculty_Count=("Faculty Name", lambda s: s.dropna().nunique()),
        )
        .reset_index()
        .rename(columns={"Course_Count": "Course Count", "Faculty_Count": "Faculty Count"})
        if not preview_df.empty
        else pd.DataFrame(columns=["Semester", "Department", "Course Count", "Faculty Count"])
    )
    return ReportResult(
        report_name="SmartSchedule Demo Export",
        preview_df=preview_df,
        excel_bytes=workbook_to_excel_bytes({"schedule_rows": preview_df, "schedule_summary": summary_df}),
        output_filename="smart_schedule_demo_from_donsheet.xlsx",
        worksheet_name="schedule_rows",
        metadata={"rows": str(len(preview_df))},
    )


def _budget_report_master(working: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Source Workbook",
        "Department",
        "Semester",
        "Fiscal Year",
        "Faculty Name",
        "Course",
        "Catalog Title",
        "Program",
        "Mode",
        "Location",
        "load/contract",
        "Cr.",
        "rate per credit",
        "total contract amount",
        "account to be charged",
    ]
    # A workbook can legitimately normalize to no usable course rows (for
    # example, when every row is marked for deletion).  Keep the Master Data
    # report structurally valid in that case instead of sorting by columns that
    # are not present on an empty source frame.
    if working.empty:
        return pd.DataFrame(columns=columns)

    available = [column for column in columns if column in working.columns]
    report = working.loc[:, available].copy()
    sort_columns = [column for column in ["Department", "Semester", "Faculty Name", "Course"] if column in report]
    return report.sort_values(sort_columns, na_position="last") if sort_columns else report


def generate_budget_reports_bundle(df: pd.DataFrame, source_label: str = "donsheet") -> ReportBundle:
    working = _prep_working_df(df, department_from_sem=True)
    sheet_frames = {
        "report_000": _budget_report_master(working),
        "report_001a": _overview_report_001a(working),
        "report_001b": _overview_report_001b(working),
        "report_001c": _overview_report_001c(working),
        "report_001f": _overview_report_001f_all_departments(working),
        "report_002a": _overview_report_002a(working),
        "report_002b": _budget_report_002b(working),
        "report_003a": _overview_report_003a(working),
        "report_003b": _budget_report_003b(working),
        "report_004a": _budget_report_004a(working),
        "report_004b": _budget_report_004b(working),
        "report_005": _budget_report_005(working),
        "report_006": _budget_report_006(working),
        "report_007": _budget_report_007(working),
        "report_008a": _budget_report_008a(working),
        "report_008b": _budget_report_008b(working),
        "report_009": _budget_report_009(working),
        "report_010": _budget_report_010(working),
    }
    report_titles = {
        "report_000": "Master Data",
        "report_001a": "Department Overview",
        "report_001b": "Department Contract Overview",
        "report_001c": "Department Overview by Load Type",
        "report_001f": "Department Faculty Year Overview",
        "report_002a": "Department Credit Cost Overview",
        "report_002b": "Department Courses by Program Credits",
        "report_003a": "Department Totals Load Contract",
        "report_003b": "Subject Totals Load Contract",
        "report_004a": "Program Credits by Department",
        "report_004b": "Program Load Credits by Department",
        "report_005": "Program Totals Load Contract",
        "report_006": "Program Contract Cost per Department",
        "report_007": "FY Comparison Department Totals",
        "report_008a": "Program Credits Year Pairs",
        "report_008b": "FY Comparison Department Program Credit",
        "report_009": "FY Comparison Program Load Credits",
        "report_010": "Teaching per Program",
    }
    reports: dict[str, ReportResult] = {}
    for key, frame in sheet_frames.items():
        worksheet_name = "Master Data" if key == "report_000" else key
        reports[key] = ReportResult(
            report_name=report_titles[key],
            preview_df=frame,
            excel_bytes=workbook_to_excel_bytes({worksheet_name: frame}, style_plan={worksheet_name: key}),
            output_filename=f"{key}_from_{source_label}.xlsx",
            worksheet_name=worksheet_name,
            metadata={"rows": str(len(frame))},
        )

    combined_frames = {
        "Master Data": sheet_frames["report_000"],
        **{key: frame for key, frame in sheet_frames.items() if key != "report_000"},
        "report_999": _budget_report_master(working),
    }
    combined_style_plan = {
        "Master Data": "report_000",
        **{key: key for key in combined_frames if key not in {"Master Data", "report_999"}},
        "report_999": "report_999",
    }
    return ReportBundle(
        bundle_name="Budget Reports",
        reports=reports,
        combined_workbook_bytes=workbook_to_excel_bytes(
            combined_frames,
            style_plan=combined_style_plan,
        ),
        combined_output_filename=f"report_999_budget_reports_from_{source_label}.xlsx",
        preview_report_key="report_001a",
        metadata={"sheet_count": str(len(combined_frames))},
    )
