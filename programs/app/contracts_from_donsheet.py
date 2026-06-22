from __future__ import annotations

import math
import re
import shutil
import tempfile
import warnings
import zipfile
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject
from reportlab.lib.colors import blue
from reportlab.pdfgen import canvas

from .config import HELPER_SHEETS, PROJECT_ROOT


DEFAULT_RATE_PER_CREDIT = 1050
OPENPYXL_WARNING = "Data Validation extension is not supported and will be removed"


@dataclass
class GeneratedContract:
    filename: str
    path: Path
    semester: str
    contract_type: str
    instructor: str
    course_code: str
    source_file: str = ""


@dataclass
class ContractGenerationResult:
    output_dir: Path
    zip_path: Path
    contracts: list[GeneratedContract]
    extracted_rows: pd.DataFrame
    records_processed: int


@dataclass
class ContractAppPaths:
    root: Path
    standard_template: Path
    intensive_template: Path


def resolve_app_paths(base_dir: Path | None = None) -> ContractAppPaths:
    root = (base_dir or PROJECT_ROOT).resolve()
    pdf_dir = root / "0_source_files" / "default_PDFcontract"
    return ContractAppPaths(
        root=root,
        standard_template=_resolve_pdf_template(pdf_dir, "standard"),
        intensive_template=_resolve_pdf_template(pdf_dir, "intensive"),
    )


def _resolve_pdf_template(pdf_dir: Path, template_kind: str) -> Path:
    direct_matches = _matching_pdf_templates(pdf_dir, template_kind)
    if direct_matches:
        return direct_matches[0]

    old_matches = _matching_pdf_templates(pdf_dir / "old", template_kind)
    if old_matches:
        return old_matches[0]

    raise FileNotFoundError(f"No PDF template with {template_kind!r} in its filename found in {pdf_dir}")


def _matching_pdf_templates(folder: Path, template_kind: str) -> list[Path]:
    if not folder.exists():
        return []
    kind = template_kind.lower()
    matches = [
        path
        for path in folder.glob("*.pdf")
        if path.is_file()
        and not path.name.startswith(("._", "~$"))
        and kind in path.name.lower()
    ]
    return sorted(matches, key=lambda path: (path.stat().st_mtime, path.name.lower()), reverse=True)


def generate_contracts_from_uploads(
    uploads: list[tuple[str, bytes]],
    base_dir: Path | None = None,
) -> ContractGenerationResult:
    paths = resolve_app_paths(base_dir)
    work_root = Path(tempfile.mkdtemp(prefix="smartseminary_contracts_"))
    output_dir = work_root / "generated_contracts"
    input_dir = work_root / "uploaded_donsheets"
    output_dir.mkdir(parents=True, exist_ok=True)
    input_dir.mkdir(parents=True, exist_ok=True)

    all_contracts: list[GeneratedContract] = []
    all_rows: list[pd.DataFrame] = []

    for uploaded_name, uploaded_bytes in uploads:
        input_path = input_dir / sanitize_filename(uploaded_name or "uploaded_donsheet.xlsx")
        input_path.write_bytes(uploaded_bytes)
        contract_rows = extract_contract_rows_from_donsheet(input_path)
        if not contract_rows.empty:
            contract_rows = contract_rows.copy()
            contract_rows["Source File"] = input_path.name
            all_rows.append(contract_rows)
            all_contracts.extend(
                generate_contract_pdfs(
                    contract_rows,
                    paths,
                    output_dir,
                    filename_prefix=f"{sanitize_filename(input_path.stem)}__",
                    source_file=input_path.name,
                )
            )

    extracted_rows = pd.concat(all_rows, ignore_index=True) if all_rows else pd.DataFrame()
    zip_path = create_zip_from_contracts(all_contracts, output_dir / "contracts.zip")
    return ContractGenerationResult(
        output_dir=output_dir,
        zip_path=zip_path,
        contracts=all_contracts,
        extracted_rows=extracted_rows,
        records_processed=len(extracted_rows),
    )


def cleanup_result(result: ContractGenerationResult) -> None:
    temp_root = result.output_dir.parent
    if temp_root.exists():
        shutil.rmtree(temp_root, ignore_errors=True)


def extract_contract_rows_from_donsheet(excel_path: Path) -> pd.DataFrame:
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=OPENPYXL_WARNING)
        excel_file = pd.ExcelFile(excel_path)

    rows: list[dict[str, object]] = []
    for sheet_name in excel_file.sheet_names:
        if sheet_name in HELPER_SHEETS:
            continue
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message=OPENPYXL_WARNING)
            sheet_df = pd.read_excel(excel_path, sheet_name=sheet_name)
        if sheet_df.empty:
            continue
        for row_index, row in sheet_df.iterrows():
            contract_row = extract_single_contract_row(row, sheet_name, row_index)
            if contract_row is not None:
                rows.append(contract_row)
    return pd.DataFrame(rows)


def extract_single_contract_row(row: pd.Series, sheet_name: str, row_index: int) -> dict[str, object] | None:
    load_contract = _normalize_load_contract(
        _coalesce(
            row,
            "load/contract {not in Banner}",
            "Reason for Contract {not in Banner}",
            "Enrollment Cap {Max Enrl}",
        )
    )
    if load_contract != "contract":
        return None

    instructor_name = _pick_instructor_name(row)
    instructor_email = _pick_email(row)
    instructor_id = _pick_identifier(row)
    semester, year = _pick_semester_year(row)
    department = _pick_department(row, sheet_name)
    account = _pick_account(row)

    prework_start = _normalize_date_value(_value(row, "Pre-work Start Date {not in Banner}"))
    prework_end = _normalize_date_value(_value(row, "Pre-work End Date {not in Banner}"))
    intensive_start = _normalize_date_value(_value(row, "Intensive Period Start Date {Meet Start Date}"))
    intensive_end = _normalize_date_value(_value(row, "Intensive Period End Date {Meet End Date}"))
    postwork_start = _normalize_date_value(_value(row, "Post-work Start Date {not in Banner}"))
    postwork_end = _normalize_date_value(_value(row, "Post-work End Date {not in Banner}"))
    begin_date = _first_date(
        prework_start,
        _normalize_date_value(_value(row, "Semester Start Date {Soaterm Start Date}")),
        intensive_start,
    )
    end_date = _last_date(
        postwork_end,
        intensive_end,
        _normalize_date_value(_value(row, "Semester End Date {Soaterm End Date}")),
    )

    dept = _stringify(_value(row, "Subject")) or department
    course_number = _stringify(_value(row, "Course Number {Crse Num}"))
    course_section = _format_section(_value(row, "Course Section {Seq Crse Num}"))
    contract_reason = _stringify(_value(row, "Reason for Contract {not in Banner}"))

    credits = _number(_value(row, "Credits {Sect Crs}"))
    rate = _number(_value(row, "costs per credit {not in Banner}"))
    total_contract_amount = _number(_value(row, "total costs {not in Banner}"))
    if credits and rate and not total_contract_amount:
        total_contract_amount = credits * rate
    explicit_total_hours = _number(_value(row, "total contract hours {not in Banner}"))

    course_row = {
        "Instructor": instructor_name,
        "ID": instructor_id,
        "Email": instructor_email,
        "Telephone": _stringify(_value(row, "telephone")),
        "Account": account,
        "Remote": _stringify(_value(row, "Remote Employee {not in Banner}")),
        "International": _stringify(_value(row, "international")),
        "DeptContact": _stringify(_value(row, "Dept contact name")),
        "DeptContactID": _stringify(_value(row, "Dept contact ID#")),
        "DeanID": _stringify(_value(row, "dean ID#")),
        "VPFinanceID": _stringify(_value(row, "VP finance ID#")),
        "Reason": contract_reason,
        "Semester": semester,
        "Year": year,
        "CourseRaw": f"{dept}{course_number}".strip(),
        "Dept": dept,
        "CourseNum": course_number,
        "CourseSection": course_section,
        "CourseTitle": _coalesce_text(row, "Catalog Title", "Section Title"),
        "Credits": credits,
        "Rate": rate,
        "DeptBudget": _stringify(_value(row, "dept budget {not in Banner}")),
        "BeginDate": begin_date,
        "EndDate": end_date,
        "prework_start": prework_start,
        "prework_end": prework_end,
        "prework_weeks": _number(_value(row, "pre-work # of weeks {not in Banner}")),
        "prework_hours_week": _number(_value(row, "pre-work hours/week {not in Banner}")),
        "prework_hours_period": _number(_value(row, "pre-work hours/period {not in Banner}")),
        "intensive_start": intensive_start,
        "intensive_end": intensive_end,
        "intensive_weeks": _number(_value(row, "intensive # of weeks {not in Banner}")),
        "intensive_hours_week": _number(_value(row, "intensive hours/week {not in Banner}")),
        "intensive_hours_period": _number(_value(row, "intensive hours/period {not in Banner}")),
        "postwork_start": postwork_start,
        "postwork_end": postwork_end,
        "postwork_weeks": _number(_value(row, "post-work # of weeks {not in Banner}")),
        "postwork_hours_week": _number(_value(row, "post-work hours/week {not in Banner}")),
        "postwork_hours_period": _number(_value(row, "post-work hours/period {not in Banner}")),
        "Enrollment": _stringify(_value(row, "Enrollment Cap {Max Enrl}")),
        "TotalContractHours": explicit_total_hours,
        "Source Sheet": sheet_name,
        "Source Row": row_index + 2,
        "Department": department,
        "load/contract": load_contract,
        "total contract amount": total_contract_amount,
    }
    return course_row


def _value(row: pd.Series, column: str) -> object:
    return row.get(column)


def _coalesce(row: pd.Series, *columns: str) -> object:
    for column in columns:
        value = row.get(column)
        if _has_value(value):
            return value
    return None


def _coalesce_text(row: pd.Series, *columns: str) -> str:
    return _stringify(_coalesce(row, *columns))


def _has_value(value: object) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except Exception:
        pass
    return str(value).strip() != ""


def _stringify(value: object) -> str:
    if not _has_value(value):
        return ""
    if isinstance(value, float) and float(value).is_integer():
        return str(int(value))
    if isinstance(value, pd.Timestamp):
        return value.strftime("%m/%d/%Y")
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def _number(value: object) -> float:
    if not _has_value(value):
        return 0
    try:
        return float(value)
    except Exception:
        text_value = str(value).strip()
        if "," in text_value:
            numeric_parts = re.findall(r"-?\d+(?:\.\d+)?", text_value)
            if numeric_parts:
                try:
                    return float(numeric_parts[-1])
                except Exception:
                    pass
        text = re.sub(r"[^\d.\-]", "", str(value))
        try:
            return float(text)
        except Exception:
            return 0


def _normalize_load_contract(value: object) -> str:
    text = _stringify(value).lower()
    if text in {"contract", "load"}:
        return text
    return ""


def _pick_instructor_name(row: pd.Series) -> str:
    for column in [
        "Instructor Email {Instr Email}",
        "Instructor Name {Instr Name}",
    ]:
        value = _stringify(_value(row, column))
        if "," in value and "@" not in value:
            return value
    return _stringify(_value(row, "Instructor Email {Instr Email}"))


def _pick_email(row: pd.Series) -> str:
    for column in [
        "Instructor ID {Instr ID}",
        "Instructor Email {Instr Email}",
    ]:
        value = _stringify(_value(row, column))
        if "@" in value:
            return value
    return ""


def _pick_identifier(row: pd.Series) -> str:
    for column in [
        "Remote Employee {not in Banner}",
        "Instructor ID {Instr ID}",
    ]:
        value = _stringify(_value(row, column))
        if value and "@" not in value and "," not in value:
            return value
    return ""


def _pick_semester_year(row: pd.Series) -> tuple[str, str]:
    season = _stringify(_value(row, "Semester {not in Banner}"))
    year = _stringify(_value(row, "Fiscal Year {not in Banner}"))
    if season in {"Spring", "Summer", "Fall"} and year.isdigit():
        if season in {"Summer", "Fall"}:
            return season, str(int(year) - 1)
        return season, year

    term_code = _stringify(_value(row, "Term Code"))
    if len(term_code) >= 5 and term_code[:4].isdigit():
        season_map = {"2": "Spring", "3": "Summer", "4": "Fall"}
        return season_map.get(term_code[4], ""), term_code[:4]
    return "", ""


def _pick_department(row: pd.Series, sheet_name: str) -> str:
    fee_amount = _stringify(_value(row, "Fee Amount {Fees Amt}"))
    if fee_amount:
        return fee_amount
    account_field = _stringify(_value(row, "SEM Department {Scacrse Dept}"))
    if ":" in account_field:
        return account_field.split(":", 1)[0].strip()
    return sheet_name


def _pick_account(row: pd.Series) -> str:
    return _coalesce_text(row, "account to be charged {not in Banner}", "SEM Department {Scacrse Dept}")


def _format_section(value: object) -> str:
    text = _stringify(value)
    if not text:
        return ""
    if text.isdigit() and len(text) < 3:
        return text.zfill(3)
    return text


def _normalize_date_value(value: object) -> object:
    if not _has_value(value):
        return None
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime(warn=False)


def _first_date(*values: object) -> object:
    dates = [pd.to_datetime(value, errors="coerce") for value in values if _has_value(value)]
    dates = [value for value in dates if not pd.isna(value)]
    if not dates:
        return None
    return min(dates).to_pydatetime()


def _last_date(*values: object) -> object:
    dates = [pd.to_datetime(value, errors="coerce") for value in values if _has_value(value)]
    dates = [value for value in dates if not pd.isna(value)]
    if not dates:
        return None
    return max(dates).to_pydatetime()


def create_zip_from_contracts(contracts: list[GeneratedContract], zip_path: Path) -> Path:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for contract in contracts:
            zip_file.write(contract.path, arcname=contract.filename)
    return zip_path


def split_name(name: str) -> tuple[str, str]:
    if not isinstance(name, str) or "," not in name:
        return str(name), ""
    last, first = name.split(",", 1)
    return last.strip(), first.strip()


def sanitize_filename(value: str) -> str:
    value = str(value or "").strip()
    value = re.sub(r"[^\w\-. ]+", "", value)
    value = re.sub(r"\s+", "_", value).strip("_")
    return value or "UNKNOWN"


def course_token(course: dict[str, object]) -> str:
    dept = _stringify(course.get("Dept"))
    number = _stringify(course.get("CourseNum"))
    section = _stringify(course.get("CourseSection"))
    token = f"{dept}{number}".strip()
    return f"{token}-{section}" if token and section else token


def has_intensive_period(course: dict[str, object]) -> bool:
    return _has_value(course.get("intensive_start")) or _has_value(course.get("intensive_end"))


def safe_int_string(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number)
    except Exception:
        return ""


def safe_number(value: object) -> float:
    if value is None:
        return 0
    try:
        if pd.isna(value):
            return 0
    except Exception:
        pass
    try:
        return float(str(value).strip())
    except Exception:
        return 0


def normalize_date(raw: object) -> str:
    if raw is None or raw == "" or (isinstance(raw, float) and math.isnan(raw)):
        return ""
    if hasattr(raw, "strftime") and not pd.isna(raw):
        return raw.strftime("%m/%d/%Y")
    s = str(raw).strip()
    parsed = pd.to_datetime(s, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%m/%d/%Y")
    if "-" in s and len(s.split("-")) >= 3:
        try:
            year, month, day = s.split(" ", 1)[0].split("-")[:3]
            return f"{month.zfill(2)}/{day.zfill(2)}/{year}"
        except Exception:
            return ""
    if "/" in s:
        return s
    return ""


def split_mmddyyyy(raw: str) -> tuple[str, str, str]:
    if not raw:
        return "", "", ""
    parts = str(raw).strip().split("/")
    if len(parts) != 3:
        return "", "", ""
    month, day, year = parts
    return month.zfill(2), day.zfill(2), year[-2:]


def format_phone(raw: object) -> str:
    if raw is None:
        return ""
    digits = re.sub(r"\D", "", str(raw))
    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    return ""


def format_year_two_digits(raw: object) -> str:
    if raw is None:
        return ""
    s = str(raw).strip()
    return s[-2:] if s[-2:].isdigit() else ""


def format_id(raw: object) -> str:
    if raw is None:
        return ""
    s = str(raw).strip()
    if s.lower() == "nan":
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if re.fullmatch(r"\d+\.\d+", s):
        try:
            return str(int(float(s)))
        except Exception:
            return s
    return s


def parse_account_number(raw: object) -> tuple[str, str, str, str, str]:
    if not isinstance(raw, str):
        return "", "", "", "", ""
    if ":" in raw:
        raw = raw.split(":", 1)[1].strip()
    parts = raw.split("-")
    parts += [""] * (5 - len(parts))
    return tuple(part.strip() for part in parts[:5])


def map_course_fields(course: dict[str, object]) -> dict[str, object]:
    dept = str(course.get("Dept", "") or "").strip()
    num = str(course.get("CourseNum", "") or "").strip()
    sec = str(course.get("CourseSection", "") or "").strip()
    course_code = f"{dept}{num}" if (dept or num) else str(course.get("CourseRaw", "") or "").strip()
    if course_code and sec:
        course_code = f"{course_code}-{sec}"
    return {
        "course1_code": course_code,
        "course1_title": course.get("CourseTitle", ""),
        "course1_credits": course.get("Credits"),
        "course1_enrollment": course.get("Enrollment"),
        "course2_code": "",
        "course2_title": "",
        "course2_credits": "",
        "course2_enrollment": "",
        "course3_code": "",
        "course3_title": "",
        "course3_credits": "",
        "course3_enrollment": "",
    }


def make_pdf_data_for_contract(instructor_name: str, semester: str, year: str, course_row: dict[str, object]) -> dict[str, object]:
    last, first = split_name(instructor_name)
    total_credits = safe_number(course_row.get("Credits"))
    weekly_hours = total_credits * 3
    raw_rate = course_row.get("Rate")
    rate_per_credit = DEFAULT_RATE_PER_CREDIT if not _has_value(raw_rate) else safe_number(raw_rate)
    amount_value = rate_per_credit * total_credits if rate_per_credit and total_credits else course_row.get("total contract amount", "")
    amount_str = str(int(amount_value)) if isinstance(amount_value, (int, float)) and amount_value else ""

    begin_m, begin_d, begin_y = split_mmddyyyy(normalize_date(course_row.get("BeginDate")))
    end_m, end_d, end_y = split_mmddyyyy(normalize_date(course_row.get("EndDate")))
    fund, orgn, acct_num, program, activity = parse_account_number(course_row.get("Account", ""))

    data: dict[str, object] = {
        "last_name": last,
        "first_name": first,
        "id": format_id(course_row.get("ID")),
        "email": course_row.get("Email", ""),
        "telephone": format_phone(course_row.get("Telephone")),
        "remote_yes": "/Yes" if str(course_row.get("Remote", "")).strip().lower() in ("yes", "y", "true", "1") else "/Off",
        "dept_contact_name": course_row.get("DeptContact", ""),
        "dept_contact_id": format_id(course_row.get("DeptContactID")),
        "reason": course_row.get("Reason", ""),
        "spring_yes": NameObject("/Yes") if str(semester).strip().lower() == "spring" else NameObject("/Off"),
        "summer_yes": NameObject("/Yes") if str(semester).strip().lower() == "summer" else NameObject("/Off"),
        "fall_yes": NameObject("/Yes") if str(semester).strip().lower() == "fall" else NameObject("/Off"),
        "year": format_year_two_digits(year),
        "begin_month": begin_m,
        "begin_day": begin_d,
        "begin_year": begin_y,
        "end_month": end_m,
        "end_day": end_d,
        "end_year": end_y,
        "Begin Date month": begin_m,
        "Begin Date date": begin_d,
        "Begin Date year": begin_y,
        "End Date month": end_m,
        "End Date date": end_d,
        "End Date year": end_y,
        "credits_total": safe_int_string(total_credits),
        "week_srv_hrs": safe_int_string(weekly_hours),
        "rate_per_credit": safe_int_string(rate_per_credit),
        "hourly_rate": "",
        "amount_total": amount_str,
        "fund": fund,
        "orgn": orgn,
        "acct": acct_num,
        "program": program,
        "activity": activity,
        "Signature2ID": format_id(course_row.get("DeptContactID")),
        "Signature3ID": format_id(course_row.get("DeanID")),
        "Signature4ID": format_id(course_row.get("VPFinanceID")),
        "international": "/Yes" if str(course_row.get("International", "")).strip().lower() in ("yes", "y", "true", "1") else "/Off",
        "dept_budget_yes": "/Yes" if str(course_row.get("DeptBudget", "")).strip().lower() in ("yes", "y", "true", "1") else "/Off",
        "dept_budget_no": "/Yes" if str(course_row.get("DeptBudget", "")).strip().lower() in ("no", "n", "false", "0") else "/Off",
        "prework_start": normalize_date(course_row.get("prework_start")),
        "prework_end": normalize_date(course_row.get("prework_end")),
        "prework_of_weeks": safe_int_string(course_row.get("prework_weeks")),
        "prework_hours_weeks": safe_int_string(course_row.get("prework_hours_week")),
        "prework_hours_period": safe_int_string(course_row.get("prework_hours_period")),
        "intensive_start": normalize_date(course_row.get("intensive_start")),
        "intensive_end": normalize_date(course_row.get("intensive_end")),
        "intensive_of_weeks": safe_int_string(course_row.get("intensive_weeks")),
        "intensive_hours_weeks": safe_int_string(course_row.get("intensive_hours_week")),
        "intensive_hours_period": safe_int_string(course_row.get("intensive_hours_period")),
        "postwork_start": normalize_date(course_row.get("postwork_start")),
        "postwork_end": normalize_date(course_row.get("postwork_end")),
        "postwork_of_weeks": safe_int_string(course_row.get("postwork_weeks")),
        "postwork_hours_week": safe_int_string(course_row.get("postwork_hours_week")),
        "postwork_hours_period": safe_int_string(course_row.get("postwork_hours_period")),
    }
    pre_h = safe_number(course_row.get("prework_hours_period"))
    int_h = safe_number(course_row.get("intensive_hours_period"))
    post_h = safe_number(course_row.get("postwork_hours_period"))
    explicit_total_hours = safe_number(course_row.get("TotalContractHours"))
    data["total_contract_hours"] = safe_int_string(explicit_total_hours or (pre_h + int_h + post_h))
    data.update(map_course_fields(course_row))
    return data


def sanitize_pdf_data(data: dict[str, object]) -> dict[str, object]:
    clean: dict[str, object] = {}
    for key, value in data.items():
        if isinstance(value, NameObject):
            clean[key] = value
        elif value is None:
            clean[key] = ""
        else:
            try:
                clean[key] = "" if isinstance(value, float) and math.isnan(value) else str(value)
            except Exception:
                clean[key] = ""
    return clean


def _fit_font_size_for_width(text: str, max_width: float, font_name: str = "Helvetica", start_size: float = 10, min_size: float = 5) -> float:
    from reportlab.pdfbase.pdfmetrics import stringWidth

    size = start_size
    while size > min_size and stringWidth(text, font_name, size) > max_width:
        size -= 0.5
    return max(size, min_size)


def fill_pdf_contract(template_path: Path, output_path: Path, data: dict[str, object]) -> None:
    clean_data = sanitize_pdf_data(data)
    reader = PdfReader(str(template_path), strict=False)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    page_field_map: dict[int, list[dict[str, object]]] = {}
    for page_num, page in enumerate(reader.pages):
        page_field_map[page_num] = []
        for annot_ref in page.get("/Annots", []):
            try:
                annot = annot_ref.get_object()
            except Exception:
                continue
            if annot.get("/Subtype") != "/Widget":
                continue
            field_name = annot.get("/T")
            parent = None
            if field_name is None and annot.get("/Parent"):
                try:
                    parent = annot["/Parent"].get_object()
                    field_name = parent.get("/T")
                except Exception:
                    parent = None
            if not field_name:
                continue
            rect = annot.get("/Rect")
            if not rect or len(rect) != 4:
                continue
            try:
                x1, y1, x2, y2 = [float(v) for v in rect]
            except Exception:
                continue
            field_type = annot.get("/FT")
            if field_type is None and parent is not None:
                field_type = parent.get("/FT")
            page_field_map[page_num].append(
                {
                    "name": str(field_name),
                    "rect": (x1, y1, x2, y2),
                    "ft": str(field_type) if field_type is not None else "",
                }
            )

    for page_num, page in enumerate(reader.pages):
        packet = BytesIO()
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        pdf_canvas = canvas.Canvas(packet, pagesize=(page_width, page_height))
        pdf_canvas.setFillColor(blue)
        for field in page_field_map.get(page_num, []):
            name = str(field["name"])
            if name not in clean_data:
                continue
            value = clean_data[name]
            if value in (None, "", "/Off"):
                continue
            x1, y1, x2, y2 = field["rect"]
            width = x2 - x1
            height = y2 - y1
            field_type = field["ft"]
            if field_type == "/Btn":
                if value in {"/Yes", "Yes", True, "True", "true", "1", 1}:
                    mark_size = max(8, min(height * 0.8, 12))
                    pdf_canvas.setFont("Helvetica", mark_size)
                    pdf_canvas.drawString(x1 + 1, y1 + max(1, (height - mark_size) / 2), "X")
                continue
            text = str(value)
            max_text_width = max(10, width - 4)
            font_size = _fit_font_size_for_width(
                text,
                max_width=max_text_width,
                font_name="Helvetica",
                start_size=max(6, min(height * 0.7, 10)),
                min_size=5,
            )
            pdf_canvas.setFont("Helvetica", font_size)
            text_y = y1 + max(1, (height - font_size) / 2)
            pdf_canvas.drawString(x1 + 2, text_y, text[:500])
        pdf_canvas.save()
        packet.seek(0)
        overlay_reader = PdfReader(packet)
        writer.pages[page_num].merge_page(overlay_reader.pages[0])

    for page in writer.pages:
        annots = page.get("/Annots")
        if not annots:
            continue
        kept = []
        for annot_ref in annots:
            try:
                annot = annot_ref.get_object()
                if annot.get("/Subtype") != "/Widget":
                    kept.append(annot_ref)
            except Exception:
                kept.append(annot_ref)
        if kept:
            page[NameObject("/Annots")] = kept
        elif "/Annots" in page:
            del page["/Annots"]

    if "/AcroForm" in writer._root_object:
        del writer._root_object["/AcroForm"]

    with output_path.open("wb") as file_obj:
        writer.write(file_obj)


def generate_contract_pdfs(
    df_course: pd.DataFrame,
    paths: ContractAppPaths,
    output_dir: Path,
    filename_prefix: str = "",
    source_file: str = "",
) -> list[GeneratedContract]:
    contracts: list[GeneratedContract] = []
    if df_course.empty:
        return contracts

    for row_i, row in df_course.reset_index(drop=False).iterrows():
        course = row.to_dict()
        instructor = str(course.get("Instructor", "") or "")
        semester = str(course.get("Semester", "") or "")
        year = str(course.get("Year", "") or "")
        contract_type = "intensive" if has_intensive_period(course) else "standard"
        template_path = paths.intensive_template if contract_type == "intensive" else paths.standard_template

        data = make_pdf_data_for_contract(instructor, semester, year, course)
        last, first = split_name(instructor)
        token_str = course_token(course)
        safe_name = f"{first}_{last}".replace(" ", "_")
        base = f"{year}_{semester}_{contract_type}_{safe_name}_{token_str}"
        unique_suffix = course.get("Source Row", row_i)
        filename = f"{filename_prefix}{base}_row{unique_suffix}.pdf"
        output_path = output_dir / sanitize_filename(filename)
        fill_pdf_contract(template_path, output_path, data)
        contracts.append(
            GeneratedContract(
                filename=output_path.name,
                path=output_path,
                semester=semester,
                contract_type=contract_type,
                instructor=instructor,
                course_code=token_str,
                source_file=source_file,
            )
        )
    return contracts
