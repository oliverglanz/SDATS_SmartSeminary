from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import BinaryIO

import pandas as pd
from openpyxl import load_workbook as load_openpyxl_workbook
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.worksheet.worksheet import Worksheet

from .config import HELPER_SHEETS
from .models import WorkbookSummary


def _read_source_bytes(source: str | Path | bytes | BinaryIO) -> bytes:
    if isinstance(source, bytes):
        return source
    if isinstance(source, Path):
        return source.read_bytes()
    if isinstance(source, str):
        return Path(source).read_bytes()
    if hasattr(source, "read"):
        content = source.read()
        if hasattr(source, "seek"):
            source.seek(0)
        return content
    return Path(source).read_bytes()


def _load_excel_file(source: str | Path | bytes | BinaryIO) -> pd.ExcelFile:
    return pd.ExcelFile(BytesIO(_read_source_bytes(source)))


def _load_openpyxl_workbook(source: str | Path | bytes | BinaryIO):
    return load_openpyxl_workbook(BytesIO(_read_source_bytes(source)), data_only=True)


def _color_to_rgb(color) -> str | None:
    if color is None or color.type is None:
        return None
    if color.type == "rgb" and color.rgb:
        rgb = str(color.rgb)
        return rgb[-6:].upper()
    if color.type == "indexed" and color.indexed is not None:
        # Excel may store automatic/custom palette references outside
        # openpyxl's built-in COLOR_INDEX tuple (for example index 64).
        # Those colors cannot be resolved reliably, but they must not abort
        # DonSheet validation.
        try:
            indexed = COLOR_INDEX[color.indexed]
        except (IndexError, TypeError):
            return None
        return str(indexed)[-6:].upper()
    return None


def _is_red_rgb(rgb: str | None) -> bool:
    if not rgb or len(rgb) != 6:
        return False
    try:
        red = int(rgb[0:2], 16)
        green = int(rgb[2:4], 16)
        blue = int(rgb[4:6], 16)
    except ValueError:
        return False
    # DonSheets use both strong red (for example FF0000) and Excel's common
    # pale-red fill (FFC7CE).  Detect a red-dominant color instead of only a
    # dark red, so visually red rows are consistently excluded.
    return red >= 160 and red - green >= 30 and red - blue >= 20


def _cell_has_red_fill(cell) -> bool:
    """Return whether a cell has a direct red fill marking an excluded row."""
    fill = cell.fill
    if fill is not None and fill.patternType is not None:
        if any(
            _is_red_rgb(_color_to_rgb(color))
            for color in (fill.fgColor, fill.start_color, fill.bgColor, fill.end_color)
        ):
            return True
    return False


def _red_highlighted_row_indices(sheet: Worksheet, column_count: int) -> set[int]:
    red_rows: set[int] = set()
    for row_number in range(2, sheet.max_row + 1):
        row_cells = sheet[row_number][: max(column_count, 1)]
        if any(_cell_has_red_fill(cell) for cell in row_cells):
            red_rows.add(row_number - 2)
    return red_rows


def _source_name(source: str | Path | bytes | BinaryIO, excel_file: pd.ExcelFile) -> str:
    explicit_name = getattr(source, "name", None)
    if explicit_name:
        return str(explicit_name)
    if isinstance(source, Path):
        return source.name
    if isinstance(source, str):
        return Path(source).name
    excel_io = getattr(excel_file, "io", None)
    if isinstance(excel_io, (str, Path)):
        return Path(excel_io).name
    return "uploaded workbook"


def load_workbook(source: str | Path | bytes | BinaryIO) -> tuple[pd.ExcelFile, WorkbookSummary]:
    excel_file = _load_excel_file(source)
    sheet_names = list(excel_file.sheet_names)
    data_sheet_names = [name for name in sheet_names if name not in HELPER_SHEETS]
    helper_sheet_names = [name for name in sheet_names if name in HELPER_SHEETS]

    row_counts: dict[str, int] = {}
    warnings: list[str] = []

    for sheet_name in data_sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        row_counts[sheet_name] = len(df)
        if df.empty:
            warnings.append(f"Sheet '{sheet_name}' is empty.")

    summary = WorkbookSummary(
        source_name=_source_name(source, excel_file),
        sheet_names=sheet_names,
        data_sheet_names=data_sheet_names,
        helper_sheet_names=helper_sheet_names,
        row_counts=row_counts,
        warnings=warnings,
    )
    return excel_file, summary


def read_department_sheets(
    excel_file: pd.ExcelFile,
    sheet_names: list[str],
    source: str | Path | bytes | BinaryIO | None = None,
) -> dict[str, pd.DataFrame]:
    department_frames: dict[str, pd.DataFrame] = {}
    styled_workbook = _load_openpyxl_workbook(source) if source is not None else None
    for sheet_name in sheet_names:
        frame = pd.read_excel(excel_file, sheet_name=sheet_name)
        if styled_workbook is not None and sheet_name in styled_workbook.sheetnames and not frame.empty:
            red_rows = _red_highlighted_row_indices(styled_workbook[sheet_name], len(frame.columns))
            if red_rows:
                frame = frame.drop(index=[row for row in red_rows if row in frame.index]).reset_index(drop=True)
        department_frames[sheet_name] = frame
    return department_frames


def read_helper_sheet(excel_file: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    if sheet_name not in excel_file.sheet_names:
        return pd.DataFrame()
    return pd.read_excel(excel_file, sheet_name=sheet_name)
