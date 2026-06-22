from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

import pandas as pd


@dataclass
class WorkbookSummary:
    source_name: str
    sheet_names: List[str]
    data_sheet_names: List[str]
    helper_sheet_names: List[str]
    row_counts: Dict[str, int]
    warnings: List[str] = field(default_factory=list)


@dataclass
class NormalizationResult:
    canonical_df: pd.DataFrame
    summary: WorkbookSummary
    warnings: List[str] = field(default_factory=list)


@dataclass
class ReportResult:
    report_name: str
    preview_df: pd.DataFrame
    excel_bytes: bytes
    output_filename: str
    worksheet_name: str
    metadata: Optional[Dict[str, str]] = None


@dataclass
class ReportBundle:
    bundle_name: str
    reports: Dict[str, ReportResult]
    combined_workbook_bytes: bytes
    combined_output_filename: str
    preview_report_key: str
    metadata: Optional[Dict[str, str]] = None
