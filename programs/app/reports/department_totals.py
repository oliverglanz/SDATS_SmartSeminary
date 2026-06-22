from __future__ import annotations

import pandas as pd
import numpy as np

from ..excel_utils import dataframe_to_excel_bytes
from ..models import ReportResult


def generate_department_totals_report(df: pd.DataFrame) -> ReportResult:
    working = df.copy()
    working["load/contract"] = working["load/contract"].astype("string").str.lower().str.strip()
    working["Cr."] = pd.to_numeric(working["Cr."], errors="coerce").fillna(0)
    working["total contract amount"] = pd.to_numeric(
        working["total contract amount"],
        errors="coerce",
    ).fillna(0)

    grouped = (
        working.groupby("Department", dropna=False)
        .apply(
            lambda group: pd.Series(
                {
                    "Contract Cost": float(group.loc[group["load/contract"].eq("contract"), "total contract amount"].sum()),
                    "Load_credits": float(group.loc[group["load/contract"].eq("load"), "Cr."].sum()),
                    "Contract_credits": float(group.loc[group["load/contract"].eq("contract"), "Cr."].sum()),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )

    grouped["Load/Contract Ratio"] = np.where(
        grouped["Load_credits"].eq(0),
        np.nan,
        grouped["Contract_credits"] / grouped["Load_credits"],
    )
    grouped["Sum of all Credits"] = grouped["Load_credits"] + grouped["Contract_credits"]
    grouped = grouped.sort_values("Department").reset_index(drop=True)

    excel_bytes = dataframe_to_excel_bytes(grouped, "Department Totals")
    return ReportResult(
        report_name="Department Totals",
        preview_df=grouped,
        excel_bytes=excel_bytes,
        output_filename="department_totals_from_donsheet.xlsx",
        worksheet_name="Department Totals",
        metadata={"rows": str(len(grouped))},
    )
