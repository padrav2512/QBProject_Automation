from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd


SUPPORTED_MAPPING_EXTENSIONS = {".csv", ".xlsx"}


def xlsx_sheet_names(data: bytes) -> list[str]:
    """Return the visible worksheet names in an XLSX mapping workbook."""
    with pd.ExcelFile(BytesIO(data), engine="openpyxl") as workbook:
        return list(workbook.sheet_names)


def read_mapping_table(data: bytes, filename: str, sheet_name: str | None = None) -> pd.DataFrame:
    """Read a CSV or one XLSX worksheet without coercing identifiers to numbers."""
    extension = Path(filename).suffix.lower()
    if extension not in SUPPORTED_MAPPING_EXTENSIONS:
        raise ValueError("Upload a CSV or XLSX mapping file.")

    source = BytesIO(data)
    if extension == ".csv":
        frame = pd.read_csv(source, dtype=str)
    else:
        sheets = xlsx_sheet_names(data)
        if not sheets:
            raise ValueError("The XLSX workbook has no readable worksheets.")
        selected_sheet = sheet_name or sheets[0]
        if selected_sheet not in sheets:
            raise ValueError(f"Worksheet '{selected_sheet}' was not found in the XLSX workbook.")
        frame = pd.read_excel(source, sheet_name=selected_sheet, dtype=str, engine="openpyxl")

    frame = frame.dropna(how="all").fillna("")
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame
