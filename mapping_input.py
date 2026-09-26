from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd


SUPPORTED_MAPPING_EXTENSIONS = {".csv", ".xlsx"}
CANONICAL_MAPPING_COLUMNS = ["Question", "Snippet ID", "Mapping Type", "Path"]


def _header_key(value: object) -> str:
    return "".join(character for character in str(value).casefold() if character.isalnum())


def _find_column(frame: pd.DataFrame, *aliases: str) -> str | None:
    by_key = {_header_key(column): str(column) for column in frame.columns}
    for alias in aliases:
        match = by_key.get(_header_key(alias))
        if match is not None:
            return match
    return None


def _paths(value: object) -> list[str]:
    return [line.strip() for line in str(value or "").replace("\r\n", "\n").replace("\r", "\n").split("\n") if line.strip()]


def normalize_mapping_table(frame: pd.DataFrame) -> pd.DataFrame:
    """Convert supported long or wide team mapping sheets to one path per row."""
    frame = frame.dropna(how="all").fillna("").copy()
    frame.columns = [str(column).strip() for column in frame.columns]

    question_column = _find_column(frame, "Question", "Project", "Project with qno", "Project with question number")
    snippet_column = _find_column(frame, "Snippet ID", "Snippet id", "Snippet IDs", "Snippet ids")
    type_column = _find_column(frame, "Mapping Type", "Section")
    path_column = _find_column(frame, "Path", "Mapping Path", "Additional mapping path")
    curriculum_column = _find_column(frame, "Curriculum", "Curriculum mapping", "Curriculum mappings")
    taxonomy_column = _find_column(frame, "Taxonomy", "Taxonomy mapping", "Taxonomy mappings")

    if question_column is None and snippet_column is None:
        raise ValueError("Include either a Snippet ID column or a Project with qno/Project column.")

    records: list[dict[str, str]] = []
    if type_column is not None and path_column is not None:
        for _, row in frame.iterrows():
            question = str(row[question_column]).strip() if question_column else ""
            snippet = str(row[snippet_column]).strip() if snippet_column else ""
            mapping_type = str(row[type_column]).strip().title()
            for path in _paths(row[path_column]):
                records.append(
                    {"Question": question, "Snippet ID": snippet, "Mapping Type": mapping_type, "Path": path}
                )
    elif curriculum_column is not None or taxonomy_column is not None:
        for _, row in frame.iterrows():
            question = str(row[question_column]).strip() if question_column else ""
            snippet = str(row[snippet_column]).strip() if snippet_column else ""
            for mapping_type, column in (("Curriculum", curriculum_column), ("Taxonomy", taxonomy_column)):
                if column is None:
                    continue
                for path in _paths(row[column]):
                    records.append(
                        {"Question": question, "Snippet ID": snippet, "Mapping Type": mapping_type, "Path": path}
                    )
    else:
        raise ValueError(
            "Use either Mapping Type + Path columns, or separate Curriculum and Taxonomy columns."
        )

    normalized = pd.DataFrame(records, columns=CANONICAL_MAPPING_COLUMNS).fillna("")
    if normalized.empty:
        raise ValueError("The selected worksheet contains no Curriculum or Taxonomy paths.")
    normalized["Question"] = normalized["Question"].astype(str).str.strip()
    normalized["Snippet ID"] = normalized["Snippet ID"].astype(str).str.strip()
    normalized["Mapping Type"] = normalized["Mapping Type"].astype(str).str.strip().str.title()
    normalized["Path"] = normalized["Path"].astype(str).str.strip()

    invalid_types = normalized[~normalized["Mapping Type"].isin(["Curriculum", "Taxonomy"])]
    invalid_snippets = normalized[
        (normalized["Snippet ID"] != "") & ~normalized["Snippet ID"].str.fullmatch(r"[1-9]\d*")
    ]
    missing_identifiers = normalized[(normalized["Snippet ID"] == "") & (normalized["Question"] == "")]
    if not invalid_types.empty:
        raise ValueError("Every Mapping Type must be Curriculum or Taxonomy.")
    if not invalid_snippets.empty:
        raise ValueError("Every supplied Snippet ID must be a positive whole number.")
    if not missing_identifiers.empty:
        raise ValueError("Every mapping row needs either a Snippet ID or a Project with qno value.")
    return normalized.drop_duplicates().reset_index(drop=True)


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
