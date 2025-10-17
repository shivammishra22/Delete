# pip install python-docx pandas
import re
from pathlib import Path
from typing import Any, List
from dataclasses import dataclass

import pandas as pd
from docx import Document


EXPECTED_COLUMNS = [
    "Signal term",
    "Date detected (month/ year)",
    "Status (ongoing or closed)",
    "Date closed (for closed signals) (month/ year)",
    "Source of signal",
    "Reason for evaluation & summary of key data",
    "Method of signal evaluation",
    "Action(s) taken or planned",
]


def _normalize(text: str) -> str:
    if text is None:
        return ""
    s = text.replace("\n", " ").strip().lower()
    s = s.replace("&", "and")
    s = re.sub(r"\s+", " ", s)
    return s


_NORM_EXPECTED = [_normalize(c) for c in EXPECTED_COLUMNS]


def _coerce_path(docx_path: Any) -> str:
    """Accept str | Path | sequence[str] and return a usable path string.

    Guards against accidental tuple like (path,) seen in some configurations.
    """
    if isinstance(docx_path, (list, tuple)):
        if not docx_path:
            return ""
        docx_path = docx_path[0]
    return str(docx_path) if docx_path is not None else ""


def extract_signal_table(docx_path: str | Path | Any) -> tuple[pd.DataFrame, list[str]]:
    """Open a .docx, find the signal table, and return (DataFrame, closed_signal_terms)."""
    path_str = _coerce_path(docx_path)
    if not path_str:
        return pd.DataFrame(columns=["Signal term", "Date detected (month/ year)", "Status (ongoing or closed)"]), []
    try:
        doc = Document(path_str)
    except Exception:
        return pd.DataFrame(columns=["Signal term", "Date detected (month/ year)", "Status (ongoing or closed)"]), []
    for tbl in doc.tables:
        header_cells = [cell.text for cell in tbl.rows[0].cells]
        header_norm = [_normalize(t) for t in header_cells]
        if set(_NORM_EXPECTED).issubset(set(header_norm)):
            idx_map = {h: i for i, h in enumerate(header_norm)}
            records = []
            for row in tbl.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                record = []
                for norm_col in _NORM_EXPECTED:
                    i = idx_map[norm_col]
                    record.append(cells[i] if i < len(cells) else "")
                records.append(record)
            df = pd.DataFrame(records, columns=EXPECTED_COLUMNS)
            df = (
                df.replace(r"\s+", " ", regex=True)
                  .replace(r"^\s*$", pd.NA, regex=True)
                  .dropna(how="all")
            )
            closed_signals = df[df["Status (ongoing or closed)"].str.strip().str.lower() == "closed"]["Signal term"].tolist()
            return df[["Signal term", "Date detected (month/ year)", "Status (ongoing or closed)"]], closed_signals
    return pd.DataFrame(columns=["Signal term", "Date detected (month/ year)", "Status (ongoing or closed)"]), []


# -----------------------
# Accumulation utilities
# -----------------------

@dataclass
class Section15Data:
    table: pd.DataFrame
    closed_signals: List[str]


def accumulate_section15(docx_path: str | Path) -> Section15Data:
    table, closed = extract_signal_table(docx_path)
    return Section15Data(table=table, closed_signals=closed)
