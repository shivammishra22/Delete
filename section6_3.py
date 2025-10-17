from __future__ import annotations

import os
import re
from typing import Tuple, Optional
from dataclasses import dataclass

import pandas as pd
from docx import Document
from striprtf.striprtf import rtf_to_text


def read_rtf_file_safe(rtf_path: str) -> str:
    """Safely read RTF content using common encodings."""
    for enc in ("utf-8", "cp1252", "latin1"):
        try:
            with open(rtf_path, 'r', encoding=enc) as fh:
                return fh.read()
        except UnicodeDecodeError:
            continue
    # Fallback
    with open(rtf_path, 'r', encoding='latin1') as fh:
        return fh.read()


def extract_text_from_document(file_path: str) -> str:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file was not found: {file_path}")
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.rtf':
        return rtf_to_text(read_rtf_file_safe(file_path))
    if ext == '.docx':
        doc = Document(file_path)
        return "\n".join(p.text for p in doc.paragraphs)
    raise ValueError(f"Unsupported file type: {file_path}. Provide .rtf or .docx.")


def extract_table_with_multilevel_header(document_path: str, start_marker: str, end_marker: str) -> Tuple[pd.DataFrame, int, int]:
    """Extract a multi-level table from plain text and compute labeled/unlabeled counts."""
    plain_text = extract_text_from_document(document_path)
    lines = plain_text.splitlines()

    last_line = next((line for line in reversed(lines) if line.strip()), '')
    parts = [p.strip() for p in last_line.split('|') if p.strip().isdigit()]
    labeled = int(parts[0]) if len(parts) > 0 else 0
    unlabeled = int(parts[1]) if len(parts) > 1 else 0

    start_idx = end_idx = None
    for i, line in enumerate(lines):
        if start_marker.lower() in line.lower():
            start_idx = i
        elif end_marker.lower() in line.lower() and start_idx is not None:
            end_idx = i
            break

    if start_idx is None or end_idx is None:
        return pd.DataFrame(), labeled, unlabeled

    table_lines = [line.strip() for line in lines[start_idx:end_idx] if line.strip()]

    header_row = subheader_row = None
    data_start = None
    for i, line in enumerate(table_lines):
        if "System Organ Class" in line and "Preferred Term" in line:
            header_row = re.split(r'\s*\|\s*', line.strip())
        elif "No" in line and "Yes" in line and "Total" in line:
            subheader_row = re.split(r'\s*\|\s*', line.strip())
            data_start = i + 1
            break

    if not header_row or not subheader_row or data_start is None:
        return pd.DataFrame(), labeled, unlabeled

    columns = [
        (header_row[0].strip(), ''),
        (header_row[1].strip(), ''),
        (header_row[2].strip(), subheader_row[1].strip()),
        (header_row[2].strip(), subheader_row[2].strip()),
        (header_row[2].strip(), subheader_row[3].strip()),
    ]
    multi_columns = pd.MultiIndex.from_tuples(columns)

    data_rows = []
    current_soc = ''
    for line in table_lines[data_start:]:
        parts = [p.strip() for p in re.split(r'\s*\|\s*', line) if p.strip()]
        if len(parts) == 5:
            current_soc = parts[0]
            data_rows.append(parts)
        elif len(parts) == 4:
            data_rows.append([current_soc] + parts)

    if not data_rows:
        return pd.DataFrame(columns=multi_columns), labeled, unlabeled

    df = pd.DataFrame(data_rows, columns=multi_columns)
    return df, labeled, unlabeled


# -----------------------
# Accumulation utilities
# -----------------------

@dataclass
class Section6_3Data:
    cumulative_text: str
    interval_text: Optional[str]


def _excel_summary(file_path: str) -> str:
    """Simple summary from Excel/CSV: count rows as proxy for ICSR count."""
    if not file_path:
        return ""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(file_path)
    else:
        df = pd.read_csv(file_path)
    total = len(df.index)
    return f"A total of {total} ICSRs were identified in the dataset."


def _rtf_soc_summary(rtf_path: str) -> str:
    df2, label, unlabel = extract_table_with_multilevel_header(
        rtf_path,
        start_marker="For the cases in this report",
        end_marker="Main Summary Tabulation: by System Organ Class (SOC)",
    )
    if df2.empty:
        return (
            f"A total of {label + unlabel} ADRs were found, with {label} labeled and {unlabel} unlabeled. "
            f"No detailed event data could be parsed."
        )

    df2 = df2.iloc[1:].reset_index(drop=True)
    df2.columns = ['Blanket diseases', 'Sub-diseases', 'No', 'Yes', 'Total']
    df2['Total'] = pd.to_numeric(df2['Total'], errors='coerce').fillna(0).astype(int)
    last = df2.iloc[-1]
    non_serious_ADR = int(last.get("No", 0))
    serious_ADR = int(last.get("Yes", 0))
    total_ADR = int(last.get("Total", 0))
    if total_ADR == 0:
        return (
            f"{serious_ADR} serious ADRs and {non_serious_ADR} non-serious ADRs were reported. "
            f"Out of {label + unlabel} ADRs, {label} were labeled and {unlabel} unlabeled as per current RSI."
        )

    subtotals_df = df2[df2['Sub-diseases'] == 'SubTotal'].copy()
    subtotals_df['Total'] = pd.to_numeric(subtotals_df['Total'], errors='coerce').fillna(0)
    top3 = subtotals_df.sort_values(by='Total', ascending=False).head(3)[['Blanket diseases', 'Total']]
    blanket_totals = list(zip(top3['Blanket diseases'].tolist(), top3['Total'].astype(int).tolist()))
    if len(blanket_totals) < 3:
        return (
            f"{serious_ADR} serious ADRs and {non_serious_ADR} non-serious ADRs were reported. "
            f"Insufficient data to generate detailed SOC summary."
        )

    sub_totals: dict[str, list[tuple[str, int]]] = {}
    for blanket, _ in blanket_totals:
        choices = df2[(df2['Blanket diseases'] == blanket) & (df2['Sub-diseases'] != 'SubTotal')].copy()
        choices['Total'] = pd.to_numeric(choices['Total'], errors='coerce').fillna(0)
        top4 = choices.sort_values(by='Total', ascending=False).head(4)[['Sub-diseases', 'Total']]
        sub_totals[blanket] = list(zip(top4['Sub-diseases'].tolist(), top4['Total'].astype(int).tolist()))

    para2 = (
        f" {serious_ADR} serious ADRs and {non_serious_ADR} non-serious ADRs. Out of {label + unlabel} ADRs, "
        f"{label} were labeled and {unlabel} were unlabeled as per current RSI."
    )
    b0, b1, b2 = blanket_totals[0], blanket_totals[1], blanket_totals[2]

    def pct(n):
        return round((n / max(total_ADR, 1)) * 100)

    st0 = sub_totals[b0[0]]
    st1 = sub_totals[b1[0]]
    st2 = sub_totals[b2[0]]
    para3 = (
        f"Of these {total_ADR} ADRs, most fell under SOC {b0[0]} (n={b0[1]} i.e. {pct(b0[1])}% of total) "
        f"including {st0[0][0]} (n={st0[0][1]}), {st0[1][0]} (n={st0[1][1]}), {st0[2][0]} (n={st0[2][1]}), "
        f"{st0[3][0]} (n={st0[3][1]}) etc."
    )
    para4 = (
        f"Second most common SOC was {b1[0]} (n={b1[1]}, i.e. {pct(b1[1])}% of total) including {st1[0][0]} "
        f"(n={st1[0][1]}), {st1[1][0]} (n={st1[1][1]}), {st1[2][0]} (n={st1[2][1]}), and {st1[3][0]} "
        f"(n={st1[3][1]}) etc."
    )
    para5 = (
        f"Third most common SOC was {b2[0]} (n={b2[1]}, i.e. {pct(b2[1])}% of total) including {st2[0][0]} "
        f"(n={st2[0][1]}), {st2[1][0]} (n={st2[1][1]}), {st2[2][0]} (n={st2[2][1]}), and {st2[3][0]} (n={st2[3][1]}) etc."
    )
    return f"{para2}\n\n{para3}\n\n{para4}\n\n{para5}"


def accumulate_section6_3(*, cumulative_excel: str, cumulative_rtf: str,
                          interval_excel: Optional[str] = None, interval_rtf: Optional[str] = None) -> Section6_3Data:
    cum1 = _excel_summary(cumulative_excel)
    cum2 = _rtf_soc_summary(cumulative_rtf)
    cumulative_text = f"{cum1} {cum2}".strip()

    interval_text: Optional[str] = None
    if interval_excel and interval_rtf:
        int1 = _excel_summary(interval_excel)
        int2 = _rtf_soc_summary(interval_rtf)
        interval_text = f"{int1} {int2}".strip()

    return Section6_3Data(cumulative_text=cumulative_text, interval_text=interval_text)

