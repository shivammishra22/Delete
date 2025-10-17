  import os
import re
from dataclasses import dataclass
from typing import Any, Dict

import pandas as pd
from docx import Document

# HTML conversion imports
import mammoth
from bs4 import BeautifulSoup


# Optimized column width mapping kept for reference if needed by writers
COLUMN_WIDTH_MAP: Dict[str, float] = {
    'Molecule/Product': 1.0,
    'Study Number': 0.8,
    'Study Title': 1.5,
    'Test product name': 1.2,
    'Active comparator name': 1.2,
    'Test Product': 0.6,
    'Active Comparator': 0.6,
    'Placebo': 0.5,
    'Total': 0.5,
    'Male': 0.4,
    'Female': 0.4,
    '<18 years': 0.5,
    '18-65 years': 0.5,
    '>65 years': 0.5,
    'Asian': 0.4,
    'Black': 0.4,
    'Caucasian': 0.5,
    'Other': 0.4,
    'Unknown': 0.5,
}


def extract_specific_table(docx_path: str) -> list[list[str]]:
    """Find and extract the clinical trial demographics table from a DOCX.

    Returns a list of rows (each row is list of cell texts). Empty list if not found.
    """
    if not os.path.exists(docx_path):
        return []

    doc = Document(docx_path)
    keywords = [
        "Molecular Product", "Study Number", "Test Product Name",
        "Active comparator name", "TestProduct", "Active Comparator",
        "Placebo", "Total", "Gender", "Age", "Racial",
    ]
    keyword_patterns = [re.compile(re.escape(kw), re.IGNORECASE) for kw in keywords]

    for table in doc.tables:
        for row in table.rows[:3]:  # Only check first 3 rows for speed
            row_text = ' '.join(cell.text.strip() for cell in row.cells)
            if any(p.search(row_text) for p in keyword_patterns):
                extracted = [[cell.text.strip() for cell in r.cells] for r in table.rows]
                if len(extracted) > 1 and extracted[0] == extracted[1]:
                    extracted.pop(1)
                return extracted
    return []


def process_table_structure(table_data: list[list[str]]) -> pd.DataFrame | None:
    """Convert extracted table rows to a typed DataFrame ready for analysis."""
    if not table_data or len(table_data) < 3:
        return None

    expected_columns = [
        'Molecule/Product', 'Study Number', 'Study Title', 'Test product name',
        'Active comparator name', 'Test Product', 'Active Comparator', 'Placebo', 'Total',
        'Male', 'Female', '<18 years', '18-65 years', '>65 years',
        'Asian', 'Black', 'Caucasian', 'Other', 'Unknown',
    ]

    num_cols = len(table_data[0])
    columns = expected_columns[:num_cols]
    df = pd.DataFrame(table_data[2:], columns=columns)

    numeric_columns = [col for col in columns if col in [
        'Test Product', 'Active Comparator', 'Placebo', 'Total',
        'Male', 'Female', '<18 years', '18-65 years', '>65 years',
        'Asian', 'Black', 'Caucasian', 'Other', 'Unknown',
    ]]

    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    return df


def parse_html_table_structure(html_table: Any) -> dict[str, Any] | None:
    """Parse an HTML table into a simple structure for Word reconstruction."""
    table_structure: dict[str, Any] = {'rows': [], 'max_cols': 0}
    rows = html_table.find_all('tr')
    try:
        for row_idx, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            row_data: list[dict[str, Any]] = []
            col_idx = 0
            for cell in cells:
                colspan = int(cell.get('colspan', 1))
                rowspan = int(cell.get('rowspan', 1))
                row_data.append({
                    'text': cell.get_text(strip=True),
                    'colspan': colspan,
                    'rowspan': rowspan,
                    'row': row_idx,
                    'col': col_idx,
                })
                col_idx += colspan
            table_structure['rows'].append(row_data)
            table_structure['max_cols'] = max(table_structure['max_cols'], col_idx)
        return table_structure
    except Exception:
        return None


def copy_table_via_html_conversion(source_docx_path: str) -> dict[str, Any] | None:
    """Convert DOCX->HTML and extract the target table structure by keywords."""
    table_keywords = [
        "Molecular Product", "Study Number", "Test Product Name",
        "Active comparator name", "TestProduct", "Active Comparator",
        "Placebo", "Total", "Gender", "Age", "Racial",
    ]
    keyword_patterns = [re.compile(kw, re.IGNORECASE) for kw in table_keywords]

    try:
        with open(source_docx_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
        soup = BeautifulSoup(result.value, 'html.parser')
        tables = soup.find_all('table')
        for table in tables:
            table_text = table.get_text()
            if any(pattern.search(table_text) for pattern in keyword_patterns):
                return parse_html_table_structure(table)
    except Exception:
        return None
    return None


# -----------------------
# Accumulation utilities
# -----------------------

@dataclass
class Section5_2Data:
    nstudies: int
    medname: str
    reporting_period: str
    total_subjects: int
    gender_text: str
    age_text: str
    race_text: str
    table_structure: dict | None


def _generate_demographic_summary(df, medname: str, nstudies: int, reporting_period: str):
    total_subjects = int(df['Total'].sum()) if 'Total' in df.columns else 0

    gender_data = {}
    for gender in ['Male', 'Female']:
        if gender in df.columns:
            gender_data[gender] = int(df[gender].sum())

    age_columns = ['<18 years', '18-65 years', '>65 years']
    age_data = {col: int(df[col].sum()) for col in age_columns if col in df.columns}

    race_columns = ['Asian', 'Black', 'Caucasian', 'Other', 'Unknown']
    race_data = {col: int(df[col].sum()) for col in race_columns if col in df.columns}

    gender_text = f"{gender_data.get('Male', 0)} were male, {gender_data.get('Female', 0)} were female" if gender_data else "gender distribution unknown"
    age_parts = [f"{count} in {age}" for age, count in age_data.items() if count > 0]
    age_text = f"age distribution: {', '.join(age_parts)}" if age_parts else "age distribution unknown"
    race_parts = [f"{count} {race}" for race, count in race_data.items() if count > 0]
    race_text = f"racial distribution: {', '.join(race_parts)}" if race_parts else "racial distribution unknown"

    return nstudies, medname, reporting_period, total_subjects, gender_text, age_text, race_text


def accumulate_section5_2(section5_docx_path: str, reporting_period: str, medname: str = "Olanzapine") -> Section5_2Data:
    tbl = extract_specific_table(section5_docx_path)
    df = process_table_structure(tbl)
    nstudies = len(df) if df is not None else 0

    if df is None:
        # Minimal defaults
        return Section5_2Data(
            nstudies=0,
            medname=medname,
            reporting_period=reporting_period,
            total_subjects=0,
            gender_text="gender distribution unknown",
            age_text="age distribution unknown",
            race_text="racial distribution unknown",
            table_structure=None,
        )

    nstudies, medname, reporting_period, total_subjects, gender_text, age_text, race_text = _generate_demographic_summary(
        df, medname, nstudies, reporting_period
    )
    table_structure = copy_table_via_html_conversion(section5_docx_path)

    return Section5_2Data(
        nstudies=nstudies,
        medname=medname,
        reporting_period=reporting_period,
        total_subjects=total_subjects,
        gender_text=gender_text,
        age_text=age_text,
        race_text=race_text,
        table_structure=table_structure,
    )


