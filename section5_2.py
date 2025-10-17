from __future__ import annotations
from os.path import abspath, dirname, join
import sys
sys.path.insert(0, abspath(join(dirname(__file__), '..')))
from docx.shared import Inches, Pt
from docx import Document

from config.styling_config import DocumentStyling


def create_word_table_from_html_structure(output_doc: Document, table_structure: dict | None, title: str) -> None:
    """Render an HTML-like table structure into the Word document."""
    title_para = output_doc.add_paragraph()
    DocumentStyling.create_split_subheading(title_para, title)

    if not table_structure or not table_structure.get('rows'):
        p = output_doc.add_paragraph("Table structure not available")
        for run in p.runs:
            DocumentStyling.apply_content_style(run)
        return

    max_rows = len(table_structure['rows'])
    max_cols = table_structure['max_cols']

    word_table = output_doc.add_table(rows=max_rows, cols=max_cols)
    word_table.style = 'Table Grid'
    word_table.autofit = False

    # Column widths: simple uniform distribution across page width
    available_width_inches = 10.0
    col_width_inches = available_width_inches / max_cols if max_cols else available_width_inches
    try:
        col_width = Inches(col_width_inches)
        for col in word_table.columns:
            col.width = col_width
    except Exception:
        pass

    merged_cells: set[tuple[int, int]] = set()
    for row_idx, row_data in enumerate(table_structure['rows']):
        word_row = word_table.rows[row_idx]
        current_col = 0
        for cell_info in row_data:
            while (row_idx, current_col) in merged_cells:
                current_col += 1
            if current_col >= max_cols:
                break
            target_cell = word_row.cells[current_col]
            target_cell.text = cell_info['text']
            colspan = int(cell_info.get('colspan', 1))
            rowspan = int(cell_info.get('rowspan', 1))

            if colspan > 1 or rowspan > 1:
                for r in range(row_idx, min(row_idx + rowspan, max_rows)):
                    for c in range(current_col, min(current_col + colspan, max_cols)):
                        if r != row_idx or c != current_col:
                            merged_cells.add((r, c))
                try:
                    end_row = min(row_idx + rowspan - 1, max_rows - 1)
                    end_col = min(current_col + colspan - 1, max_cols - 1)
                    if end_row > row_idx or end_col > current_col:
                        target_cell.merge(word_table.rows[end_row].cells[end_col])
                except Exception:
                    pass

            for paragraph in target_cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = DocumentStyling.FONT_NAME
                    run.font.size = Pt(7) if row_idx > 1 else Pt(8)
                    if row_idx <= 1:
                        run.font.bold = True
            current_col += colspan


def write_section_5_2(doc: Document, *, nstudies: int, medname: str, reporting_period: str,
                      total_subjects: int, gender_text: str, age_text: str, race_text: str,
                      table_structure: dict | None) -> None:
    section_5_heading = "5 ESTIMATED EXPOSURE AND USE PATTERNS"
    section5_1_subheading = "5.1 General considerations"
    section5_1_para = (
        "For clinical trials, patient exposure can be accurately calculated because dosage and duration of "
        "treatment are clearly known. In terms of post-marketing use, patient exposure cannot be accurately "
        "calculated for certain reasons such as varying dosage and duration of treatment as well as changing "
        "or unknown patient compliance."
    )
    section5_2_subheading = "5.2 Cumulative subject exposure in clinical trials"

    study_word = "study" if nstudies == 1 else "studies"

    # Map provided values to the variable names used in the paragraph template
    race_data = race_text
    gender_data = gender_text
    age_data = age_text

    section5_2_para = (
        f"Jubilant, as MAH, has not conducted any clinical trials. However, Jubilant has conducted "
        f"{nstudies:02d} BA/BE {study_word} with {medname} till the DLP of the PSUR "
        f"{reporting_period.split(' to ')[-1]}, and cumulative subject exposure in the completed clinical trials "
        f"were {total_subjects} subjects.\n\n"
        f"Of these {total_subjects} subjects, all were {race_data} {gender_data} of age  distribution between {age_data}. "
        f"Cumulative subject exposure to {medname} in BA/BE studies is given in the table below:"
    )

    p = doc.add_paragraph(section_5_heading)
    p.style = 'Heading 1'

    p = doc.add_paragraph(section5_1_subheading)
    p.style = 'Heading 2'
    p = doc.add_paragraph(section5_1_para)
    for run in p.runs:
        DocumentStyling.apply_content_style(run)

    p = doc.add_paragraph(section5_2_subheading)
    p.style = 'Heading 2'
    p = doc.add_paragraph(section5_2_para)
    for run in p.runs:
        DocumentStyling.apply_content_style(run)

    title = f"Cumulative subject exposure to {medname} in BA/BE studies"
    create_word_table_from_html_structure(doc, table_structure, title)
