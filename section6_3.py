from __future__ import annotations

from typing import Optional

from docx import Document


def write_section_6_3(doc: Document, *, cumulative_text: str, interval_text: Optional[str] = None) -> None:
    # 6 DATA IN SUMMARY TABULATIONS
    doc.add_heading("6 DATA IN SUMMARY TABULATIONS", level=1)

    # 6.1 Reference Information
    doc.add_heading("6.1 Reference Information", level=2)
    doc.add_paragraph(
        "The Medical Dictionary for Regulatory Activities (MedDRA) versions from 24.1 to 27.1, is valid in the reporting period of this PSUR and was used for coding adverse events. The summary tabulations are sorted alphabetically by primary System Organ Class (SOC) and Preferred Term (PT) level."
    )

    # 6.2 Cumulative Summary Tabulations of Serious Adverse Events from Clinical Trials
    doc.add_heading("6.2 Cumulative Summary Tabulations of Serious Adverse Events from Clinical Trials", level=2)
    doc.add_paragraph(
        "No information was available as no clinical trials have been conducted by the MAH since obtaining MA for Levetiracetam."
    )

    # 6.3 Cumulative and interval summary tabulations from Post-Marketing Data Sources
    doc.add_heading(
        "6.3 Cumulative and interval summary tabulations from Post-Marketing Data Sources",
        level=2,
    )
    doc.add_paragraph(
        "The Safety Database was searched for all individual case safety reports (ICSRs) (also called cases) meeting the criteria mentioned below."
    )
    doc.add_paragraph(
        "Serious and non-serious Adverse Drug Reactions (ADRs) from spontaneous ICSRs, including reports from healthcare professionals, consumers, scientific literature and regulatory authorities."
    )
    doc.add_paragraph(
        "Note: As described in ICH Guideline E2D, spontaneously reported AEs usually imply at least a suspicion of causality by the reporter and should be considered to be adverse reactions for regulatory reporting purposes."
    )
    doc.add_paragraph(
        "Serious adverse reactions from non-interventional studies and from any other non-interventional solicited sources.",
        style="List Bullet",
    )
    doc.add_paragraph(
        "From initial MA worldwide by the MAH to the data lock point (DLP) of this report (= cumulative data set)",
        style="List Bullet",
    )
    doc.add_paragraph(
        "Received during the period of this report (= interval data set)",
        style="List Bullet",
    )

    # Now add generated cumulative and interval narratives
    if cumulative_text:
        p_cum = doc.add_paragraph("")
        run_cum = p_cum.add_run("Cumulative summary tabulations:")
        run_cum.bold = True
        run_cum.underline = True
        doc.add_paragraph(cumulative_text)
    if interval_text:
        p_int = doc.add_paragraph("")
        run_int = p_int.add_run("Interval summary tabulations:")
        run_int.bold = True
        run_int.underline = True
        doc.add_paragraph(interval_text)

    # Closing narrative
    doc.add_paragraph("No patterns or clusters were observed from these cases.")
    doc.add_paragraph("")
    
    doc.add_paragraph(
        "A single table of summary tabulation of serious and non-serious reactions is presented side-by-side and is organized by MedDRA SOC. A summary tabulation of adverse reactions for esomeprazole as extracted from company safety database has been appended as Appendix 20.3 (Table B)."
    )
