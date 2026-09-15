"""The "sra-brief" output format: A4 portrait, a SINGLE flat Word table
covering every SRA finding, with only 3 columns: #, Findings, Risk Level.

Reuses the SAME data-extraction logic as "landscape-detail" (the
(section_title, [Finding, ...]) groups produced by
efw.extraction.extract_findings()) - this format is simply a lower-detail
subset of that same data, similar in spirit to "sa-brief" (a condensed
counterpart of "sa-detail").

Layout (per the reference table provided):
    #  | Findings | Risk Level
    -------------------------------------
    **General Control Review**            <- merged section-header row
    G1 | ...      | Medium
    G2 | ...      | Low
    ...
    **Vulnerability Scanning**             <- merged section-header row
    V1 | ...      | Critical
    ...

Styling:
    - The "#" / "Findings" / "Risk Level" header row: text is BOLD, but -
      unlike every other format's blue/white header row - this row is
      explicitly LEFT UN-FILLED (no shading at all), per user
      specification.
    - Section-header rows (e.g. "General Control Review") are merged across
      all 3 columns, BOLD text, and filled with the standard MS Word theme
      swatch "Blue, Accent 1, Lighter 80%" (#D9E2F3).
    - Risk Level VALUE cells are colored using the EXACT SAME color scheme
      as "portrait-detail"/"landscape-detail" (RISK_LEVEL_COLORS).
    - Data rows (# / Findings / Risk Level) are NOT bold, no fill.
    - "Repeat Header Rows" is enabled on the #/Findings/Risk Level header
      row only (not on the section-header rows).
"""

from typing import Optional

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Cm

from ..constants import BLACK_AUTO_COLOR, RISK_LEVEL_COLORS, SRA_BRIEF_SECTION_FILL
from ..diagnostics import warn
from ..docx_utils import (
    _apply_base_styles,
    _set_cell_fill,
    _set_cell_text,
    _set_repeat_header_row,
    _set_table_column_widths,
    _setup_a4_portrait,
)
from ..models import Finding
from ..text_utils import normalize

# Content width for A4 portrait with 1.5cm margins on each side (same
# convention as "sa-brief"/"veri-summary"): 21cm - 3cm = 18cm.
SRA_BRIEF_COLUMN_WIDTHS = [Cm(1.5), Cm(13.0), Cm(3.5)]  # #, Findings, Risk Level
SRA_BRIEF_HEADER_LABELS = ["#", "Findings", "Risk Level"]


def build_sra_brief_document(
    groups: list[tuple[Optional[str], list[Finding]]],
    title: str = "Follow-up Findings (Brief)",
) -> Document:
    """Build the "sra-brief" output: A4 portrait, ONE flat table covering
    every SRA finding (grouped under bold, merged section-header rows),
    with only #, Findings, and Risk Level columns."""
    document = Document()
    _apply_base_styles(document)
    _setup_a4_portrait(document)
    document.add_heading(title, level=1)

    table = document.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    table.autofit = False
    _set_table_column_widths(table, SRA_BRIEF_COLUMN_WIDTHS)

    # ---- Header row (#, Findings, Risk Level) - BOLD text, but explicitly
    # LEFT UN-FILLED (no shading), unlike every other format's blue/white
    # header row convention. ----
    header_row = table.rows[0]
    for i, label in enumerate(SRA_BRIEF_HEADER_LABELS):
        cell = header_row.cells[i]
        cell.width = SRA_BRIEF_COLUMN_WIDTHS[i]
        _set_cell_text(cell, [label], bold=True, font_color=BLACK_AUTO_COLOR)
        _set_cell_fill(cell, None)  # explicitly un-filled - NOT blue
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_repeat_header_row(header_row)

    total = 0
    for section_title, findings in groups:
        if not findings:
            continue

        # ---- Section header row: merged across all 3 columns, bold,
        # filled with "Blue, Accent 1, Lighter 80%" (#D9E2F3). ----
        section_row = table.add_row()
        section_cell = section_row.cells[0]
        for c in section_row.cells[1:]:
            section_cell = section_cell.merge(c)
        _set_cell_text(section_cell, [section_title or "Findings"], bold=True, font_color=BLACK_AUTO_COLOR)
        _set_cell_fill(section_cell, SRA_BRIEF_SECTION_FILL)
        section_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        # ---- Data rows: #, Findings (title), Risk Level (colored) ----
        for finding in findings:
            row = table.add_row()
            cells = row.cells
            for i, width in enumerate(SRA_BRIEF_COLUMN_WIDTHS):
                cells[i].width = width

            _set_cell_text(cells[0], [finding.finding_id], bold=False, font_color=BLACK_AUTO_COLOR)
            _set_cell_fill(cells[0], None)
            cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP

            _set_cell_text(cells[1], [finding.finding_title], bold=False, font_color=BLACK_AUTO_COLOR)
            _set_cell_fill(cells[1], None)
            cells[1].vertical_alignment = WD_ALIGN_VERTICAL.TOP

            _set_cell_text(cells[2], [finding.risk_level], bold=False, font_color=BLACK_AUTO_COLOR)
            risk_hex = RISK_LEVEL_COLORS.get(normalize(finding.risk_level))
            _set_cell_fill(cells[2], risk_hex)  # None if unrecognized -> no fill
            cells[2].vertical_alignment = WD_ALIGN_VERTICAL.TOP

            total += 1

    if total == 0:
        warn("No findings were extracted - the output document will be empty of tables.")

    return document
