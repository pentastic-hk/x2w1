"""The "sa-detail" and "sa-brief" output formats: render the "SA Follow-up"
(Security Audit) half of the SRAA report, reusing as much of the SRA
data-extraction/rendering machinery as makes sense.

    - "sa-detail": A4 LANDSCAPE. ONE flat Word table (no per-section
      grouping - the SA table has no section-header rows) with columns:
      #, Items on the checklist, Affected S17 Security Domain, Findings,
      Recommendation, Rectification Status as of <date>. Styling/fonts
      match "landscape-detail": Times New Roman 12pt black (Automatic),
      standard blue (#0070C0)/white/bold header row with "Repeat Header
      Rows" enabled. All cell content is LEFT-aligned (not justified).
      The rectification date is derived the same way as "landscape-detail"
      (latest/rightmost non-empty verification column across ALL items).

    - "sa-brief": Identical styling/color-scheme/layout to "sa-detail",
      except A4 PORTRAIT and only 3 columns: #, Affected S17 Security
      Domain, Findings.
"""

from typing import Optional

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm

from ..constants import HEADER_ROW_FILL, HEADER_ROW_FONT_COLOR
from ..diagnostics import warn
from ..docx_utils import (
    _apply_base_styles,
    _set_cell_fill,
    _set_cell_text,
    _set_repeat_header_row,
    _set_table_column_widths,
    _setup_a4_landscape,
    _setup_a4_portrait,
)
from ..extraction import group_verification_status_label
from ..models import SAFinding, SAHeaderMap
from .landscape_detail import _lines_from_text

# =============================================================================
# "sa-detail": A4 landscape, one flat summary table (no sections)
# =============================================================================

# Content width for A4 landscape with 1.5cm margins on each side (same
# margins as "landscape-detail"): 29.7cm - 3cm = 26.7cm.
SA_DETAIL_COLUMNS: list[tuple[str, Optional[str], object]] = [
    ("index", "#", Cm(1.0)),
    ("items", "Items on the checklist", Cm(6.5)),
    ("affected", "Affected S17 Security Domain", Cm(4.0)),
    ("findings", "Findings", Cm(6.2)),
    ("recommendation", "Recommendation", Cm(6.0)),
    # Dynamic header ("Rectification Status as of <date>"), filled in at
    # render time - same convention as "landscape-detail".
    ("rectification", None, Cm(3.0)),
]

# Content width for A4 portrait with 1.5cm margins on each side (same
# margins as "veri-summary" formats): 21cm - 3cm = 18cm.
SA_BRIEF_COLUMNS: list[tuple[str, Optional[str], object]] = [
    ("index", "#", Cm(1.5)),
    ("affected", "Affected S17 Security Domain", Cm(6.5)),
    ("findings", "Findings", Cm(10.0)),
]


def _sa_row_values(f: SAFinding) -> dict[str, list[str]]:
    """Map each SA column `key` to the paragraph(s) to render for a given
    SA item. Only keys actually present in the chosen column list are
    looked up by the caller, so it's fine for this dict to always include
    every key regardless of which format is being rendered."""
    return {
        "index": [f.item_id],
        "items": f.items_to_check or [""],
        "affected": _lines_from_text(f.affected),
        "findings": f.findings or [""],
        "recommendation": f.recommended_safeguards or [""],
        "rectification": [f.verification_status],
    }


def _add_sa_table(
    document: Document,
    findings: list[SAFinding],
    columns: list[tuple[str, Optional[str], object]],
    rectification_header: Optional[str] = None,
):
    """Render ALL SA items as a single flat summary table (one row per
    item, no section grouping - the SA table has no section-header rows).
    Every cell (header AND data) is explicitly LEFT-aligned (not justified),
    per spec. Uses the exact same font/color/shading conventions as
    "landscape-detail" (Times New Roman 12pt black, blue/white/bold header
    row with Repeat Header Rows enabled)."""
    n_cols = len(columns)
    column_widths = [width for _, _, width in columns]
    table = document.add_table(rows=1, cols=n_cols)
    table.style = "Table Grid"
    table.autofit = False
    _set_table_column_widths(table, column_widths)

    # ---- Header row ----
    header_row = table.rows[0]
    for i, (key, label, width) in enumerate(columns):
        cell = header_row.cells[i]
        cell.width = width
        header_text = label if label is not None else (rectification_header or "Rectification Status")
        _set_cell_text(
            cell, [header_text], bold=True, font_color=HEADER_ROW_FONT_COLOR,
            alignment=WD_ALIGN_PARAGRAPH.LEFT,
        )
        _set_cell_fill(cell, HEADER_ROW_FILL)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_repeat_header_row(header_row)

    # ---- Data rows (one per SA item) ----
    for f in findings:
        row = table.add_row()
        cells = row.cells
        for i, (_, _, width) in enumerate(columns):
            cells[i].width = width

        row_values = _sa_row_values(f)
        for i, (key, _, _) in enumerate(columns):
            cell = cells[i]
            _set_cell_text(cell, row_values[key], bold=False, alignment=WD_ALIGN_PARAGRAPH.LEFT)
            _set_cell_fill(cell, None)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    return table


def build_sa_detail_document(
    findings: list[SAFinding],
    hmap: SAHeaderMap,
    title: str = "Security Audit Findings",
) -> Document:
    """Build the "sa-detail" output: A4 landscape, ONE flat table covering
    every SA item, reusing the same Finding-style data produced by
    extract_sa_findings()."""
    document = Document()
    _apply_base_styles(document)
    _setup_a4_landscape(document)
    document.add_heading(title, level=1)

    # Same convention as "landscape-detail": derive a SINGLE dynamic
    # "Rectification Status as of <date>" header from the rightmost
    # verification column that has any data anywhere in the table.
    rectification_header = group_verification_status_label(findings, hmap)

    _add_sa_table(document, findings, SA_DETAIL_COLUMNS, rectification_header=rectification_header)

    if not findings:
        warn("No SA items were extracted - the output document will be empty of tables.")

    return document


def build_sa_brief_document(
    findings: list[SAFinding],
    title: str = "Security Audit Findings (Brief)",
) -> Document:
    """Build the "sa-brief" output: A4 portrait, ONE flat table with only
    #, Affected S17 Security Domain, and Findings columns. All other
    styling/color-scheme/layout follows "sa-detail" (which in turn follows
    "landscape-detail")."""
    document = Document()
    _apply_base_styles(document)
    _setup_a4_portrait(document)
    document.add_heading(title, level=1)

    _add_sa_table(document, findings, SA_BRIEF_COLUMNS, rectification_header=None)

    if not findings:
        warn("No SA items were extracted - the output document will be empty of tables.")

    return document
