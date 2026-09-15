"""The "veri-summary-by-section" and "veri-summary-executive" output
formats: both render Word table(s) (styled as the built-in "Grid Table 4 -
Accent 6" style) containing Risk-Level x Rectification-Status count blocks
for the SRA data, PLUS a Compliance-Status x Rectification-Status count
block for the SA ("Security Audit") data:

    - "veri-summary-by-section": ONE SINGLE combined table - one block PER
      SRA SECTION, stacked in order, followed by the SA block ("Security
      Audit") APPENDED to that SAME table, so all blocks (SRA and SA
      alike) share an identical column structure - i.e. a single
      "SRA and SA" verification summary table as a whole.

    - "veri-summary-executive": the SRA data becomes ONE combined block in
      ONE table (as before), and the SA data becomes a SEPARATE, second
      table (titled "Security Audit") immediately below it, separated by a
      single blank paragraph line.

They share every styling/formatting rule and the row/column inclusion
logic (Critical row, Partially Completed column - decided ONCE, globally,
across BOTH the SRA and SA data so every block/table has an identical
column set), differing only in how findings are grouped into
blocks/tables.
"""

from typing import Optional

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Mm

from ..constants import BLACK_AUTO_COLOR, RISK_LEVEL_DISPLAY
from ..diagnostics import warn
from ..docx_utils import (
    _add_blank_separator_paragraph,
    _apply_base_styles,
    _set_cell_fill,
    _set_cell_text,
    _set_repeat_header_row,
    _set_row_height,
    _set_table_column_widths,
)
from ..models import Finding, SAFinding
from ..text_utils import normalize

# =============================================================================
# "veri-summary-by-section" format: ONE combined table (Word style
# "Grid Table 4 - Accent 6"), one Risk-Level x Rectification-Status count
# block per section, A4 portrait ("Verification Summary per Section")
# =============================================================================

# Base (always-shown) risk levels, in display order. "critical" is inserted
# at the front of this list at render time ONLY if used anywhere in the
# workbook (see compute_veri_summary_flags()).
_VERI_SUMMARY_BASE_RISK_LEVELS = ["high", "medium", "low", "ofi"]

# Base (always-shown) status columns, in display order. "Partially Completed"
# is inserted between "Completed" and "Incomplete" ONLY if used anywhere in
# the workbook (see compute_veri_summary_flags()).
_VERI_SUMMARY_BASE_STATUS_COLUMNS = ["Completed", "Incomplete", "Scheduled", "Accepted"]

# (label, width) for the two always-present leading columns.
VERI_SUMMARY_LEADING_COLUMNS = [
    ("Risk Level", Cm(2.6)),
    ("Total", Cm(1.6)),
]
# width used for EVERY status column (Completed/Partially Completed/
# Incomplete/Scheduled/Accepted) - kept uniform for a clean grid, wide
# enough that "Incomplete"/"Partially Completed" (the two longest labels)
# don't wrap awkwardly mid-word.
VERI_SUMMARY_STATUS_COL_WIDTH = Cm(2.6)

# Fixed row height for ONLY the section title row ("Security Risk
# Assessment - <Section>" / "Security Audit"). Uses hRule="atLeast" so the
# row is guaranteed to be at least this tall, but will still grow if the
# title text ever needs more vertical space than that (e.g. a very long
# section name that wraps onto 2 lines) - avoiding any risk of
# clipped/overlapping text.
VERI_SUMMARY_TITLE_ROW_HEIGHT = Cm(0.9)

# ---- "Grid Table 4 - Accent 6" built-in Word table style ----
# python-docx's default template does not ship this style (only "Table
# Grid" and "Normal Table" are included), so we inject a full definition,
# translated from the OOXML produced by real Word / the OOXML SDK for this
# exact built-in style. All colors are expressed as theme (accent6)
# references (with a literal fallback matching python-docx's bundled
# "Office" theme, where accent6 = F79646, an orange) so the table
# automatically follows the *scheme*, not a hardcoded color choice.
GRID_TABLE_4_ACCENT6_STYLE_ID = "GridTable4Accent6"
GRID_TABLE_4_ACCENT6_STYLE_NAME = "Grid Table 4 Accent 6"
THEME_ACCENT6_HEX = "F79646"  # accent6 in python-docx's bundled "Office" theme

# Hardcoded fill color for ONLY the section/table title row ("Security Risk
# Assessment - <Section>" / "Security Audit"). Explicitly hardcoded per
# user request (rather than derived from the table style/theme like the
# borders and banding are), so it does NOT change if the attached table
# style's theme color changes.
VERI_SUMMARY_HEADER_FILL = "FFC000"


def _theme_tint_hex(hex_color: str, tint_255: int) -> str:
    """Approximate the Office 'theme tint' lightening transform: a simple
    per-channel linear blend toward white. `tint_255` is 0-255, matching the
    2-hex-digit value used by OOXML's w:themeTint/w:themeFillTint (e.g.
    0x33 = 51 -> a light/pale tint; 0xFF = 255 -> no change)."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    factor = tint_255 / 255.0
    r2 = round(255 - (255 - r) * factor)
    g2 = round(255 - (255 - g) * factor)
    b2 = round(255 - (255 - b) * factor)
    return f"{r2:02X}{g2:02X}{b2:02X}"


def _ensure_grid_table_4_accent6_style(document: Document) -> None:
    """Inject the "Grid Table 4 - Accent 6" built-in table style definition
    into the document's styles part, if not already present. Idempotent."""
    styles_element = document.styles.element
    style_id_attr = qn("w:styleId")
    for existing in styles_element.findall(qn("w:style")):
        if existing.get(style_id_attr) == GRID_TABLE_4_ACCENT6_STYLE_ID:
            return  # already injected

    band_tint_hex = _theme_tint_hex(THEME_ACCENT6_HEX, 0x33)

    style_xml = f"""
    <w:style {nsdecls('w')} w:type="table" w:styleId="{GRID_TABLE_4_ACCENT6_STYLE_ID}">
      <w:name w:val="{GRID_TABLE_4_ACCENT6_STYLE_NAME}"/>
      <w:basedOn w:val="TableNormal"/>
      <w:uiPriority w:val="49"/>
      <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <w:tblPr>
        <w:tblStyleRowBandSize w:val="1"/>
        <w:tblStyleColBandSize w:val="1"/>
        <w:tblBorders>
          <w:top w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:themeTint="99" w:sz="4" w:space="0"/>
          <w:left w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:themeTint="99" w:sz="4" w:space="0"/>
          <w:bottom w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:themeTint="99" w:sz="4" w:space="0"/>
          <w:right w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:themeTint="99" w:sz="4" w:space="0"/>
          <w:insideH w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:themeTint="99" w:sz="4" w:space="0"/>
          <w:insideV w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:themeTint="99" w:sz="4" w:space="0"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblStylePr w:type="firstRow">
        <w:rPr>
          <w:b/>
          <w:bCs/>
          <w:color w:val="FFFFFF" w:themeColor="background1"/>
        </w:rPr>
        <w:tblPr/>
        <w:tcPr>
          <w:tcBorders>
            <w:top w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:sz="4" w:space="0"/>
            <w:left w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:sz="4" w:space="0"/>
            <w:bottom w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:sz="4" w:space="0"/>
            <w:right w:val="single" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:sz="4" w:space="0"/>
            <w:insideH w:val="nil"/>
            <w:insideV w:val="nil"/>
          </w:tcBorders>
          <w:shd w:val="clear" w:color="auto" w:fill="{THEME_ACCENT6_HEX}" w:themeFill="accent6"/>
        </w:tcPr>
      </w:tblStylePr>
      <w:tblStylePr w:type="lastRow">
        <w:rPr>
          <w:b/>
          <w:bCs/>
        </w:rPr>
        <w:tblPr/>
        <w:tcPr>
          <w:tcBorders>
            <w:top w:val="double" w:color="{THEME_ACCENT6_HEX}" w:themeColor="accent6" w:sz="4" w:space="0"/>
          </w:tcBorders>
        </w:tcPr>
      </w:tblStylePr>
      <w:tblStylePr w:type="firstCol">
        <w:rPr>
          <w:b/>
          <w:bCs/>
        </w:rPr>
      </w:tblStylePr>
      <w:tblStylePr w:type="lastCol">
        <w:rPr>
          <w:b/>
          <w:bCs/>
        </w:rPr>
      </w:tblStylePr>
      <w:tblStylePr w:type="band1Vert">
        <w:tcPr>
          <w:shd w:val="clear" w:color="auto" w:fill="{band_tint_hex}" w:themeFill="accent6" w:themeFillTint="33"/>
        </w:tcPr>
      </w:tblStylePr>
      <w:tblStylePr w:type="band1Horz">
        <w:tcPr>
          <w:shd w:val="clear" w:color="auto" w:fill="{band_tint_hex}" w:themeFill="accent6" w:themeFillTint="33"/>
        </w:tcPr>
      </w:tblStylePr>
    </w:style>
    """
    style_el = parse_xml(style_xml)
    styles_element.append(style_el)


def _setup_a4_portrait(document: Document) -> None:
    """Configure the document's first section as A4, portrait orientation."""
    section = document.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)


def compute_veri_summary_flags(
    groups: list[tuple[Optional[str], list[Finding]]],
    sa_findings: Optional[list[SAFinding]] = None,
) -> tuple[bool, bool]:
    """Scan ALL findings across ALL SRA sections, PLUS all SA items (if
    provided), globally (not per-section/per-table) to decide once, for
    the WHOLE output:
        - has_critical: whether the "Critical" risk-level row should be
          shown (above "High") in every SRA block. SA items have no risk
          level concept at all, so `sa_findings` never affects this flag.
        - has_partial: whether the "Partially Completed" status column
          should be shown (between "Completed" and "Incomplete") in every
          block/table - SRA AND SA alike, since they must share an
          identical column structure (required for
          "veri-summary-by-section", where SRA and SA blocks live in the
          SAME physical table; kept consistent for "veri-summary-executive"
          too, even though its SRA/SA tables are physically separate, so
          both formats always agree on the same column set for the same
          input data).
    Deciding this globally keeps every block/table the same shape."""
    has_critical = False
    has_partial = False
    for _, findings in groups:
        for f in findings:
            if normalize(f.risk_level) == "critical":
                has_critical = True
            if f.verification_status == "Partially Completed":
                has_partial = True
    if sa_findings:
        for f in sa_findings:
            if f.verification_status == "Partially Completed":
                has_partial = True
    return has_critical, has_partial


def _veri_summary_risk_levels(has_critical: bool) -> list[str]:
    return (["critical"] if has_critical else []) + _VERI_SUMMARY_BASE_RISK_LEVELS


def _veri_summary_status_columns(has_partial: bool) -> list[str]:
    cols = list(_VERI_SUMMARY_BASE_STATUS_COLUMNS)
    if has_partial:
        cols.insert(1, "Partially Completed")  # between Completed and Incomplete
    return cols


def _count_section_by_risk_and_status(
    findings: list[Finding], risk_levels: list[str], status_columns: list[str]
) -> tuple[dict, dict]:
    """Returns (counts, totals):
        - counts[risk_level][status_column] = number of findings in this
          section with that risk level AND that (canonical) verification
          status.
        - totals[risk_level] = total number of findings in this section
          with that risk level (regardless of verification status,
          including unmatched/invalid statuses - those are already flagged
          via warnings elsewhere and simply don't add to any status column
          here, but DO still count towards the risk level's Total)."""
    counts = {rl: {col: 0 for col in status_columns} for rl in risk_levels}
    totals = {rl: 0 for rl in risk_levels}

    for f in findings:
        rl_norm = normalize(f.risk_level)
        if rl_norm not in risk_levels:
            # Either an unrecognized risk level (already warned elsewhere)
            # or - for "critical" - a level not shown because has_critical
            # was globally False (shouldn't happen, since has_critical is
            # computed FROM these same findings, but guarded defensively).
            continue
        totals[rl_norm] += 1
        if f.verification_status in status_columns:
            counts[rl_norm][f.verification_status] += 1

    return counts, totals


# Column indices (within the single merged table) that hold NUMERIC counts
# and their corresponding header labels, used to decide center-alignment.
def _veri_summary_numeric_col_indices(n_status_cols: int) -> list[int]:
    # Column 0 = Risk Level (label, never numeric); column 1 = Total; then
    # one column per status.
    return list(range(1, 2 + n_status_cols))


def add_veri_summary_section_block(
    table,
    title_text: str,
    findings: list[Finding],
    risk_levels: list[str],
    status_columns: list[str],
) -> None:
    """Append ONE "Verification Summary" block (title row + 2 header rows +
    risk-level rows + Total row) onto the END of an existing, shared table.

    `title_text` is used VERBATIM as the title row's text - callers are
    responsible for formatting it (e.g. "Security Risk Assessment - <Section>"
    for the per-section format, or plain "Security Risk Assessment" for the
    executive/combined format), so this function can be shared by both
    "veri-summary-by-section" (one block per section) and
    "veri-summary-executive" (a single block covering ALL findings combined,
    ignoring section boundaries).

    ONLY the title row gets the solid header fill (hardcoded to
    VERI_SUMMARY_HEADER_FILL, per user request), a FIXED row height
    (VERI_SUMMARY_TITLE_ROW_HEIGHT), and is marked as a repeating header
    row. The 2 header-label rows below it ("Risk Level" / "Number of items
    by Rectification Status" / individual status column names) are
    deliberately left WITHOUT any explicit shading, WITHOUT a fixed height,
    and WITHOUT the repeating-header flag, so they - like the risk-level
    data rows and the Total row - simply inherit the attached "Grid Table 4
    - Accent 6" table style's own alternating band1Horz shading (pale
    orange / no-fill) based on their absolute position in the table. This
    shading continues seamlessly across section boundaries (verified
    empirically - see module docstring)."""
    counts, totals = _count_section_by_risk_and_status(findings, risk_levels, status_columns)
    n_status_cols = len(status_columns)
    n_cols = 2 + n_status_cols
    numeric_cols = set(_veri_summary_numeric_col_indices(n_status_cols))

    # ---- Title row (merged across all columns) - THE ONLY row in this
    # block that gets the solid header fill, fixed height, and repeats as
    # a page header. ----
    title_row = table.add_row()
    title_cell = title_row.cells[0]
    for c in title_row.cells[1:]:
        title_cell = title_cell.merge(c)
    _set_cell_text(
        title_cell, [title_text],
        bold=True, font_color=BLACK_AUTO_COLOR,
    )
    _set_cell_fill(title_cell, VERI_SUMMARY_HEADER_FILL)
    title_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_repeat_header_row(title_row)
    _set_row_height(title_row, VERI_SUMMARY_TITLE_ROW_HEIGHT, rule="atLeast")

    # ---- Header row A + B ----
    # No explicit shading, no fixed height, and no repeat-header flag on
    # either row: they fall back to the table style's normal alternating
    # band shading (and natural/auto height), same as any other data row.
    header_row_a = table.add_row()
    header_row_b = table.add_row()

    # "Risk Level" - UNMERGED: placed only in the upper cell (row A). The
    # lower cell (row B, column 0) is left empty. Bold (per requirement 5),
    # black text.
    _set_cell_text(header_row_a.cells[0], ["Risk Level"], bold=True, font_color=BLACK_AUTO_COLOR)
    header_row_a.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_cell_text(header_row_b.cells[0], [""], bold=False, font_color=BLACK_AUTO_COLOR)
    header_row_b.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Column 1: blank in row A, "Total" in row B. "Total" is a numbering
    # column label -> center-aligned, not bold (only "Risk Level" and
    # Total-ROW cells are bold per requirement 5).
    _set_cell_text(header_row_a.cells[1], [""], bold=False, font_color=BLACK_AUTO_COLOR)
    _set_cell_text(
        header_row_b.cells[1], ["Total"], bold=False, font_color=BLACK_AUTO_COLOR,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    # "Number of items by Rectification Status" - merged horizontally across
    # all status columns, in row A only. Not bold, black text, left-aligned
    # (not a numbering/status-label cell itself).
    super_header_cell = header_row_a.cells[2]
    for c in header_row_a.cells[3:]:
        super_header_cell = super_header_cell.merge(c)
    _set_cell_text(
        super_header_cell, ["Number of items by Rectification Status"],
        bold=False, font_color=BLACK_AUTO_COLOR,
    )

    # Individual status column labels in row B - these ARE the "verification
    # status cells that label the columns" -> center-aligned, not bold.
    for i, col_name in enumerate(status_columns):
        cell = header_row_b.cells[2 + i]
        _set_cell_text(
            cell, [col_name], bold=False, font_color=BLACK_AUTO_COLOR,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # ---- Risk-level data rows ----
    # Shading intentionally NOT set here: the attached table style's
    # band1Horz/band2Horz conditional formatting supplies the alternating
    # pale/unfilled fill automatically, based on each row's absolute
    # position in the whole (multi-section) table.
    for rl in risk_levels:
        row = table.add_row()

        label_cell = row.cells[0]
        _set_cell_text(label_cell, [RISK_LEVEL_DISPLAY[rl]], bold=False, font_color=BLACK_AUTO_COLOR)
        label_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        total_cell = row.cells[1]
        _set_cell_text(
            total_cell, [str(totals[rl])], bold=False, font_color=BLACK_AUTO_COLOR,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        total_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        for j, col in enumerate(status_columns):
            cell = row.cells[2 + j]
            _set_cell_text(
                cell, [str(counts[rl][col])], bold=False, font_color=BLACK_AUTO_COLOR,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
            )
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # ---- Final "Total" row: EVERY cell bold (per requirement 5); numeric
    # cells center-aligned; shading again left to the table style. ----
    total_row = table.add_row()
    _set_cell_text(total_row.cells[0], ["Total"], bold=True, font_color=BLACK_AUTO_COLOR)
    total_row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP

    grand_total = sum(totals.values())
    cell = total_row.cells[1]
    _set_cell_text(
        cell, [str(grand_total)], bold=True, font_color=BLACK_AUTO_COLOR,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    for j, col in enumerate(status_columns):
        col_total = sum(counts[rl][col] for rl in risk_levels)
        cell = total_row.cells[2 + j]
        _set_cell_text(
            cell, [str(col_total)], bold=True, font_color=BLACK_AUTO_COLOR,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP


# =============================================================================
# SA (Security Audit) summary block - shares the SAME column
# structure/styling as the SRA blocks above, so it can be appended to the
# SAME shared table ("veri-summary-by-section") or rendered into its own,
# separately-styled-but-structurally-identical table
# ("veri-summary-executive").
# =============================================================================


def compute_sa_summary_counts(
    sa_findings: list[SAFinding], status_columns: list[str]
) -> tuple[dict, int]:
    """Returns (counts, total) for the SA "Non-compliance" row:
        - counts[status_column] = number of SA items with that (canonical)
          verification status.
        - total = total number of SA items (EVERY item in the "SA
          Follow-up" table is inherently a non-compliance finding - there
          is no separate "not found"/"pass" category - so this is simply
          len(sa_findings))."""
    counts = {col: 0 for col in status_columns}
    total = 0
    for f in sa_findings:
        total += 1
        if f.verification_status in status_columns:
            counts[f.verification_status] += 1
    return counts, total


def add_sa_summary_block(
    table,
    title_text: str,
    sa_findings: list[SAFinding],
    status_columns: list[str],
) -> None:
    """Append the SA (Security Audit) "Verification Summary" block onto the
    END of an existing table: a title row ("Security Audit") + 2 header
    rows ("Compliance Status" / "Number of items by Rectification Status" /
    individual status column names) + a SINGLE "Non-compliance" data row.

    Unlike the SRA blocks (which have several risk-level rows PLUS a
    separate, final "Total" row), the SA table has only ONE compliance
    category - every item extracted from the "SA Follow-up" sheet IS
    (by definition) a non-compliance finding - so that single row doubles
    as BOTH the category row AND the grand-total row, and therefore (like
    the SRA format's Total row) has EVERY cell bold, per the reference
    layout provided by the user.

    Column count/widths/fonts/shading conventions are IDENTICAL to
    add_veri_summary_section_block()'s SRA blocks (reusing the same
    VERI_SUMMARY_HEADER_FILL / VERI_SUMMARY_TITLE_ROW_HEIGHT / table-style
    banding), so this can be appended to the SAME shared table
    ("veri-summary-by-section", producing one combined "SRA and SA" table)
    or rendered into an equally-styled but separate table
    ("veri-summary-executive")."""
    counts, total = compute_sa_summary_counts(sa_findings, status_columns)

    # ---- Title row (merged across all columns) - "Security Audit" ----
    title_row = table.add_row()
    title_cell = title_row.cells[0]
    for c in title_row.cells[1:]:
        title_cell = title_cell.merge(c)
    _set_cell_text(title_cell, [title_text], bold=True, font_color=BLACK_AUTO_COLOR)
    _set_cell_fill(title_cell, VERI_SUMMARY_HEADER_FILL)
    title_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_repeat_header_row(title_row)
    _set_row_height(title_row, VERI_SUMMARY_TITLE_ROW_HEIGHT, rule="atLeast")

    # ---- Header row A + B ----
    header_row_a = table.add_row()
    header_row_b = table.add_row()

    # "Compliance Status" - mirrors "Risk Level": bold, upper cell only.
    _set_cell_text(header_row_a.cells[0], ["Compliance Status"], bold=True, font_color=BLACK_AUTO_COLOR)
    header_row_a.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_cell_text(header_row_b.cells[0], [""], bold=False, font_color=BLACK_AUTO_COLOR)
    header_row_b.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.TOP

    _set_cell_text(header_row_a.cells[1], [""], bold=False, font_color=BLACK_AUTO_COLOR)
    _set_cell_text(
        header_row_b.cells[1], ["Total"], bold=False, font_color=BLACK_AUTO_COLOR,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    super_header_cell = header_row_a.cells[2]
    for c in header_row_a.cells[3:]:
        super_header_cell = super_header_cell.merge(c)
    _set_cell_text(
        super_header_cell, ["Number of items by Rectification Status"],
        bold=False, font_color=BLACK_AUTO_COLOR,
    )

    for i, col_name in enumerate(status_columns):
        cell = header_row_b.cells[2 + i]
        _set_cell_text(
            cell, [col_name], bold=False, font_color=BLACK_AUTO_COLOR,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # ---- Single "Non-compliance" row: EVERY cell bold (it doubles as the
    # category AND the grand-total row) - shading left to the table style,
    # exactly like the SRA Total row. ----
    data_row = table.add_row()
    label_cell = data_row.cells[0]
    _set_cell_text(label_cell, ["Non-compliance"], bold=True, font_color=BLACK_AUTO_COLOR)
    label_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    total_cell = data_row.cells[1]
    _set_cell_text(
        total_cell, [str(total)], bold=True, font_color=BLACK_AUTO_COLOR,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )
    total_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    for j, col in enumerate(status_columns):
        cell = data_row.cells[2 + j]
        _set_cell_text(
            cell, [str(counts[col])], bold=True, font_color=BLACK_AUTO_COLOR,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP


# =============================================================================
# Shared document/table setup
# =============================================================================


def _create_veri_summary_table(document: Document, status_columns: list[str]):
    """Create a new, empty (correctly-styled and correctly-sized) "Grid
    Table 4 - Accent 6" table on the given document, ready for blocks to be
    appended via add_veri_summary_section_block() / add_sa_summary_block().

    Used both for the FIRST table added to a document, and - in
    "veri-summary-executive" - for the SEPARATE second table that holds the
    SA block below the SRA table."""
    n_status_cols = len(status_columns)
    n_cols = 2 + n_status_cols
    column_widths = [w for _, w in VERI_SUMMARY_LEADING_COLUMNS] + [VERI_SUMMARY_STATUS_COL_WIDTH] * n_status_cols

    # Start with a single throwaway row (python-docx requires >=1 row to
    # create a table); we remove it immediately since every block appends
    # its own rows via add_veri_summary_section_block()/add_sa_summary_block().
    table = document.add_table(rows=1, cols=n_cols)
    table.style = GRID_TABLE_4_ACCENT6_STYLE_NAME
    table.autofit = False
    _set_table_column_widths(table, column_widths)
    placeholder_row_element = table.rows[0]._tr
    placeholder_row_element.getparent().remove(placeholder_row_element)

    return table


def _setup_veri_summary_document_and_table(
    groups: list[tuple[Optional[str], list[Finding]]],
    title: str,
    sa_findings: Optional[list[SAFinding]] = None,
) -> tuple[Document, object, list[str], list[str]]:
    """Shared setup: creates the A4-portrait document, applies base
    styling, injects the "Grid Table 4 - Accent 6" table style, adds the
    document title heading, computes the (globally-decided, across BOTH
    SRA and SA data) risk-level rows and status columns to use, and
    creates the single shared table (with correct column widths) that the
    first block(s) will be appended onto.

    Returns (document, table, risk_levels, status_columns) so callers can
    append one or more blocks via add_veri_summary_section_block() /
    add_sa_summary_block(), and/or create additional tables via
    _create_veri_summary_table(document, status_columns) using the SAME
    status_columns (so every table in the document stays aligned)."""
    document = Document()
    _apply_base_styles(document)
    _setup_a4_portrait(document)
    _ensure_grid_table_4_accent6_style(document)

    document.add_heading(title, level=1)

    has_critical, has_partial = compute_veri_summary_flags(groups, sa_findings)
    risk_levels = _veri_summary_risk_levels(has_critical)
    status_columns = _veri_summary_status_columns(has_partial)

    table = _create_veri_summary_table(document, status_columns)

    return document, table, risk_levels, status_columns


def build_veri_summary_document(
    groups: list[tuple[Optional[str], list[Finding]]],
    title: str,
    sa_findings: Optional[list[SAFinding]] = None,
) -> Document:
    """Build the "veri-summary-by-section" output: A4 portrait, ONE SINGLE
    combined table (styled as Word's built-in "Grid Table 4 - Accent 6")
    containing one compact Risk-Level x Rectification-Status count block
    PER SRA SECTION, reusing the same Finding data produced by
    extract_findings() - PLUS, if `sa_findings` is provided, the SA
    ("Security Audit") Compliance-Status x Rectification-Status block
    APPENDED onto that SAME table (so the whole "SRA and SA" verification
    summary is rendered as a single table with an identical column
    structure throughout)."""
    document, table, risk_levels, status_columns = _setup_veri_summary_document_and_table(
        groups, title, sa_findings=sa_findings
    )

    total = 0
    for section_title, findings in groups:
        if not findings:
            continue
        block_title = f"Security Risk Assessment - {section_title or 'Findings'}"
        add_veri_summary_section_block(table, block_title, findings, risk_levels, status_columns)
        total += len(findings)

    if sa_findings:
        add_sa_summary_block(table, "Security Audit", sa_findings, status_columns)
        total += len(sa_findings)

    if total == 0:
        warn("No findings were extracted - the output document will be empty of tables.")

    return document


def build_veri_summary_executive_document(
    groups: list[tuple[Optional[str], list[Finding]]],
    title: str,
    sa_findings: Optional[list[SAFinding]] = None,
) -> Document:
    """Build the "veri-summary-executive" output: identical styling,
    coloring, and row/column logic to "veri-summary-by-section", EXCEPT
    that findings from ALL SRA sections are combined into a SINGLE
    Risk-Level x Rectification-Status count block (section boundaries are
    ignored - the counts simply cover every SRA finding in the workbook),
    rendered in its own table titled plainly "Security Risk Assessment"
    (no " - <Section>" suffix).

    If `sa_findings` is provided, the SA ("Security Audit") summary is
    rendered as a SECOND, SEPARATE table (NOT appended to the SRA table),
    titled "Security Audit", with a single blank paragraph line separating
    the two tables.

    The "Critical" row / "Partially Completed" column inclusion decision
    (via compute_veri_summary_flags()) is computed ONCE across BOTH the SRA
    and SA data, so the two separate tables still share an identical
    column structure."""
    document, table, risk_levels, status_columns = _setup_veri_summary_document_and_table(
        groups, title, sa_findings=sa_findings
    )

    # Combine every finding from every section into one flat list - section
    # boundaries are intentionally NOT preserved for this format.
    all_findings = [f for _, findings in groups for f in findings]

    if all_findings:
        add_veri_summary_section_block(table, "Security Risk Assessment", all_findings, risk_levels, status_columns)
    elif not sa_findings:
        warn("No findings were extracted - the output document will be empty of tables.")

    if sa_findings:
        # SEPARATE table, one blank line below the SRA table.
        _add_blank_separator_paragraph(document)
        sa_table = _create_veri_summary_table(document, status_columns)
        add_sa_summary_block(sa_table, "Security Audit", sa_findings, status_columns)

    return document
