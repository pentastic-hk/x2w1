#!/usr/bin/env python3
"""
excel_to_word_findings.py
==========================

Convert a cybersecurity "Follow-up Plan" Excel workbook into a formal Word
report, rendering ONE Word table per finding (General Control Review,
Vulnerability Scanning, Web/API Penetration Testing, Source Code Review, etc).

-------------------------------------------------------------------------
HOW IT LOCATES THE DATA
-------------------------------------------------------------------------
1. Opens the worksheet named "Follow-up Items" (falls back to the 3rd sheet
   in the workbook if that exact name isn't found, with a warning).
2. Scans every cell in the worksheet and finds the bounding box of ALL cells
   that have at least one visible border side set. This bounding box is
   assumed to be the findings table (borders are the only reliable signal,
   per the source workbook's convention, since the table doesn't always
   start at A1).
3. The first row of that bounding box is treated as the header row. Column
   purposes are identified by matching header text (case-insensitive,
   whitespace-normalized) against known labels:
       - 1st column (leftmost)      -> Finding ID (e.g. "W1", "MP1")
       - "Finding"                  -> Finding title
       - "Affected*"                -> Affected asset / URL / endpoint
       - "Risk Level"                -> Risk Level
       - "Impact" (optional)         -> Impact
       - "Likelihood" (optional)     -> Likelihood
       - "Risk Description"          -> Risk Description
       - "Recommend*"/"Safeguard*"   -> Recommended Safeguards
       - "Verification*"             -> one or more verification columns
                                         (rightmost non-empty per row wins)
   Any other column immediately to the left of the first "Verification*"
   column is treated as the (unused-in-output) Vendor Response column.
4. Rows that are fully merged across the table width are treated as SECTION
   HEADER rows (e.g. "General Control Review", "Penetration Testing") and
   are rendered as Word Heading 2 paragraphs instead of finding tables.
5. Every other non-blank row within the bounding box is treated as a single
   finding and validated + rendered as its own 2-column Word table.

-------------------------------------------------------------------------
VALIDATION
-------------------------------------------------------------------------
The script warns (stderr) - but does NOT fail - on unexpected values for:
    - Risk Level     : Critical / High / Medium / Low / OFI
    - Impact         : Critical / High / Medium / Low / Very Low
    - Likelihood     : High / Medium / Low / Very Low
    - Verification   : must START WITH one of Completed / Partially
                        completed / Incomplete / Scheduled / Accepted
All warnings are also collected and summarized at the end of the run.

-------------------------------------------------------------------------
WORD TABLE STYLING
-------------------------------------------------------------------------
    - All text: Times New Roman, 12pt.
    - First row of each finding table (Finding ID / Finding title):
      standard blue shading (#0070C0), white font, bold.
    - Risk Level VALUE cell: shaded according to its risk level
      (Critical=#FF0000, High=#F4B083, Medium=#FFFF00, Low=#00FFFF,
      OFI=#92D050).
    - All other cells: no fill (transparent / white background).

-------------------------------------------------------------------------
USAGE
-------------------------------------------------------------------------
    python excel_to_word_findings.py input.xlsx
    python excel_to_word_findings.py input.xlsx output.docx
    python excel_to_word_findings.py input.xlsx --sheet "Follow-up Items"
    python excel_to_word_findings.py input.xlsx --debug

Manual overrides (use if auto-detection of the table picks the wrong
region - e.g. if other bordered cells exist elsewhere on the sheet):
    --top-row N --left-col N --bottom-row N --right-col N
(1-based row/column numbers, as shown in Excel's row/column headers.)
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import openpyxl
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl.worksheet.worksheet import Worksheet

# =============================================================================
# Constants / accepted vocabularies
# =============================================================================

DEFAULT_SHEET_NAME = "Follow-up Items"
DEFAULT_SHEET_INDEX_FALLBACK = 2  # zero-based -> 3rd sheet

RISK_LEVELS = {"critical", "high", "medium", "low", "ofi"}
IMPACT_LEVELS = {"critical", "high", "medium", "low", "very low"}
LIKELIHOOD_LEVELS = {"high", "medium", "low", "very low"}
VERIFICATION_PREFIXES = (
    "completed",
    "partially completed",
    "incomplete",
    "scheduled",
    "accepted",
)

# ---- Styling ----
FONT_NAME = "Times New Roman"
FONT_SIZE = 12

HEADER_ROW_FILL = "0070C0"     # standard blue
HEADER_ROW_FONT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)  # white

RISK_LEVEL_COLORS = {
    "critical": "FF0000",
    "high": "F4B083",
    "medium": "FFFF00",
    "low": "00FFFF",
    "ofi": "92D050",
}

WARNINGS: list[str] = []


def warn(message: str) -> None:
    """Record + immediately print a validation warning."""
    WARNINGS.append(message)
    print(f"WARNING: {message}", file=sys.stderr)


# =============================================================================
# Step 1-2: locate worksheet + table bounding box (via borders)
# =============================================================================


def load_sheet(wb: openpyxl.Workbook, sheet_name: str) -> Worksheet:
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]

    # case-insensitive match
    for name in wb.sheetnames:
        if name.strip().lower() == sheet_name.strip().lower():
            return wb[name]

    if len(wb.sheetnames) > DEFAULT_SHEET_INDEX_FALLBACK:
        fallback = wb.sheetnames[DEFAULT_SHEET_INDEX_FALLBACK]
        warn(
            f'Sheet named "{sheet_name}" not found. '
            f'Falling back to the 3rd sheet: "{fallback}".'
        )
        return wb[fallback]

    raise ValueError(
        f'Sheet named "{sheet_name}" not found, and the workbook has fewer '
        f"than 3 sheets to fall back to. Available sheets: {wb.sheetnames}"
    )


def _cell_has_border(cell) -> bool:
    b = cell.border
    for side in (b.top, b.bottom, b.left, b.right):
        if side is not None and side.style is not None:
            return True
    return False


def find_table_bounds(ws: Worksheet) -> tuple[int, int, int, int]:
    """Return (min_row, min_col, max_row, max_col) of all bordered cells."""
    min_row = min_col = max_row = max_col = None
    for row in ws.iter_rows():
        for cell in row:
            if _cell_has_border(cell):
                r, c = cell.row, cell.column
                min_row = r if min_row is None else min(min_row, r)
                max_row = r if max_row is None else max(max_row, r)
                min_col = c if min_col is None else min(min_col, c)
                max_col = c if max_col is None else max(max_col, c)

    if min_row is None:
        raise ValueError(
            "No bordered cells were found on the sheet - cannot locate the "
            "findings table. Use --top-row/--left-col/--bottom-row/--right-col "
            "to specify the table location manually."
        )
    return min_row, min_col, max_row, max_col


# =============================================================================
# Text cleaning helpers
# =============================================================================

_TAG_RE = re.compile(r"<[^>]+>")
_DATE_RE = re.compile(
    r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})"
)
_MONTHS = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]


def clean_text(value) -> str:
    """Convert a raw cell value to plain text, stripping any HTML markup
    that may have been pasted into the cell (e.g. anchor tags)."""
    if value is None:
        return ""
    text = str(value)
    if "<" in text and ">" in text:
        try:
            soup = BeautifulSoup(text, "html.parser")
            a_tag = soup.find("a")
            if a_tag is not None:
                visible = a_tag.get_text(strip=True)
                if visible:
                    return visible
                href = a_tag.get("href")
                if href:
                    return href
            plain = soup.get_text(separator="\n").strip()
            if plain:
                return plain
        except Exception:
            return _TAG_RE.sub("", text).strip()
    return text.strip()


def split_paragraphs(value) -> list[str]:
    """Clean a (possibly multi-line) cell value and split it into a list of
    non-empty paragraphs, preserving the author's paragraph breaks."""
    text = clean_text(value)
    if not text:
        return []
    parts = [p.strip() for p in text.replace("\r\n", "\n").split("\n")]
    return [p for p in parts if p]


def extract_date(header_text: str) -> Optional[str]:
    """Try to pull a date out of a verification column header such as
    'Verification by Pentastic on 29.08.2026' and format it as 'dd MMM yyyy'.
    Returns None if no recognizable date pattern is found."""
    m = _DATE_RE.search(header_text)
    if not m:
        return None
    d, mth, y = m.groups()
    try:
        d, mth = int(d), int(mth)
        y = int(y)
        if y < 100:
            y += 2000
        if not (1 <= mth <= 12 and 1 <= d <= 31):
            return None
        return f"{d:02d} {_MONTHS[mth - 1]} {y}"
    except (ValueError, IndexError):
        return None


def normalize(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


# =============================================================================
# Step 3: header mapping
# =============================================================================


@dataclass
class HeaderMap:
    id_col: int
    finding_col: Optional[int] = None
    affected_col: Optional[int] = None
    risk_level_col: Optional[int] = None
    impact_col: Optional[int] = None
    likelihood_col: Optional[int] = None
    risk_description_col: Optional[int] = None
    recommended_safeguards_col: Optional[int] = None
    vendor_response_col: Optional[int] = None
    # list of (column_index, raw_header_text), left-to-right
    verification_cols: list[tuple[int, str]] = field(default_factory=list)


def map_headers(ws: Worksheet, header_row: int, min_col: int, max_col: int) -> HeaderMap:
    hmap = HeaderMap(id_col=min_col)
    assigned: set[int] = {min_col}

    for col in range(min_col + 1, max_col + 1):
        raw = ws.cell(row=header_row, column=col).value
        text = clean_text(raw)
        norm = normalize(text)
        if not norm:
            continue

        if norm.startswith("verification"):
            hmap.verification_cols.append((col, text))
            assigned.add(col)
        elif "finding" in norm:
            hmap.finding_col = col
            assigned.add(col)
        elif "affected" in norm:
            hmap.affected_col = col
            assigned.add(col)
        elif "risk level" in norm or norm == "risk":
            hmap.risk_level_col = col
            assigned.add(col)
        elif "risk description" in norm or ("description" in norm and "risk" in norm):
            hmap.risk_description_col = col
            assigned.add(col)
        elif "likelihood" in norm:
            hmap.likelihood_col = col
            assigned.add(col)
        elif norm == "impact" or ("impact" in norm and "risk" not in norm):
            hmap.impact_col = col
            assigned.add(col)
        elif "recommend" in norm or "safeguard" in norm:
            hmap.recommended_safeguards_col = col
            assigned.add(col)
        # else: leave unassigned for now (candidate for vendor response)

    # Vendor Response: header text/name is inconsistent across projects.
    # It's positionally the column immediately before the first
    # Verification* column (and is not consumed by the final Word output,
    # but we identify it so it isn't mistaken for anything else).
    if hmap.verification_cols:
        first_verif_col = hmap.verification_cols[0][0]
        candidate = first_verif_col - 1
        if candidate >= min_col and candidate not in assigned:
            hmap.vendor_response_col = candidate
            assigned.add(candidate)
    else:
        warn(
            "No 'Verification...' column was detected in the header row. "
            "Rectification status will be left blank for all findings."
        )

    if hmap.finding_col is None:
        warn('Could not find a "Finding" column in the header row.')
    if hmap.risk_description_col is None:
        warn('Could not find a "Risk Description" column in the header row.')
    if hmap.risk_level_col is None:
        warn('Could not find a "Risk Level" column in the header row.')
    if hmap.recommended_safeguards_col is None:
        warn('Could not find a "Recommended Safeguards" column in the header row.')

    return hmap


# =============================================================================
# Section header row detection (fully merged rows, e.g. "Penetration Testing")
# =============================================================================


def section_title_for_row(ws: Worksheet, row: int, min_col: int, max_col: int) -> Optional[str]:
    width = max_col - min_col + 1
    for mc in ws.merged_cells.ranges:
        if mc.min_row == row and mc.max_row == row:
            covered = mc.max_col - mc.min_col + 1
            # Treat as a section header if the merge spans (almost) the
            # whole width of the table.
            if mc.min_col <= min_col + 1 and covered >= width - 1:
                anchor = ws.cell(row=mc.min_row, column=mc.min_col).value
                return clean_text(anchor)
    return None


# =============================================================================
# Step 3-4: extract + validate findings
# =============================================================================


@dataclass
class Finding:
    row: int
    finding_id: str
    finding_title: str
    affected: str
    risk_level: str
    impact: str
    likelihood: str
    risk_description: list[str]
    recommended_safeguards: list[str]
    verification_status: str
    verification_date_label: str


def get_val(ws: Worksheet, row: int, col: Optional[int]) -> str:
    if col is None:
        return ""
    return clean_text(ws.cell(row=row, column=col).value)


def validate_risk_level(value: str, finding_id: str, row: int) -> None:
    if value and normalize(value) not in RISK_LEVELS:
        warn(
            f'Row {row} (Finding "{finding_id}"): unexpected Risk Level '
            f'"{value}". Expected one of {sorted(RISK_LEVELS)}.'
        )


def validate_impact(value: str, finding_id: str, row: int) -> None:
    if value and normalize(value) not in IMPACT_LEVELS:
        warn(
            f'Row {row} (Finding "{finding_id}"): unexpected Impact '
            f'"{value}". Expected one of {sorted(IMPACT_LEVELS)}.'
        )


def validate_likelihood(value: str, finding_id: str, row: int) -> None:
    if value and normalize(value) not in LIKELIHOOD_LEVELS:
        warn(
            f'Row {row} (Finding "{finding_id}"): unexpected Likelihood '
            f'"{value}". Expected one of {sorted(LIKELIHOOD_LEVELS)}.'
        )


def validate_verification(value: str, finding_id: str, row: int) -> None:
    if value and not normalize(value).startswith(VERIFICATION_PREFIXES):
        warn(
            f'Row {row} (Finding "{finding_id}"): unexpected verification '
            f'status "{value}". Expected it to start with one of '
            f"{VERIFICATION_PREFIXES}."
        )


def extract_findings(
    ws: Worksheet, hmap: HeaderMap, header_row: int, min_col: int, max_col: int, max_row: int
) -> list[tuple[Optional[str], list[Finding]]]:
    """Returns a list of (section_title_or_None, [Finding, ...]) groups, in
    top-to-bottom order, matching the layout of the source sheet."""

    groups: list[tuple[Optional[str], list[Finding]]] = []
    current_section: Optional[str] = None
    current_findings: list[Finding] = []

    for row in range(header_row + 1, max_row + 1):
        section = section_title_for_row(ws, row, min_col, max_col)
        if section is not None:
            if current_findings or current_section is not None:
                groups.append((current_section, current_findings))
            current_section = section
            current_findings = []
            continue

        finding_id = get_val(ws, row, hmap.id_col)
        finding_title = get_val(ws, row, hmap.finding_col)

        # Skip fully blank rows
        row_values = [
            get_val(ws, row, c)
            for c in (
                hmap.id_col,
                hmap.finding_col,
                hmap.affected_col,
                hmap.risk_level_col,
                hmap.risk_description_col,
            )
        ]
        if not any(row_values):
            continue

        risk_level = get_val(ws, row, hmap.risk_level_col)
        impact = get_val(ws, row, hmap.impact_col)
        likelihood = get_val(ws, row, hmap.likelihood_col)
        affected = get_val(ws, row, hmap.affected_col)
        risk_description = split_paragraphs(ws.cell(row=row, column=hmap.risk_description_col).value) if hmap.risk_description_col else []
        recommended_safeguards = split_paragraphs(ws.cell(row=row, column=hmap.recommended_safeguards_col).value) if hmap.recommended_safeguards_col else []

        # Verification: take the LAST non-empty column, left-to-right.
        verification_status = ""
        verification_header = ""
        for col, header_text in hmap.verification_cols:
            val = get_val(ws, row, col)
            if val:
                verification_status = val
                verification_header = header_text

        date_str = extract_date(verification_header) if verification_header else None
        if date_str:
            verification_date_label = f"Rectification status as of {date_str}"
        elif verification_header:
            verification_date_label = f"Rectification status ({verification_header})"
        else:
            verification_date_label = "Rectification status as of dd MMM yyyy"

        validate_risk_level(risk_level, finding_id, row)
        validate_impact(impact, finding_id, row)
        validate_likelihood(likelihood, finding_id, row)
        validate_verification(verification_status, finding_id, row)

        current_findings.append(
            Finding(
                row=row,
                finding_id=finding_id,
                finding_title=finding_title,
                affected=affected,
                risk_level=risk_level,
                impact=impact,
                likelihood=likelihood,
                risk_description=risk_description,
                recommended_safeguards=recommended_safeguards,
                verification_status=verification_status,
                verification_date_label=verification_date_label,
            )
        )

    if current_findings or current_section is not None:
        groups.append((current_section, current_findings))

    return groups


# =============================================================================
# Step 5: build the Word document
# =============================================================================

LABEL_COL_WIDTH = Cm(4.2)
VALUE_COL_WIDTH = Cm(12.8)


def _set_cell_fill(cell, hex_color: Optional[str]) -> None:
    """Set the cell background fill. Pass hex_color=None for 'no fill'
    (transparent / default white background)."""
    tcPr = cell._tc.get_or_add_tcPr()
    # Remove any existing shading element first
    for existing in tcPr.findall(qn("w:shd")):
        tcPr.remove(existing)

    if hex_color is None:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "auto")
        tcPr.append(shd)
    else:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tcPr.append(shd)


def _set_cell_text(
    cell,
    paragraphs: list[str],
    bold: bool = False,
    font_color: Optional[RGBColor] = None,
) -> None:
    """Write one or more paragraphs of text into a table cell, applying the
    document-wide font (Times New Roman, 12pt)."""
    if not paragraphs:
        paragraphs = [""]
    cell.text = ""  # clear default empty paragraph's placeholder run
    first = True
    for text in paragraphs:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        run = p.add_run(text)
        run.font.name = FONT_NAME
        run.font.size = Pt(FONT_SIZE)
        run.font.bold = bold
        if font_color is not None:
            run.font.color.rgb = font_color
        p.paragraph_format.space_after = Pt(0)


def add_finding_table(document: Document, finding: Finding) -> None:
    table = document.add_table(rows=10, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    for row in table.rows:
        row.cells[0].width = LABEL_COL_WIDTH
        row.cells[1].width = VALUE_COL_WIDTH

    rows_content = [
        (finding.finding_id, [finding.finding_title]),
        ("Risk Description", finding.risk_description),
        ("Risk Level", [finding.risk_level]),
        ("Impact/Likelihood", [f"{finding.impact}/{finding.likelihood}" if (finding.impact or finding.likelihood) else ""]),
        ("OWASP Top 10", [""]),
        ("Affected Asset", [finding.affected]),
        ("Evidence for the finding", [""]),
        ("Recommended Safeguards", finding.recommended_safeguards),
        ("Evidence for the remedial actions", [""]),
        (finding.verification_date_label, [finding.verification_status]),
    ]

    RISK_LEVEL_ROW_INDEX = 2

    for i, (label, value_paragraphs) in enumerate(rows_content):
        row = table.rows[i]
        label_cell, value_cell = row.cells[0], row.cells[1]

        if i == 0:
            # First row: standard blue shading, white bold font (both cells)
            _set_cell_text(label_cell, [label], bold=True, font_color=HEADER_ROW_FONT_COLOR)
            _set_cell_text(value_cell, value_paragraphs, bold=True, font_color=HEADER_ROW_FONT_COLOR)
            _set_cell_fill(label_cell, HEADER_ROW_FILL)
            _set_cell_fill(value_cell, HEADER_ROW_FILL)
        else:
            _set_cell_text(label_cell, [label], bold=False)
            _set_cell_text(value_cell, value_paragraphs, bold=False)
            _set_cell_fill(label_cell, None)  # no fill

            if i == RISK_LEVEL_ROW_INDEX:
                risk_hex = RISK_LEVEL_COLORS.get(normalize(finding.risk_level))
                _set_cell_fill(value_cell, risk_hex)  # None if unrecognized -> no fill
            else:
                _set_cell_fill(value_cell, None)  # no fill

        label_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        value_cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP


def build_document(groups: list[tuple[Optional[str], list[Finding]]], title: str) -> Document:
    document = Document()

    # python-docx's default template omits the required w:percent attribute
    # on <w:zoom>; add it so the resulting file passes strict OOXML validation.
    zoom = document.settings.element.find(qn("w:zoom"))
    if zoom is not None:
        zoom.set(qn("w:percent"), "100")

    style = document.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = Pt(FONT_SIZE)

    # Apply the same font/size to heading styles too, so ALL text in the
    # document (including section headings) uses Times New Roman 12pt.
    for heading_style_id in ("Heading 1", "Heading 2"):
        if heading_style_id in document.styles:
            hstyle = document.styles[heading_style_id]
            hstyle.font.name = FONT_NAME
            hstyle.font.size = Pt(FONT_SIZE)

    document.add_heading(title, level=1)

    total = 0
    for section_title, findings in groups:
        if not findings:
            continue
        if section_title:
            document.add_heading(section_title, level=2)
        for finding in findings:
            add_finding_table(document, finding)
            total += 1
            # Separate each table with 2 empty lines
            document.add_paragraph("")
            document.add_paragraph("")

    if total == 0:
        warn("No findings were extracted - the output document will be empty of tables.")

    return document


# =============================================================================
# Main
# =============================================================================


def convert(
    input_path: Path,
    output_path: Path,
    sheet_name: str = DEFAULT_SHEET_NAME,
    top_row: Optional[int] = None,
    left_col: Optional[int] = None,
    bottom_row: Optional[int] = None,
    right_col: Optional[int] = None,
    debug: bool = False,
) -> None:
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = load_sheet(wb, sheet_name)

    if None not in (top_row, left_col, bottom_row, right_col):
        min_row, min_col, max_row, max_col = top_row, left_col, bottom_row, right_col
    else:
        min_row, min_col, max_row, max_col = find_table_bounds(ws)

    if debug:
        print(
            f"[debug] Sheet: {ws.title!r} | Table bounds: "
            f"row {min_row}-{max_row}, col {min_col}-{max_col}",
            file=sys.stderr,
        )

    hmap = map_headers(ws, header_row=min_row, min_col=min_col, max_col=max_col)

    if debug:
        print(f"[debug] Header map: {hmap}", file=sys.stderr)

    groups = extract_findings(ws, hmap, header_row=min_row, min_col=min_col, max_col=max_col, max_row=max_row)

    if debug:
        n = sum(len(f) for _, f in groups)
        print(f"[debug] Extracted {n} findings across {len(groups)} section group(s).", file=sys.stderr)

    document = build_document(groups, title="Follow-up Findings")
    document.save(output_path)

    print(f"Saved: {output_path}")
    if WARNINGS:
        print(f"\n{len(WARNINGS)} warning(s) were raised during conversion:", file=sys.stderr)
        for w in WARNINGS:
            print(f"  - {w}", file=sys.stderr)
    else:
        print("No validation warnings.", file=sys.stderr)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert a cybersecurity follow-up plan Excel workbook into a Word report of findings."
    )
    parser.add_argument("input", type=Path, help="Path to the source .xlsx workbook")
    parser.add_argument("output", type=Path, nargs="?", default=None, help="Path to the output .docx file (default: <input>.docx)")
    parser.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help='Worksheet name to read (default: "Follow-up Items")')
    parser.add_argument("--top-row", type=int, default=None, help="Manually override: 1-based header row")
    parser.add_argument("--left-col", type=int, default=None, help="Manually override: 1-based leftmost column")
    parser.add_argument("--bottom-row", type=int, default=None, help="Manually override: 1-based last data row")
    parser.add_argument("--right-col", type=int, default=None, help="Manually override: 1-based rightmost column")
    parser.add_argument("--debug", action="store_true", help="Print diagnostic information while converting")
    args = parser.parse_args()

    output = args.output or args.input.with_suffix(".docx")

    convert(
        input_path=args.input,
        output_path=output,
        sheet_name=args.sheet,
        top_row=args.top_row,
        left_col=args.left_col,
        bottom_row=args.bottom_row,
        right_col=args.right_col,
        debug=args.debug,
    )


if __name__ == "__main__":
    main()
