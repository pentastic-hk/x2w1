"""Worksheet/table discovery: locating the correct sheet, finding the
bordered findings table's bounding box, mapping header columns to their
semantic roles, and detecting section-header rows.
"""

from typing import Optional

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from .constants import DEFAULT_SHEET_NAME_CANDIDATES, DEFAULT_SHEET_INDEX_FALLBACK
from .diagnostics import warn
from .models import HeaderMap
from .text_utils import clean_text, normalize

# =============================================================================
# Step 1-2: locate worksheet + table bounding box (via borders)
# =============================================================================


def load_sheet(wb: openpyxl.Workbook, sheet_name: Optional[str] = None) -> Worksheet:
    """Locate the worksheet to read from.

    - If `sheet_name` is given (e.g. via --sheet), it is the ONLY candidate
      tried (case-insensitive exact match), before falling back to the 3rd
      sheet.
    - If `sheet_name` is None (the default - no explicit --sheet override),
      candidates from DEFAULT_SHEET_NAME_CANDIDATES are tried IN ORDER
      (case-insensitive): "SRA Follow-up" first, then "Follow-up Items",
      before falling back to the 3rd sheet.
    """
    candidates = [sheet_name] if sheet_name else DEFAULT_SHEET_NAME_CANDIDATES

    for candidate in candidates:
        if candidate in wb.sheetnames:
            return wb[candidate]
        # case-insensitive match
        for name in wb.sheetnames:
            if name.strip().lower() == candidate.strip().lower():
                return wb[name]

    if len(wb.sheetnames) > DEFAULT_SHEET_INDEX_FALLBACK:
        fallback = wb.sheetnames[DEFAULT_SHEET_INDEX_FALLBACK]
        tried = " / ".join(f'"{c}"' for c in candidates)
        warn(
            f"None of the candidate sheet name(s) {tried} were found. "
            f'Falling back to the 3rd sheet: "{fallback}".'
        )
        return wb[fallback]

    tried = " / ".join(f'"{c}"' for c in candidates)
    raise ValueError(
        f"None of the candidate sheet name(s) {tried} were found, and the "
        f"workbook has fewer than 3 sheets to fall back to. "
        f"Available sheets: {wb.sheetnames}"
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
        elif "owasp" in norm:
            # Optional column. There are no predefined/canonical values for
            # this field, so its contents are never validated/normalized -
            # they are simply passed through verbatim (after the usual HTML
            # cleanup) to the output.
            hmap.owasp_top_10_col = col
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
# Section header row detection (rows with at least the first 6 columns
# merged, e.g. "Penetration Testing")
# =============================================================================

SECTION_HEADER_MIN_MERGED_COLS = 6


def section_title_for_row(ws: Worksheet, row: int, min_col: int, max_col: int) -> Optional[str]:
    width = max_col - min_col + 1
    # A table narrower than the usual minimum can never satisfy the "first 6
    # columns merged" rule, so fall back to requiring the full width instead.
    required_merged_cols = min(SECTION_HEADER_MIN_MERGED_COLS, width)

    for mc in ws.merged_cells.ranges:
        if mc.min_row == row and mc.max_row == row:
            covered = mc.max_col - mc.min_col + 1
            # Treat as a section header if the merge starts at (or right at)
            # the leftmost column of the table AND spans at least the first
            # 6 columns (rather than requiring the entire row to be merged).
            # A row merged across the FULL table width also satisfies this,
            # since it necessarily covers at least the first 6 columns too.
            if mc.min_col <= min_col + 1 and covered >= required_merged_cols:
                anchor = ws.cell(row=mc.min_row, column=mc.min_col).value
                return clean_text(anchor)
    return None
