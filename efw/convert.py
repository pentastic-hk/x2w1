"""Orchestrates the full pipeline for a single conversion run: load the
workbook, locate the sheet/table, map headers, extract findings, build
the requested output-format Document, and save it (applying the
format-suffix filename convention when appropriate).
"""

import sys
from pathlib import Path
from typing import Optional

import openpyxl

from .constants import (
    DEFAULT_OUTPUT_FORMAT,
    DEFAULT_SECTION_NUMBER,
    FORMAT_LANDSCAPE_DETAIL,
    FORMAT_VERI_SUMMARY_BY_SECTION,
    FORMAT_VERI_SUMMARY_EXECUTIVE,
)
from .diagnostics import WARNINGS, warn
from .excel_source import find_table_bounds, load_sheet, map_headers
from .extraction import extract_findings
from .formats.landscape_detail import build_landscape_document
from .formats.portrait_detail import build_document
from .formats.veri_summary import (
    build_veri_summary_document,
    build_veri_summary_executive_document,
)

def append_format_suffix(output_path: Path, output_format: str) -> Path:
    """Append the output format id as a suffix to the output filename, right
    before the file extension - e.g. "QC FUP v1.2.docx" with
    output_format="landscape-detail" becomes "QC FUP v1.2-landscape-detail.docx".

    Only applied when the output filename was AUTO-DERIVED from the input
    filename (i.e. the user did not explicitly provide an output filename).
    If the user explicitly names the output file, that name is used exactly
    as given, with no suffix appended - see the `output_explicit` parameter
    of convert()."""
    return output_path.with_name(f"{output_path.stem}-{output_format}{output_path.suffix}")


# =============================================================================
# Main
# =============================================================================


def convert(
    input_path: Path,
    output_path: Path,
    sheet_name: Optional[str] = None,
    top_row: Optional[int] = None,
    left_col: Optional[int] = None,
    bottom_row: Optional[int] = None,
    right_col: Optional[int] = None,
    output_format: str = DEFAULT_OUTPUT_FORMAT,
    section_number: str = DEFAULT_SECTION_NUMBER,
    debug: bool = False,
    output_explicit: bool = False,
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

    if output_format == FORMAT_LANDSCAPE_DETAIL:
        document = build_landscape_document(
            groups, hmap, title="Follow-up Findings", section_base_number=section_number
        )
    elif output_format == FORMAT_VERI_SUMMARY_BY_SECTION:
        document = build_veri_summary_document(groups, title="Verification Summary by Section")
    elif output_format == FORMAT_VERI_SUMMARY_EXECUTIVE:
        document = build_veri_summary_executive_document(groups, title="Verification Summary (Executive)")
    else:
        document = build_document(groups, title="Follow-up Findings")

    # Only auto-append the "-<format>" suffix when the output filename was
    # NOT explicitly provided by the user (i.e. it was auto-derived from the
    # input filename). If the user explicitly named the output file, honor
    # that name exactly as given.
    if not output_explicit:
        output_path = append_format_suffix(output_path, output_format)
    document.save(output_path)

    print(f"Saved: {output_path}")
    if WARNINGS:
        print(f"\n{len(WARNINGS)} warning(s) were raised during conversion:", file=sys.stderr)
        for w in WARNINGS:
            print(f"  - {w}", file=sys.stderr)
    else:
        print("No validation warnings.", file=sys.stderr)

