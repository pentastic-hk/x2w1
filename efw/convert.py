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
    DEFAULT_SA_SHEET_NAME_CANDIDATES,
    DEFAULT_SECTION_NUMBER,
    FORMAT_LANDSCAPE_DETAIL,
    FORMAT_SA_BRIEF,
    FORMAT_SA_DETAIL,
    FORMAT_SRA_BRIEF,
    FORMAT_VERI_SUMMARY_BY_SECTION,
    FORMAT_VERI_SUMMARY_EXECUTIVE,
    SA_SHEET_INDEX_FALLBACK,
)
from .diagnostics import WARNINGS, warn
from .excel_source import find_table_bounds, load_sheet, map_headers, map_sa_headers
from .extraction import extract_findings, extract_sa_findings
from .formats.landscape_detail import build_landscape_document
from .formats.portrait_detail import build_document
from .formats.sa import build_sa_brief_document, build_sa_detail_document
from .formats.sra_brief import build_sra_brief_document
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


# SA (Security Audit) STANDALONE output formats - these use a DIFFERENT
# sheet (default "SA Follow-up", falling back to the 4th sheet) and a
# DIFFERENT table structure/data pipeline than the SRA formats.
_SA_FORMATS = {FORMAT_SA_DETAIL, FORMAT_SA_BRIEF}

# The two "veri-summary-*" formats now read BOTH the SRA sheet AND (if
# present) the SA sheet, combining both halves of the SRAA report into a
# single output (see efw.formats.veri_summary for exactly how they're
# combined/separated in each format).
_VERI_SUMMARY_FORMATS = {FORMAT_VERI_SUMMARY_BY_SECTION, FORMAT_VERI_SUMMARY_EXECUTIVE}


def _load_sa_findings(wb, debug: bool):
    """Best-effort load of the SA ("SA Follow-up") table, for use by the
    "veri-summary-*" formats. Returns None (with a warning, NOT a hard
    failure) if the SA sheet/table cannot be located - e.g. because the
    workbook only contains the SRA half of the report - so those formats
    can gracefully fall back to SRA-only output, preserving their
    pre-existing behavior for workbooks that don't have an SA sheet."""
    try:
        sa_ws = load_sheet(
            wb, None,
            candidates=DEFAULT_SA_SHEET_NAME_CANDIDATES,
            fallback_index=SA_SHEET_INDEX_FALLBACK,
        )
        sa_min_row, sa_min_col, sa_max_row, sa_max_col = find_table_bounds(sa_ws)
        sa_hmap = map_sa_headers(sa_ws, header_row=sa_min_row, min_col=sa_min_col, max_col=sa_max_col)
        sa_findings = extract_sa_findings(
            sa_ws, sa_hmap, header_row=sa_min_row, min_col=sa_min_col, max_col=sa_max_col, max_row=sa_max_row
        )
        if debug:
            print(
                f"[debug] SA Sheet: {sa_ws.title!r} | Extracted {len(sa_findings)} SA item(s).",
                file=sys.stderr,
            )
        return sa_findings
    except Exception as exc:
        warn(
            "Could not locate/extract the SA ('SA Follow-up') table - the "
            f"verification summary will only include SRA data. ({exc})"
        )
        return None


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

    if output_format in _SA_FORMATS:
        # ---- SA (Security Audit) STANDALONE pipeline ----
        # Locate "SA Follow-up" (case-insensitive), falling back to the 4th
        # sheet - reusing the SAME load_sheet()/find_table_bounds() logic
        # as the SRA pipeline (bordered-cell bounding box detection is
        # generic and works identically for either table).
        ws = load_sheet(
            wb, sheet_name,
            candidates=DEFAULT_SA_SHEET_NAME_CANDIDATES,
            fallback_index=SA_SHEET_INDEX_FALLBACK,
        )

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

        hmap = map_sa_headers(ws, header_row=min_row, min_col=min_col, max_col=max_col)

        if debug:
            print(f"[debug] SA Header map: {hmap}", file=sys.stderr)

        findings = extract_sa_findings(ws, hmap, header_row=min_row, min_col=min_col, max_col=max_col, max_row=max_row)

        if debug:
            print(f"[debug] Extracted {len(findings)} SA item(s).", file=sys.stderr)

        if output_format == FORMAT_SA_DETAIL:
            document = build_sa_detail_document(findings, hmap, title="Security Audit Findings")
        else:
            document = build_sa_brief_document(findings, title="Security Audit Findings (Brief)")

    else:
        # ---- SRA (Security Risk Assessment) pipeline ----
        # --sheet / --top-row / --left-col / --bottom-row / --right-col all
        # continue to apply ONLY to the SRA table, as before, regardless of
        # output format (including the "veri-summary-*" formats below,
        # where the SA table's bounds are always auto-detected).
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
        elif output_format == FORMAT_SRA_BRIEF:
            document = build_sra_brief_document(groups, title="Follow-up Findings (Brief)")
        elif output_format in _VERI_SUMMARY_FORMATS:
            # ---- Combined SRA + SA pipeline for the verification summary
            # formats: additionally load the SA half of the report (best
            # effort - gracefully degrades to SRA-only if not found). ----
            sa_findings = _load_sa_findings(wb, debug)

            if output_format == FORMAT_VERI_SUMMARY_BY_SECTION:
                document = build_veri_summary_document(
                    groups, title="Verification Summary by Section", sa_findings=sa_findings
                )
            else:
                document = build_veri_summary_executive_document(
                    groups, title="Verification Summary (Executive)", sa_findings=sa_findings
                )
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
