"""Command-line entry point: argument parsing only - all actual work is
delegated to convert().
"""

import argparse
from pathlib import Path

from .constants import DEFAULT_OUTPUT_FORMAT, DEFAULT_SECTION_NUMBER, OUTPUT_FORMATS
from .convert import convert


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert a cybersecurity follow-up plan Excel workbook into a Word report of findings."
    )
    parser.add_argument("input", type=Path, help="Path to the source .xlsx workbook")
    parser.add_argument("output", type=Path, nargs="?", default=None, help="Path to the output .docx file (default: <input>.docx)")
    parser.add_argument(
        "--sheet",
        default=None,
        help=(
            'Worksheet name to read. If omitted, the sheet is auto-detected '
            '(case-insensitive): for SRA formats, by trying "SRA Follow-up" '
            'first, then "Follow-up Items", before falling back to the 3rd '
            'sheet; for SA formats ("sa-detail"/"sa-brief"), by trying '
            '"SA Follow-up", before falling back to the 4th sheet.'
        ),
    )
    parser.add_argument("--top-row", type=int, default=None, help="Manually override: 1-based header row")
    parser.add_argument("--left-col", type=int, default=None, help="Manually override: 1-based leftmost column")
    parser.add_argument("--bottom-row", type=int, default=None, help="Manually override: 1-based last data row")
    parser.add_argument("--right-col", type=int, default=None, help="Manually override: 1-based rightmost column")
    parser.add_argument(
        "--format",
        dest="output_format",
        choices=OUTPUT_FORMATS,
        default=DEFAULT_OUTPUT_FORMAT,
        help=(
            'Output format (default: "portrait-detail"). '
            '"portrait-detail" = A4 portrait, one detailed table per finding (SRA). '
            '"landscape-detail" = A4 landscape, one summary table per section (SRA). '
            '"veri-summary-by-section" = A4 portrait, one combined verification-status-count table, split by section (SRA). '
            '"veri-summary-executive" = A4 portrait, same table but combining ALL sections into a single count block (SRA). '
            '"sa-detail" = A4 landscape, one flat summary table of all Security Audit items. '
            '"sa-brief" = A4 portrait, same Security Audit items with fewer columns.'
        ),
    )
    parser.add_argument(
        "--section-number",
        default=DEFAULT_SECTION_NUMBER,
        help=(
            'Base report section number used to auto-number section headings '
            'in "landscape-detail" (default: "9", producing "9.1", "9.2", ...). '
            'Ignored for all other formats.'
        ),
    )
    parser.add_argument("--debug", action="store_true", help="Print diagnostic information while converting")
    args = parser.parse_args()

    output_explicit = args.output is not None
    output = args.output or args.input.with_suffix(".docx")

    convert(
        input_path=args.input,
        output_path=output,
        sheet_name=args.sheet,
        top_row=args.top_row,
        left_col=args.left_col,
        bottom_row=args.bottom_row,
        right_col=args.right_col,
        output_format=args.output_format,
        output_explicit=output_explicit,
        section_number=args.section_number,
        debug=args.debug,
    )


if __name__ == "__main__":
    main()
