"""Shared constants: output-format enum, sheet-name auto-detection
candidates, validation vocabularies, and styling constants (fonts,
colors) used across every output format.
"""

from docx.shared import RGBColor

# =============================================================================
# Constants / accepted vocabularies
# =============================================================================

# Ordered list of candidate sheet names to try (case-insensitive), used when
# the user hasn't explicitly overridden --sheet: "SRA Follow-up" is tried
# FIRST, then "Follow-up Items", before falling back to the 3rd sheet.
DEFAULT_SHEET_NAME_CANDIDATES = ["SRA Follow-up", "Follow-up Items"]
DEFAULT_SHEET_INDEX_FALLBACK = 2  # zero-based -> 3rd sheet

# Candidate sheet name(s) for the SA (Security Audit) half of the report,
# tried case-insensitively; falls back to the 4th sheet (zero-based index 3)
# if not found.
DEFAULT_SA_SHEET_NAME_CANDIDATES = ["SA Follow-up"]
SA_SHEET_INDEX_FALLBACK = 3  # zero-based -> 4th sheet

# Output format enum (string values, NOT a boolean, so more formats can be
# added later without changing the CLI shape).
FORMAT_PORTRAIT_DETAIL = "portrait-detail"
FORMAT_LANDSCAPE_DETAIL = "landscape-detail"
FORMAT_SRA_BRIEF = "sra-brief"
FORMAT_VERI_SUMMARY_BY_SECTION = "veri-summary-by-section"
FORMAT_VERI_SUMMARY_EXECUTIVE = "veri-summary-executive"
FORMAT_SA_DETAIL = "sa-detail"
FORMAT_SA_BRIEF = "sa-brief"
OUTPUT_FORMATS = [
    FORMAT_PORTRAIT_DETAIL,
    FORMAT_LANDSCAPE_DETAIL,
    FORMAT_SRA_BRIEF,
    FORMAT_VERI_SUMMARY_BY_SECTION,
    FORMAT_VERI_SUMMARY_EXECUTIVE,
    FORMAT_SA_DETAIL,
    FORMAT_SA_BRIEF,
]
DEFAULT_OUTPUT_FORMAT = FORMAT_PORTRAIT_DETAIL
DEFAULT_SECTION_NUMBER = "9"

RISK_LEVELS = {"critical", "high", "medium", "low", "ofi"}
IMPACT_LEVELS = {"critical", "high", "medium", "low", "very low"}
LIKELIHOOD_LEVELS = {"high", "medium", "low", "very low"}

# Canonical verification status labels (output values). The source cell text
# only needs to START WITH one of these (case-insensitive) - any trailing
# text (e.g. explanatory notes) is dropped from the Word output.
VERIFICATION_CANONICAL_LABELS = [
    "Completed",
    "Partially Completed",
    "Incomplete",
    "Scheduled",
    "Accepted",
]
# Sort longest-first so "Partially Completed" is checked before "Completed"
# would otherwise never incorrectly match as a prefix of it (they don't
# overlap as prefixes of one another, but this keeps matching unambiguous
# and future-proof if labels are edited).
_VERIFICATION_MATCH_ORDER = sorted(
    VERIFICATION_CANONICAL_LABELS, key=len, reverse=True
)

# ---- Styling ----
FONT_NAME = "Times New Roman"
FONT_SIZE = 12

HEADER_ROW_FILL = "0070C0"     # standard blue
HEADER_ROW_FONT_COLOR = RGBColor(0xFF, 0xFF, 0xFF)  # white
BLACK_AUTO_COLOR = RGBColor(0x00, 0x00, 0x00)  # "black (Automatic)"

RISK_LEVEL_COLORS = {
    "critical": "FF0000",
    "high": "F4B083",
    "medium": "FFFF00",
    "low": "00FFFF",
    "ofi": "92D050",
}

# Display order/labels for risk levels, shared by every format that breaks
# counts down by risk level.
RISK_LEVEL_DISPLAY = {
    "critical": "Critical",
    "high": "High",
    "medium": "Medium",
    "low": "Low",
    "ofi": "OFI",
}

# Standard MS Word theme swatch "Blue, Accent 1, Lighter 80%" (Office theme:
# accent1 = 4472C4, tinted 80% toward white). Used ONLY for the section
# header rows in the "sra-brief" format.
SRA_BRIEF_SECTION_FILL = "D9E2F3"

WARNINGS: list[str] = []
