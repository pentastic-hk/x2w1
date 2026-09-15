"""Data models shared across the whole pipeline: HeaderMap (resolved
column positions for one worksheet) and Finding (one fully-parsed,
validated row from the source workbook).
"""

from dataclasses import dataclass, field
from typing import Optional


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
    owasp_top_10_col: Optional[int] = None
    vendor_response_col: Optional[int] = None
    # list of (column_index, raw_header_text), left-to-right
    verification_cols: list[tuple[int, str]] = field(default_factory=list)


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
    # Optional, free-text column with no predefined/canonical values, so it
    # is never validated or normalized - just cleaned (HTML stripped) and
    # split into paragraphs like Risk Description / Recommended Safeguards.
    # Empty list if the source workbook has no "OWASP Top 10" column.
    owasp_top_10: list[str]
    verification_status: str
    verification_date_label: str
    # Column index (1-based, openpyxl-style) of the verification column that
    # was actually used (last non-empty, left-to-right) for THIS finding, or
    # None if no verification column had a value. Used by the
    # "landscape-detail" format to derive a single, section-level
    # "Rectification Status as of <date>" column header.
    verification_col_used: Optional[int] = None


@dataclass
class SAHeaderMap:
    """Column mapping for the "SA Follow-up" (Security Audit) table, which
    has a different structure than the SRA "Follow-up Items"/"SRA
    Follow-up" table: #, Items to check, Affected, Findings, Recommended
    Safeguards, Client Response, Planned Completion Date, and one or more
    (possibly non-contiguous) Verification columns."""
    id_col: int
    items_col: Optional[int] = None
    affected_col: Optional[int] = None
    findings_col: Optional[int] = None
    recommended_safeguards_col: Optional[int] = None
    # Recorded for completeness/diagnostics but NOT used in either SA
    # output format (per spec, only # / Items / Affected / Findings /
    # Recommendation / Rectification Status are rendered).
    client_response_col: Optional[int] = None
    planned_completion_date_col: Optional[int] = None
    # list of (column_index, raw_header_text), left-to-right. May be
    # non-contiguous (e.g. interleaved with "Client Response" columns).
    verification_cols: list[tuple[int, str]] = field(default_factory=list)


@dataclass
class SAFinding:
    row: int
    item_id: str
    items_to_check: list[str]
    affected: str
    findings: list[str]
    recommended_safeguards: list[str]
    verification_status: str
    verification_date_label: str
    verification_col_used: Optional[int] = None
