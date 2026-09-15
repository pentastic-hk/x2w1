"""Text-cleaning and normalization helpers shared by every format:
HTML stripping (clean_text), paragraph splitting, date extraction from
verification-column headers, and verification-status normalization.
"""

import re
from typing import Optional

from bs4 import BeautifulSoup

from .constants import _VERIFICATION_MATCH_ORDER

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
    that may have been pasted into the cell (e.g. anchor tags), while
    PRESERVING any surrounding plain text and the cell's original line
    breaks.

    Handles both:
        - A cell whose ENTIRE value is just an anchor tag, e.g.
          '<a href="https://x">https://x</a>' -> "https://x"
        - A cell with an anchor EMBEDDED within a larger block of text,
          e.g. 'Update the packages.\\n\\nReference: <a href="...">...</a>'
          -> 'Update the packages.\\n\\nReference: https://...' (all
          surrounding text is preserved; only the <a> tag itself is
          replaced by its visible text, or its href if it has none).
    """
    if value is None:
        return ""
    text = str(value)
    if "<" in text and ">" in text:
        try:
            soup = BeautifulSoup(text, "html.parser")
            for a_tag in soup.find_all("a"):
                visible = a_tag.get_text(strip=True)
                href = a_tag.get("href")
                a_tag.replace_with(visible or href or "")
            # No separator: real line breaks already exist as literal "\n"
            # characters WITHIN the original text nodes. Using a separator
            # here would incorrectly insert extra line breaks at every tag
            # boundary (e.g. splitting "Reference: " from the URL that
            # replaced the <a> tag onto two different lines).
            plain = soup.get_text().strip()
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


def _depluralize_word(word: str) -> str:
    """Naive singularization: strip a single trailing 's' from a word, but
    only when doing so is unlikely to change its meaning/matchability -
    i.e. the word is longer than 3 characters and doesn't already end in
    'ss' (e.g. "safeguards" -> "safeguard", "items" -> "item", but "class"
    stays "class", and short words like "is"/"as" are left untouched)."""
    if len(word) > 3 and word.endswith("s") and not word.endswith("ss"):
        return word[:-1]
    return word


def normalize_singular(s) -> str:
    """Like normalize(), but ALSO naively singularizes every word (see
    _depluralize_word). Used for matching column header labels that may
    freely include or omit a trailing plural 's' (e.g. "Items to check" vs
    "Item to check", "Recommended Safeguards" vs "Recommended Safeguard"),
    so keyword matching can be written once in singular form and match
    either variant."""
    norm = normalize(s)
    if not norm:
        return norm
    return " ".join(_depluralize_word(w) for w in norm.split(" "))


def normalize_verification_status(value: str) -> tuple[str, bool]:
    """Match `value` against the canonical verification labels as a
    case-insensitive PREFIX match. Returns (output_value, matched):
        - If a canonical label matches as a prefix, returns
          (canonical_label, True) - trailing text (e.g. explanatory notes)
          is dropped.
        - If no canonical label matches, returns (original_value, False) so
          the original text is preserved in the output (nothing is silently
          lost) and the caller can raise a warning.
        - Empty input returns ("", True) - no warning for blank cells.
    """
    if not value:
        return "", True

    norm_value = normalize(value)
    for label in _VERIFICATION_MATCH_ORDER:
        if norm_value.startswith(label.lower()):
            return label, True

    return value, False
