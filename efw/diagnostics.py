"""Collects and prints non-fatal validation warnings raised while
reading/converting a workbook (e.g. unexpected Risk Level values).
"""

import sys

WARNINGS: list[str] = []


def warn(message: str) -> None:
    """Record + immediately print a validation warning."""
    WARNINGS.append(message)
    print(f"WARNING: {message}", file=sys.stderr)
