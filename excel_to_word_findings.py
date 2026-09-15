#!/usr/bin/env python3
"""Backward-compatible entry point.

Usage is UNCHANGED from before this refactor:
    python excel_to_word_findings.py input.xlsx [output.docx] [options]

All implementation now lives in the `efw` package (in this same directory) -
see `efw/__init__.py` (or `efw/cli.py`) for the full tool documentation.
This file is intentionally a thin shim so existing invocations/scripts that
call `python excel_to_word_findings.py ...` keep working exactly as before.
"""
from efw.cli import main

if __name__ == "__main__":
    main()
