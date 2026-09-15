# Convert Excel to Word Tables

Convert Follow-up Plan to Detailed Findings Tables in Businesses Reporting Template.

Specifications provided by David, vibe-coded with Claude Sonnet.

## Installation

### Installing Python

You should install Python version 3 or above on your machine.

### Cloning the Repo

Then, clone this repo to your machine.
```sh
git clone https://github.com/pentastic-hk/x2w1.git
```

### Setting Up the Virtual Environment

Python virtual environment is used to keep this project isolated.
Run the following command to setup your virtual environment:
```sh
python3 -m venv .venv
```
This creates a folder `.venv` in your project folder.

Then activate your virtual environment with the following command:
```sh
source .venv/bin/active
```

Or the following command if you're using Windows:
```sh
.venv\scripts\Activate
```

### Installing Python Dependencies

Run the following command to install the project dependencies to your virtual environment:
```sh
pip install -r requirements.txt
```

## Usage

```sh
python3 excel_to_word_findings.py path/to/follow-up-plan.xlsx
python3 excel_to_word_findings.py path/to/follow-up-plan.xlsx output-table.docx
python3 excel_to_word_findings.py path/to/follow-up-plan.xlsx --sheet "Follow-up Items"
python3 excel_to_word_findings.py path/to/follow-up-plan.xlsx --debug
```

If you're too lazy to type the full path to your Excel file,
drag your file from your Files Explorer onto the command line.

## Refactoring

See `efw/cli.py` (or run with `--help`) for the full option reference, and
`efw/__init__.py` for the complete tool documentation (data-extraction
rules, validation rules, and per-format styling specs) - this was
previously all one big docstring at the top of the monolithic script; it's
preserved verbatim there.

## Why this layout

The original tool was a single ~2,000-line script. It has been split into
the `efw/` package below so that each file has one clear responsibility and
stays easy to navigate/extend. **No behavior changed** - only the code
organization. `excel_to_word_findings.py` is now a thin shim that just
calls into the package, so any existing invocation of the old script keeps
working exactly as before.

## Layout

```
excel_to_word_findings.py      <- thin backward-compatible CLI entry point
efw/                           <- the actual implementation, as a package
    __init__.py                 <- full tool docstring + re-exports main()
    __main__.py                  <- enables `python -m efw ...`
    cli.py                        <- argparse setup (the CLI surface)
    convert.py                     <- convert(): orchestrates one full run
    constants.py                    <- format enum, styling/color constants,
                                       validation vocabularies, sheet-name
                                       auto-detection candidates
    diagnostics.py                   <- WARNINGS list + warn() helper
    text_utils.py                     <- HTML cleanup, paragraph splitting,
                                       date extraction, status normalization
    models.py                          <- HeaderMap / Finding dataclasses
    excel_source.py                     <- sheet lookup, bordered-table
                                       bounding box, header-column mapping,
                                       section-header-row detection
    extraction.py                        <- row -> Finding extraction +
                                       validation (Risk Level/Impact/
                                       Likelihood) + section grouping
    docx_utils.py                         <- low-level python-docx/OOXML
                                       helpers shared by every format
                                       (cell shading, row height, table
                                       grid widths, base font styling, ...)
    formats/
        __init__.py
        portrait_detail.py                 <- "portrait-detail" format
        landscape_detail.py                 <- "landscape-detail" format
        veri_summary.py                      <- "veri-summary-by-section"
                                       AND "veri-summary-executive" (they
                                       share ~95% of their logic)
```

## Dependencies

See `requirements.txt`:
```
openpyxl
python-docx
beautifulsoup4
```

## Extending

To add a new `--format` value:
1. Add a new `FORMAT_...` constant + append it to `OUTPUT_FORMATS` in
   `efw/constants.py`.
2. Add a new module under `efw/formats/` (or extend an existing one, as
   `veri_summary.py` does) with a `build_..._document(groups, title=...)`
   function that returns a `docx.Document`.
3. Add a matching `elif output_format == FORMAT_...:` branch in
   `efw/convert.py`'s `convert()`.

The shared data pipeline (`excel_source.py` + `extraction.py`, producing a
list of `(section_title, [Finding, ...])` groups) and the shared styling
helpers (`docx_utils.py`) require no changes for most new formats.
