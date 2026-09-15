"""
efw - Excel-to-Word Findings converter (package)
==================================================

This package implements the tool described below. It was refactored from
a single monolithic script into focused modules for maintainability; the
CLI/behavior is 100% unchanged - see the root-level
`excel_to_word_findings.py` script for the unchanged command-line entry
point (`python excel_to_word_findings.py input.xlsx ...`).


Convert a cybersecurity "Follow-up Plan" Excel workbook into a formal Word
report. The SAME data-extraction logic (locating the sheet/table, mapping
columns, validating values, grouping by section) is shared across THREE
output FORMATS, selected via --format:

    - "portrait-detail" (default): A4 portrait. Renders ONE 2-column Word
      table PER FINDING (General Control Review, Vulnerability Scanning,
      Web/API Penetration Testing, Source Code Review, etc), matching the
      detailed per-finding write-up layout used elsewhere in the report.

    - "landscape-detail": A4 landscape. Renders ONE summary Word table PER
      SECTION (e.g. "9.1 General Control Review", "9.2 Vulnerability
      Scanning", ...), with one row per finding and columns for #, Findings,
      Affected, Risk Description, Risk Level, Impact / Likelihood,
      Recommendation, and Rectification Status.

    - "veri-summary-by-section": A4 portrait. Renders a SINGLE table (styled
      as Word's built-in "Grid Table 4 - Accent 6") containing one compact
      Risk-Level x Rectification-Status count block PER SECTION, stacked in
      order - i.e. a "Verification Summary per Section" table.

    - "veri-summary-executive": Identical to "veri-summary-by-section" in
      every respect (styling, coloring, row/column logic) EXCEPT that ALL
      findings from every section are combined into a SINGLE count block
      (section boundaries are ignored), with its title simply
      "Security Risk Assessment" (no " - <Section>" suffix) - i.e. an
      executive-level, workbook-wide "Verification Summary" table.

-------------------------------------------------------------------------
HOW IT LOCATES THE DATA
-------------------------------------------------------------------------
1. Locates the worksheet to read by trying, in order (case-insensitive):
   "SRA Follow-up", then "Follow-up Items", then falling back to the 3rd
   sheet in the workbook (with a warning) if neither name is found. Use
   --sheet to override this auto-detection with an exact sheet name (which
   is then the only candidate tried, still falling back to the 3rd sheet
   if not found).
2. Scans every cell in the worksheet and finds the bounding box of ALL cells
   that have at least one visible border side set. This bounding box is
   assumed to be the findings table (borders are the only reliable signal,
   per the source workbook's convention, since the table doesn't always
   start at A1).
3. The first row of that bounding box is treated as the header row. Column
   purposes are identified by matching header text (case-insensitive,
   whitespace-normalized) against known labels:
       - 1st column (leftmost)      -> Finding ID (e.g. "W1", "MP1")
       - "Finding"                  -> Finding title
       - "Affected*"                -> Affected asset / URL / endpoint
       - "Risk Level"                -> Risk Level
       - "Impact" (optional)         -> Impact
       - "Likelihood" (optional)     -> Likelihood
       - "Risk Description"          -> Risk Description
       - "Recommend*"/"Safeguard*"   -> Recommended Safeguards
       - "OWASP*" (optional)         -> OWASP Top 10 (free text, no
                                         predefined/canonical values, so it
                                         is never validated or normalized -
                                         just cleaned and passed through)
       - "Verification*"             -> one or more verification columns
                                         (rightmost non-empty per row wins)
   Any other column immediately to the left of the first "Verification*"
   column is treated as the (unused-in-output) Vendor Response column.
4. Rows where at least the first 6 columns of the table are merged into one
   cell are treated as SECTION HEADER rows (e.g. "General Control Review",
   "Vulnerability Scanning", "Web Penetration Testing") and are rendered as
   Word Heading 2 paragraphs instead of finding tables. The merge does not
   need to span the entire table width - only the first 6 columns (a merge
   spanning the FULL table width also still counts, since it necessarily
   covers at least the first 6 columns too).
5. Every other non-blank row within the bounding box is treated as a single
   finding and validated + rendered as its own 2-column Word table.

-------------------------------------------------------------------------
VALIDATION
-------------------------------------------------------------------------
The script warns (stderr) - but does NOT fail - on unexpected values for:
    - Risk Level     : Critical / High / Medium / Low / OFI
    - Impact         : Critical / High / Medium / Low / Very Low
    - Likelihood     : High / Medium / Low / Very Low
    - Verification   : the cell value must BEGIN WITH (case-insensitive) one
                        of: Completed / Partially Completed / Incomplete /
                        Scheduled / Accepted. If it does, the output value
                        is NORMALIZED to just that canonical label (the rest
                        of the cell's text, e.g. "Completed. Vendor added a
                        strict allow-list." -> "Completed", is dropped from
                        the Word table). If no match is found, a warning is
                        raised and the ORIGINAL (un-normalized) text is kept
                        in the output so nothing is silently lost.
All warnings are also collected and summarized at the end of the run.

-------------------------------------------------------------------------
WORD TABLE STYLING - "portrait-detail" format
-------------------------------------------------------------------------
    - All text (including section headings): Times New Roman, 12pt, black
      (Automatic) font color.
    - All cell content is top-aligned (vertically) and has no extra spacing
      between wrapped paragraphs within a cell.
    - First row of each finding table (Finding ID / Finding title):
      standard blue shading (#0070C0), white font, bold.
    - First column, all rows EXCEPT the first row: not bold.
    - "OWASP Top 10" row: populated from the optional "OWASP*" source
      column when present (left blank otherwise, as before).
    - Risk Level VALUE cell: shaded according to its risk level
      (Critical=#FF0000, High=#F4B083, Medium=#FFFF00, Low=#00FFFF,
      OFI=#92D050).
    - All other cells: no fill (transparent / white background).
    - Each finding's table is separated from the next by 2 blank lines
      (with no extra paragraph spacing added below them).

-------------------------------------------------------------------------
WORD TABLE STYLING - "landscape-detail" format
-------------------------------------------------------------------------
    - Page: A4, landscape orientation.
    - All text (including section headings and the "The following issues
      are identified:" intro paragraph): Times New Roman, 12pt, black
      (Automatic) font color.
    - Each section becomes its own bold heading, auto-numbered as
      "<section-number>.<n> <Section Title>" (e.g. "9.1 General Control
      Review"), followed by one summary table for that section's findings.
    - Table header row: standard blue shading (#0070C0), white font, bold,
      and "Repeat Header Rows" enabled (repeats on every page the table
      spans).
    - Table columns: #, Findings, Affected, Risk Description, Risk Level,
      Impact / Likelihood, [OWASP Top 10 (optional - only included if the
      source workbook has that column)], Recommendation, Rectification
      Status as of <date> (date is derived per-section from the rightmost
      verification column that has any data in that section, using the
      same date extraction/normalization logic as the portrait format).
      Whether the "OWASP Top 10" column is included is decided ONCE for
      the whole document (not per-section), so every section's table has
      an identical column set.
    - Risk Level value cells are colored using the EXACT SAME color scheme
      as the portrait format (Critical=#FF0000, High=#F4B083,
      Medium=#FFFF00, Low=#00FFFF, OFI=#92D050) - the coloring logic is
      shared/reused, not duplicated.
    - Data rows are NOT bold (only the header row and section headings are
      bold).

-------------------------------------------------------------------------
WORD TABLE STYLING - "veri-summary-by-section" format
-------------------------------------------------------------------------
    - Page: A4, portrait orientation.
    - ONE single Word table for the ENTIRE document (all sections stacked
      into the same table, in order) - NOT one table per section.
    - The table uses Word's built-in "Grid Table 4 - Accent 6" style
      (injected into the document, since python-docx's default template
      does not ship it), which drives:
        - the table's border color (a themed orange/accent6 grid), and
        - the alternating pale/unfilled row shading, computed AUTOMATICALLY
          by Word's own table-style banding engine based on each row's
          absolute position in the table (continues seamlessly across
          section boundaries - verified empirically that intervening rows
          do not reset or disrupt the alternating sequence).
      Risk Level value cells (High/Medium/Low/OFI/Critical) are therefore
      NOT colored with the portrait/landscape formats' risk-level color
      scheme - they simply inherit whichever alternating band color falls
      on their row, exactly like every other data cell.
    - ONLY each section's title row ("Security Risk Assessment - <Section>",
      merged across all columns) gets a solid, HARDCODED fill color
      (#FFC000) and is marked as a repeating header row. It is also given a
      FIXED row height of 0.9cm (VERI_SUMMARY_TITLE_ROW_HEIGHT). The 2
      header-label rows below it ("Risk Level" / "Number of items by
      Rectification Status" / individual status column names) are
      deliberately left un-shaded, at their natural (auto) height, and
      non-repeating, so they - like the risk-level rows and the Total row -
      simply follow the table style's normal alternating band shading
      (pale orange / no-fill) based on row position.
    - The "Risk Level" header label is UNMERGED across its two header rows:
      it appears only in the upper cell; the lower cell beneath it is left
      empty.
    - ALL text in the table (including the title row) is black (Automatic)
      - no white text anywhere, even on the shaded title row.
    - Bold is applied ONLY to: each section's title cell, the "Risk Level"
      header cell, and every cell in each section's "Total" row. All other
      cells (including the Risk Level VALUE cells like "High"/"Medium" and
      all other header cells) are NOT bold.
    - Center-aligned: every numeric count cell (Total/Completed/Partially
      Completed/Incomplete/Scheduled/Accepted values, in both the
      risk-level rows and the Total row) AND the column-label header cells
      for those same columns (i.e. the second header row's cells: "Total",
      "Completed", "Partially Completed", "Incomplete", "Scheduled",
      "Accepted"). Everything else (title, "Risk Level" label, risk level
      values, "Number of items by Rectification Status") stays left-aligned.
    - The "Critical" risk-level row is shown (above "High") only if used
      anywhere in the workbook; the "Partially Completed" status column is
      shown (between "Completed" and "Incomplete") only if used anywhere in
      the workbook. This is decided ONCE, globally, so every section's
      block has an identical column/row shape.

-------------------------------------------------------------------------
WORD TABLE STYLING - "veri-summary-executive" format
-------------------------------------------------------------------------
    - Identical to "veri-summary-by-section" in EVERY styling/formatting
      respect (page setup, table style, borders/banding, bold placement,
      alignment, Critical row / Partially Completed column inclusion
      logic) - see above.
    - The ONLY difference: instead of one block per section, ALL findings
      across ALL sections are combined into a SINGLE Risk-Level x
      Rectification-Status count block (section boundaries such as
      "General Control Review", "Vulnerability Scanning", "Web Penetration
      Testing", "API Penetration Testing", "Agent Penetration Testing",
      etc. are NOT used to split the counts).
    - That single block's title row reads plainly "Security Risk
      Assessment" - WITHOUT the hyphen and section name suffix used by
      "veri-summary-by-section" (e.g. NOT "Security Risk Assessment -
      General Control Review").

-------------------------------------------------------------------------
USAGE
-------------------------------------------------------------------------
    python excel_to_word_findings.py input.xlsx
    python excel_to_word_findings.py input.xlsx output.docx
    python excel_to_word_findings.py input.xlsx --sheet "SRA Follow-up"
    python excel_to_word_findings.py input.xlsx --format landscape-detail
    python excel_to_word_findings.py input.xlsx --format landscape-detail --section-number 9
    python excel_to_word_findings.py input.xlsx --format veri-summary-by-section
    python excel_to_word_findings.py input.xlsx --format veri-summary-executive
    python excel_to_word_findings.py input.xlsx --debug

--format accepts a string enum (not a boolean), so more formats can be
added later without breaking the CLI:
    - "portrait-detail"          (default) - one detailed table per finding, A4 portrait.
    - "landscape-detail"         - one summary table per section, A4 landscape.
    - "veri-summary-by-section"  - one combined verification-status-count table, split by section, A4 portrait.
    - "veri-summary-executive"   - same table, but combining ALL sections into a single count block, A4 portrait.

--section-number sets the base report section number used to auto-number
section headings in "landscape-detail" (default: "9", producing "9.1",
"9.2", "9.3", ... in the order sections appear in the workbook). Ignored
for "portrait-detail", "veri-summary-by-section", and
"veri-summary-executive".

OUTPUT FILENAME: the --format id is appended as a suffix to the output
filename, immediately before the file extension - but ONLY when the output
filename is auto-derived from the input filename (i.e. no explicit output
argument was given). For example, "QC FUP v1.2.xlsx" run with
--format landscape-detail (no output argument) produces
"QC FUP v1.2-landscape-detail.docx". If you explicitly provide an output
filename (e.g. "python excel_to_word_findings.py input.xlsx report.docx
--format landscape-detail"), that exact filename ("report.docx") is used
as-is, with NO suffix appended.

Manual overrides (use if auto-detection of the table picks the wrong
region - e.g. if other bordered cells exist elsewhere on the sheet):
    --top-row N --left-col N --bottom-row N --right-col N
(1-based row/column numbers, as shown in Excel's row/column headers.)
"""

from .cli import main

__all__ = ["main"]
