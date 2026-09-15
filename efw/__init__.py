"""efw - excel_to_word_findings library.

Converts a cybersecurity "Follow-up Plan" Excel workbook (containing BOTH
the SRA "SRA Follow-up"/"Follow-up Items" table and the SA "SA Follow-up"
table, which together form an SRAA - Security Risk Assessment and Audit -
report) into a formal Word report, in one of several output formats. See
efw.cli for the full description of each format and the CLI options.
"""

from .convert import convert  # noqa: F401
from .cli import main  # noqa: F401
