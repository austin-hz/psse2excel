"""psse2excel: read PSS/E raw files (v33, v34, v35) without psspy and export them to Excel or CSV.

    from psse2excel import PSSEParser
    data = PSSEParser().parse_file("case.raw")
    data.to_excel("case.xlsx")
    data.to_csv("case_csv")      # one CSV file per tab

or from the command line:

    python -m psse2excel case.raw -o case.xlsx
    python -m psse2excel case.raw --csv
"""

from .model import ParsedData, TABS
from .parser import PSSEParser
from .validate import PSSEDataValidator

__version__ = "1.0.0"
__all__ = ["PSSEParser", "ParsedData", "PSSEDataValidator", "TABS", "parse_file", "__version__"]


def parse_file(filepath: str, validate_data: bool = True) -> ParsedData:
    """Parse a raw file with default settings and no progress bar."""
    return PSSEParser(validate_data=validate_data, progress_bar=False).parse_file(filepath)
