"""Write a ParsedData as CSV files, one per tab, into a folder.

The files hold the same columns and values as the Excel tabs:
- Each file is named after its tab (Buses.csv, 2W_Transformers.csv, ...).
- A section with no records gets a file with just the header row.
- Empty cells stay empty, and infinite values are written as inf and -inf.

Files are UTF-8 with a byte order mark, so Excel shows accented names
correctly when it opens them.
"""

import csv
import logging
import os
from typing import List

import pandas as pd

from .model import TABS, ParsedData

logger = logging.getLogger('psse2excel')

ENCODING = 'utf-8-sig'
_INF = float('inf')


def write_csv_folder(data: ParsedData, folder: str, include_summary: bool = True) -> List[str]:
    """Write one CSV per tab into folder (created if needed). Returns the paths written.

    Files of the same name are overwritten; other files in the folder are left alone."""
    logger.info(f"Exporting data to CSV files in: {folder}")
    os.makedirs(folder, exist_ok=True)
    written = []
    if include_summary:
        for name, (columns, rows) in (('Summary', data.summary_table()),
                                      ('System_Info', data.system_info_table())):
            written.append(_write(folder, name, columns, rows))
    for tab in TABS:
        df = getattr(data, tab.attr)
        columns = [str(c) for c in df.columns]
        values = [df[c].tolist() for c in df.columns]
        written.append(_write(folder, tab.sheet, columns, zip(*values)))
        logger.debug(f"Exported {len(df)} records to {tab.sheet}.csv")
    logger.info(f"Successfully exported {len(written)} CSV files to {folder}")
    return written


def _write(folder: str, name: str, columns: List[str], rows) -> str:
    path = os.path.join(folder, name + '.csv')
    with open(path, 'w', encoding=ENCODING, newline='') as f:
        w = csv.writer(f)
        w.writerow(columns)
        w.writerows([_text(v) for v in row] for row in rows)
    return path


def _text(v) -> str:
    """The cell text, matching what the Excel writer puts in the cell."""
    if v is None:
        return ''
    kind = type(v)
    if kind is str:
        return v
    if kind is float:
        if v != v:                  # NaN: an empty cell
            return ''
        if v in (_INF, -_INF):
            return 'inf' if v > 0 else '-inf'
        return repr(v)              # shortest text that reads back as the same number
    if kind is int or kind is bool:
        return str(v)
    if pd.isna(v):
        return ''
    return str(v)
