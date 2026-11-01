"""Write a ParsedData to an Excel workbook, one tab per section.

Cells are written directly with xlsxwriter, which is several times faster than
DataFrame.to_excel. Every tab gets a frozen, filterable header row and column
widths that fit its contents. A section with no records still gets its tab,
with just the header row, so "no records" can't be mistaken for "not parsed".
"""

import logging
from typing import List, Sequence

import pandas as pd
import xlsxwriter

from .model import TABS, ParsedData

logger = logging.getLogger('psse2excel')

_WIDTH_SAMPLE = 500     # rows looked at to size each column
_MIN_WIDTH, _MAX_WIDTH = 6, 60


def write_workbook(data: ParsedData, filename: str, include_summary: bool = True) -> None:
    logger.info(f"Exporting data to Excel file: {filename}")
    wb = xlsxwriter.Workbook(filename, {'constant_memory': True})
    header = wb.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'top'})
    try:
        if include_summary:
            _write_rows(wb, 'Summary', *data.summary_table(), header)
            _write_rows(wb, 'System_Info', *data.system_info_table(), header)
        for tab in TABS:
            df = getattr(data, tab.attr)
            _write_frame(wb, tab.sheet[:31], df, header)
            logger.debug(f"Exported {len(df)} records to sheet '{tab.sheet}'")
    finally:
        wb.close()
    logger.info(f"Successfully exported data to {filename}")


def _write_frame(wb, sheet: str, df: pd.DataFrame, header) -> None:
    columns = [str(c) for c in df.columns]
    values = [df[c].tolist() for c in df.columns] if len(df) else [[] for _ in columns]
    _write_columns(wb, sheet, columns, values, header)


def _write_rows(wb, sheet: str, columns: List[str], rows: List[list], header) -> None:
    values = [list(col) for col in zip(*rows)] if rows else [[] for _ in columns]
    _write_columns(wb, sheet, columns, values, header)


def _write_columns(wb, sheet: str, columns: List[str], values: Sequence[list], header) -> None:
    ws = wb.add_worksheet(sheet)
    n_rows = len(values[0]) if values else 0
    for c, (name, col) in enumerate(zip(columns, values)):
        ws.set_column(c, c, _width(name, col))
    ws.freeze_panes(1, 0)
    if columns:
        ws.autofilter(0, 0, n_rows, len(columns) - 1)
    for c, name in enumerate(columns):
        ws.write_string(0, c, name, header)

    write_number, write_string = ws.write_number, ws.write_string
    for r, row in enumerate(zip(*values), start=1):
        for c, v in enumerate(row):
            if v is None:
                continue
            kind = type(v)
            if kind is str:
                if v:
                    write_string(r, c, v)
            elif kind is float:
                if v != v:          # NaN: an empty cell
                    continue
                if v in (_INF, -_INF):
                    write_string(r, c, 'inf' if v > 0 else '-inf')
                else:
                    write_number(r, c, v)
            elif kind is int:
                write_number(r, c, v)
            elif kind is bool:
                ws.write_boolean(r, c, v)
            elif pd.isna(v):
                continue
            else:
                ws.write(r, c, v)


_INF = float('inf')


def _width(name: str, col: list) -> float:
    longest = len(name)
    for v in col[:_WIDTH_SAMPLE]:
        if v is None or (type(v) is float and v != v):
            continue
        longest = max(longest, len(str(v)))
    return min(max(longest + 2, _MIN_WIDTH), _MAX_WIDTH)
