"""Turn records (lists of field strings) into tables, using the layouts.

Conversion goes column by column, which is fast. If any value in a section
can't be converted, the section is converted again record by record, so that
only the bad records are skipped, each with a log message.
"""

import logging
from itertools import zip_longest
from typing import List, Optional, Sequence, Tuple

import pandas as pd

from .layouts import REQUIRED, Field

logger = logging.getLogger('psse2excel')

_CONVERTERS = {'i': int, 'r': float}
_BAD_VALUE = (ValueError, TypeError, AttributeError)


def pad(parts: List[str], length: int) -> List[Optional[str]]:
    """Give a record line exactly `length` fields. Missing fields are None, so
    "not there" stays different from "empty"."""
    if len(parts) >= length:
        return parts[:length]
    return parts + [None] * (length - len(parts))


def to_frame(records: Sequence[Sequence[Optional[str]]], layout: List[Field],
             what: str, where: Optional[Sequence[int]] = None) -> Tuple[pd.DataFrame, int]:
    """Convert records to a DataFrame with the layout's columns.

    Args:
        records: one list of field strings per record
        layout: the fields, in order
        what: record type, for log messages (e.g. "bus")
        where: file line number of each record, for log messages

    Returns:
        (table, number of records skipped because a value was unreadable)
    """
    names = [f.name for f in layout]
    if not records:
        return pd.DataFrame(columns=names), 0
    try:
        columns = _convert_columns(records, layout)
        skipped = 0
    except _BAD_VALUE:
        good = []
        for n, record in enumerate(records):
            try:
                good.append(_convert_row(record, layout))
            except _BAD_VALUE as exc:
                at = f" at line {where[n]}" if where else ""
                shown = ', '.join('' if v is None else v for v in record if v is not None)
                logger.warning(f"Skipped {what} record{at}: {_reason(exc)}. Fields: {shown[:200]}")
        skipped = len(records) - len(good)
        columns = [list(col) for col in zip(*good)] if good else [[] for _ in names]
    return pd.DataFrame(dict(zip(names, columns)), columns=names), skipped


def _reason(exc: Exception) -> str:
    if isinstance(exc, (TypeError, AttributeError)):
        return "a required field is missing"
    return str(exc)


def _token_index(layout: List[Field]) -> List[Optional[int]]:
    """Field position in the record for each layout field (None for constants)."""
    out, idx = [], 0
    for f in layout:
        if f.kind == 'c':
            out.append(None)
        else:
            out.append(idx)
            idx += 1
    return out


def _convert_columns(records, layout: List[Field]) -> List[list]:
    n = len(records)
    positions = _token_index(layout)
    width = max(p for p in positions if p is not None) + 1 if any(
        p is not None for p in positions) else 0
    cols = list(zip_longest(*records, fillvalue=None))[:width]
    while len(cols) < width:
        cols.append((None,) * n)
    out: List[list] = []
    first: Optional[list] = None
    for f, pos in zip(layout, positions):
        if f.kind == 'c':
            values = [f.default] * n
        elif f.kind == 'rest':
            values = [', '.join(v for v in r[pos:] if v is not None) for r in records]
        else:
            col = cols[pos]
            if f.kind == 's':
                if f.default_from_first:
                    fmt = f.default_from_first
                    values = [fmt.format(first[k]) if v is None else v.strip('"\'')
                              for k, v in enumerate(col)]
                elif f.default is REQUIRED:
                    values = [v.strip('"\'') for v in col]
                else:
                    d = f.default
                    values = [d if v is None else v.strip('"\'') for v in col]
            else:
                conv = _CONVERTERS[f.kind]
                if f.default is REQUIRED:
                    values = [conv(v) for v in col]
                else:
                    d = f.default
                    values = [conv(v) if v else d for v in col]
        if first is None:
            first = values
        out.append(values)
    return out


def _convert_row(record, layout: List[Field]) -> list:
    out = []
    pos = 0
    n = len(record)
    for f in layout:
        if f.kind == 'c':
            out.append(f.default)
            continue
        v = record[pos] if pos < n else None
        if f.kind == 'rest':
            out.append(', '.join(x for x in record[pos:] if x is not None))
        elif f.kind == 's':
            if v is not None:
                out.append(v.strip('"\''))
            elif f.default_from_first:
                out.append(f.default_from_first.format(out[0]))
            elif f.default is REQUIRED:
                raise TypeError(f.name)
            else:
                out.append(f.default)
        else:
            if f.default is REQUIRED:
                if not v:
                    raise TypeError(f.name)
                try:
                    out.append(_CONVERTERS[f.kind](v))
                except ValueError:
                    raise ValueError(f"{f.name} = {v!r} is not a number") from None
            elif v:
                try:
                    out.append(_CONVERTERS[f.kind](v))
                except ValueError:
                    raise ValueError(f"{f.name} = {v!r} is not a number") from None
            else:
                out.append(f.default)
        pos += 1
    return out
