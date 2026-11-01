"""Split one raw-file record into its fields.

Fields are separated by commas or blanks. Names may be in single or double
quotes and may contain commas, slashes or the other quote character. Anything
after a "/" outside quotes is a comment. Fields come back stripped and without
their quotes.
"""

import csv
import re
from typing import Iterable, List

# A line that csv can split correctly: comma-separated fields, each either a
# single-quoted name with no apostrophe inside, or unquoted text with no blank.
# A line that doesn't match (blank delimiters, double quotes, apostrophes in
# names, unclosed quotes) goes through tokenize() instead.
_FIELD = r"""[ \t]*(?:'[^']*'|[^\s,'"]*)[ \t]*"""
_SIMPLE_LINE = re.compile(rf"{_FIELD}(?:,{_FIELD})*")


def _needs_tokenizer(line: str) -> bool:
    """True when csv's simple rules would split this line wrongly."""
    return _SIMPLE_LINE.fullmatch(line) is None


def split_fields(line: str) -> List[str]:
    """Split one record into fields."""
    if '/' in line:
        line = strip_comment(line)
    line = line.strip()
    if not _needs_tokenizer(line):
        # Common case: commas only, single-quoted names. csv is fast.
        parts = next(csv.reader([line], quotechar="'", skipinitialspace=True), [])
        return [part.strip() for part in parts]
    return tokenize(line)


def split_all(lines: Iterable[str]) -> List[List[str]]:
    """Split many records. Same result as split_fields on each line, but faster:
    the lines that csv can handle go through one csv reader together. That is
    safe because those lines have no unclosed quote that could run on into the
    next line."""
    lines = [strip_comment(line).strip() if '/' in line else line.strip() for line in lines]
    out: List[List[str]] = [None] * len(lines)  # type: ignore[list-item]
    simple_idx = []
    for idx, line in enumerate(lines):
        if _needs_tokenizer(line):
            out[idx] = tokenize(line)
        else:
            simple_idx.append(idx)
    reader = csv.reader((lines[idx] for idx in simple_idx), quotechar="'", skipinitialspace=True)
    for idx, parts in zip(simple_idx, reader):
        out[idx] = [part.strip() for part in parts]
    # csv yields nothing for an empty line; split_fields gives [] there too
    for idx in simple_idx:
        if out[idx] is None:
            out[idx] = []
    return out


def strip_comment(line: str) -> str:
    """Drop a "/ comment" that is outside quotes."""
    quote = None
    n = len(line)
    for idx, ch in enumerate(line):
        if quote:
            if ch == quote and (idx + 1 == n or line[idx + 1] in ' \t,/'):
                quote = None
        elif ch in "'\"":
            quote = ch
        elif ch == '/':
            return line[:idx]
    return line


def tokenize(line: str) -> List[str]:
    """General field splitter: commas or blanks as delimiters, either quote style.

    A quote character closes a name only when it is followed by the end of the
    line, a blank, a comma or a "/". Otherwise it is part of the name, as in
    'O'Brien Wi~1'.
    """
    fields = []
    current = []
    quote = None
    pending_blank = False   # blanks seen after a field, before the next delimiter
    have_field = False      # something (even an empty quoted name) since the last comma
    n = len(line)
    for idx, ch in enumerate(line):
        if quote:
            if ch == quote and (idx + 1 == n or line[idx + 1] in ' \t,/'):
                quote = None
            else:
                current.append(ch)
            continue
        if ch == ',':
            fields.append(''.join(current).strip())
            current, pending_blank, have_field = [], False, False
        elif ch in ' \t':
            if have_field:
                pending_blank = True
        else:
            if pending_blank:
                # a blank between two fields is a delimiter
                fields.append(''.join(current).strip())
                current = []
            pending_blank = False
            have_field = True
            if ch in "'\"":
                quote = ch
            else:
                current.append(ch)
    fields.append(''.join(current).strip())
    return fields
