"""PSSEParser: read a PSS/E raw file (v33, v34 or v35) into a ParsedData."""

import logging
import math
import os
import re
from typing import Callable, Dict, List, NamedTuple, Optional, Tuple

import pandas as pd

from . import layouts as L
from .model import TAB_BY_ATTR, ParsedData
from .records import pad, to_frame
from .tokenizer import split_all, split_fields
from .validate import PSSEDataValidator

try:
    from tqdm import tqdm
except ImportError:  # the progress bar is optional
    tqdm = None

logger = logging.getLogger('psse2excel')

# A section terminator is a record that is just "0", optionally followed by a
# "/ comment" such as "/ END OF BUS DATA, BEGIN LOAD DATA".
_TERMINATOR = re.compile(r'^0\s*(/.*)?$')
_END_OF = re.compile(r'END OF (.+? DATA)', re.IGNORECASE)


class Section(NamedTuple):
    lines: List[str]
    line_numbers: List[int]   # 1-based line number in the file of each line


def read_text(filepath: str) -> str:
    """Read the file as UTF-8, falling back to Windows-1252 and then Latin-1."""
    with open(filepath, 'rb') as f:
        raw = f.read()
    for encoding in ('utf-8', 'cp1252'):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            logger.info(f"File is not {encoding}, trying the next encoding")
    return raw.decode('latin-1')  # never fails


def _is_int(text: Optional[str]) -> bool:
    if not text:
        return False
    try:
        int(text)
        return True
    except ValueError:
        return False


def _quoted(line: str) -> bool:
    """True when the record's first field is a quoted name."""
    s = line.lstrip()
    return s[:1] in ("'", '"')


class PSSEParser:
    """
    Parser for PSS/E raw files, versions 33, 34 and 35.

    Older versions are rearranged into v35 field order, so every version gives
    the same tabs and columns.
    """

    def __init__(self, validate_data: bool = True, progress_bar: bool = True):
        """
        Args:
            validate_data: check the parsed data for common problems (results are logged)
            progress_bar: show a progress bar while parsing sections (needs tqdm)
        """
        self.logger = logger
        self.validate_data = validate_data
        self.progress_bar = progress_bar
        self.validator = PSSEDataValidator()
        self._version = 35
        self._data: Optional[ParsedData] = None
        self.parsing_stats: Dict[str, int] = {}

        # Section name -> method that parses it
        self.section_parsers: Dict[str, Callable[[Section], None]] = {
            'SYSTEM-WIDE DATA': self._parse_system_wide,
            'BUS DATA': self._simple('buses', L.BUS, 'bus'),
            'LOAD DATA': self._simple('loads', L.LOAD, 'load', 'LOAD'),
            'FIXED SHUNT DATA': self._simple('fixed_shunts', L.FIXED_SHUNT, 'fixed shunt'),
            'GENERATOR DATA': self._simple('generators', L.GENERATOR, 'generator', 'GENERATOR'),
            'BRANCH DATA': self._simple('branches', L.BRANCH, 'branch', 'BRANCH'),
            'SYSTEM SWITCHING DEVICE DATA': self._simple(
                'system_switching_devices', L.SYSTEM_SWITCHING_DEVICE, 'system switching device'),
            'TRANSFORMER DATA': self._parse_transformers,
            'AREA DATA': self._simple('areas', L.AREA, 'area'),
            'TWO-TERMINAL DC DATA': self._parse_two_terminal_dc,
            'VSC DC LINE DATA': self._parse_vsc_dc,
            'IMPEDANCE CORRECTION DATA': self._parse_impedance_correction,
            'MULTI-TERMINAL DC DATA': self._parse_multi_terminal_dc,
            'MULTI-SECTION LINE DATA': self._simple(
                'multi_section_lines', L.MULTI_SECTION_LINE, 'multi-section line'),
            'ZONE DATA': self._simple('zones', L.ZONE, 'zone'),
            'INTER-AREA TRANSFER DATA': self._simple(
                'inter_area_transfers', L.INTER_AREA_TRANSFER, 'inter-area transfer'),
            'OWNER DATA': self._simple('owners', L.OWNER, 'owner'),
            'FACTS DEVICE DATA': self._simple('facts', L.FACTS, 'FACTS device', 'FACTS'),
            'SWITCHED SHUNT DATA': self._simple(
                'switched_shunts', L.SWITCHED_SHUNT, 'switched shunt', 'SWITCHED_SHUNT'),
            'GNE DATA': self._parse_gne,
            'INDUCTION MACHINE DATA': self._simple(
                'induction_machines', L.INDUCTION_MACHINE, 'induction machine'),
            'SUBSTATION DATA': self._parse_substations,
        }

    # ------------------------------------------------------------------ entry points

    def parse_file(self, filepath: str) -> ParsedData:
        """Parse a raw file.

        Raises:
            FileNotFoundError: if the file doesn't exist
        """
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        self.logger.info(f"Starting to parse PSS/E file: {filepath}")
        self.logger.info(f"File size: {os.path.getsize(filepath) / 1024 / 1024:.2f} MB")
        return self.parse_text(read_text(filepath), source=os.path.basename(filepath))

    def parse_text(self, content: str, source: str = '') -> ParsedData:
        """Parse the text of a raw file."""
        lines = content.splitlines()
        data = ParsedData(source_file=source)
        self._data = data
        self.parsing_stats = {'total_lines_processed': len(lines), 'successful_records': 0,
                              'failed_records': 0, 'sections_found': 0, 'validation_issues': 0}

        header_idx = self._header_index(lines)
        self._read_header(lines, header_idx, data)
        data.file_version = self._read_rev(lines, header_idx)
        self._version = data.version = self._detect_version(data.file_version)
        sections = self._split_sections(lines, header_idx, data.version)

        order = L.SECTION_ORDER[data.version]
        bar = None
        if self.progress_bar and tqdm is not None:
            bar = tqdm(total=len(order), desc="Parsing sections")
        for name in order:
            section = sections.get(name, Section([], []))
            try:
                self.section_parsers[name](section)
            except Exception as e:  # keep going with the other sections
                self.logger.error(f"Error parsing {name}: {e}")
            if bar:
                bar.update(1)
        if bar:
            bar.close()

        # Sections the file doesn't have still get their columns, for the header row
        for tab in TAB_BY_ATTR.values():
            if len(getattr(data, tab.attr).columns) == 0:
                setattr(data, tab.attr, pd.DataFrame(columns=tab.columns))
        self.parsing_stats['successful_records'] = sum(
            len(getattr(data, t.attr)) for t in TAB_BY_ATTR.values())
        self.parsing_stats['failed_records'] = data.total_skipped
        self.parsing_stats['sections_found'] = sum(
            1 for t in TAB_BY_ATTR.values() if not getattr(data, t.attr).empty)

        if self.validate_data:
            self._validate(data)
        self._log_summary(data)
        if data.total_skipped or data.lines_not_read:
            self.logger.warning(
                f"Parsing finished with {data.total_skipped} records skipped and "
                f"{data.lines_not_read} lines not read; see the warnings above")
        else:
            self.logger.info("File parsing completed successfully")
        return data

    def get_parsing_statistics(self) -> Dict[str, int]:
        return dict(self.parsing_stats)

    # ------------------------------------------------------------------ header and sections

    @staticmethod
    def _header_index(lines: List[str]) -> int:
        """Index of the case identification line (the first line that is not an @! comment)."""
        for idx, line in enumerate(lines):
            if not line.lstrip().startswith('@!'):
                return idx
        return 0

    def _read_header(self, lines: List[str], header_idx: int, data: ParsedData) -> None:
        """Base MVA from the case identification line; title from the two lines after it."""
        if len(lines) > header_idx:
            header_parts = split_fields(lines[header_idx].split('/')[0])
            if len(header_parts) >= 2:
                try:
                    data.base_mva = float(header_parts[1])
                    self.logger.info(f"System base MVA: {data.base_mva}")
                except ValueError:
                    data.base_mva = 100.0
                    self.logger.warning("Could not parse base MVA, using default 100.0")
        title_lines = [line.strip().strip('"\'').strip()
                       for line in lines[header_idx + 1:header_idx + 3]]
        data.title = ' | '.join(t for t in title_lines if t)
        self.logger.info(f"System title: {data.title}")

    @staticmethod
    def _read_rev(lines: List[str], header_idx: int) -> Optional[int]:
        """REV, the 3rd field of the case identification line, or None if it can't be read."""
        if len(lines) > header_idx:
            parts = split_fields(lines[header_idx].split('/')[0])
            if len(parts) > 2 and parts[2]:
                try:
                    return int(float(parts[2]))
                except ValueError:
                    pass
        return None

    def _detect_version(self, rev: Optional[int]) -> int:
        """The version to read the file as: REV if supported, else the nearest supported one."""
        if rev is None:
            self.logger.warning("Could not read the PSS/E version (REV) from the header; assuming 35")
            return 35
        if rev in L.SECTION_ORDER:
            self.logger.info(f"PSS/E version: {rev}")
            return rev
        nearest = min(L.SUPPORTED_VERSIONS, key=lambda v: abs(v - rev))
        self.logger.warning(f"PSS/E version {rev} is not supported; reading it as version {nearest}")
        return nearest

    def _split_sections(self, lines: List[str], header_idx: int, version: int) -> Dict[str, Section]:
        """Split the data lines into sections, keyed by section name.

        Data starts after the case identification line and the two title lines.
        Every "0" record closes the current section and the next one in
        SECTION_ORDER begins. A "Q" record ends the data. Comment lines ("/..."
        and "@!...") and blank lines are dropped. Substation data has "0"
        records of its own, so everything from its start up to "Q" belongs to it.
        """
        order = L.SECTION_ORDER[version]
        sections = {name: Section([], []) for name in order}
        pos = 0
        for idx in range(header_idx + 3, len(lines)):
            line = lines[idx].strip()
            if pos >= len(order) or line == 'Q' or line.startswith('Q '):
                break
            if not line or line.startswith('/') or line.startswith('@!'):
                continue
            if _TERMINATOR.match(line) and order[pos] != 'SUBSTATION DATA':
                said = _END_OF.search(line)
                said_name = said.group(1).upper() if said else None
                said_name = L.SECTION_ALIASES.get(said_name, said_name)
                if said_name and said_name != order[pos]:
                    self.logger.warning(f"Line {idx + 1}: terminator says 'END OF {said.group(1)}', "
                                        f"expected end of {order[pos]}")
                pos += 1
                continue
            sections[order[pos]].lines.append(line)
            sections[order[pos]].line_numbers.append(idx + 1)
        if pos < len(order) and order[pos] != 'SUBSTATION DATA':
            self.logger.debug(f"Data ended before {order[pos]}")
        return sections

    # ------------------------------------------------------------------ helpers

    def _store(self, attr: str, records, layout, what: str, where, found: int,
               orphans: int = 0, skipped_early: int = 0) -> pd.DataFrame:
        """Convert records, store the table and log the record count."""
        df, skipped = to_frame(records, layout, what, where)
        if list(df.columns) != TAB_BY_ATTR[attr].columns:
            df = df[TAB_BY_ATTR[attr].columns]  # the tab's column order (3W transformers)
        setattr(self._data, attr, df)
        total_skipped = skipped + skipped_early
        if total_skipped:
            self._data.skipped[attr] = self._data.skipped.get(attr, 0) + total_skipped
        self._data.lines_not_read += orphans
        self._log_count(what, found, len(df), total_skipped, orphans)
        return df

    def _log_count(self, what: str, found: int, parsed: int, skipped: int, orphans: int = 0):
        if not found and not orphans:
            return
        msg = f"{what}: {found} records in file, {parsed} parsed"
        if skipped:
            msg += f", {skipped} skipped"
        if orphans:
            msg += f", {orphans} lines that belong to no record"
        (self.logger.warning if skipped or orphans else self.logger.info)(msg)

    def _remap(self, parts: List[str], record: Optional[str]) -> List[str]:
        return L.to_v35_order(parts, record, self._version) if record else parts

    def _simple(self, attr: str, layout, what: str, record: Optional[str] = None):
        """Parser for a section with one line per record."""
        def parse(section: Section) -> None:
            records = [self._remap(t, record) for t in split_all(section.lines)]
            self._store(attr, records, layout, what, section.line_numbers, len(records))
        return parse

    def _orphans(self, what: str, section: Section, start: int, end: int) -> int:
        """Log lines start..end-1 that are not part of any record; return how many."""
        n = end - start
        self.logger.warning(
            f"Line {section.line_numbers[start]}: {n} {what} line(s) that are not part of a "
            f"record were skipped: {section.lines[start][:120]}")
        return n

    # ------------------------------------------------------------------ multi-line records

    @staticmethod
    def _transformer_start(t: List[str]) -> bool:
        """Line 1 of a transformer: I, J, K are integers, and so are CW, CZ, CM if present."""
        return (len(t) >= 3 and _is_int(t[0]) and _is_int(t[1]) and _is_int(t[2])
                and all(_is_int(x) for x in t[4:7] if x))

    def _parse_transformers(self, section: Section) -> None:
        """Transformers: 4 lines for two-winding, 5 lines for three-winding (K != 0).

        If a record has too few lines or a line doesn't belong to a record, it is
        logged and parsing carries on at the next transformer's first line, so
        one bad record doesn't shift all the records after it.
        """
        toks = split_all(section.lines)
        n = len(toks)
        start = [self._transformer_start(t) for t in toks]
        rec2, where2, rec3, where3 = [], [], [], []
        orphans = short2 = short3 = 0
        i = 0
        while i < n:
            if not start[i]:
                j = i + 1
                while j < n and not start[j]:
                    j += 1
                orphans += self._orphans('transformer', section, i, j)
                i = j
                continue
            t = toks[i]
            three = int(t[2]) != 0
            size = 5 if three else 4
            end = i + 1
            while end < n and end < i + size and not start[end]:
                end += 1
            if end - i < size:
                self.logger.warning(
                    f"Line {section.line_numbers[i]}: {'3' if three else '2'}-winding transformer "
                    f"has {end - i} of its {size} lines; skipped: {section.lines[i][:120]}")
                if three:
                    short3 += 1
                else:
                    short2 += 1
                i = end
                continue
            windings = [self._remap(w, 'WINDING') for w in toks[i + 2:i + (5 if three else 3)]]
            if three:
                rec3.append(pad(t, L.XFMR_LINE1_LEN) + pad(toks[i + 1], 11)
                            + pad(windings[0], L.WINDING_LEN) + pad(windings[1], L.WINDING_LEN)
                            + pad(windings[2], L.WINDING_LEN))
                where3.append(section.line_numbers[i])
            else:
                rec2.append(pad(t, L.XFMR_LINE1_LEN) + pad(toks[i + 1], 3)
                            + pad(windings[0], L.WINDING_LEN) + pad(toks[i + 3], 2))
                where2.append(section.line_numbers[i])
            i += size
        self._store('two_winding_transformers', rec2, L.TWO_WINDING, '2-winding transformer',
                    where2, len(rec2) + short2, orphans, short2)
        self._store('three_winding_transformers', rec3, L.THREE_WINDING, '3-winding transformer',
                    where3, len(rec3) + short3, 0, short3)

    def _three_line_records(self, section: Section, what: str, line1_len: int,
                            conv_len: int, conv_record: str):
        """Group two-terminal DC or VSC lines into 3-line records.

        Line 1 starts with the quoted NAME; lines 2 and 3 (the converters) start
        with a bus number. A record that is cut short is logged and skipped, and
        parsing carries on at the next line 1.
        """
        toks = split_all(section.lines)
        n = len(toks)
        # converter-like: first field is an unquoted integer (a bus number)
        conv = [(not _quoted(line)) and bool(t) and _is_int(t[0])
                for line, t in zip(section.lines, toks)]
        records, where = [], []
        orphans = short = 0
        i = 0
        while i < n:
            if conv[i] and len(toks[i]) > line1_len:
                # a converter line where line 1 should be
                j = i + 1
                while j < n and conv[j]:
                    j += 1
                orphans += self._orphans(what, section, i, j)
                i = j
                continue
            if i + 2 < n and conv[i + 1] and conv[i + 2]:
                records.append(pad(toks[i], line1_len)
                               + pad(self._remap(toks[i + 1], conv_record), conv_len)
                               + pad(self._remap(toks[i + 2], conv_record), conv_len))
                where.append(section.line_numbers[i])
                i += 3
                continue
            self.logger.warning(f"Line {section.line_numbers[i]}: {what} record has fewer than "
                                f"3 lines; skipped: {section.lines[i][:120]}")
            short += 1
            j = i + 1
            while j < n and conv[j]:
                j += 1
            i = j
        return records, where, orphans, short

    def _parse_two_terminal_dc(self, section: Section) -> None:
        records, where, orphans, short = self._three_line_records(
            section, 'two-terminal DC', L.DC_LINE1_LEN, L.DC_CONVERTER_LEN, 'DC_CONVERTER')
        self._store('two_terminal_dc_lines', records, L.TWO_TERMINAL_DC, 'two-terminal DC',
                    where, len(records) + short, orphans, short)

    def _parse_vsc_dc(self, section: Section) -> None:
        records, where, orphans, short = self._three_line_records(
            section, 'VSC DC', L.VSC_LINE1_LEN, L.VSC_CONVERTER_LEN, 'VSC_CONVERTER')
        self._store('vsc_dc_lines', records, L.VSC_DC, 'VSC DC', where,
                    len(records) + short, orphans, short)

    def _parse_multi_terminal_dc(self, section: Section) -> None:
        """Multi-terminal DC: header line, then NCONV converter, NDCBS DC bus and
        NDCLN DC link lines."""
        toks = split_all(section.lines)
        n = len(toks)
        lines_, conv, buses, links = [], [], [], []
        w_lines, w_conv, w_buses, w_links = [], [], [], []
        orphans = short = 0
        i = 0
        while i < n:
            t = toks[i]
            counts = t[1:4] if len(t) >= 4 else []
            if not (len(counts) == 3 and all(_is_int(c) for c in counts)
                    and (_quoted(section.lines[i]) or not _is_int(t[0]))):
                j = i + 1
                while j < n and not _quoted(section.lines[j]):
                    j += 1
                orphans += self._orphans('multi-terminal DC', section, i, j)
                i = j
                continue
            nconv, ndcbs, ndcln = (int(c) for c in counts)
            size = 1 + nconv + ndcbs + ndcln
            end = i + 1
            while end < n and end < i + size and not _quoted(section.lines[end]):
                end += 1
            if end - i < size:
                self.logger.warning(
                    f"Line {section.line_numbers[i]}: multi-terminal DC line has {end - i} of "
                    f"its {size} lines; skipped: {section.lines[i][:120]}")
                short += 1
                i = end
                continue
            name = t[0]
            lines_.append(t)
            w_lines.append(section.line_numbers[i])
            k = i + 1
            for out, where, count in ((conv, w_conv, nconv), (buses, w_buses, ndcbs),
                                      (links, w_links, ndcln)):
                for _ in range(count):
                    out.append([name] + toks[k])
                    where.append(section.line_numbers[k])
                    k += 1
            i += size
        name_field = [L.S('NAME')]
        self._store('mtdc_lines', lines_, L.MTDC_LINE, 'multi-terminal DC line', w_lines,
                    len(lines_) + short, orphans, short)
        self._store('mtdc_converters', conv, name_field + L.MTDC_CONVERTER,
                    'multi-terminal DC converter', w_conv, len(conv))
        self._store('mtdc_dc_buses', buses, name_field + L.MTDC_DC_BUS,
                    'multi-terminal DC bus', w_buses, len(buses))
        self._store('mtdc_dc_links', links, name_field + L.MTDC_DC_LINK,
                    'multi-terminal DC link', w_links, len(links))

    def _parse_gne(self, section: Section) -> None:
        """GNE devices: line 1 with NAME, MODEL, NTERM, the buses and NREAL, NINTG,
        NCHAR; line 2 with STATUS, OWNER, NMETR; then the values, 10 per line."""
        toks = split_all(section.lines)
        n = len(toks)
        records, where = [], []
        orphans = short = 0
        i = 0
        while i < n:
            t = toks[i]
            sizes = None
            if _quoted(section.lines[i]) and len(t) >= 3 and _is_int(t[2]):
                nterm = int(t[2])
                counts = t[3 + nterm:6 + nterm]
                if len(counts) == 3 and all(_is_int(c) for c in counts):
                    sizes = [math.ceil(int(c) / 10) for c in counts]
            if sizes is None:
                j = i + 1
                while j < n and not _quoted(section.lines[j]):
                    j += 1
                orphans += self._orphans('GNE', section, i, j)
                i = j
                continue
            size = 2 + sum(sizes)
            if i + size > n:
                self.logger.warning(f"Line {section.line_numbers[i]}: GNE device is cut short; "
                                    f"skipped: {section.lines[i][:120]}")
                short += 1
                break
            nterm = int(t[2])
            values = []
            k = i + 2
            for count in sizes:
                group = [v for row in toks[k:k + count] for v in row]
                values.append(', '.join(v.strip('"\'') for v in group))
                k += count
            line2 = pad(toks[i + 1], 3)
            records.append([t[0], t[1], t[2], ', '.join(t[3:3 + nterm])]
                           + t[3 + nterm:6 + nterm] + line2 + values)
            where.append(section.line_numbers[i])
            i += size
        self._store('gne_devices', records, L.GNE, 'GNE device', where,
                    len(records) + short, orphans, short)

    def _parse_substations(self, section: Section) -> None:
        """Substations (v35): each substation record is followed by node,
        switching device and equipment connection blocks, each ended by "0"."""
        blocks = {'substations': [], 'substation_nodes': [],
                  'substation_switching_devices': [], 'substation_connections': []}
        where = {k: [] for k in blocks}
        order = ['substations', 'substation_nodes', 'substation_switching_devices',
                 'substation_connections']
        state = 0
        current = None
        for line, lineno in zip(section.lines, section.line_numbers):
            if _TERMINATOR.match(line):
                if state == 0:
                    break  # end of substation data
                state = (state + 1) % 4
                continue
            t = split_fields(line)
            target = order[state]
            if state == 0:
                current = t[0] if t else ''
                blocks[target].append(t)
                state = 1
            else:
                blocks[target].append([current] + t)
            where[target].append(lineno)
        layouts = {'substations': L.SUBSTATION,
                   'substation_nodes': [L.I('IS')] + L.SUBSTATION_NODE,
                   'substation_switching_devices': [L.I('IS')] + L.SUBSTATION_SWITCHING_DEVICE,
                   'substation_connections': [L.I('IS')] + L.SUBSTATION_CONNECTION}
        for attr in order:
            what = attr.replace('_', ' ').rstrip('s')
            self._store(attr, blocks[attr], layouts[attr], what, where[attr], len(blocks[attr]))

    # ------------------------------------------------------------------ other sections

    def _parse_system_wide(self, section: Section) -> None:
        """System-wide data (v34+): GENERAL, GAUSS, NEWTON, ADJUST, TYSL and SOLVER
        records with KEY=VALUE settings, and RATING records."""
        rows, where = [], []
        for t, lineno in zip(split_all(section.lines), section.line_numbers):
            if not t:
                continue
            record = t[0].upper()
            if record == 'RATING':
                vals = pad(t[1:], 3)
                rows.append(['RATING', _number(vals[0] or ''), (vals[1] or '').strip(),
                             (vals[2] or '').strip()])
                where.append(lineno)
                continue
            for item in t[1:]:
                if not item:
                    continue
                if '=' in item:
                    key, value = item.split('=', 1)
                    rows.append([record, key.strip(), _number(value.strip()), ''])
                else:
                    rows.append([record, 'METHOD' if record == 'SOLVER' else '', item, ''])
                where.append(lineno)
        df = pd.DataFrame(rows, columns=L.SYSTEM_WIDE_COLUMNS)
        self._data.system_wide = df
        self._log_count('system-wide setting', len(rows), len(rows), 0)

    def _parse_impedance_correction(self, section: Section) -> None:
        if self._version < 34:
            self._parse_impedance_correction_v33(section)
            return
        # v34/v35: line 1 is I, T1, Re(F1), Im(F1), ... T6, Re(F6), Im(F6); more points
        # follow on continuation lines, 6 per line. A new table starts with the integer I.
        toks = split_all(section.lines)
        tables: List[Tuple[int, List[List[str]]]] = []
        orphans = 0
        for k, t in enumerate(toks):
            if len(t) < 3:
                self.logger.warning(f"Line {section.line_numbers[k]}: skipped short impedance "
                                    f"correction line: {section.lines[k]}")
                orphans += 1
            elif t[0].isdigit():
                tables.append((k, [t]))
            elif tables:
                tables[-1][1].append(t)
            else:
                orphans += self._orphans('impedance correction', section, k, k + 1)
        rows = []
        skipped = 0
        for k, table in tables:
            try:
                row = {'I': int(table[0][0])}
                row.update(_impedance_points(table[0][1:], 1))
                for j, cont in enumerate(table[1:]):
                    row.update(_impedance_points(cont, 7 + 6 * j))
                rows.append(row)
            except ValueError as e:
                self.logger.warning(f"Line {section.line_numbers[k]}: skipped impedance "
                                    f"correction table: {e}")
                skipped += 1
        self._store_rows('impedance_correction', rows, 'impedance correction table',
                         len(tables), skipped, orphans)

    def _parse_impedance_correction_v33(self, section: Section) -> None:
        """v33 impedance correction: one line per table, I, T1, F1, ..., T11, F11.

        v33 factors are real numbers. They go in the v35 Re(F) columns, with
        Im(F) = 0, the same way PSS/E converts them when it reads a v33 case.
        """
        rows = []
        skipped = 0
        for t, lineno, line in zip(split_all(section.lines), section.line_numbers, section.lines):
            try:
                row = {'I': int(t[0])}
                for k in range((len(t) - 1) // 2):
                    temp, factor = t[1 + 2 * k], t[2 + 2 * k]
                    if not temp.strip() and not factor.strip():
                        continue
                    row[f'T{k + 1}'] = float(temp)
                    row[f'Re(F{k + 1})'] = float(factor)
                    row[f'Im(F{k + 1})'] = 0.0
                rows.append(row)
            except (ValueError, IndexError) as e:
                self.logger.warning(f"Line {lineno}: skipped impedance correction table: {e}. "
                                    f"Line: {line[:120]}")
                skipped += 1
        self._store_rows('impedance_correction', rows, 'impedance correction table',
                         len(rows) + skipped, skipped, 0)

    def _store_rows(self, attr, rows, what, found, skipped, orphans):
        df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=TAB_BY_ATTR[attr].columns)
        setattr(self._data, attr, df)
        if skipped:
            self._data.skipped[attr] = self._data.skipped.get(attr, 0) + skipped
        self._data.lines_not_read += orphans
        self._log_count(what, found, len(df), skipped, orphans)

    # ------------------------------------------------------------------ checks and summary

    def _validate(self, data: ParsedData) -> None:
        self.logger.info("Validating parsed data...")
        v = self.validator
        issues = []
        issues += [f"Bus validation: {x}" for x in v.validate_bus_data(data.buses)]
        issues += [f"Generator validation: {x}"
                   for x in v.validate_generator_data(data.generators, data.buses)]
        issues += [f"Branch validation: {x}"
                   for x in v.validate_branch_data(data.branches, data.buses)]
        issues += [f"Load validation: {x}" for x in v.validate_load_data(data.loads, data.buses)]
        issues += [f"2W Transformer validation: {x}"
                   for x in v.validate_transformer_data(data.two_winding_transformers, data.buses)]
        issues += [f"3W Transformer validation: {x}"
                   for x in v.validate_transformer_data(data.three_winding_transformers, data.buses)]
        self.parsing_stats['validation_issues'] = len(issues)
        if issues:
            self.logger.warning(f"Validation found {len(issues)} issues:")
            for issue in issues[:10]:
                self.logger.warning(f"  - {issue}")
            if len(issues) > 10:
                self.logger.warning(f"  ... and {len(issues) - 10} more issues")
        else:
            self.logger.info("Data validation passed successfully")

    def _log_summary(self, data: ParsedData) -> None:
        self.logger.info("=== PARSING SUMMARY ===")
        self.logger.info(f"System: {data.title}")
        self.logger.info(f"Base MVA: {data.base_mva}, PSS/E version {data.version}")
        self.logger.info(f"Total lines processed: {self.parsing_stats['total_lines_processed']}")
        for tab in TAB_BY_ATTR.values():
            count = len(getattr(data, tab.attr))
            if count:
                self.logger.info(f"{tab.sheet}: {count}")
        if self.parsing_stats.get('validation_issues'):
            self.logger.warning(f"Validation issues: {self.parsing_stats['validation_issues']}")
        self.logger.info("=====================")


def _number(text: str):
    """Text as int or float when it is a number, else the text itself."""
    for conv in (int, float):
        try:
            return conv(text)
        except ValueError:
            pass
    return text


def _impedance_points(t: List[str], first: int) -> Dict[str, Optional[float]]:
    """Six T, Re(F), Im(F) points from one impedance correction line. Points the
    line doesn't have are None, except that the first point defaults to 0.0."""
    row: Dict[str, Optional[float]] = {}
    for p in range(6):
        k = first + p
        default = 0.0 if p == 0 else None
        for name, idx in ((f'T{k}', 3 * p), (f'Re(F{k})', 3 * p + 1), (f'Im(F{k})', 3 * p + 2)):
            v = t[idx] if idx < len(t) else ''
            row[name] = float(v) if v else default
    return row
