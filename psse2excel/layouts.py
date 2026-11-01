"""Field layouts of the PSS/E raw-file records.

Each layout is the list of columns of one Excel tab, in v35 field order. Older
versions are rearranged into v35 order first (see FIELD_ORDER_FROM), so every
version gives the same tabs and columns.

A field is read from the record's fields one after the other:
    I(name, default)  integer
    R(name, default)  real number
    S(name, default)  text; quotes are removed
    C(name, value)    a constant column that does not read a field
    REST(name)        text: all remaining fields, joined with ", "
A field without a default is required: a record where it is missing or
unreadable is skipped and logged. For an optional number, an empty field gives
the default. For optional text, an empty field gives '' and only a field that
is not there at all gives the default.
"""

from typing import Any, Dict, List, NamedTuple, Optional

REQUIRED = object()


class Field(NamedTuple):
    name: str
    kind: str              # 'i', 'r', 's', 'c' (constant) or 'rest'
    default: Any = REQUIRED
    # For text: a default built from the record's first field, e.g. 'AREA_{}'
    default_from_first: Optional[str] = None


def I(name, default=REQUIRED):
    return Field(name, 'i', default)


def R(name, default=REQUIRED):
    return Field(name, 'r', default)


def S(name, default=REQUIRED, default_from_first=None):
    return Field(name, 's', default, default_from_first)


def C(name, value):
    return Field(name, 'c', value)


def REST(name):
    return Field(name, 'rest', '')


def _owners():
    """O1, F1, ... O4, F4 ownership fields."""
    out = [I('O1', None), R('F1', 1.0)]
    for k in range(2, 5):
        out += [I(f'O{k}', None), R(f'F{k}', None)]
    return out


def _winding(w: int):
    """One transformer winding line (v35): WINDV, NOMV, ANG, RATE1..12, COD, ..."""
    return ([R(f'WINDV{w}', 1.0), R(f'NOMV{w}', 0.0), R(f'ANG{w}', 0.0)]
            + [R(f'RATE{w}-{k}', 0.0) for k in range(1, 13)]
            + [I(f'COD{w}', 0), I(f'CONT{w}', 0), I(f'NOD{w}', 0),
               R(f'RMA{w}', 1.1), R(f'RMI{w}', 0.9), R(f'VMA{w}', 1.1), R(f'VMI{w}', 0.9),
               I(f'NTP{w}', 33), I(f'TAB{w}', 0),
               R(f'CR{w}', 0.0), R(f'CX{w}', 0.0), R(f'CNXA{w}', 0.0)])


BUS = [I('BUS_NUMBER'), S('BUS_NAME'), R('BASE_KV'), I('BUS_TYPE'),
       I('AREA', 1), I('ZONE', 1), I('OWNER', 1), R('VM', 1.0), R('VA', 0.0),
       R('NVHI', 1.1), R('NVLO', 0.9), R('EVHI', 1.1), R('EVLO', 0.9)]

LOAD = [I('BUS_NUMBER'), S('LOAD_ID'), I('STATUS', 1), I('AREA', 1), I('ZONE', 1),
        R('PL', 0.0), R('QL', 0.0), R('IP', 0.0), R('IQ', 0.0), R('YP', 0.0), R('YQ', 0.0),
        I('OWNER', 1), I('SCALE', 1), I('INTRPT', 0),
        R('DGENP', 0.0), R('DGENQ', 0.0), I('DGENM', 0), S('LD_TYPE', '')]

FIXED_SHUNT = [I('BUS_NUMBER'), S('SHUNT_ID'), I('STATUS', 1), R('GL', 0.0), R('BL', 0.0)]

GENERATOR = ([I('BUS_NUMBER'), S('GEN_ID'), R('PG', 0.0), R('QG', 0.0),
              R('QMAX', 9999.0), R('QMIN', -9999.0), R('VS', 1.0), I('IREG', 0), I('NREG', 0),
              R('MBASE', 100.0), R('ZR', 0.0), R('ZX', 1.0), R('RT', 0.0), R('XT', 0.0),
              R('GTAP', 1.0), I('STAT', 1), R('RMPCT', 100.0), R('PMAX', 9999.0),
              R('PMIN', 0.0), I('BASELD', 0)]
             + _owners() + [I('WMOD', None), R('WPF', None)])

BRANCH = ([I('FROM_BUS'), I('TO_BUS'), S('CKT', '1'), R('R', 0.0), R('X', 0.001), R('B', 0.0),
           S('NAME', '')]
          + [R(f'RATE{k}', 0.0) for k in range(1, 13)]
          + [R('GI', 0.0), R('BI', 0.0), R('GJ', 0.0), R('BJ', 0.0),
             I('STAT', 1), I('MET', 1), R('LEN', 0.0)]
          + _owners())

SYSTEM_SWITCHING_DEVICE = ([I('FROM_BUS'), I('TO_BUS'), S('CKT', '1'), R('X', 0.0001)]
                           + [R(f'RATE{k}', 0.0) for k in range(1, 13)]
                           + [I('STAT', 1), I('NSTAT', 1), I('MET', 1), I('STYPE', 2),
                              S('NAME', '')])

# Transformer line 1, the same for two- and three-winding transformers
_XFMR_LINE1 = ([I('FROM_BUS'), I('TO_BUS'), I('K', 0), S('CKT', '1'),
                I('CW', 1), I('CZ', 1), I('CM', 1), R('MAG1', 0.0), R('MAG2', 0.0),
                I('NMETR', 2), S('NAME', ''), I('STAT', 1)]
               + _owners() + [S('VECGRP', ''), I('ZCOD', None)])
XFMR_LINE1_LEN = len(_XFMR_LINE1)        # 22
WINDING_LEN = len(_winding(1))           # 27

TWO_WINDING = ([C('TRANSFORMER_TYPE', '2W')] + _XFMR_LINE1
               + [R('R1_2', 0.0), R('X1_2', 0.001), R('SBASE1_2', 100.0)]
               + _winding(1)
               + [R('WINDV2', 1.0), R('NOMV2', 0.0)])

THREE_WINDING = ([C('TRANSFORMER_TYPE', '3W')]
                 + [I('FROM_BUS'), I('TO_BUS'), I('K')] + _XFMR_LINE1[3:]
                 + [R('R1_2', 0.0), R('X1_2', 0.001), R('SBASE1_2', 100.0),
                    R('R2_3', 0.0), R('X2_3', 0.001), R('SBASE2_3', 100.0),
                    R('R3_1', 0.0), R('X3_1', 0.001), R('SBASE3_1', 100.0),
                    R('VMSTAR', 1.0), R('ANSTAR', 0.0)]
                 + _winding(1) + _winding(2) + _winding(3))

AREA = [I('AREA_NUMBER'), I('ISWING', 0), R('PDES', 0.0), R('PTOL', 10.0),
        S('ARNAME', default_from_first='AREA_{}')]


def _dc_converter(t: str):
    """Two-terminal DC converter line (v35); t is 'R' (rectifier) or 'I' (inverter)."""
    return [I(f'IP{t}', 0), I(f'NB{t}', 0), R(f'ANMX{t}', 0.0), R(f'ANMN{t}', 0.0),
            R(f'RC{t}', 0.0), R(f'XC{t}', 0.0), R(f'EBAS{t}', 0.0), R(f'TR{t}', 0.0),
            R(f'TAP{t}', 1.0), R(f'TMX{t}', 1.0), R(f'TMN{t}', 1.0), R(f'STP{t}', 0.01),
            I(f'IC{t}', 0), I(f'ND{t}', 0), I(f'IF{t}', 0), I(f'IT{t}', 0),
            S(f'ID{t}', '1'), R(f'XCAP{t}', 0.0)]


_DC_LINE1 = [S('NAME'), I('MDC', 1), R('RDC', 0.0), R('SETVL', 0.0), R('VSCHD', 0.0),
             R('VCMOD', 0.0), R('RCOMP', 0.0), R('DELTI', 0.1), S('METER', ''),
             R('DCVMIN', 0.0), I('CCCITMX', 20), R('CCCACC', 1.0)]
DC_LINE1_LEN = len(_DC_LINE1)            # 12
DC_CONVERTER_LEN = len(_dc_converter('R'))  # 18
TWO_TERMINAL_DC = _DC_LINE1 + _dc_converter('R') + _dc_converter('I')


def _vsc_converter(t: str):
    """VSC converter line (v35); t is 'I' (first) or 'J' (second converter)."""
    return [I(f'{t}BUS', 0), I(f'{t}TYPE', 0), I(f'{t}MODE', 0), R(f'{t}DCSET', 0.0),
            R(f'{t}ACSET', 0.0), R(f'{t}ALOSS', 0.0), R(f'{t}BLOSS', 0.0),
            R(f'{t}MINLOSS', 0.0), R(f'{t}SMAX', 1.0), R(f'{t}IMAX', 0.0),
            R(f'{t}PWF', 0.5), R(f'{t}MAXQ', 0.01), R(f'{t}MINQ', 0),
            I(f'{t}VSREG', 0), I(f'{t}NREG', 0), R(f'{t}RMPCT', 100.0)]


_VSC_LINE1 = [S('NAME'), I('MDC', 1), R('RDC', 0.0)] + _owners()
VSC_LINE1_LEN = len(_VSC_LINE1)          # 11
VSC_CONVERTER_LEN = len(_vsc_converter('I'))  # 16
VSC_DC = _VSC_LINE1 + _vsc_converter('I') + _vsc_converter('J')

MULTI_SECTION_LINE = ([I('FROM_BUS'), I('TO_BUS'), S('ID', '&1'), I('MET', 1)]
                      + [I(f'DUM{k}', None) for k in range(1, 10)])

ZONE = [I('ZONE_NUMBER'), S('ZONAME', default_from_first='ZONE_{}')]

OWNER = [I('OWNER_NUMBER'), S('OWNAME', default_from_first='OWNER_{}')]

FACTS = [S('FACTS_NAME'), I('I'), I('J', 0), I('MODE', 1), R('PDES', 0.0), R('QDES', 0.0),
         R('VSET', 1.0), R('SHMX', 9999.0), R('TRMX', 1.1), R('VTMN', 0.9), R('VTMX', 1.1),
         R('VSMX', 1.0), R('IMAX', 0.0), R('LINX', 0.05), R('RMPCT', 1.0), I('OWNER', 1),
         R('SET1', 1.0), R('SET2', 1.0), R('VSREF', 0.0), I('FCREG', 0), I('NREG', 0),
         S('MNAME', '')]

SWITCHED_SHUNT = ([I('BUS_NUMBER'), S('ID', ''), I('MODSW', 1), I('ADJM', 0), I('STAT', 1),
                   R('VSWHI', 1.0), R('VSWLO', 1.0), I('SWREG', None), I('NREG', 0),
                   R('RMPCT', 100.0), S('RMIDNT', ''), R('BINIT', 0.0),
                   I('S1', 0), I('N1', 0), R('B1', 0.0)]
                  + [f for k in range(2, 9)
                     for f in (I(f'S{k}', None), I(f'N{k}', None), R(f'B{k}', None))])

INTER_AREA_TRANSFER = [I('ARFROM'), I('ARTO'), S('TRID', '1'), R('PTRAN', 0.0)]

INDUCTION_MACHINE = [I('BUS_NUMBER'), S('MACHINE_ID', '1'), I('STAT', 1), I('SCODE', 1),
                     I('DCODE', 2), I('AREA', None), I('ZONE', None), I('OWNER', None),
                     I('TCODE', 1), I('BCODE', 1), R('MBASE', None), R('RATEKV', 0.0),
                     I('PCODE', 1), R('PSET', 0.0), R('H', 1.0), R('A', 1.0), R('B', 1.0),
                     R('D', 1.0), R('E', 1.0), R('RA', 0.0), R('XA', 0.0), R('XM', 2.5),
                     R('R1', 999.0), R('X1', 999.0), R('R2', 999.0), R('X2', 999.0),
                     R('X3', 0.0), R('E1', 1.0), R('SE1', 0.0), R('E2', 1.2), R('SE2', 0.0),
                     R('IA1', 0.0), R('IA2', 0.0), R('XAMULT', 1.0)]

# Multi-terminal DC: one header line, then NCONV converter lines, NDCBS DC bus
# lines and NDCLN DC link lines. Each goes to its own tab; NAME links them.
MTDC_LINE = [S('NAME'), I('NCONV'), I('NDCBS'), I('NDCLN'), I('MDC', 0), I('VCONV', 0),
             R('VCMOD', 0.0), I('VCONVN', 0)]
MTDC_CONVERTER = [I('IB'), I('N'), R('ANGMX'), R('ANGMN'), R('RC'), R('XC'), R('EBAS'),
                  R('TR', 1.0), R('TAP', 1.0), R('TPMX', 1.5), R('TPMN', 0.51),
                  R('TSTP', 0.00625), R('SETVL'), R('DCPF', 1.0), R('MARG', 0.0),
                  I('CNVCOD', 1)]
MTDC_DC_BUS = [I('IDC'), I('IB', 0), I('AREA', 1), I('ZONE', 1), S('DCNAME', ''),
               I('IDC2', 0), R('RGRND', 0.0), I('OWNER', 1)]
MTDC_DC_LINK = [I('IDC'), I('JDC'), S('DCCKT', '1'), I('MET', 1), R('RDC'), R('LDC', 0.0)]

# Substation data (v35). A substation record is followed by blocks of node,
# switching device and equipment terminal records, each block ended by "0".
SUBSTATION = [I('IS'), S('NAME', ''), R('LATI', 0.0), R('LONG', 0.0), R('SRG', 0.0)]
SUBSTATION_NODE = [I('NI'), S('NAME', ''), I('I'), I('STATUS', 1), R('VM', 1.0), R('VA', 0.0)]
SUBSTATION_SWITCHING_DEVICE = [I('NI'), I('NJ'), S('CKT', '1'), S('NAME', ''), I('TYPE', 2),
                               I('STATUS', 1), I('NSTAT', 1), R('X', 0.0001), REST('RATINGS')]
# Equipment connections: I, NI, TYPE, then the equipment's other buses and its
# ID, which depend on TYPE. They are kept as one text column.
SUBSTATION_CONNECTION = [I('I'), I('NI'), S('TYPE', ''), REST('EQUIPMENT')]

# GNE devices: line 1 'NAME','MODEL',NTERM,BUS1..BUSNTERM,NREAL,NINTG,NCHAR;
# line 2 STATUS,OWNER,NMETR; then the REAL, INTG and CHAR values, up to 10 per
# line. Buses and values are kept as text lists.
GNE = [S('NAME'), S('MODEL', ''), I('NTERM', 1), S('BUSES', ''), I('NREAL', 0), I('NINTG', 0),
       I('NCHAR', 0), I('STATUS', 1), I('OWNER', None), I('NMETR', None),
       S('REAL', ''), S('INTG', ''), S('CHAR', '')]

SYSTEM_WIDE_COLUMNS = ['RECORD', 'PARAMETER', 'VALUE', 'DESCRIPTION']


def columns(layout: List[Field]) -> List[str]:
    return [f.name for f in layout]


# Field order of records whose layout differs from v35, as the v35 position
# each field should go to. Each list has one entry per v35 field: the index of
# that field in the older record, or None when the older version has no such
# field (the parser then uses its default). Records not listed here have the
# same layout as v35.
FIELD_ORDER_FROM: Dict[int, Dict[str, list]] = {
    33: {
        # v33 has no NREG and no BASLOD
        'GENERATOR': [0, 1, 2, 3, 4, 5, 6, 7, None, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
                      None, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27],
        # v33 has no DGENP, DGENQ, DGENM or LOADTYPE
        'LOAD': list(range(14)) + [None] * 4,
        # v33 has no NAME and 3 ratings (RATEA-C) instead of 12
        'BRANCH': [0, 1, 2, 3, 4, 5, None, 6, 7, 8] + [None] * 9 + list(range(9, 24)),
        # v33 has no ID, no NREG and no block status; SWREM is v35's SWREG
        'SWITCHED_SHUNT': [0, None, 1, 2, 3, 4, 5, 6, None, 7, 8, 9]
                          + [x for k in range(8) for x in (None, 10 + 2 * k, 11 + 2 * k)],
        # transformer winding lines: 3 ratings instead of 12, and no NOD
        'WINDING': [0, 1, 2, 3, 4, 5] + [None] * 9 + [6, 7, None] + list(range(8, 17)),
        # two-terminal DC converter lines: no NDR/NDI
        'DC_CONVERTER': list(range(13)) + [None, 13, 14, 15, 16],
        # VSC converter lines: REMOT is v35's VSREG; no NREG
        'VSC_CONVERTER': list(range(14)) + [None, 14],
        # FACTS: REMOT is v35's FCREG; no NREG
        'FACTS': list(range(20)) + [None, 20],
    },
    34: {
        # v34 has no BASLOD, and NREG is the last field instead of after IREG
        'GENERATOR': [0, 1, 2, 3, 4, 5, 6, 7, 28, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17,
                      None, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27],
        # v34 has no LOADTYPE
        'LOAD': list(range(17)) + [None],
        # v34 has no ID and no block status S1..S8; NREG is the last field
        'SWITCHED_SHUNT': [0, None, 1, 2, 3, 4, 5, 6, 26, 7, 8, 9]
                          + [x for k in range(8) for x in (None, 10 + 2 * k, 11 + 2 * k)],
        # transformer winding lines: NOD is the last field instead of after CONT
        'WINDING': list(range(17)) + [26] + list(range(17, 26)),
        # two-terminal DC converter lines: NDR/NDI is the last field instead of after ICR/ICI
        'DC_CONVERTER': list(range(13)) + [17, 13, 14, 15, 16],
        # VSC converter lines: NREG comes after RMPCT instead of before it
        'VSC_CONVERTER': list(range(14)) + [15, 14],
        # FACTS: MNAME comes before NREG instead of after it
        'FACTS': list(range(20)) + [21, 20],
    },
}


def to_v35_order(parts: List[str], record: str, version: int) -> List[str]:
    """Rearrange an older-version record's fields into v35 order (no-op for v35)."""
    order = FIELD_ORDER_FROM.get(version, {}).get(record)
    if order is None:
        return parts
    out = [parts[i] if i is not None and i < len(parts) else '' for i in order]
    if record == 'SWITCHED_SHUNT':
        # Before v35 every block is in service; mark the blocks that are present.
        for k in range(8):
            if out[13 + 3 * k].strip():
                out[12 + 3 * k] = '1'
    return out


# Data sections after the header, in file order, per PSS/E version. Each ends
# with a "0" record. SYSTEM-WIDE DATA, system switching devices and substations
# exist from v34 on.
SECTION_ORDER = {
    33: ['BUS DATA', 'LOAD DATA', 'FIXED SHUNT DATA', 'GENERATOR DATA',
         'BRANCH DATA', 'TRANSFORMER DATA', 'AREA DATA',
         'TWO-TERMINAL DC DATA', 'VSC DC LINE DATA', 'IMPEDANCE CORRECTION DATA',
         'MULTI-TERMINAL DC DATA', 'MULTI-SECTION LINE DATA', 'ZONE DATA',
         'INTER-AREA TRANSFER DATA', 'OWNER DATA', 'FACTS DEVICE DATA',
         'SWITCHED SHUNT DATA', 'GNE DATA', 'INDUCTION MACHINE DATA'],
    34: ['SYSTEM-WIDE DATA', 'BUS DATA', 'LOAD DATA', 'FIXED SHUNT DATA',
         'GENERATOR DATA', 'BRANCH DATA', 'SYSTEM SWITCHING DEVICE DATA',
         'TRANSFORMER DATA', 'AREA DATA', 'TWO-TERMINAL DC DATA',
         'VSC DC LINE DATA', 'IMPEDANCE CORRECTION DATA', 'MULTI-TERMINAL DC DATA',
         'MULTI-SECTION LINE DATA', 'ZONE DATA', 'INTER-AREA TRANSFER DATA',
         'OWNER DATA', 'FACTS DEVICE DATA', 'SWITCHED SHUNT DATA', 'GNE DATA',
         'INDUCTION MACHINE DATA', 'SUBSTATION DATA'],
}
SECTION_ORDER[35] = SECTION_ORDER[34]
SUPPORTED_VERSIONS = sorted(SECTION_ORDER)

# Other names that exporters such as PowerWorld write in terminator comments
SECTION_ALIASES = {
    'VOLTAGE SOURCE CONVERTER DATA': 'VSC DC LINE DATA',
    'FACTS CONTROL DEVICE DATA': 'FACTS DEVICE DATA',
    'GNE DEVICE DATA': 'GNE DATA',
    'GEN DEVICE DATA': 'GNE DATA',
}
