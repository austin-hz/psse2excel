"""
Edge cases on small, made-up raw files: field splitting, broken multi-line
records, encodings and section endings.

Run from the project folder:
    python tests/test_edge_cases.py
or with pytest:
    pytest tests/test_edge_cases.py
"""

import logging
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from psse2excel import PSSEParser  # noqa: E402
from psse2excel.layouts import SECTION_ORDER  # noqa: E402
from psse2excel.tokenizer import split_all, split_fields  # noqa: E402

logging.getLogger('psse2excel').setLevel(logging.CRITICAL)

WINDING = "1.0, 345.0, 0.0, 100.0, 110.0, 120.0, 0,0,0,0,0,0,0,0,0, 1, 0, 0, 1.1, 0.9, 1.1, 0.9, 33, 0, 0.0, 0.0, 0.0"
XFMR_2W = ["{i}, {j}, 0, '1', 1, 1, 1, 0.0, 0.0, 2, 'T{i}{j}', 1, 1, 1.0",
           "0.001, 0.05, 100.0", WINDING, "1.0, 138.0"]
XFMR_3W = ["{i}, {j}, {k}, '1', 1, 1, 1, 0.0, 0.0, 2, 'T3', 1, 1, 1.0",
           "0.001, 0.05, 100.0, 0.001, 0.05, 100.0, 0.001, 0.05, 100.0, 1.0, 0.0",
           WINDING, WINDING, WINDING]
DC_LINE1 = "'{name}', 1, 5.0, 100.0, 500.0, 0.0, 0.0, 0.1, 'I', 0.0, 20, 1.0"
DC_CONV = "{bus}, 4, 20.0, 10.0, 0.1, 5.0, 345.0, 0.5, 1.0, 1.1, 0.9, 0.0125, 0, 0, 0, 0, '1', 0.0"


def make_raw(sections, version=35, title=("TEST CASE", "SECOND TITLE LINE")):
    """Build a raw file. sections maps section name -> list of record lines."""
    out = [f"0, 100.00, {version}, 0, 1, 60.00 / PSS(R)E-{version}", *title]
    for name in SECTION_ORDER[version]:
        out += sections.get(name, [])
        out.append(f"0 / END OF {name}")
    out.append("Q")
    return "\n".join(out) + "\n"


def parse(text, encoding='utf-8'):
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, 'case.raw')
        with open(path, 'w', encoding=encoding, newline='\n') as f:
            f.write(text)
        return PSSEParser(validate_data=False, progress_bar=False).parse_file(path)


BUSES = ["1, 'ONE', 345.0, 3", "2, 'TWO', 345.0, 1", "3, 'THREE', 138.0, 1",
         "4, 'FOUR', 138.0, 1", "5, 'FIVE', 138.0, 1"]


# ---------------------------------------------------------------- field splitting

def test_comment_after_slash_is_dropped():
    assert split_fields("1, 'A', 2 / a comment, with commas") == ['1', 'A', '2']


def test_slash_inside_quotes_is_kept():
    assert split_fields("1, 'A/B', 2 / comment") == ['1', 'A/B', '2']


def test_quoted_name_with_comma():
    assert split_fields("1, 'SMITH, JOHN', 345.0") == ['1', 'SMITH, JOHN', '345.0']


def test_double_quotes():
    assert split_fields('1, "NAME 1", 2') == ['1', 'NAME 1', '2']


def test_apostrophe_inside_name():
    assert split_fields("1, 'O'Brien Wi~1', 345.0") == ['1', "O'Brien Wi~1", '345.0']


def test_blank_delimiters():
    assert split_fields("1 'ONE TWO' 345.0   3") == ['1', 'ONE TWO', '345.0', '3']


def test_empty_fields_are_kept():
    assert split_fields("1,,'',3") == ['1', '', '', '3']


def test_unclosed_quote_does_not_swallow_the_next_line():
    lines = ["1, 'OPEN, 2", "3, 'B', 4"]
    assert split_all(lines) == [split_fields(lines[0]), ['3', 'B', '4']]


def test_split_all_matches_split_fields():
    lines = ["1, 'A', 2", "1 2 3", "'X', \"Y Z\"", "", "4, 'O'Hara', 5 / c", "  7 ,  8  "]
    assert split_all(lines) == [split_fields(line) for line in lines]


# ---------------------------------------------------------------- sections and header

def test_title_and_version():
    data = parse(make_raw({'BUS DATA': BUSES}))
    assert data.title == 'TEST CASE | SECOND TITLE LINE'
    assert data.version == 35
    assert len(data.buses) == 5


def test_bare_zero_terminators_and_names_with_commas():
    text = make_raw({'BUS DATA': ["1, 'A, B', 345.0, 3"], 'LOAD DATA': ["1, '1', 1, 1, 1, 10.0, 5.0"]})
    text = text.replace("0 / END OF BUS DATA", "0")
    data = parse(text)
    assert data.buses['BUS_NAME'].tolist() == ['A, B']
    assert data.loads['PL'].tolist() == [10.0]


def test_file_that_ends_early():
    text = make_raw({'BUS DATA': BUSES})
    text = text[:text.index('0 / END OF LOAD DATA')]
    data = parse(text)
    assert len(data.buses) == 5 and data.loads.empty


def test_cp1252_file():
    data = parse(make_raw({'BUS DATA': ["1, 'GUIÑAZU', 345.0, 3"]}), encoding='cp1252')
    assert data.buses['BUS_NAME'].tolist() == ['GUIÑAZU']


def test_bad_record_is_skipped_and_counted():
    data = parse(make_raw({'BUS DATA': BUSES + ["6, 'SIX', not_a_number, 1"]}))
    assert len(data.buses) == 5
    assert data.skipped == {'buses': 1}


# ---------------------------------------------------------------- multi-line records

def _xfmr(template, **buses):
    return [line.format(**buses) for line in template]


def test_transformers_parse():
    lines = _xfmr(XFMR_2W, i=1, j=2) + _xfmr(XFMR_3W, i=1, j=2, k=3) + _xfmr(XFMR_2W, i=3, j=4)
    data = parse(make_raw({'BUS DATA': BUSES, 'TRANSFORMER DATA': lines}))
    assert data.two_winding_transformers['TO_BUS'].tolist() == [2, 4]
    assert data.three_winding_transformers['K'].tolist() == [3]
    assert data.two_winding_transformers['RATE1-2'].tolist() == [110.0, 110.0]


def test_transformer_with_a_missing_line_does_not_shift_the_rest():
    broken_3w = _xfmr(XFMR_3W, i=1, j=2, k=3)[:-1]           # winding 3 line missing
    lines = (_xfmr(XFMR_2W, i=1, j=2) + broken_3w
             + _xfmr(XFMR_2W, i=3, j=4) + _xfmr(XFMR_2W, i=4, j=5))
    data = parse(make_raw({'BUS DATA': BUSES, 'TRANSFORMER DATA': lines}))
    assert data.two_winding_transformers['FROM_BUS'].tolist() == [1, 3, 4]
    assert data.three_winding_transformers.empty
    assert data.skipped == {'three_winding_transformers': 1}


def test_stray_line_between_transformers():
    lines = _xfmr(XFMR_2W, i=1, j=2) + ["9.99, 1.0, 7.0"] + _xfmr(XFMR_2W, i=3, j=4)
    data = parse(make_raw({'BUS DATA': BUSES, 'TRANSFORMER DATA': lines}))
    assert data.two_winding_transformers['FROM_BUS'].tolist() == [1, 3]
    assert data.lines_not_read == 1


def test_dc_line_with_a_missing_converter_does_not_shift_the_rest():
    lines = ([DC_LINE1.format(name='DC1'), DC_CONV.format(bus=1), DC_CONV.format(bus=2)]
             + [DC_LINE1.format(name='DC2'), DC_CONV.format(bus=1)]
             + [DC_LINE1.format(name='DC3'), DC_CONV.format(bus=3), DC_CONV.format(bus=4)])
    data = parse(make_raw({'BUS DATA': BUSES, 'TWO-TERMINAL DC DATA': lines}))
    assert data.two_terminal_dc_lines['NAME'].tolist() == ['DC1', 'DC3']
    assert data.two_terminal_dc_lines['IPI'].tolist() == [2, 4]
    assert data.skipped == {'two_terminal_dc_lines': 1}


def test_multi_terminal_dc():
    lines = ["'MT1', 2, 2, 1, 1, 1, 0.0, 0",
             "1, 2, 21.0, 18.0, 0.19, 7.68, 345.0, 0.54, 1.0, 1.1, 0.9, 0.0125, 450.0, 0.0, 0.0, 1",
             "2, 2, 21.0, 17.0, 0.17, 7.77, 315.0, 0.60, 0.9, 1.1, 0.8, 0.0125, -450.0, 0.0, 0.0, 2",
             "1, 1, 1, 1, 'DCB1', 0, 0.0, 1",
             "2, 2, 1, 1, 'DCB2', 0, 0.0, 1",
             "1, 2, '1', 1, 11.19, 812.6"]
    data = parse(make_raw({'BUS DATA': BUSES, 'MULTI-TERMINAL DC DATA': lines}))
    assert data.mtdc_lines['NCONV'].tolist() == [2]
    assert data.mtdc_converters['SETVL'].tolist() == [450.0, -450.0]
    assert data.mtdc_dc_buses['DCNAME'].tolist() == ['DCB1', 'DCB2']
    assert data.mtdc_dc_links['RDC'].tolist() == [11.19]


def test_gne_device():
    lines = ["'GNE1', 'MODEL1', 2, 1, 2, 3, 1, 0", "1, 1, 1",
             "1.5, 2.5, 3.5", "7"]
    data = parse(make_raw({'BUS DATA': BUSES, 'GNE DATA': lines}))
    assert data.gne_devices['BUSES'].tolist() == ['1, 2']
    assert data.gne_devices['REAL'].tolist() == ['1.5, 2.5, 3.5']
    assert data.gne_devices['INTG'].tolist() == ['7']


def test_substations_with_their_own_zero_records():
    lines = ["1, 'SUB1', 40.0, -80.0, 0.1",
             "1, 'NODE1', 1, 1, 1.0, 0.0", "2, 'NODE2', 2, 1, 1.0, 0.0",
             "0 / END OF SUBSTATION NODE DATA",
             "1, 2, '1', 'BRK1', 2, 1, 1, 0.0001, 100.0, 100.0, 100.0",
             "0 / END OF SUBSTATION SWITCHING DEVICE DATA",
             "1, 1, 'L', '1'",
             "0 / END OF SUBSTATION EQUIPMENT CONNECTION DATA",
             "2, 'SUB2', 41.0, -81.0, 0.1",
             "0", "0", "0"]
    data = parse(make_raw({'BUS DATA': BUSES, 'SUBSTATION DATA': lines}))
    assert data.substations['NAME'].tolist() == ['SUB1', 'SUB2']
    assert data.substation_nodes['NAME'].tolist() == ['NODE1', 'NODE2']
    assert data.substation_switching_devices['NAME'].tolist() == ['BRK1']
    assert data.substation_connections['EQUIPMENT'].tolist() == ['1']


def test_v34_generator_fields_move_to_v35_places():
    # v34 has NREG as the last field and no BASLOD
    gen = ("1, '1', 100.0, 0.0, 50.0, -50.0, 1.0, 0, 120.0, 0.0, 1.0, 0.0, 0.0, 1.0, 1, 100.0, "
           "200.0, 0.0, 1, 1.0, 0, 0.0, 0, 0.0, 0, 0.0, 0, 1.0, 7")
    data = parse(make_raw({'BUS DATA': BUSES, 'GENERATOR DATA': [gen]}, version=34))
    row = data.generators.iloc[0]
    assert (row['NREG'], row['MBASE'], row['PMAX'], row['BASELD']) == (7, 120.0, 200.0, 0)


def test_v33_and_v34_give_v35_columns():
    v33 = parse(make_raw({'BUS DATA': BUSES, 'LOAD DATA': ["1, '1', 1, 1, 1, 10.0, 5.0"]},
                         version=33))
    v35 = parse(make_raw({'BUS DATA': BUSES, 'LOAD DATA': ["1, '1', 1, 1, 1, 10.0, 5.0"]}))
    assert list(v33.loads.columns) == list(v35.loads.columns)
    assert v33.loads['PL'].tolist() == [10.0]


def test_excel_has_every_tab_even_when_empty():
    import openpyxl
    from psse2excel import TABS
    data = parse(make_raw({'BUS DATA': BUSES}))
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, 'out.xlsx')
        data.to_excel(out)
        wb = openpyxl.load_workbook(out, read_only=True)
        assert wb.sheetnames == ['Summary', 'System_Info'] + [t.sheet for t in TABS]
        header = next(wb['Loads'].iter_rows(values_only=True))
        assert header[:2] == ('BUS_NUMBER', 'LOAD_ID')
        wb.close()


def test_csv_has_one_file_per_tab_with_the_same_values():
    import csv
    from psse2excel import TABS
    data = parse(make_raw({'BUS DATA': ["1, 'GUIÑAZU', 345.0, 3", "2, 'A, B', 138.0, 1"],
                           'LOAD DATA': ["1, '1', 1, 1, 1, 10.5, 5.0"]}))
    data.loads.loc[0, 'QL'] = float('inf')
    with tempfile.TemporaryDirectory() as tmp:
        written = data.to_csv(os.path.join(tmp, 'case_csv'))
        names = [os.path.basename(p) for p in written]
        assert names == ['Summary.csv', 'System_Info.csv'] + [t.sheet + '.csv' for t in TABS]

        def rows(name):
            with open(os.path.join(tmp, 'case_csv', name), encoding='utf-8-sig', newline='') as f:
                return list(csv.reader(f))

        buses = rows('Buses.csv')
        assert buses[0][:4] == ['BUS_NUMBER', 'BUS_NAME', 'BASE_KV', 'BUS_TYPE']
        assert buses[1][:4] == ['1', 'GUIÑAZU', '345.0', '3']
        assert buses[2][1] == 'A, B'
        loads = rows('Loads.csv')
        ql = loads[0].index('QL')
        assert loads[1][loads[0].index('PL')] == '10.5' and loads[1][ql] == 'inf'
        assert rows('Generators.csv') == [list(data.generators.columns)]  # header only
        summary = rows('Summary.csv')
        assert ['buses', 'Buses', '2', '0'] == summary[[r[0] for r in summary].index('buses')][:4]


def test_unsupported_version_is_shown_as_read_as():
    text = make_raw({'BUS DATA': BUSES})
    assert parse(text).version_text() == 35
    data = parse(text.replace('0, 100.00, 35,', '0, 100.00, 36,', 1))
    assert data.version == 35 and data.file_version == 36
    info = dict(data.system_info_table()[1])
    assert info['PSS/E Version'] == '36 (read as 35)'
    assert len(data.buses) == 5


if __name__ == '__main__':
    tests = [(name, fn) for name, fn in sorted(globals().items()) if name.startswith('test_')]
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print(f"PASS  {name}")
        except Exception as e:  # report every failure, not just the first
            failed += 1
            print(f"FAIL  {name}: {type(e).__name__}: {e}")
    print(f"\n{len(tests) - failed} passed, {failed} failed")
    sys.exit(1 if failed else 0)
