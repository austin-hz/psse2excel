# psse2excel

Reads PSS/E raw files (versions 33, 34 and 35) in pure Python, without psspy,
and writes an Excel workbook with one tab per equipment type, or CSV files
with one file per equipment type.

PSS®E is a trademark of Siemens. This project is not affiliated with or
endorsed by Siemens.

## Install

From the project folder:

    pip install .                 # or: pip install -e .  to use your working copy directly
    pip install ".[progress]"     # also installs tqdm for a progress bar

This installs the `psse2excel` package and command only. The test cases,
baselines and tests stay in the project folder.

## Use

After installing, from any folder:

    psse2excel C:\cases\mycase.raw

Without installing, run `python -m psse2excel` from the project folder instead
of `psse2excel`. Everything else is the same.

### Where the output is saved

- By default, in a folder named `output` inside the folder you run the command
  from (it is created if needed), with the raw file's name and `.xlsx`. Run
  from `C:\work`, the example above writes `C:\work\output\mycase.xlsx`.
- `-d` picks another output folder: `psse2excel C:\cases\mycase.raw -d D:\reports`
  writes `D:\reports\mycase.xlsx`.
- `-o` picks the file itself: `psse2excel C:\cases\mycase.raw -o D:\reports\report.xlsx`.
- Several raw files at once (`psse2excel a.raw b.raw c.raw`) give one workbook
  each in the output folder. `-o` only works with one raw file, and two raw
  files with the same name can't be run together.
- An existing workbook is overwritten without asking. A workbook that is open
  in Excel can't be written: close it and run again.
- The log goes to `psse_parser.log` in the folder you run the command from;
  `--log FILE` puts it elsewhere.

### CSV files

    psse2excel C:\cases\mycase.raw --csv

- This writes CSV files instead of the workbook, into the folder
  `output\mycase_csv\` (inside the output folder above, so `-d` moves it too).
  `--csv D:\reports\mycase` picks the folder itself (only with one raw file).
- There is one file per tab, named after it: `Buses.csv`, `Generators.csv`,
  ..., plus `Summary.csv` and `System_Info.csv`. Columns and values are the
  same as in the workbook, and a section with no records gets a file with just
  the header row.
- `--csv --excel` writes both. So does `--csv` with `-o`.
- Files are UTF-8 with a byte order mark, so Excel shows accented names
  correctly. Files of the same name are overwritten; other files in the folder
  are left alone.

Other options: `-q` (only warnings and errors), `--no-summary` (no Summary and
System_Info tabs or files), `--no-validate` (skip the data checks). `psse2excel --help`
lists them all.

From Python:

    from psse2excel import PSSEParser
    data = PSSEParser().parse_file("case.raw")
    data.buses            # one pandas DataFrame per tab
    data.to_excel("case.xlsx")
    data.to_csv("case_csv")   # one CSV file per tab

## What it reads

Every section of the raw file, each into its own tab. Records of v33 and v34
files are rearranged into v35 field order, so all versions give the same tabs
and columns; fields an older version doesn't have get their PSS/E defaults.

A file of another version (for example v36) is read as the nearest supported
version, with a warning in the log, and System_Info shows `36 (read as 35)`.
Check such output carefully: fields that moved in the new version would land
in the wrong columns.

A record that can't be read is skipped and logged with its line number, and
the Summary tab counts skipped records per tab. A broken multi-line record
(transformer, DC line) doesn't shift the records after it.

## Limits

- Versions 33, 34 and 35 are supported. Version 36 isn't yet: no v36 file or
  format description was available to build and test it.
- No test file has GNE devices, induction machines or substation data, so
  those sections are only tested on small made-up records
  (`tests/test_edge_cases.py`).
- Files are read whole into memory: a 100,000-bus case needs about 1.5 GB.

Bug reports and raw files that break it are welcome as GitHub issues. Please
don't attach files with CEII or other confidential data.

## Code layout

    psse2excel/tokenizer.py   split a record line into fields (quotes, "/" comments, blanks)
    psse2excel/layouts.py     the fields of every record type, and older-version field orders
    psse2excel/parser.py      read the file, find the sections, group multi-line records
    psse2excel/records.py     turn records into tables using the layouts
    psse2excel/model.py       ParsedData and the list of tabs
    psse2excel/excel.py       write the workbook
    psse2excel/csv_export.py  write the CSV files
    psse2excel/validate.py    data checks (logged)
    psse2excel/cli.py         command line

## Tests

    python tests/test_edge_cases.py      # small made-up cases, seconds
    python tests/test_regression.py      # every file in testcases/ against baseline/, cell by cell,
                                         # for the workbook and for the CSV files
    python tests/test_regression.py --quick

When a change to the output is intended, regenerate the affected baseline
workbooks and commit them with the change.

### Test data

The cases in `testcases/` are synthetic grids from the Texas A&M
[Electric Grid Test Case Repository](https://electricgrids.engr.tamu.edu/),
which provides them free for commercial or non-commercial use. They do not
represent any real grid and contain no CEII.

| File | Case | PSS/E version |
|---|---|---|
| `ACTIVSg10k.RAW` | ACTIVSg10k | 33 |
| `ACTIVSg25k.RAW` | ACTIVSg25k | 33 |
| `ACTIVSg70k.RAW` | ACTIVSg70k | 33 |
| `SyntheticUSA.RAW` | ACTIVSg82k (synthetic USA) | 33 |
| `Midwest24k_20220923.RAW` | Midwest24K, ARPA-E PERFORM program (F. Safdarian et al.) | 35 |

If you use these cases, please cite:

- A. B. Birchfield, T. Xu, K. M. Gegner, K. S. Shetye and T. J. Overbye,
  "Grid Structural Characteristics as Validation Criteria for Synthetic
  Networks," IEEE Transactions on Power Systems, vol. 32, no. 4,
  pp. 3258-3265, July 2017.
- A. B. Birchfield, T. Xu, K. S. Shetye and T. J. Overbye, "Building synthetic
  power transmission networks of many voltage levels, spanning multiple
  areas," 51st Hawaii International Conference on System Sciences (HICSS),
  2018.

## License

MIT; see [LICENSE](LICENSE).
