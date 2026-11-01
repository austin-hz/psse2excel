"""
Regression test: parse every .raw file in testcases/ and compare the
resulting workbook with baseline/<same name>.xlsx, tab by tab and cell by cell.
The CSV files (one per tab) are compared with the same baseline workbook.

The baselines are the output of the parser before refactoring. Any difference
is a failure unless we have agreed it is a bug fix and updated the baseline.

Run from the project folder:
    python tests/test_regression.py            # all files, in parallel
    python tests/test_regression.py --quick    # ACTIVSg10k (v33), Midwest24k (v35), and sadi (v34) if present
    python tests/test_regression.py 70k        # only files whose name contains "70k"
    python tests/test_regression.py --keep     # keep the generated workbooks and CSV folders in tests/output/
    python tests/test_regression.py --no-csv   # only check the workbooks
or with pytest (if installed):
    pytest tests/test_regression.py

tests/test_edge_cases.py has the small hand-made cases (broken records,
quotes, comments, encodings). It runs in seconds.
"""

import argparse
import itertools
import os
import shutil
import sys
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "testcases"
BASELINE_DIR = ROOT / "baseline"
QUICK_CASES = ["ACTIVSg10k", "sadi", "Midwest24k_20220923"]  # one each of v33, v34, v35
MAX_REPORTED_DIFFS = 20  # per tab
# Rows whose value changes on every run, as (tab, first cell): not compared
NOT_COMPARED = {("System_Info", "Run Date")}


def build_outputs(raw_path: Path, out_path: Path, csv_dir=None) -> None:
    """Run the parser exactly the way main() does and write the workbook.

    This is the only place that knows the parser's API. Update it if the
    entry point changes during the refactor.
    """
    sys.path.insert(0, str(ROOT))
    import logging
    import psse2excel

    logging.getLogger('psse2excel').setLevel(logging.ERROR)  # keep the test output readable
    parser = psse2excel.PSSEParser(validate_data=True, progress_bar=False)
    data = parser.parse_file(str(raw_path))
    data.to_excel(str(out_path), include_summary=True)
    if csv_dir is not None:
        data.to_csv(str(csv_dir), include_summary=True)


def compare_workbooks(expected_path: Path, actual_path: Path) -> list:
    """Return a list of human-readable differences (empty list = identical)."""
    problems = []
    exp_wb = openpyxl.load_workbook(expected_path, read_only=True)
    act_wb = openpyxl.load_workbook(actual_path, read_only=True)
    try:
        if exp_wb.sheetnames != act_wb.sheetnames:
            missing = [s for s in exp_wb.sheetnames if s not in act_wb.sheetnames]
            extra = [s for s in act_wb.sheetnames if s not in exp_wb.sheetnames]
            problems.append(f"tab list differs: missing={missing} extra={extra}")
            if not missing and not extra:
                problems.append(f"tab order differs: expected {exp_wb.sheetnames}, got {act_wb.sheetnames}")

        for name in exp_wb.sheetnames:
            if name not in act_wb.sheetnames:
                continue
            problems.extend(_compare_sheet(name, exp_wb[name].iter_rows(values_only=True),
                                           act_wb[name].iter_rows(values_only=True)))
    finally:
        exp_wb.close()
        act_wb.close()
    return problems


def compare_csv_folder(expected_path: Path, csv_dir: Path) -> list:
    """Compare every tab of the baseline workbook with the CSV file of the same name."""
    import csv
    sys.path.insert(0, str(ROOT))
    from psse2excel.csv_export import ENCODING

    problems = []
    exp_wb = openpyxl.load_workbook(expected_path, read_only=True)
    try:
        expected_files = {name + ".csv" for name in exp_wb.sheetnames}
        actual_files = {p.name for p in csv_dir.glob("*.csv")}
        if expected_files != actual_files:
            problems.append(f"CSV file list differs: missing={sorted(expected_files - actual_files)} "
                            f"extra={sorted(actual_files - expected_files)}")
        for name in exp_wb.sheetnames:
            path = csv_dir / (name + ".csv")
            if not path.exists():
                continue
            with open(path, encoding=ENCODING, newline="") as f:
                rows = ([v if v != "" else None for v in row] for row in csv.reader(f))
                problems.extend(_compare_sheet("csv:" + name, exp_wb[name].iter_rows(values_only=True),
                                               rows, same=_same_as_text, key=name))
    finally:
        exp_wb.close()
    return problems


def _compare_sheet(name, exp_rows, act_rows, same=None, key=None) -> list:
    """Compare two sequences of rows. key is the tab name used for NOT_COMPARED."""
    same = same or _same
    key = key or name
    problems = []
    n_diff = 0
    n_exp_rows = n_act_rows = 0
    for row_idx, (exp_row, act_row) in enumerate(itertools.zip_longest(exp_rows, act_rows), start=1):
        if exp_row is not None:
            n_exp_rows += 1
        if act_row is not None:
            n_act_rows += 1
        if exp_row is None or act_row is None:
            continue  # one sheet is longer; reported once as a row-count difference below
        exp_row = _trim(exp_row or ())
        act_row = _trim(act_row or ())
        if exp_row and (key, exp_row[0]) in NOT_COMPARED and act_row[:1] == exp_row[:1]:
            continue
        if row_idx == 1 and list(exp_row) != list(act_row):
            problems.append(f"[{name}] header differs:\n    expected {list(exp_row)}\n    got      {list(act_row)}")
            return problems  # columns don't line up, cell diffs would be noise
        if exp_row == act_row:
            continue
        for col_idx, (e, a) in enumerate(itertools.zip_longest(exp_row, act_row), start=1):
            if not same(e, a):
                n_diff += 1
                if n_diff <= MAX_REPORTED_DIFFS:
                    cell = f"{get_column_letter(col_idx)}{row_idx}"
                    problems.append(f"[{name}] {cell}: expected {e!r}, got {a!r}")
    if n_exp_rows != n_act_rows:
        shorter = min(n_exp_rows, n_act_rows)
        problems.append(f"[{name}] row count: expected {n_exp_rows}, got {n_act_rows} "
                        f"(rows after {shorter} exist in only one workbook)")
    if n_diff > MAX_REPORTED_DIFFS:
        problems.append(f"[{name}] ... {n_diff - MAX_REPORTED_DIFFS} more differing cells ({n_diff} total)")
    return problems


def _trim(row):
    """Drop trailing empty cells so rows of different stored width compare equal."""
    row = tuple(row)
    end = len(row)
    while end and row[end - 1] is None:
        end -= 1
    return row[:end]


def _same(e, a) -> bool:
    # Exact match on purpose: the baseline is our own earlier output.
    # 1 == 1.0 is accepted because Excel itself does not distinguish them.
    if type(e) is bool or type(a) is bool:
        return type(e) is type(a) and e == a
    return e == a


def _same_as_text(e, a) -> bool:
    """e is a workbook cell, a is the CSV text (None for an empty cell)."""
    if e is None or a is None:
        return e is None and a is None
    if type(e) is bool:
        return a == str(e)
    if isinstance(e, (int, float)):
        try:
            return float(a) == e
        except ValueError:
            return False
    return a == e


def find_cases(filters=None, quick=False) -> list:
    raws = sorted(p for p in RAW_DIR.iterdir() if p.suffix.lower() == ".raw")  # .raw or .RAW
    if quick:
        raws = [r for r in raws if r.stem in QUICK_CASES]
    if filters:
        raws = [r for r in raws if any(f.lower() in r.stem.lower() for f in filters)]
    return raws


def run_case(raw_path: Path, keep_dir=None, check_csv=True):
    """Parse one raw file and compare it with its baseline. Returns (name, problems, seconds)."""
    start = time.time()
    baseline = BASELINE_DIR / (raw_path.stem + ".xlsx")
    if not baseline.exists():
        return raw_path.name, [f"no baseline file {baseline.relative_to(ROOT)}"], 0.0

    # Run from a scratch folder, so that nothing the parser writes lands in the project.
    # The folder lives as long as the process (Windows can't delete a file still open).
    work = _work_dir()
    old_cwd = os.getcwd()
    os.chdir(work)
    try:
        out_dir = Path(keep_dir) if keep_dir else work
        out_dir.mkdir(parents=True, exist_ok=True)
        actual = out_dir / (raw_path.stem + ".xlsx")
        csv_dir = out_dir / (raw_path.stem + "_csv") if check_csv else None
        try:
            build_outputs(raw_path, actual, csv_dir)
        except Exception as exc:  # report, don't crash the other cases
            return raw_path.name, [f"parser raised {type(exc).__name__}: {exc}"], time.time() - start
        problems = compare_workbooks(baseline, actual)
        if csv_dir is not None:
            problems += compare_csv_folder(baseline, csv_dir)
        if not keep_dir:
            actual.unlink(missing_ok=True)
            if csv_dir is not None:
                shutil.rmtree(csv_dir, ignore_errors=True)
    finally:
        os.chdir(old_cwd)
    return raw_path.name, problems, time.time() - start


_WORK_DIR = None


def _work_dir() -> Path:
    global _WORK_DIR
    if _WORK_DIR is None:
        _WORK_DIR = Path(tempfile.mkdtemp(prefix="psse2excel_regression_"))
    return _WORK_DIR


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("filters", nargs="*", help="only run raw files whose name contains one of these")
    ap.add_argument("--quick", action="store_true", help=f"only run {QUICK_CASES}")
    ap.add_argument("-j", "--jobs", type=int, default=min(4, os.cpu_count() or 1),
                    help="files to process in parallel (each large case needs ~1.5 GB of memory)")
    ap.add_argument("--keep", action="store_true",
                    help="keep generated workbooks and CSV folders in tests/output/")
    ap.add_argument("--no-csv", action="store_true", help="don't check the CSV export")
    args = ap.parse_args(argv)

    cases = find_cases(args.filters, args.quick)
    if not cases:
        print("No matching .raw files in", RAW_DIR)
        return 2
    keep_dir = str(ROOT / "tests" / "output") if args.keep else None

    print(f"Comparing {len(cases)} case(s) against {BASELINE_DIR.name}/ ...")
    failed = 0
    t0 = time.time()
    with ProcessPoolExecutor(max_workers=max(1, args.jobs)) as pool:
        futures = [pool.submit(run_case, raw, keep_dir, not args.no_csv) for raw in cases]
        for fut in futures:
            name, problems, secs = fut.result()
            status = "PASS" if not problems else "FAIL"
            print(f"{status}  {name}  ({secs:.0f} s)")
            for p in problems:
                print("      " + p)
            failed += bool(problems)
    print(f"\n{len(cases) - failed} passed, {failed} failed in {time.time() - t0:.0f} s")
    return 1 if failed else 0


# --- pytest entry point (optional) ---------------------------------------------
try:
    import pytest

    @pytest.mark.parametrize("raw_path", find_cases(), ids=lambda p: p.stem)
    def test_matches_baseline(raw_path):
        _, problems, _ = run_case(raw_path)
        assert not problems, "\n".join(problems)
except ImportError:
    pass


if __name__ == "__main__":
    sys.exit(main())
