"""Command line: python -m psse2excel case.raw [-d FOLDER] [-o case.xlsx] [--csv [FOLDER]] [--excel]"""

import argparse
import logging
import os
import sys
import time
from typing import List, Optional

from . import __version__
from .parser import PSSEParser


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        prog='psse2excel',
        description='Convert a PSS/E raw file (v33, v34 or v35) to an Excel workbook, '
                    'one tab per equipment type, or to CSV files, one per equipment type.')
    ap.add_argument('raw_files', nargs='+', help='raw file(s) to convert')
    ap.add_argument('-d', '--output-dir', default='output', metavar='FOLDER',
                    help='folder for the workbook and CSV folders, created if needed '
                         '(default: %(default)s, in the folder you run the command from)')
    ap.add_argument('-o', '--output',
                    help='output workbook (only with one raw file; default: the raw file '
                         'name with .xlsx, in the output folder)')
    ap.add_argument('--csv', nargs='?', const='', metavar='FOLDER',
                    help='write one CSV file per tab instead of the workbook, into FOLDER '
                         '(only with one raw file; default: the raw file name with _csv, '
                         'in the output folder)')
    ap.add_argument('--excel', action='store_true',
                    help='with --csv, also write the workbook (implied by -o)')
    ap.add_argument('--no-summary', action='store_true', help='leave out the Summary and '
                    'System_Info tabs (or files)')
    ap.add_argument('--no-validate', action='store_true', help='skip the data checks')
    ap.add_argument('--log', default='psse_parser.log', help='log file (default: %(default)s)')
    ap.add_argument('-q', '--quiet', action='store_true', help='only show warnings and errors')
    ap.add_argument('--version', action='version', version=f'psse2excel {__version__}')
    args = ap.parse_args(argv)
    if args.output and len(args.raw_files) > 1:
        ap.error('-o/--output can only be used with one raw file')
    if args.csv and len(args.raw_files) > 1:
        ap.error('a --csv folder can only be given with one raw file')
    stems = [os.path.splitext(os.path.basename(r))[0].lower() for r in args.raw_files]
    if len(set(stems)) < len(stems):
        ap.error('two raw files have the same name, so their outputs would overwrite each other; '
                 'run them separately with different -d folders')
    write_csv = args.csv is not None
    write_excel = not write_csv or args.excel or bool(args.output)

    handlers = [logging.StreamHandler()]
    if args.log:
        handlers.append(logging.FileHandler(args.log))
    logging.basicConfig(level=logging.WARNING if args.quiet else logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s', handlers=handlers)
    log = logging.getLogger('psse2excel')

    failed = 0
    for raw in args.raw_files:
        stem = os.path.splitext(os.path.basename(raw))[0]
        out = args.output or os.path.join(args.output_dir, stem + '.xlsx')
        csv_dir = args.csv or os.path.join(args.output_dir, stem + '_csv')
        try:
            start = time.time()
            parser = PSSEParser(validate_data=not args.no_validate, progress_bar=not args.quiet)
            data = parser.parse_file(raw)
            parsed = time.time()
            if write_excel:
                os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
                data.to_excel(out, include_summary=not args.no_summary)
                log.info(f"{raw}: parsed in {parsed - start:.1f} s, Excel written in "
                         f"{time.time() - parsed:.1f} s -> {out}")
            if write_csv:
                t = time.time()
                data.to_csv(csv_dir, include_summary=not args.no_summary)
                log.info(f"{raw}: parsed in {parsed - start:.1f} s, CSV files written in "
                         f"{time.time() - t:.1f} s -> {csv_dir}")
        except Exception as e:
            log.error(f"{raw}: {e}")
            failed += 1
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
