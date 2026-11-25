#!/usr/bin/env python3
"""
run_pipeline.py - run the whole COPD cleaning pipeline in one command.

    1. Finds every raw extract in evaluation_data/raw/ (the GP code is taken
       from the file name, e.g. WHC_GP01.xlsx -> GP01).
    2. Runs the pre-run checks (check_inputs.py) on all of them and stops,
       with one report listing everything to fix, if any file has a problem.
    3. Flattens each extract to evaluation_data/flattened/<name>.csv.
    4. Runs clean_copd_dataset.py on all of them and saves the CSVs + workbook.
       With --deprivation, LSOA + IMD fields are added to the Patients table
       (see deprivation.py). The deprivation files are checked BEFORE step 1,
       so a missing lookup/IMD file stops the run straight away.

To add a surgery: drop its .xlsx into evaluation_data/raw/ with its GP code in
the file name. No code changes needed.

Usage:
    python run_pipeline.py
    python run_pipeline.py --check-only          # stop after the checks
    python run_pipeline.py --raw-dir some/folder
    python run_pipeline.py --sheet "Sheet1"
    python run_pipeline.py --deprivation        # also add IMD 2025 fields
    python run_pipeline.py --deprivation --imd-version 2025
"""

import argparse
import sys
from pathlib import Path

import check_inputs
import clean_copd_dataset as clean
import deprivation as dep
from flatten_headers import flatten_workbook

BASE_DIR = Path(__file__).resolve().parent
FLATTENED_DIR = BASE_DIR / "evaluation_data" / "flattened"


def output_is_locked(path):
    """True if the output workbook is open in Excel (Windows locks open files)."""
    if not path.exists():
        return False
    try:
        with open(path, "a"):
            return False
    except PermissionError:
        return True


def main():
    parser = argparse.ArgumentParser(description="Run the full COPD cleaning pipeline.")
    parser.add_argument("--raw-dir", default=str(check_inputs.DEFAULT_RAW_DIR),
                        help="Folder of raw .xlsx extracts (default: evaluation_data/raw/)")
    parser.add_argument("--sheet", default=None, help="Sheet name (default: active sheet)")
    parser.add_argument("--check-only", action="store_true", help="Run the checks only")
    parser.add_argument("--deprivation", action="store_true",
                        help="Add LSOA + Index of Multiple Deprivation fields to the Patients table")
    parser.add_argument("--imd-version", type=int, default=dep.IMD_VERSION,
                        help=f"IMD version for --deprivation (default: {dep.IMD_VERSION})")
    args = parser.parse_args()

    if args.deprivation:
        try:
            lookup_path, imd_path = dep.preflight(args.imd_version)
        except dep.DeprivationSetupError as e:
            sys.exit(f"Stopped: deprivation setup problem - {e}")
        print(f"Deprivation: IMD {args.imd_version} "
              f"(lookup: {lookup_path.name}, IMD file: {imd_path.name})\n")

    if not args.check_only and output_is_locked(clean.OUTPUT_XLSX):
        sys.exit(f"{clean.OUTPUT_XLSX.name} is open in Excel - close it and run again.")

    files = check_inputs.discover_raw_files([args.raw_dir])
    if not files:
        sys.exit(f"No .xlsx files found in {args.raw_dir}")

    print(f"STEP 1/3  Checking {len(files)} raw file(s)\n")
    _, ok = check_inputs.run_checks(files, sheet=args.sheet)
    if not ok:
        sys.exit("\nStopped: fix the blocking issues above, then run again.")
    if args.check_only:
        return

    print("\nSTEP 2/3  Flattening headers")
    FLATTENED_DIR.mkdir(parents=True, exist_ok=True)
    gp_sources = []
    for f in files:
        gp_id = check_inputs.gp_id_from_filename(f)
        out = FLATTENED_DIR / f"{f.stem}.csv"
        info = flatten_workbook(f, out, sheet=args.sheet)
        print(f"  [{gp_id}] {f.name} -> {out.relative_to(BASE_DIR)} ({info['n_rows_written']} rows)")
        gp_sources.append({"gp_id": gp_id, "source_file": str(out.relative_to(BASE_DIR))})

    print("\nSTEP 3/3  Cleaning and combining\n")
    try:
        clean.run(gp_sources, deprivation=args.deprivation, imd_version=args.imd_version)
    except ValueError as e:
        sys.exit(f"\nPIPELINE STOPPED:\n{e}")
    print(f"\nDone. Workbook: {clean.OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
