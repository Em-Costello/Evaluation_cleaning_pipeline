#!/usr/bin/env python3
"""
check_inputs.py

Pre-run checks on raw GP extracts (.xlsx). Replaces diagnose_structure.py.

For each file it flattens the headers (in memory, exactly as flatten_headers.py
would) and then runs every check the cleaning script would run, but reports ALL
problems at once instead of stopping at the first one:

  - Columns: missing, unexpected, renamed (with "did you mean" suggestions)
    or out of order, compared with NEW_EXPECTED_COLUMNS.
  - Row structure: anchor rows, NHS Number presence/checksum/duplicates,
    patient-level fields leaking onto detail rows.
  - Mapping files: every raw value with no mapping, across every mapped field.
  - Dates: values that can't be read as dates.
  - A short profile: patients, rows, and completeness of key patient fields.

Outputs (in qa_reports/):
  - preflight_report.txt   - the full report (also printed to the terminal)
  - unmapped_values.csv    - one row per unmapped value, in the same
                             raw_text_lowercase,category layout as the mapping
                             files, so rows can be filled in and copied across.

No NHS numbers or other identifiers are written to either file.

Usage:
    python check_inputs.py                          # every .xlsx in evaluation_data/raw/
    python check_inputs.py path/to/folder
    python check_inputs.py file1.xlsx file2.xlsx
"""

import argparse
import difflib
import io
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

import clean_copd_dataset as clean
from flatten_headers import flatten_workbook

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_RAW_DIR = BASE_DIR / "evaluation_data" / "raw"
REPORT_DIR = BASE_DIR / "qa_reports"

# Patient-level fields whose completeness is worth eyeballing per surgery.
PROFILE_FIELDS = ["DOB", "Gender", "Ethnicity", "Postcode", "MRC_Code",
                  "COPD_Diagnosis", "Smoking_Status"]
EVENT_FIELDS = {
    "Inhaler issues": "Inhaler_Issue_Date",
    "Prednisolone issues": "Prednisolone_Issue_Date",
    "Antibiotic issues": "Antibiotic_Issue_Date",
    "Consultations": "Consultation_Date",
    "Pulm rehab entries": "Referral_to_Pulmonary_Rehab",
    "Influenza vaccs": "Influenza_Vaccine",
    "Pneumococcal vaccs": "Pneumococcal_Vaccine",
    "COVID vaccs": "COVID_Vaccine",
}


# --------------------------------------------------------------------------
# File discovery
# --------------------------------------------------------------------------
def gp_id_from_filename(path):
    """'WHC_GP01.xlsx' -> 'GP01'. Returns None if the name has no GPnn code."""
    m = re.search(r"GP\d+", Path(path).stem, flags=re.IGNORECASE)
    return m.group(0).upper() if m else None


def discover_raw_files(paths):
    """Expand folders into .xlsx files, skipping Excel's ~$ lock files."""
    files = []
    for p in paths:
        p = Path(p)
        if p.is_dir():
            files.extend(sorted(f for f in p.glob("*.xlsx") if not f.name.startswith("~$")))
        elif p.is_file():
            files.append(p)
        else:
            sys.exit(f"Not found: {p}")
    return files


# --------------------------------------------------------------------------
# Checks
# --------------------------------------------------------------------------
def compare_columns(found, expected):
    """Return a list of human-readable column problems (empty = match)."""
    if found == expected:
        return []
    problems = []
    missing = [c for c in expected if c not in found]
    unexpected = [c for c in found if c not in expected]

    for col in unexpected:
        guess = difflib.get_close_matches(col, missing, n=1, cutoff=0.6)
        hint = f" - did you mean {guess[0]!r}?" if guess else ""
        problems.append(f"Unexpected column {col!r}{hint}")
    for col in missing:
        problems.append(f"Missing column {col!r}")

    if not missing and not unexpected:
        if len(found) != len(set(found)):
            dupes = sorted({c for c in found if found.count(c) > 1})
            problems.append(f"Duplicate column name(s): {dupes}")
        else:
            first_diff = next(i for i, (a, b) in enumerate(zip(found, expected)) if a != b)
            problems.append(
                f"All expected columns are present but in a different order "
                f"(first difference at position {first_diff + 1}: found {found[first_diff]!r}, "
                f"expected {expected[first_diff]!r}). Some renaming is done by position, "
                f"so the order must match.")
    problems.append(f"({len(found)} columns found, {len(expected)} expected)")
    return problems


def check_file(gp_id, xlsx_path, sheet=None):
    """Run every pre-run check on one raw extract. Returns a report dict."""
    report = {"gp_id": gp_id, "file": Path(xlsx_path).name, "blocking": [],
              "warnings": [], "unmapped": [], "profile": {}}

    # 1. Flatten in memory (same code path as flatten_headers.py), then read
    #    back with pandas exactly as clean_copd_dataset.py will.
    buffer = io.StringIO()
    try:
        info = flatten_workbook(xlsx_path, buffer, sheet=sheet)
    except Exception as e:  # unreadable workbook, missing sheet, etc.
        report["blocking"].append(f"Could not read/flatten the workbook: {e}")
        return report
    report["flatten"] = info
    buffer.seek(0)
    df = pd.read_csv(buffer, dtype={"NHS Number": str})
    df.columns = df.columns.str.strip()

    # 2. Columns
    column_problems = compare_columns(list(df.columns), clean.NEW_EXPECTED_COLUMNS)
    if column_problems:
        report["blocking"].extend(column_problems)
        # Later checks need the expected columns; only carry on if they're all there.
        if any(p.startswith(("Missing", "Duplicate")) for p in column_problems):
            return report
        df = df[clean.NEW_EXPECTED_COLUMNS]

    df = df.rename(columns=clean.RENAME_MAP)

    # 3. Unmapped values (checked on raw values, before any cleaning)
    report["unmapped"] = clean.find_unmapped_values(df)
    if report["unmapped"]:
        report["blocking"].append(
            f"{len(report['unmapped'])} unmapped value(s) - see the list below and "
            f"qa_reports/unmapped_values.csv")

    # 4. MRC codes
    mrc = pd.to_numeric(df["MRC_Code"].astype("string").str.extract(r"(\d+)")[0], errors="coerce")
    bad_mrc = df["MRC_Code"].notna() & ~mrc.isin([1, 2, 3, 4, 5])
    if bad_mrc.any():
        examples = sorted(df.loc[bad_mrc, "MRC_Code"].astype(str).unique())[:10]
        report["blocking"].append(f"{bad_mrc.sum()} MRC_Code value(s) not graded 1-5, e.g. {examples}")

    # 5. Row structure. Mapped fields are run through their mappings first so
    #    this matches what clean_copd_dataset.py sees at the same point.
    structured = df.copy()
    for col, mapping, _, strip_commas, out_col in clean.MAPPED_FIELDS:
        structured[out_col] = clean.normalise_for_mapping(structured[col], strip_commas).map(mapping)
    structured["Smoking_Status"] = structured["Smoking_Status"].apply(clean.classify_smoking_status)
    structured["MRC_Code"] = mrc
    report["blocking"].extend(clean.check_structure(structured))

    # 6. Dates (warning only - the cleaning step leaves these blank)
    for col in [c for c in df.columns if "Date" in c or c == "DOB"]:
        n_bad = clean.count_unparseable_dates(df[col], clean.parse_dates(df[col]))
        if n_bad:
            examples = df.loc[df[col].notna() & clean.parse_dates(df[col]).isna(), col].astype(str).unique()[:5]
            report["warnings"].append(
                f"{n_bad} value(s) in {col} can't be read as dates (e.g. {list(examples)}) "
                f"- they will be left blank.")

    dob = clean.parse_dates(df["DOB"])
    diag = clean.parse_dates(df["COPD_Diagnosis_Date"])
    if (diag < dob).any():
        report["warnings"].append(f"{(diag < dob).sum()} COPD diagnosis date(s) fall before DOB.")
    if (dob > pd.Timestamp.today()).any():
        report["warnings"].append(f"{(dob > pd.Timestamp.today()).sum()} DOB(s) are in the future.")

    # 7. Profile
    anchors = clean.find_anchor_rows(df)
    n_patients = int(anchors.sum())
    report["profile"]["Source rows"] = len(df)
    report["profile"]["Patients"] = n_patients
    if n_patients:
        for f in PROFILE_FIELDS:
            pct = 100 * df.loc[anchors, f].notna().mean()
            report["profile"][f"{f} recorded"] = f"{pct:.0f}%"
    for label, col in EVENT_FIELDS.items():
        report["profile"][label] = int(df[col].notna().sum())
    return report


# --------------------------------------------------------------------------
# Reporting
# --------------------------------------------------------------------------
def format_report(reports):
    out = [f"PRE-RUN CHECKS - {datetime.now():%d/%m/%Y %H:%M}", "=" * 60]
    for r in reports:
        status = "BLOCKING ISSUES" if r["blocking"] else ("OK with warnings" if r["warnings"] else "OK")
        out += ["", f"[{r['gp_id']}] {r['file']}  ->  {status}", "-" * 60]
        if "flatten" in r:
            f = r["flatten"]
            out.append(f"  Sheet '{f['sheet']}': header at row {f['header_start_row']}, "
                       f"{f['n_columns_kept']}/{f['n_columns_in_sheet']} columns kept"
                       + (f" (dropped {', '.join(f['dropped_columns'])})" if f["dropped_columns"] else "")
                       + f", {f['n_rows_written']} data rows.")
        for p in r["blocking"]:
            out.append(f"  X {p}")
        for u in r["unmapped"]:
            out.append(f"      {u['field']}: {u['raw_text_lowercase']!r} ({u['n_rows']} rows) "
                       f"-> {u['mapping_file']}")
        for w in r["warnings"]:
            out.append(f"  ! {w}")
        if r["profile"]:
            out.append("  Profile:")
            items = list(r["profile"].items())
            for i in range(0, len(items), 3):
                out.append("    " + "   ".join(f"{k}: {v}" for k, v in items[i:i + 3]))

    n_block = sum(bool(r["blocking"]) for r in reports)
    out += ["", "=" * 60,
            f"{len(reports)} file(s) checked, {n_block} with blocking issues."
            + ("" if n_block else " Ready to run.")]
    return "\n".join(out)


def write_unmapped_csv(reports, path):
    """Merge unmapped values across all GPs into one review CSV."""
    rows = [{**u, "gp_id": r["gp_id"]} for r in reports for u in r["unmapped"]]
    if not rows:
        if path.exists():
            path.unlink()
        return None
    df = pd.DataFrame(rows)
    merged = (df.groupby(["mapping_file", "field", "raw_text_lowercase"], as_index=False)
                .agg(n_rows=("n_rows", "sum"), found_in=("gp_id", lambda s: ", ".join(sorted(set(s))))))
    merged.insert(3, "category", "")  # fill in, then copy raw_text_lowercase,category into the mapping file
    merged.to_csv(path, index=False)
    return path


def run_checks(files, sheet=None):
    """Check every file; print + save the report. Returns (reports, ok)."""
    gp_ids = {}
    for f in files:
        gp_id = gp_id_from_filename(f)
        if gp_id is None:
            sys.exit(f"Can't find a GP code (e.g. GP01) in the file name {f.name!r}. "
                     f"Rename it like WHC_GP01.xlsx.")
        if gp_id in gp_ids:
            sys.exit(f"Two files share the GP code {gp_id}: {gp_ids[gp_id].name} and {f.name}")
        gp_ids[gp_id] = f

    reports = [check_file(gp_id, f, sheet=sheet) for gp_id, f in gp_ids.items()]

    REPORT_DIR.mkdir(exist_ok=True)
    text = format_report(reports)
    (REPORT_DIR / "preflight_report.txt").write_text(text, encoding="utf-8")
    print(text)
    unmapped_path = write_unmapped_csv(reports, REPORT_DIR / "unmapped_values.csv")
    print(f"\nReport saved to {REPORT_DIR / 'preflight_report.txt'}")
    if unmapped_path:
        print(f"Unmapped values saved to {unmapped_path} - fill in the 'category' column, "
              f"then copy the raw_text_lowercase,category pairs into the named mapping file.")
    ok = not any(r["blocking"] for r in reports)
    return reports, ok


def main():
    parser = argparse.ArgumentParser(description="Pre-run checks on raw GP extracts.")
    parser.add_argument("paths", nargs="*", default=[str(DEFAULT_RAW_DIR)],
                        help="Raw .xlsx files or folders (default: evaluation_data/raw/)")
    parser.add_argument("--sheet", default=None, help="Sheet name (default: active sheet)")
    args = parser.parse_args()

    files = discover_raw_files(args.paths)
    if not files:
        sys.exit(f"No .xlsx files found in {args.paths}")
    _, ok = run_checks(files, sheet=args.sheet)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
