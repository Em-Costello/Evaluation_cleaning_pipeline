'''
deprivation.py

Adds Index of Multiple Deprivation (IMD) fields to a dataset that holds a
postcode or LSOA column.

    postcode / LSOA  ->  LSOA 2021 code  ->  IMD fields

Standalone for now (runs separately from the main pipeline) so the methods can
be tested on the test dataset:

    python deprivation.py
    python deprivation.py --input path/to/file.xlsx --column Postcode --id-column none

Every input row is kept. Rows that can't be matched get a reason in
Deprivation_Match_Status rather than stopping the run. Only setup problems
(missing lookup/IMD file, missing columns, unsupported IMD version) stop it.

The summary printed to the terminal is aggregate counts only - no postcodes,
LSOAs or other row-level values are printed.
'''

# =============================================================================
# LOGIC (step numbers are referenced in the code below)
# =============================================================================
#
# -----------------------------------------------------------------------------
# 0. CONFIG
# -----------------------------------------------------------------------------
#   - IMD_VERSION = 2025   (allowed: 2019, 2025) - only 2025 works for now,
#     see step 7.
#   - INPUT_FILE / INPUT_COLUMN  -> which file and which column holds the
#     postcode or LSOA (default column: 'Postcode')
#   - ID_COLUMN -> patient ID column ('NHS Number'). If set, the input is
#     collapsed to one row per patient before matching (the flattened GP file
#     has several rows per patient). Set to None for a plain list of
#     postcodes/LSOAs.
#   - IMD fields -> chosen in mapping_files/deprivation_indices/
#     imd_field_mapping.csv (include = Yes/No). Default is the overall IMD
#     score, rank and decile; switch on more (income, employment, health,
#     etc.) by changing 'No' to 'Yes' - no code change needed.
#   - OUTPUT_FILE -> where the enriched CSV is written.
#   - Later, when merged into the pipeline, these become command line
#     arguments (argparse), e.g. --deprivation --imd-version 2025. A basic
#     argparse is already included for testing.
#
# -----------------------------------------------------------------------------
# 1. LOCATE THE POSTCODE LOOKUP FILE (version-agnostic)
# -----------------------------------------------------------------------------
#   Goal: dropping a newer ONS lookup into mapping_files/deprivation_indices/
#   should work with no code changes (given that the structure is the same).
#   - Search the folder with a glob pattern.
#   - If more than one match, pick the newest. The release date is in the
#     filename (e.g. FEB25), so parse month+year from it and sort. Print which
#     file was chosen so the run is auditable.
#   - If no match: hard error (fail loud).
#
# -----------------------------------------------------------------------------
# 2. READ THE POSTCODE LOOKUP (handle zip + changing column names)
# -----------------------------------------------------------------------------
#   - pd.read_csv can read a .zip directly if the zip contains exactly one
#     CSV. ONS zips often also contain user guides / metadata, so use the
#     zipfile module: list contents, pick the single file ending in .csv
#     that looks like the lookup (not a 'metadata' / 'names' file), open it.
#   - ONS has changed column-name case between releases (PCD7 vs pcd7,
#     LSOA21CD vs lsoa21cd). Lower-case all column names on read so the rest
#     of the code doesn't care.
#   - Only read the columns we need (usecols) and force dtype=str - the file
#     is ~2.7m rows, so this saves a lot of memory and time:
#         pcds (or pcd7/pcd8), lsoa21cd, doterm
#   - Required-column check: if lsoa21cd or no postcode column is present,
#     hard error naming the missing column(s).
#
# -----------------------------------------------------------------------------
# 3. BUILD A SINGLE NORMALISED POSTCODE KEY (instead of matching by length)
# -----------------------------------------------------------------------------
#   Why: the ONS formats are
#       pcd7 = fixed 7 chars   e.g. 'A1  2BC', 'AB1 2CD', 'AB123CD'
#       pcd8 = fixed 8 chars   e.g. 'A1   2BC', 'AB1  2CD', 'AB12 3CD'
#       pcds = one space       e.g. 'A1 2BC',  'AB1 2CD', 'AB12 3CD'
#   Matching on raw input length works for some inputs by coincidence but
#   misses others (e.g. a 6-char 'A1 2BC' matches neither pcd7 nor pcd8,
#   and 'AB1 2CD' typed into an 8-char field won't match pcd8).
#   Fix: create the same key on BOTH sides:
#       key = upper-case, remove ALL whitespace   -> 'AB12CD', 'AB123CD'
#   Then one dictionary lookup handles every format.
#   This switches from my original logic of using pcd length to map
#   the input to the correct ONS column.
#
# -----------------------------------------------------------------------------
# 4. HANDLE TERMINATED POSTCODES
# -----------------------------------------------------------------------------
#   - doterm is blank for live postcodes and filled for terminated ones.
#   - Keep terminated postcodes in the lookup: GP records often hold old
#     addresses, and a terminated postcode still has a valid LSOA.
#   - If the same key appears twice (live + terminated), keep the live row
#     (sort so doterm is null first, then drop_duplicates on key).
#   - Add a flag so we know a match came from a terminated postcode
#     (e.g. status 'Matched (terminated postcode)').
#   - Final lookup structure: a pandas DataFrame indexed by key, with
#     lsoa21cd and a terminated True/False column.
#
# -----------------------------------------------------------------------------
# 5. CLEAN AND CLASSIFY EACH INPUT VALUE
# -----------------------------------------------------------------------------
#   - Keep the og value untouched in its own column for an audit trail;
#     do all cleaning in a separate working column.
#   - Clean: cast to str, strip, upper-case, remove internal whitespace.
#     Treat '', 'NAN', 'NONE', 'NULL', 'N/A' as missing.
#   - Classify with regex (on the cleaned value):
#       LSOA 2021 (England):  ^E01\d{6}$
#       LSOA (Wales):         ^W01\d{6}$   -> valid LSOA, but no English IMD
#       UK postcode (no space): ^[A-Z]{1,2}\d[A-Z\d]?\d[A-Z]{2}$
#       anything else       -> 'Invalid format'
#       missing             -> 'Missing'
#   - Store this as a Deprivation_Input_Type column:
#     Postcode / LSOA / Invalid / Missing. (Prefixed so it can't clash with
#     other columns once merged into the pipeline.)
#
# -----------------------------------------------------------------------------
# 6. POSTCODE -> LSOA
# -----------------------------------------------------------------------------
#   - For Input_Type == 'Postcode': map key -> lsoa21cd using the lookup.
#       found     -> LSOA filled
#       not found -> status 'Postcode not found in lookup' (could be a typo,
#                    a very new postcode, or a non-geographic postcode)
#   - For Input_Type == 'LSOA': copy the cleaned value straight into LSOA.
#   - Don't drop rows at any point - every input row must appear in the
#     output, with a status explaining why IMD is blank if it is.
#
# -----------------------------------------------------------------------------
# 7. LOAD THE IMD FILE FOR THE CHOSEN VERSION
# -----------------------------------------------------------------------------
#   - Keep a small config dict per version (IMD_CONFIG below).
#   - IMPORTANT: IMD 2019 is published on 2011 LSOAs; IMD 2025 on 2021 LSOAs.
#     The current postcode lookup only gives lsoa21cd. So for 2019 we would
#     need either an older lookup with lsoa11cd, or the ONS LSOA11 -> LSOA21
#     best-fit lookup. DECISION: only 2025 is supported for now; asking for
#     2019 raises a clear error explaining why.
#   - Rename the long IMD headers to short pipeline-style names, prefixed
#     with the version so outputs from different versions can't be confused:
#       'Index of Multiple Deprivation (IMD) Score'  -> IMD2025_Score
#       '... (IMD) Rank (where 1 is most deprived)'  -> IMD2025_Rank
#       '... (IMD) Decile (where 1 is most ...)'     -> IMD2025_Decile
#     Done through imd_field_mapping.csv in mapping_files/deprivation_indices/
#     (consistent with the rest of the pipeline: mappings external, not
#     hard-coded).
#   - Required-column check against the selected fields; hard error if missing.
#   - Index by LSOA code once (set_index) and reuse.
#   - Ranks and deciles are stored as whole numbers (Int64), so blanks don't
#     turn them into decimals like 3.0.
#
# -----------------------------------------------------------------------------
# 8. LSOA -> IMD
# -----------------------------------------------------------------------------
#   - Look up every row's LSOA in the IMD table in one go (reindex - same
#     result as a left-merge, but guarantees row order and row count are
#     unchanged).
#   - Rows with an LSOA but no IMD match:
#       LSOA doesn't start with E01 (Welsh W01, Scottish S01, NI, Channel
#       Islands / Isle of Man pseudo-codes) -> 'Outside England (no English IMD)'
#       otherwise  -> status 'LSOA not found in IMD file'
#
# -----------------------------------------------------------------------------
# 9. STATUS / FLAG COLUMN (flag, don't block)
# -----------------------------------------------------------------------------
#   One column, Deprivation_Match_Status, with a fixed set of values:
#       Matched
#       Matched (terminated postcode)
#       Missing
#       Invalid format
#       Postcode not found in lookup
#       Outside England (no English IMD)
#       LSOA not found in IMD file
#   - Errors are recorded here but never raised, so one bad postcode can't stop
#     the pipeline. Hard errors are reserved for setup problems only
#     (missing lookup file, missing required columns, unsupported version).
#
# -----------------------------------------------------------------------------
# 10. SUMMARY (aggregate counts only)
# -----------------------------------------------------------------------------
#   - Print value_counts of Deprivation_Match_Status and Input_Type, plus %
#     matched, and which lookup / IMD files were used.
#   - Do NOT print postcodes, LSOAs or any row-level data.
#   - Warning threshold: if match rate < 90% (of rows that had a value), print
#     a loud warning (suggests a wrong column or format issue) but continue.
#
# -----------------------------------------------------------------------------
# 11. OUTPUT
# -----------------------------------------------------------------------------
#   - Columns: original input columns (or ID + postcode if ID_COLUMN is set)
#     + LSOA21CD + Deprivation_Input_Type + Deprivation_Match_Status + chosen
#     IMD fields. Helper columns (cleaned key) are never added to the output.
#   - Written to cleaned_data/deprivation/<input name>_deprivation_IMD2025.csv;
#     once merged, also add as a sheet in the formatted Excel workbook.
#
# -----------------------------------------------------------------------------
# 12. STRUCTURE FOR MERGING INTO THE PIPELINE LATER
# -----------------------------------------------------------------------------
#   Each step is a function so the main pipeline can import them:
#       check_imd_version(version)
#       find_latest_lookup(folder)            -> Path
#       load_postcode_lookup(path)            -> DataFrame (key -> lsoa21cd)
#       load_imd(version, folder, mapping)    -> (DataFrame indexed by LSOA, Path)
#       clean_and_classify(series)            -> (cleaned, input type)
#       add_deprivation(df, column, lookup, imd, version)
#                                             -> df with IMD + status cols
#       summarise_matches(df, ...)            -> prints aggregate counts
#       read_input(path) / collapse_to_patients(df, id_col, col)
#   - Load lookup + IMD ONCE outside any per-GP loop (they are large).
#   - In clean_copd_dataset.py: call add_deprivation(patients_df, 'Postcode', ...)
#     only when the --deprivation flag is set; prefix messages with [gp_id]
#     if run per GP.
#   - Input from Excel as well as CSV: choose reader by file suffix
#     (.csv -> read_csv, .xlsx -> read_excel), always dtype=str.
#
# =============================================================================

import argparse
import re
import sys
import zipfile
from pathlib import Path
from typing import Optional

import pandas as pd


# =============================================================================
# 0. CONFIG
# =============================================================================

# All paths are resolved relative to this script's folder, so the pipeline
# works no matter which folder the terminal is in.
BASE_DIR = Path(__file__).resolve().parent
FLAT_DATA_DIR = BASE_DIR / 'evaluation_data' / 'flattened'
MAPPING_DIR = BASE_DIR / 'mapping_files'
DEPRIVATION_MAPPING_DIR = MAPPING_DIR / 'deprivation_indices'
OUTPUT_DIR = BASE_DIR / 'cleaned_data' / 'deprivation'

IMD_VERSION = 2025
INPUT_FILE = FLAT_DATA_DIR / 'TEST_GP00.csv'
INPUT_COLUMN = 'Postcode'
ID_COLUMN: Optional[str] = 'NHS Number'   # None for a plain list of postcodes/LSOAs
OUTPUT_FILE: Optional[Path] = None        # None -> OUTPUT_DIR/<input>_deprivation_IMD<version>.csv
IMD_FIELD_MAPPING_FILE = DEPRIVATION_MAPPING_DIR / 'imd_field_mapping.csv'
MATCH_RATE_WARNING = 0.90

# Postcode lookup (ONS). Glob matches both .csv and .zip.
LOOKUP_GLOB = 'PCD_OA21_LSOA21_*_LU.*'
LOOKUP_POSTCODE_COLS = ['pcds', 'pcd7', 'pcd8']   # first one present is used
LOOKUP_LSOA_COL = 'lsoa21cd'
LOOKUP_TERM_COL = 'doterm'
LOOKUP_LSOA_YEAR = 2021                           # which LSOA vintage the lookup gives

IMD_CONFIG = {
    2025: {'file_pattern': 'IoD2025_*.csv',
           'lsoa_col': 'LSOA code (2021)', 'lsoa_year': 2021},
    2019: {'file_pattern': 'File_7_*IoD2019*.csv',
           'lsoa_col': 'LSOA code (2011)', 'lsoa_year': 2011},
}

MONTHS = {m: i for i, m in enumerate(
    ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
     'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'], start=1)}

MISSING_TOKENS = {'', 'NAN', 'NONE', 'NULL', 'N/A', 'NA'}
POSTCODE_RE = r'[A-Z]{1,2}\d[A-Z\d]?\d[A-Z]{2}'
LSOA_RE = r'[EW]01\d{6}'

# Output column names
LSOA_OUT_COL = f'LSOA{str(LOOKUP_LSOA_YEAR)[2:]}CD'     # 'LSOA21CD'
INPUT_TYPE_COL = 'Deprivation_Input_Type'
STATUS_COL = 'Deprivation_Match_Status'

# Input types (step 5)
TYPE_POSTCODE, TYPE_LSOA, TYPE_INVALID, TYPE_MISSING = 'Postcode', 'LSOA', 'Invalid', 'Missing'

# Status values (step 9) - in the order they are shown in the summary
STATUS_MATCHED = 'Matched'
STATUS_MATCHED_TERM = 'Matched (terminated postcode)'
STATUS_MISSING = 'Missing'
STATUS_INVALID = 'Invalid format'
STATUS_PC_NOT_FOUND = 'Postcode not found in lookup'
STATUS_OUTSIDE_ENGLAND = 'Outside England (no English IMD)'
STATUS_LSOA_NOT_IN_IMD = 'LSOA not found in IMD file'
ALL_STATUSES = [STATUS_MATCHED, STATUS_MATCHED_TERM, STATUS_MISSING, STATUS_INVALID,
                STATUS_PC_NOT_FOUND, STATUS_OUTSIDE_ENGLAND, STATUS_LSOA_NOT_IN_IMD]


class DeprivationSetupError(Exception):
    '''Setup problem (missing file/column, unsupported version). Stops the run.'''


# =============================================================================
# 7 (part). VERSION CHECK - done first so we fail before loading big files
# =============================================================================

def check_imd_version(version: int) -> None:
    if version not in IMD_CONFIG:
        raise DeprivationSetupError(
            f'IMD version {version} not recognised. Allowed: {sorted(IMD_CONFIG)}.')
    imd_year = IMD_CONFIG[version]['lsoa_year']
    if imd_year != LOOKUP_LSOA_YEAR:
        raise DeprivationSetupError(
            f'IMD {version} is published on {imd_year} LSOAs, but the postcode lookup '
            f'gives {LOOKUP_LSOA_YEAR} LSOAs, so the codes would not match. To use IMD '
            f'{version}, add a lookup with lsoa{str(imd_year)[2:]}cd (or the ONS '
            f'LSOA11 -> LSOA21 best-fit lookup) first. Use --imd-version '
            f'{max(v for v, c in IMD_CONFIG.items() if c["lsoa_year"] == LOOKUP_LSOA_YEAR)} for now.')


# =============================================================================
# 1. LOCATE THE POSTCODE LOOKUP FILE
# =============================================================================

def _release_date(path: Path) -> Optional[tuple]:
    '''(year, month) from a filename like ..._FEB25_UK_LU.zip, or None.'''
    m = re.search(r'_([A-Z]{3})(\d{2})_', path.name.upper())
    if not m or m.group(1) not in MONTHS:
        return None
    return 2000 + int(m.group(2)), MONTHS[m.group(1)]


def find_latest_lookup(folder: Path) -> Path:
    candidates = [p for p in folder.glob(LOOKUP_GLOB)
                  if p.suffix.lower() in ('.csv', '.zip')]
    if not candidates:
        raise DeprivationSetupError(
            f'No postcode lookup matching {LOOKUP_GLOB} (.csv or .zip) in {folder}')

    dated = []
    for p in candidates:
        d = _release_date(p)
        if d is None:
            print(f'  WARNING: could not read a release date (e.g. FEB25) from '
                  f'{p.name} - skipping it.')
            continue
        # If the same release exists as .csv and .zip, prefer the .csv (faster).
        dated.append((d, p.suffix.lower() == '.csv', p))
    if not dated:
        raise DeprivationSetupError(
            f'Found postcode lookup file(s) in {folder} but none had a release date '
            f'like FEB25 in the name: {[p.name for p in candidates]}')

    dated.sort(key=lambda t: (t[0], t[1]))
    chosen = dated[-1][2]
    if len(dated) > 1:
        print(f'  {len(dated)} postcode lookups found - using the newest.')
    return chosen


# =============================================================================
# 2. READ THE POSTCODE LOOKUP
# =============================================================================

def _pick_zip_member(zf: zipfile.ZipFile, zip_path: Path) -> str:
    skip_words = ('metadata', 'names', 'readme', 'guide')
    csvs = [n for n in zf.namelist()
            if n.lower().endswith('.csv')
            and not any(w in Path(n).name.lower() for w in skip_words)]
    if len(csvs) > 1:   # narrow to files that look like the lookup itself
        csvs = [n for n in csvs if Path(n).name.lower().startswith('pcd')] or csvs
    if len(csvs) != 1:
        raise DeprivationSetupError(
            f'Expected exactly one lookup CSV inside {zip_path.name}, found '
            f'{len(csvs)}: {csvs}')
    return csvs[0]


def _read_lookup_csv(path: Path, **kwargs) -> pd.DataFrame:
    if path.suffix.lower() == '.csv':
        return pd.read_csv(path, **kwargs)
    with zipfile.ZipFile(path) as zf:
        member = _pick_zip_member(zf, path)
        with zf.open(member) as f:
            return pd.read_csv(f, **kwargs)


# =============================================================================
# 3. NORMALISED POSTCODE KEY
# =============================================================================

def normalise_postcode(s: pd.Series) -> pd.Series:
    '''Upper-case and remove ALL whitespace. Blank -> <NA>.'''
    key = s.astype('string').str.upper().str.replace(r'\s+', '', regex=True)
    return key.mask(key == '')


# =============================================================================
# 2-4. LOAD LOOKUP, BUILD KEY, HANDLE TERMINATED POSTCODES
# =============================================================================

def load_postcode_lookup(path: Path) -> pd.DataFrame:
    '''Returns a DataFrame indexed by postcode key with columns
    lsoa21cd (str) and terminated (bool).'''
    header = _read_lookup_csv(path, nrows=0)
    by_lower = {c.strip().lower(): c for c in header.columns}

    pc_col = next((c for c in LOOKUP_POSTCODE_COLS if c in by_lower), None)
    missing = []
    if pc_col is None:
        missing.append(f'a postcode column (one of {LOOKUP_POSTCODE_COLS})')
    if LOOKUP_LSOA_COL not in by_lower:
        missing.append(LOOKUP_LSOA_COL)
    if missing:
        raise DeprivationSetupError(
            f'Postcode lookup {path.name} is missing: {", ".join(missing)}. '
            f'Columns found: {list(header.columns)}')

    has_term = LOOKUP_TERM_COL in by_lower
    if not has_term:
        print(f'  WARNING: lookup has no {LOOKUP_TERM_COL} column - all postcodes '
              f'treated as live.')

    wanted = [pc_col, LOOKUP_LSOA_COL] + ([LOOKUP_TERM_COL] if has_term else [])
    lk = _read_lookup_csv(path, usecols=[by_lower[c] for c in wanted], dtype=str)
    lk.columns = [c.strip().lower() for c in lk.columns]

    lk['key'] = normalise_postcode(lk[pc_col])
    lk[LOOKUP_LSOA_COL] = lk[LOOKUP_LSOA_COL].str.strip()
    lk = lk[lk['key'].notna() & lk[LOOKUP_LSOA_COL].notna() & (lk[LOOKUP_LSOA_COL] != '')]

    if has_term:
        lk['terminated'] = lk[LOOKUP_TERM_COL].fillna('').str.strip() != ''
    else:
        lk['terminated'] = False

    # Step 4: live rows first, then keep one row per key
    lk = (lk.sort_values('terminated', kind='stable')
            .drop_duplicates('key', keep='first')
            .set_index('key'))
    return lk[[LOOKUP_LSOA_COL, 'terminated']]


# =============================================================================
# 7. LOAD THE IMD FILE
# =============================================================================

def load_imd(version: int, folder: Path, mapping_file: Path) -> tuple:
    '''Returns (DataFrame indexed by LSOA code with IMD<version>_* columns, path).'''
    cfg = IMD_CONFIG[version]

    matches = sorted(folder.glob(cfg['file_pattern']))
    if len(matches) != 1:
        raise DeprivationSetupError(
            f'Expected exactly one IMD {version} file matching {cfg["file_pattern"]} '
            f'in {folder}, found {len(matches)}: {[p.name for p in matches]}')
    imd_path = matches[0]

    # Which fields, and what to call them
    if not mapping_file.exists():
        raise DeprivationSetupError(f'IMD field mapping file not found: {mapping_file}')
    mapping = pd.read_csv(mapping_file, dtype=str)
    need = {'imd_version', 'source_column', 'output_name', 'include'}
    if not need.issubset(mapping.columns):
        raise DeprivationSetupError(
            f'{mapping_file.name} must have columns {sorted(need)}; '
            f'found {list(mapping.columns)}')
    mapping = mapping.apply(lambda c: c.str.strip())
    mapping = mapping[mapping['imd_version'] == str(version)]
    selected = mapping[mapping['include'].str.lower().isin(['yes', 'y', 'true', '1'])]
    if selected.empty:
        raise DeprivationSetupError(
            f'No fields set to include=Yes for IMD {version} in {mapping_file.name}')
    if selected['output_name'].duplicated().any():
        raise DeprivationSetupError(
            f'Duplicate output_name values in {mapping_file.name}: '
            f'{selected.loc[selected["output_name"].duplicated(), "output_name"].tolist()}')

    imd = pd.read_csv(imd_path, dtype={cfg['lsoa_col']: str})
    imd.columns = imd.columns.str.strip()

    src_cols = selected['source_column'].tolist()
    missing = [c for c in [cfg['lsoa_col']] + src_cols if c not in imd.columns]
    if missing:
        raise DeprivationSetupError(
            f'IMD file {imd_path.name} is missing column(s): {missing}')

    out = imd[[cfg['lsoa_col']] + src_cols].copy()
    out[cfg['lsoa_col']] = out[cfg['lsoa_col']].str.strip()
    if out[cfg['lsoa_col']].duplicated().any():
        raise DeprivationSetupError(
            f'IMD file {imd_path.name} has duplicate LSOA codes '
            f'({out[cfg["lsoa_col"]].duplicated().sum()} rows).')

    rename = {src: f'IMD{version}_{name}'
              for src, name in zip(selected['source_column'], selected['output_name'])}
    out = out.rename(columns=rename).set_index(cfg['lsoa_col'])
    out.index.name = None

    # Numbers: scores as decimals, ranks/deciles as whole numbers
    for col in out.columns:
        converted = pd.to_numeric(out[col], errors='coerce')
        n_bad = converted.isna().sum() - out[col].isna().sum()
        if n_bad:
            raise DeprivationSetupError(
                f'{n_bad} non-numeric value(s) in IMD column {col} ({imd_path.name}).')
        if col.endswith(('_Rank', '_Decile')):
            converted = converted.round().astype('Int64')
        out[col] = converted
    return out, imd_path


# =============================================================================
# 5. CLEAN AND CLASSIFY
# =============================================================================

def clean_and_classify(series: pd.Series) -> tuple:
    '''Returns (cleaned values, input type) - both Series on the same index.
    The original series is not modified.'''
    cleaned = normalise_postcode(series)
    cleaned = cleaned.mask(cleaned.isin(MISSING_TOKENS))

    is_pc = cleaned.str.fullmatch(POSTCODE_RE).fillna(False).astype(bool)
    is_lsoa = cleaned.str.fullmatch(LSOA_RE).fillna(False).astype(bool)

    itype = pd.Series(TYPE_INVALID, index=series.index, dtype=object)
    itype[cleaned.isna()] = TYPE_MISSING
    itype[is_pc] = TYPE_POSTCODE
    itype[is_lsoa] = TYPE_LSOA
    return cleaned, itype


# =============================================================================
# 6, 8, 9. POSTCODE -> LSOA -> IMD, WITH STATUS
# =============================================================================

def add_deprivation(df: pd.DataFrame, column: str, lookup: pd.DataFrame,
                    imd: pd.DataFrame, version: int) -> pd.DataFrame:
    '''Returns a copy of df with LSOA, input type, match status and IMD columns
    added. Never drops rows; never raises for bad values.'''
    if column not in df.columns:
        raise DeprivationSetupError(
            f'Column {column!r} not found in input. Columns: {list(df.columns)}')

    cleaned, itype = clean_and_classify(df[column])

    # Step 6: postcode -> LSOA
    lsoa = pd.Series(pd.NA, index=df.index, dtype='string')
    terminated = pd.Series(False, index=df.index)
    is_pc = itype.eq(TYPE_POSTCODE)
    is_lsoa = itype.eq(TYPE_LSOA)
    lsoa[is_pc] = cleaned[is_pc].map(lookup[LOOKUP_LSOA_COL])
    terminated[is_pc] = cleaned[is_pc].map(lookup['terminated']).fillna(False).astype(bool)
    lsoa[is_lsoa] = cleaned[is_lsoa]

    # Steps 8-9: status
    has_lsoa = lsoa.notna()
    outside = has_lsoa & ~lsoa.str.startswith('E01').fillna(False).astype(bool)
    in_imd = has_lsoa & lsoa.isin(imd.index).fillna(False).astype(bool)

    status = pd.Series(pd.NA, index=df.index, dtype=object)
    status[itype.eq(TYPE_MISSING)] = STATUS_MISSING
    status[itype.eq(TYPE_INVALID)] = STATUS_INVALID
    status[is_pc & ~has_lsoa] = STATUS_PC_NOT_FOUND
    status[outside] = STATUS_OUTSIDE_ENGLAND
    status[has_lsoa & ~outside & ~in_imd] = STATUS_LSOA_NOT_IN_IMD
    status[in_imd & ~terminated] = STATUS_MATCHED
    status[in_imd & terminated] = STATUS_MATCHED_TERM
    assert status.notna().all(), 'Internal error: some rows were given no status'

    # Step 8: attach IMD fields (blank where not matched)
    imd_vals = imd.reindex(lsoa.fillna('').to_numpy())
    imd_vals.index = df.index

    out = df.drop(columns=[c for c in [LSOA_OUT_COL, INPUT_TYPE_COL, STATUS_COL,
                                       *imd_vals.columns] if c in df.columns])
    out[LSOA_OUT_COL] = lsoa
    out[INPUT_TYPE_COL] = itype
    out[STATUS_COL] = status
    return pd.concat([out, imd_vals], axis=1)


# =============================================================================
# 10. SUMMARY (aggregate counts only)
# =============================================================================

def summarise_matches(df: pd.DataFrame, lookup_path: Path, imd_path: Path,
                      version: int) -> None:
    n = len(df)
    print('\n' + '=' * 60)
    print(f'DEPRIVATION SUMMARY (IMD {version})')
    print('=' * 60)
    print(f'Postcode lookup : {lookup_path.name}')
    print(f'IMD file        : {imd_path.name}')
    print(f'Rows            : {n}')

    print('\nInput type:')
    for t, c in df[INPUT_TYPE_COL].value_counts().reindex(
            [TYPE_POSTCODE, TYPE_LSOA, TYPE_INVALID, TYPE_MISSING], fill_value=0).items():
        print(f'  {t:<35}{c:>8}  ({c / n:6.1%})' if n else f'  {t:<35}{c:>8}')

    print('\nMatch status:')
    for s, c in df[STATUS_COL].value_counts().reindex(ALL_STATUSES, fill_value=0).items():
        print(f'  {s:<35}{c:>8}  ({c / n:6.1%})' if n else f'  {s:<35}{c:>8}')

    provided = int((df[INPUT_TYPE_COL] != TYPE_MISSING).sum())
    matched = int(df[STATUS_COL].isin([STATUS_MATCHED, STATUS_MATCHED_TERM]).sum())
    if provided:
        rate = matched / provided
        print(f'\nMatched {matched} of {provided} rows with a value ({rate:.1%}).')
        if rate < MATCH_RATE_WARNING:
            print('\n' + '!' * 60)
            print(f'WARNING: match rate is below {MATCH_RATE_WARNING:.0%}. Check the '
                  f'input column is the right one and the format looks right.')
            print('!' * 60)
    else:
        print('\nWARNING: no rows had a postcode or LSOA value.')


# =============================================================================
# INPUT HELPERS
# =============================================================================

def read_input(path: Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise DeprivationSetupError(f'Input file not found: {path}')
    suffix = path.suffix.lower()
    # Only truly empty cells become NaN - text like 'N/A' or 'NULL' is kept as
    # written so the original column stays an exact audit trail (step 5 still
    # treats those values as Missing).
    if suffix == '.csv':
        return pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[''])
    if suffix in ('.xlsx', '.xls'):
        return pd.read_excel(path, dtype=str, keep_default_na=False, na_values=[''])
    raise DeprivationSetupError(f'Unsupported input type {suffix!r} (use .csv or .xlsx)')


def collapse_to_patients(df: pd.DataFrame, id_col: str, column: str) -> pd.DataFrame:
    '''One row per patient: [id_col, column], taking the first non-missing
    postcode/LSOA for each patient. Prints aggregate counts of anything odd.'''
    for c in (id_col, column):
        if c not in df.columns:
            raise DeprivationSetupError(
                f'Column {c!r} not found in input. Columns: {list(df.columns)}')

    work = df[[id_col, column]].copy()
    cleaned, _ = clean_and_classify(work[column])
    work['_has'] = cleaned.notna()
    work['_key'] = cleaned

    no_id = work[id_col].isna() | (work[id_col].str.strip() == '')
    n_no_id_with_value = int((no_id & work['_has']).sum())
    if n_no_id_with_value:
        print(f'  WARNING: {n_no_id_with_value} row(s) have a {column} but no '
              f'{id_col} - these are excluded from the patient-level output.')
    work = work[~no_id]

    n_multi = int(work[work['_has']].groupby(id_col)['_key'].nunique().gt(1).sum())
    if n_multi:
        print(f'  WARNING: {n_multi} patient(s) have more than one different '
              f'{column} - the first one in the file is used.')

    out = (work.sort_values('_has', ascending=False, kind='stable')
               .drop_duplicates(id_col, keep='first')
               .sort_index())
    print(f'  Collapsed {len(df)} input rows to {len(out)} patients.')
    return out[[id_col, column]].reset_index(drop=True)


# =============================================================================
# 11. MAIN (standalone run)
# =============================================================================

def parse_args(argv=None):
    p = argparse.ArgumentParser(description='Add IMD deprivation fields to a dataset.')
    p.add_argument('--input', type=Path, default=INPUT_FILE)
    p.add_argument('--column', default=INPUT_COLUMN)
    p.add_argument('--id-column', default=ID_COLUMN,
                   help="Patient ID column, or 'none' to keep every row as-is.")
    p.add_argument('--imd-version', type=int, default=IMD_VERSION)
    p.add_argument('--output', type=Path, default=OUTPUT_FILE)
    args = p.parse_args(argv)
    if isinstance(args.id_column, str) and args.id_column.lower() == 'none':
        args.id_column = None
    return args


def main(argv=None) -> pd.DataFrame:
    args = parse_args(argv)

    # Fail fast on cheap checks before loading the big lookup
    check_imd_version(args.imd_version)
    print(f'Reading input: {args.input.name}')
    df = read_input(args.input)
    if args.column not in df.columns:
        raise DeprivationSetupError(
            f'Column {args.column!r} not found in input. Columns: {list(df.columns)}')
    if args.id_column:
        df = collapse_to_patients(df, args.id_column, args.column)

    lookup_path = find_latest_lookup(DEPRIVATION_MAPPING_DIR)
    print(f'Loading postcode lookup: {lookup_path.name} (can take a minute)')
    lookup = load_postcode_lookup(lookup_path)

    print(f'Loading IMD {args.imd_version}')
    imd, imd_path = load_imd(args.imd_version, DEPRIVATION_MAPPING_DIR, IMD_FIELD_MAPPING_FILE)

    result = add_deprivation(df, args.column, lookup, imd, args.imd_version)
    summarise_matches(result, lookup_path, imd_path, args.imd_version)

    out_path = args.output or OUTPUT_DIR / f'{args.input.stem}_deprivation_IMD{args.imd_version}.csv'
    out_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(out_path, index=False)
    print(f'\nWritten: {out_path}')
    return result


if __name__ == '__main__':
    try:
        main()
    except DeprivationSetupError as e:
        print(f'\nERROR: {e}', file=sys.stderr)
        sys.exit(1)