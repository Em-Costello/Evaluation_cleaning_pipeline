import pandas as pd
import numpy as np
import re
import sys
from pathlib import Path
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import deprivation as dep

NEW_EXPECTED_COLUMNS = [ # CHANGE IF THE SOURCE FILE STRUCTURE CHANGES
    'NHS Number', 'Date of Birth', 'Age', 'Gender', 'Postcode',
    'Ethnic Origin', 'Clinical Codes - Code Term', 'Clinical Codes - Date',
    'COPD Diagnosis - Code Term', 'COPD Diagnosis - Date',
    'Depression - Code Term', 'Depression - Date', 'Anxiety - Code Term',
    'Anxiety - Date', 'Referral to Pulmonary Rehab - Code Term',
    'Referral to Pulmonary Rehab - Date', 'Smoking Status - Code Term',
    'Smoking Status - Date', 'Smoking Status - Associated Text',
    'Review of Inhaler Technique - Code Term',
    'Review of Inhaler Technique - Date',
    'Medication Issues - Name, Dosage and Quantity',
    'Medication Issues - Date of Issue', 'Medication Issues - Dose',
    'Prednisolone - Name, Dosage and Quantity',
    'Prednisolone - Date of Issue', 'Prednisolone - Dose',
    'Antibiotics - Name, Dosage and Quantity',
    'Antibiotics - Date of Issue', 'Antibiotics - Dose',
    'Consultations - Date', 'Consultations - Type of Consultation',
    'Cardiovascular Disease / Heart Disease - Code Term',
    'Cardiovascular Disease / Heart Disease - Date',
    'Cardiovascular Disease / Heart Disease - Associated Text',
    'Influenza - Code Term', 'Influenza - Date', 'Pneumococcal - Code Term',
    'Pneumococcal - Date', 'COVID - Code Term', 'COVID - Date'
]

RENAME_MAP = {
    'NHS Number': 'NHS_Number',
    'Date of Birth': 'DOB',
    'Ethnic Origin': 'Ethnicity',
    'Clinical Codes - Code Term': 'MRC_Code',
    'Clinical Codes - Date': 'MRC_Code_Date',
    'COPD Diagnosis - Code Term': 'COPD_Diagnosis',
    'COPD Diagnosis - Date': 'COPD_Diagnosis_Date',
    'Depression - Code Term': 'Depression',
    'Depression - Date': 'Depression_Date',
    'Anxiety - Code Term': 'Anxiety',
    'Anxiety - Date': 'Anxiety_Date',
    'Referral to Pulmonary Rehab - Code Term': 'Referral_to_Pulmonary_Rehab',
    'Referral to Pulmonary Rehab - Date': 'Referral_to_Pulmonary_Rehab_Date',
    'Smoking Status - Code Term': 'Smoking_Status',
    'Smoking Status - Date': 'Smoking_Status_Date',
    'Smoking Status - Associated Text': 'Smoking_Status_Associated_Text',
    'Review of Inhaler Technique - Code Term': 'Review_of_Inhaler_Technique',
    'Review of Inhaler Technique - Date': 'Review_of_Inhaler_Technique_Date',
    'Medication Issues - Name, Dosage and Quantity': 'Inhaler_Name_Dosage_Quantity',
    'Medication Issues - Date of Issue': 'Inhaler_Issue_Date',
    'Medication Issues - Dose': 'Inhaler_Dose',
    'Prednisolone - Name, Dosage and Quantity': 'Prednisolone_Name_Dosage_Quantity',
    'Prednisolone - Date of Issue': 'Prednisolone_Issue_Date',
    'Prednisolone - Dose': 'Prednisolone_Dose',
    'Antibiotics - Name, Dosage and Quantity': 'Antibiotic_Name_Dosage_Quantity',
    'Antibiotics - Date of Issue': 'Antibiotic_Issue_Date',
    'Antibiotics - Dose': 'Antibiotic_Dose',
    'Consultations - Date': 'Consultation_Date',
    'Consultations - Type of Consultation': 'Consultation_Type',
    'Cardiovascular Disease / Heart Disease - Code Term': 'Cardiovascular_Disease',
    'Cardiovascular Disease / Heart Disease - Date': 'Cardiovascular_Disease_Date',
    'Cardiovascular Disease / Heart Disease - Associated Text': 'Cardiovascular_Disease_Associated_Text',
    'Influenza - Code Term': 'Influenza_Vaccine',
    'Influenza - Date': 'Influenza_Vaccine_Date',
    'Pneumococcal - Code Term': 'Pneumococcal_Vaccine',
    'Pneumococcal - Date': 'Pneumococcal_Vaccine_Date',
    'COVID - Code Term': 'COVID_Vaccine',
    'COVID - Date': 'COVID_Vaccine_Date',
}

PATIENT_LEVEL_COLS = ['MRC_Code', 'COPD_Diagnosis', 'Depression', 'Anxiety',
                       'Smoking_Status', 'Gender', 'Ethnicity', 'Cardiovascular_Disease']

PATIENT_COLS = ['Patient_ID', 'GP_ID', 'Postcode', 'Age', 'DOB', 'Gender', 'Ethnicity',
                 'MRC_Code', 'MRC_Code_Date', 'COPD_Diagnosis', 'COPD_Diagnosis_Date',
                 'Depression', 'Depression_Date', 'Anxiety', 'Anxiety_Date',
                 'Smoking_Status', 'Smoking_Status_Date',
                 'Cardiovascular_Disease', 'Cardiovascular_Disease_Date']

RESCUE_PACK_TEXT_COLS = ['Inhaler_Dose', 'Prednisolone_Dose', 'Antibiotic_Dose']

# All paths are resolved relative to this script's folder, so the pipeline
# works no matter which folder the terminal is in.
BASE_DIR = Path(__file__).resolve().parent
MAPPING_DIR = BASE_DIR / 'mapping_files'
OUTPUT_DIR = BASE_DIR / 'cleaned_data' / 'combined_data'
OUTPUT_XLSX = OUTPUT_DIR / 'copd_combined_output.xlsx'
QA_REPORT_DIR = BASE_DIR / 'qa_reports'
DEPRIVATION_REPORT = QA_REPORT_DIR / 'deprivation_report.txt'

# ---------- GP sources to process ----------
# Only used when this script is run directly. run_pipeline.py builds this list
# automatically from the files in evaluation_data/raw/.
username = 'Emily Costello'

GP_SOURCES = [
    {'gp_id': 'GP01', 'source_file': 'evaluation_data/flattened/WHC_GP01.csv'},
    {'gp_id': 'GP02', 'source_file': 'evaluation_data/flattened/HARP_GP02.csv'},
    {'gp_id': 'GP03', 'source_file': 'evaluation_data/flattened/LW_GP03.csv'},
    {'gp_id': 'GP04', 'source_file': 'evaluation_data/flattened/MGP_GP04.csv'},
    {'gp_id': 'GP05', 'source_file': 'evaluation_data/flattened/SLRP_GP05.csv'},
    {'gp_id': 'GP06', 'source_file': 'evaluation_data/flattened/VS_GP06.csv'},
    
]

# ---------- Deprivation (optional) ----------
# Only used when this script is run directly. With run_pipeline.py, use the
# --deprivation (and optionally --imd-version) flags instead.
RUN_DEPRIVATION = False
IMD_VERSION = dep.IMD_VERSION
# IMD fields (output_name in imd_field_mapping.csv) copied onto the Patients
# table. Everything else goes in the separate Deprivation table only.
PATIENT_IMD_FIELDS = ['Decile']

# ---------- Load categorical mappings once (shared across all GP files) ----------
def load_mapping(filename):
    mapping = pd.read_csv(MAPPING_DIR / filename).set_index('raw_text_lowercase')
    return mapping['category'].to_dict()


COPD_DIAGNOSIS_CATEGORY_MAP = load_mapping('copd_diagnosis_mapping.csv')
GENDER_CATEGORY_MAP = load_mapping('gender_mapping.csv')
ETHNICITY_CATEGORY_MAP = load_mapping('ethnicity_mapping.csv')
CARDIOVASCULAR_DISEASE_CATEGORY_MAP = load_mapping('cardiovascular_disease_mapping.csv')
PULMONARY_REHAB_CATEGORY_MAP = load_mapping('pulmonary_rehab_status_mapping.csv')

# Every field cleaned with a mapping file, in one place so the cleaning step
# and the pre-run checks (check_inputs.py) always use identical rules.
#   (column, mapping, mapping file, strip commas before lookup?, output column)
MAPPED_FIELDS = [
    ('COPD_Diagnosis', COPD_DIAGNOSIS_CATEGORY_MAP, 'copd_diagnosis_mapping.csv', True, 'COPD_Diagnosis'),
    ('Gender', GENDER_CATEGORY_MAP, 'gender_mapping.csv', True, 'Gender'),
    ('Ethnicity', ETHNICITY_CATEGORY_MAP, 'ethnicity_mapping.csv', True, 'Ethnicity'),
    ('Cardiovascular_Disease', CARDIOVASCULAR_DISEASE_CATEGORY_MAP, 'cardiovascular_disease_mapping.csv', True, 'Cardiovascular_Disease'),
    ('Referral_to_Pulmonary_Rehab', PULMONARY_REHAB_CATEGORY_MAP, 'pulmonary_rehab_status_mapping.csv', False, 'Pulmonary_Rehab_Status'),
]


def normalise_for_mapping(series, strip_commas=True):
    """Lower-case/strip a raw column the same way the mapping files are keyed."""
    s = series.astype(str)
    if strip_commas:
        s = s.str.replace(',', '', regex=False)
    return s.str.strip().str.lower().replace('nan', pd.NA)


def find_unmapped_values(df):
    """Return every raw value that has no mapping (or that the smoking parser
    doesn't recognise), as a list of dicts. Checks ALL fields in one pass, so
    one run shows everything that needs adding rather than one error at a time."""
    problems = []
    for col, mapping, mapping_file, strip_commas, _ in MAPPED_FIELDS:
        normalised = normalise_for_mapping(df[col], strip_commas)
        counts = normalised.dropna().value_counts()
        for raw_value, n in counts.items():
            if raw_value not in mapping:
                problems.append({'field': col, 'mapping_file': mapping_file,
                                 'raw_text_lowercase': raw_value, 'n_rows': int(n)})

    smoking_raw = df['Smoking_Status'].dropna()
    unrecognised = smoking_raw[smoking_raw.apply(classify_smoking_status).isna()]
    for raw_value, n in unrecognised.value_counts().items():
        problems.append({'field': 'Smoking_Status',
                         'mapping_file': 'classify_smoking_status() in clean_copd_dataset.py',
                         'raw_text_lowercase': str(raw_value).strip().lower(), 'n_rows': int(n)})
    return problems


def parse_dates(series):
    """Parse a date column that may hold ISO dates (YYYY-MM-DD, which is how
    real Excel dates come out of flatten_headers.py) and/or UK text dates
    (DD/MM/YYYY).

    pd.to_datetime(..., dayfirst=True) on its own swaps day and month on ISO
    dates (2023-05-04 -> 5 April) and turns any ISO date with day > 12 into NaT,
    so the two formats are parsed separately."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    s = series.astype('string').str.strip()
    is_iso = s.str.match(r'^\d{4}-\d{2}-\d{2}', na=False)
    parsed = pd.Series(pd.NaT, index=series.index, dtype='datetime64[ns]')
    if is_iso.any():
        parsed[is_iso] = pd.to_datetime(s[is_iso].str[:10], format='%Y-%m-%d', errors='coerce')
    other = ~is_iso & s.notna() & (s != '')
    if other.any():
        parsed[other] = pd.to_datetime(s[other], dayfirst=True, format='mixed', errors='coerce')
    return parsed


def count_unparseable_dates(raw, parsed):
    """Number of non-blank raw values that failed to parse as a date."""
    raw_present = raw.notna() & (raw.astype(str).str.strip() != '')
    return int((raw_present & parsed.isna()).sum())

def classify_smoking_status(raw):
    if pd.isna(raw):
        return pd.NA
    text = str(raw).strip().lower()
    if re.search(r'non[- ]smoker|never smoked', text):
        return 'Non-Smoker'

    has_smoker_word = 'smoker' in text
    has_rolls_own = 'cigarette' in text and 'roll' in text
    has_cigar = 'cigar' in text and 'cigarette' not in text
    has_consumption_or_packyears = 'cigarette' in text and ('consum' in text or 'pack' in text and 'year' in text)

    if not (has_smoker_word or has_rolls_own or has_cigar or has_consumption_or_packyears):
        return None

    status = 'Former' if re.search(r'\bex[- ]', text) else 'Current'
    if 'very heavy' in text:
        heaviness, band = 'Very Heavy', '40+/day'
    elif 'heavy' in text:
        heaviness, band = 'Heavy', '20-39/day'
    elif 'moderate' in text:
        heaviness, band = 'Moderate', '10-19/day'
    elif 'light' in text:
        heaviness, band = 'Light', '1-9/day'
    elif 'trivial' in text:
        heaviness, band = 'Trivial', '<1/day'
    else:
        return f'{status} Smoker (Amount Not Specified)'
    return f'{status} {heaviness} Smoker ({band})'


def is_valid_nhs_number(nhs_num):
    # A numeric column with blanks is read as float (9434765919.0); without
    # this, the '.0' would add an 11th digit and fail every number.
    if isinstance(nhs_num, float) and nhs_num.is_integer():
        nhs_num = int(nhs_num)
    digits = re.sub(r'\D', '', str(nhs_num))
    if len(digits) != 10:
        return False
    total = sum(int(d) * w for d, w in zip(digits[:9], range(10, 1, -1)))
    check_digit = 11 - (total % 11)
    if check_digit == 11:
        check_digit = 0
    if check_digit == 10:
        return False
    return check_digit == int(digits[9])


def find_anchor_rows(df):
    return df['Age'].notna() & df['Postcode'].notna() & df['DOB'].notna()


def check_structure(df):
    """Check the anchor/detail row structure of a renamed GP DataFrame.
    Returns a list of problem descriptions (empty list = all fine).
    Row numbers refer to lines in the flattened CSV (line 1 = header).
    NHS numbers themselves are never printed."""
    problems = []
    is_anchor = find_anchor_rows(df)
    to_lines = lambda idx: [i + 2 for i in idx][:20]  # cap the list length

    if len(df) == 0:
        return ["File has no data rows."]
    if not is_anchor.iloc[0]:
        problems.append("First data row isn't a patient anchor row (Age, Postcode and DOB all filled).")

    mismatch = df['NHS_Number'].notna() != is_anchor
    if mismatch.any():
        problems.append(
            f"{mismatch.sum()} row(s) where NHS_Number presence doesn't match the anchor-row "
            f"pattern, at CSV lines {to_lines(df.index[mismatch])}.")

    anchor_nhs = df.loc[is_anchor & df['NHS_Number'].notna(), 'NHS_Number']
    invalid = ~anchor_nhs.apply(is_valid_nhs_number)
    if invalid.any():
        problems.append(
            f"{invalid.sum()} NHS Number(s) failed the checksum, at CSV lines "
            f"{to_lines(anchor_nhs.index[invalid])}.")

    dupes = anchor_nhs.duplicated(keep=False)
    if dupes.any():
        problems.append(
            f"{anchor_nhs[dupes].nunique()} NHS Number(s) appear in more than one patient block, "
            f"at CSV lines {to_lines(anchor_nhs.index[dupes])}.")

    leaks = df.loc[~is_anchor, PATIENT_LEVEL_COLS].notna()
    leak_cols = leaks.columns[leaks.any()].tolist()
    if leak_cols:
        problems.append(
            f"Patient-level field(s) {leak_cols} filled on non-anchor rows "
            f"(e.g. CSV lines {to_lines(leaks.index[leaks.any(axis=1)])}).")
    return problems


def build_event_table(source_df, cols, subset_for_filter=None):
    subset_for_filter = subset_for_filter or cols
    mask = source_df[subset_for_filter].notna().any(axis=1)
    return source_df.loc[mask, ['Patient_ID', 'GP_ID'] + cols].reset_index(drop=True)


def process_gp_file(gp_id, source_file):
    """Run the full single-file cleaning pipeline for one GP surgery extract
    and return a dict of {table_name: DataFrame}, each tagged with GP_ID.
    Every error message is prefixed with the gp_id so failures are traceable
    back to the offending file when processing several GPs in one run."""

    def fail(msg):
        raise ValueError(f"[{gp_id}] {msg}")

    # ---------- 1. Load + validate structure ----------
    # NHS Number read as text so leading zeros survive and it isn't turned into a float.
    df = pd.read_csv(source_file, dtype={'NHS Number': str})
    df.columns = df.columns.str.strip()
    if list(df.columns) != NEW_EXPECTED_COLUMNS:
        fail("Column mismatch - structure of this file differs from the expected format. "
             "Run check_inputs.py on the raw file to see exactly which columns differ.")

    # ---------- 2. Rename to internal names ----------
    df.rename(columns=RENAME_MAP, inplace=True)
    df.drop(columns=['Review_of_Inhaler_Technique', 'Review_of_Inhaler_Technique_Date'], inplace=True)

    # ---------- 2.5 Clean inconsistent data formats ----------
    # All unmapped values across all fields are reported together.
    unmapped = find_unmapped_values(df)
    if unmapped:
        lines = [f"  {u['field']} -> {u['mapping_file']}: {u['raw_text_lowercase']!r} ({u['n_rows']} rows)"
                 for u in unmapped]
        fail("Unmapped value(s) - add rows to the mapping file(s) below "
             "(or run check_inputs.py for a CSV of these):\n" + "\n".join(lines))

    for col, mapping, _, strip_commas, out_col in MAPPED_FIELDS:
        df[out_col] = normalise_for_mapping(df[col], strip_commas).map(mapping)

    df['Smoking_Status'] = df['Smoking_Status'].apply(classify_smoking_status)

    df['MRC_Code'] = pd.to_numeric(df['MRC_Code'].str.extract(r'(\d+)')[0], errors='coerce')
    if not df['MRC_Code'].dropna().isin([1, 2, 3, 4, 5]).all():
        fail("MRC_Code contains unexpected values - check file structure.")

    # ---------- 3. Patient ID: real NHS Number (confirmed patient-row only) ----------
    structure_problems = check_structure(df)
    if structure_problems:
        fail("File doesn't match the expected anchor/detail row structure:\n  "
             + "\n  ".join(structure_problems))

    is_anchor = find_anchor_rows(df)
    anchor_rows = df.loc[is_anchor]
    temp_group = is_anchor.cumsum()
    group_to_id = dict(zip(temp_group[is_anchor], anchor_rows['NHS_Number']))
    df['Patient_ID'] = temp_group.map(group_to_id)
    df['GP_ID'] = gp_id

    # ---------- 3.5 Parse dates needed for rescue pack matching (ahead of section 6) ----------
    RESCUE_PACK_DATE_COLS = ['Inhaler_Issue_Date', 'Prednisolone_Issue_Date', 'Antibiotic_Issue_Date']
    for c in RESCUE_PACK_DATE_COLS:
        df[c] = parse_dates(df[c])

    # ---------- 4. patients_df: one row per patient ----------
    patients_df = df.loc[is_anchor, PATIENT_COLS].reset_index(drop=True)
    patients_df['Depression'] = patients_df['Depression'].notna().map({True: 'Yes', False: pd.NA})
    patients_df['Anxiety'] = patients_df['Anxiety'].notna().map({True: 'Yes', False: pd.NA})

    # ---------- 5. Event tables ----------
    inhaler_prescriptions_df = build_event_table(
        df, ['Inhaler_Name_Dosage_Quantity', 'Inhaler_Dose', 'Inhaler_Issue_Date'])
    prednisolone_courses_df = build_event_table(
        df, ['Prednisolone_Name_Dosage_Quantity', 'Prednisolone_Dose', 'Prednisolone_Issue_Date'])
    antibiotic_courses_df = build_event_table(
        df, ['Antibiotic_Name_Dosage_Quantity', 'Antibiotic_Dose', 'Antibiotic_Issue_Date'])
    consultations_df = build_event_table(
        df, ['Consultation_Type', 'Consultation_Date'])
    referrals_df = build_event_table(
        df, ['Referral_to_Pulmonary_Rehab', 'Referral_to_Pulmonary_Rehab_Date'])
    influenza_df = build_event_table(df, ['Influenza_Vaccine', 'Influenza_Vaccine_Date'])
    pneumococcal_df = build_event_table(df, ['Pneumococcal_Vaccine', 'Pneumococcal_Vaccine_Date'])
    covid_df = build_event_table(df, ['COVID_Vaccine', 'COVID_Vaccine_Date'])

    # ---------- 5.5 Prescription/course counts onto patients_df ----------
    count_specs = [
        (inhaler_prescriptions_df, 'Inhaler_Prescription_Count'),
        (prednisolone_courses_df, 'Prednisolone_Course_Count'),
        (antibiotic_courses_df, 'Antibiotic_Course_Count'),
        (consultations_df, 'Consultation_Count'),
    ]
    for event_table, count_col in count_specs:
        counts = event_table.groupby('Patient_ID').size().rename(count_col)
        patients_df = patients_df.merge(counts, on='Patient_ID', how='left')
    count_cols = [c for _, c in count_specs]
    patients_df[count_cols] = patients_df[count_cols].fillna(0).astype(int)

    # ---------- 5.6 Evidence of Rescue Pack (text mention OR same-day pred+antibiotic) ----------
    RESCUE_PACK_COL_PAIRS = [
        ('Inhaler_Dose', 'Inhaler_Issue_Date'),
        ('Prednisolone_Dose', 'Prednisolone_Issue_Date'),
        ('Antibiotic_Dose', 'Antibiotic_Issue_Date'),
    ]

    # Signal 1: explicit "rescue pack" text mention, paired with that row's (parsed) date
    text_mentioned = pd.Series(False, index=df.index)
    text_mention_date = pd.Series(pd.NaT, index=df.index)
    for dose_col, date_col in RESCUE_PACK_COL_PAIRS:
        match = df[dose_col].str.contains('rescue pack', case=False, na=False)
        text_mentioned |= match
        text_mention_date = text_mention_date.mask(match, df[date_col])

    text_events = set(
        zip(df.loc[text_mentioned, 'Patient_ID'], text_mention_date[text_mentioned])
    )

    # Signal 2: prednisolone + antibiotic issued on the same (parsed) date
    prednisolone_dates = set(
        zip(df.loc[df['Prednisolone_Issue_Date'].notna(), 'Patient_ID'],
            df.loc[df['Prednisolone_Issue_Date'].notna(), 'Prednisolone_Issue_Date'])
    )
    antibiotic_dates = set(
        zip(df.loc[df['Antibiotic_Issue_Date'].notna(), 'Patient_ID'],
            df.loc[df['Antibiotic_Issue_Date'].notna(), 'Antibiotic_Issue_Date'])
    )
    inferred_events = prednisolone_dates & antibiotic_dates

    # Union: a (Patient_ID, date) counts once even if flagged by both signals
    rescue_pack_events = text_events | inferred_events

    rescue_pack_counts = (
        pd.Series([pid for pid, _ in rescue_pack_events], name='Patient_ID')
        .value_counts()
        .rename('Evidence_of_Rescue_Pack')
    )

    patients_df = patients_df.merge(rescue_pack_counts, on='Patient_ID', how='left')
    patients_df['Evidence_of_Rescue_Pack'] = patients_df['Evidence_of_Rescue_Pack'].fillna(0).astype(int)

    # ---------- 5.7 Vaccinated (Y/N) flags onto patients_df ----------
    vaccination_specs = [
        (influenza_df, 'Influenza_Vaccinated'),
        (pneumococcal_df, 'Pneumococcal_Vaccinated'),
        (covid_df, 'COVID_Vaccinated'),
    ]
    for event_table, flag_col in vaccination_specs:
        vaccinated_ids = set(event_table['Patient_ID'])
        patients_df[flag_col] = patients_df['Patient_ID'].isin(vaccinated_ids).map({True: 'Yes', False: pd.NA})

    # ---------- 5.8 Pulmonary rehab referral counts onto patients_df ----------
    rehab_counts = (
        df.loc[df['Pulmonary_Rehab_Status'].isin(['Referred', 'Declined', 'Completed'])]
        .groupby(['Patient_ID', 'Pulmonary_Rehab_Status'])
        .size()
        .unstack(fill_value=0)
        # A surgery may have no Declined/Completed entries at all; make sure
        # every status column exists so the lookups below can't fail.
        .reindex(columns=['Referred', 'Declined', 'Completed'], fill_value=0)
    )
    for status, col_name in [('Referred', 'Number_of_Times_Referred'),
                              ('Declined', 'Number_of_Times_Declined'),
                              ('Completed', 'Number_of_Times_Completed')]:
        patients_df[col_name] = patients_df['Patient_ID'].map(rehab_counts[status]).fillna(0).astype(int)

    # ---------- 6. Parse dates in every table ----------
    gp_tables = {
        'patients': patients_df,
        'inhaler_prescriptions': inhaler_prescriptions_df,
        'prednisolone_courses': prednisolone_courses_df,
        'antibiotic_courses': antibiotic_courses_df,
        'consultations': consultations_df,
        'referrals_pulmonary_rehab': referrals_df,
        'influenza_vaccinations': influenza_df,
        'pneumococcal_vaccinations': pneumococcal_df,
        'covid_vaccinations': covid_df,
    }
    for table_name, t in gp_tables.items():
        for c in t.columns:
            if 'Date' in c or c == 'DOB':
                parsed = parse_dates(t[c])
                n_bad = count_unparseable_dates(t[c], parsed)
                if n_bad:
                    print(f"[{gp_id}] WARNING: {n_bad} value(s) in {table_name}.{c} "
                          f"couldn't be read as dates and were left blank.")
                t[c] = parsed

    # ---------- 7. Standardise categorical text in patients_df ----------
    cat_cols = ['MRC_Code', 'COPD_Diagnosis', 'Depression', 'Anxiety']
    for c in cat_cols:
        patients_df[c] = patients_df[c].astype(str).str.strip().str.title().replace('Nan', pd.NA)
    patients_df['Postcode'] = patients_df['Postcode'].astype(str).str.strip().str.upper().replace('NAN', pd.NA)

    # ---------- 8. Sanity checks ----------
    if (patients_df['COPD_Diagnosis_Date'] < patients_df['DOB']).any():
        print(f"[{gp_id}] Data quality issue: some COPD diagnosis dates are before DOB")

    print(f"[{gp_id}] {len(patients_df)} patients, {len(df)} source rows")
    return gp_tables


# =========================================================================
# Excel export for analysts (NHS workers who work in Excel, not Python)
# One workbook, one sheet per table, plus a Summary index sheet up front.
# =========================================================================

HEADER_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=11)
HEADER_FILL = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
BODY_FONT = Font(name='Arial', size=11)
TITLE_FONT = Font(name='Arial', bold=True, size=14)
DATE_FORMAT = 'DD/MM/YYYY'
MAX_COL_WIDTH = 45
MIN_COL_WIDTH = 10

# Excel sheet names are capped at 31 characters and can't repeat, so table
# names are mapped to short, analyst-friendly labels here.
SHEET_LABELS = {
    'patients': 'Patients',
    'inhaler_prescriptions': 'Inhaler Prescriptions',
    'prednisolone_courses': 'Prednisolone Courses',
    'antibiotic_courses': 'Antibiotic Courses',
    'consultations': 'Consultations',
    'referrals_pulmonary_rehab': 'Referrals - Pulm Rehab',
    'influenza_vaccinations': 'Influenza Vaccinations',
    'pneumococcal_vaccinations': 'Pneumococcal Vaccs',
    'covid_vaccinations': 'COVID Vaccinations',
    'deprivation': 'Deprivation',
}

SHEET_DESCRIPTIONS = {
    'patients': 'One row per patient: demographics, diagnosis, comorbidities and event counts.',
    'inhaler_prescriptions': 'One row per inhaler prescription issued.',
    'prednisolone_courses': 'One row per prednisolone course issued.',
    'antibiotic_courses': 'One row per antibiotic course issued.',
    'consultations': 'One row per recorded consultation.',
    'referrals_pulmonary_rehab': 'One row per pulmonary rehab referral.',
    'influenza_vaccinations': 'One row per influenza vaccination record.',
    'pneumococcal_vaccinations': 'One row per pneumococcal vaccination record.',
    'covid_vaccinations': 'One row per COVID vaccination record.',
    'deprivation': ('One row per patient: LSOA, match status and Index of Multiple '
                    'Deprivation scores, ranks and deciles (1 = most deprived). '
                    'Only present when run with --deprivation.'),
}


def _write_table_sheet(wb, sheet_title, df):
    """Write one DataFrame to its own formatted, filterable sheet."""
    ws = wb.create_sheet(title=sheet_title[:31])

    # pandas NA/NaT types aren't writable by openpyxl - swap for plain None first.
    clean_df = df.astype(object).where(df.notna(), None)

    ws.append(list(clean_df.columns))
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical='center')

    date_col_idxs = [i for i, c in enumerate(clean_df.columns, start=1) if 'Date' in c or c == 'DOB']

    for row in clean_df.itertuples(index=False, name=None):
        ws.append(list(row))

    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row_cells:
            cell.font = BODY_FONT
        for idx in date_col_idxs:
            row_cells[idx - 1].number_format = DATE_FORMAT

    for i, col in enumerate(clean_df.columns, start=1):
        sample = clean_df[col].map(lambda x: str(x) if pd.notna(x) else "").tolist()
        max_len = max([len(str(col))] + [len(v) for v in sample])
        ws.column_dimensions[get_column_letter(i)].width = min(
            max(max_len + 2, MIN_COL_WIDTH),
            MAX_COL_WIDTH
        )
        
    ws.freeze_panes = 'A2'
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    return ws


def save_excel_workbook(tables, gp_sources, output_path, notes=None):
    """Combine every cleaned table into one workbook: a Summary index sheet
    followed by one formatted, filterable sheet per table.
    notes: optional extra lines shown under the GP list on the Summary sheet."""
    wb = Workbook()
    wb.remove(wb.active)  # drop the default blank sheet

    summary = wb.create_sheet(title='Summary', index=0)
    summary.append(['COPD Evaluation - Cleaned Data Export'])
    summary['A1'].font = TITLE_FONT
    summary.append([f'Generated: {date.today().strftime("%d/%m/%Y")} by: {username}'])
    summary['A2'].font = BODY_FONT
    summary.append([f'GP surgeries included: {", ".join(g["gp_id"] for g in gp_sources)}'])
    summary['A3'].font = BODY_FONT
    for note in notes or []:
        summary.append([note])
        summary.cell(row=summary.max_row, column=1).font = BODY_FONT
    summary.append([])

    header_row = ['Sheet', 'Row Count', 'Description']
    summary.append(header_row)
    header_row_idx = summary.max_row
    for cell in summary[summary.max_row]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    for name, df in tables.items():
        label = SHEET_LABELS.get(name, name)
        summary.append([label, len(df), SHEET_DESCRIPTIONS.get(name, '')])
        for cell in summary[summary.max_row]:
            cell.font = BODY_FONT
        _write_table_sheet(wb, label, df)

    summary.column_dimensions['A'].width = 28
    summary.column_dimensions['B'].width = 12
    summary.column_dimensions['C'].width = 70
    summary.freeze_panes = f'A{header_row_idx + 1}'

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    print(f"Excel workbook saved: {output_path} ({len(tables)} data sheets + Summary)")


# ---------- Run pipeline across all configured GP surgeries ----------
TABLE_NAMES = [
    'patients', 'inhaler_prescriptions', 'prednisolone_courses', 'antibiotic_courses',
    'consultations', 'referrals_pulmonary_rehab', 'influenza_vaccinations',
    'pneumococcal_vaccinations', 'covid_vaccinations',
]


def build_deprivation_table(patients_df, imd_version):
    """Section 8.5: postcode -> LSOA -> IMD for the combined patients table.
    Loads the (large) lookup and IMD files once, for all GPs together.
    Bad/missing postcodes are flagged in Deprivation_Match_Status, never raised.

    Returns (patients_df with the PATIENT_IMD_FIELDS added,
             deprivation_df - one row per patient with LSOA, input type,
                              match status and every chosen IMD field,
             line for the workbook Summary sheet)."""
    print("\n---------- Deprivation indices ----------")
    ref = dep.load_reference_data(imd_version)
    enriched = dep.add_deprivation(patients_df[['Patient_ID', 'GP_ID', 'Postcode']], 'Postcode',
                                   ref['lookup'], ref['imd'], imd_version)
    counts = dep.summarise_matches(enriched, ref['lookup_path'], ref['imd_path'],
                                   imd_version, group_col='GP_ID')
    where = (f"Open {OUTPUT_XLSX.name} -> Deprivation sheet (or copd_deprivation.csv) and "
             f"filter Deprivation_Match_Status on the status below; add a GP_ID filter to "
             f"narrow to one surgery. The postcode for each Patient_ID is on the Patients "
             f"sheet (Postcode column) - match on Patient_ID + GP_ID.")
    dep.write_deprivation_report(enriched, ref['lookup_path'], ref['imd_path'], imd_version,
                                 DEPRIVATION_REPORT, where, group_col='GP_ID')

    # Postcode is already on the Patients table, so it isn't repeated here.
    deprivation_df = enriched.drop(columns=['Postcode'])

    patients_df = patients_df.copy()
    for field in PATIENT_IMD_FIELDS:
        col = f'IMD{imd_version}_{field}'
        patients_df[col] = enriched[col]

    note = (f"Deprivation: IMD {imd_version} - {counts['matched']} of {counts['provided']} "
            f"patients with a postcode matched (postcode lookup: {ref['lookup_path'].name}). "
            f"Decile on the Patients sheet; full fields and match status on the Deprivation sheet. "
            f"Unmatched postcodes: see qa_reports/{DEPRIVATION_REPORT.name}.")
    return patients_df, deprivation_df, note


def run(gp_sources, deprivation=False, imd_version=dep.IMD_VERSION):
    """Clean every GP file in gp_sources and save the combined CSVs + workbook.
    gp_sources: list of {'gp_id': ..., 'source_file': ...} (paths may be
    relative to the repository folder).
    deprivation: if True, add the IMD decile to the patients table and save a
    separate deprivation table (LSOA, match status, all chosen IMD fields)."""
    # Cheap check up front so a missing lookup/IMD file stops the run before
    # any cleaning, not after it.
    if deprivation:
        dep.preflight(imd_version, required_outputs=PATIENT_IMD_FIELDS)

    gp_sources = [
        {**gp, 'source_file': str(BASE_DIR / gp['source_file'])} for gp in gp_sources
    ]
    per_gp_results = [process_gp_file(gp['gp_id'], gp['source_file']) for gp in gp_sources]

    tables = {
        name: pd.concat([result[name] for result in per_gp_results], ignore_index=True)
        for name in TABLE_NAMES
    }
    patients_df = tables['patients']

    cross_gp_dupes = patients_df[patients_df.duplicated('Patient_ID', keep=False)]
    if not cross_gp_dupes.empty:
        n_patients = cross_gp_dupes['Patient_ID'].nunique()
        gp_ids_involved = sorted(cross_gp_dupes['GP_ID'].unique())
        print(
            f"NOTE: {n_patients} Patient_ID(s) appear under more than one GP_ID "
            f"({gp_ids_involved}) - review before treating Patient_ID as a unique key."
        )

    # ---------- 8.5 Deprivation indices (optional, --deprivation) ----------
    deprivation_csv = OUTPUT_DIR / 'copd_deprivation.csv'
    if deprivation:
        patients_df, deprivation_df, deprivation_note = build_deprivation_table(patients_df, imd_version)
        # Deprivation table goes straight after Patients (sheet order follows this dict).
        tables = {'patients': patients_df, 'deprivation': deprivation_df,
                  **{k: v for k, v in tables.items() if k != 'patients'}}
    else:
        deprivation_note = "Deprivation: not run (use run_pipeline.py --deprivation to add IMD fields)."
        # Don't leave deprivation outputs from an earlier run next to fresh outputs.
        for old in (deprivation_csv, DEPRIVATION_REPORT):
            if old.exists():
                old.unlink()
                print(f"Removed old {old.name} (deprivation not run this time).")

    # ---------- 9. Save + summary ----------
    # CSVs kept as-is for the pipeline/version control; the Excel workbook below
    # is the analyst-facing deliverable (NHS workers who work in Excel, not Python).
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for name, t in tables.items():
        t.to_csv(OUTPUT_DIR / f'copd_{name}.csv', index=False)
        print(f"{name}: {t.shape}")

    save_excel_workbook(tables, gp_sources, OUTPUT_XLSX, notes=[deprivation_note])

    print(f"\nTotal GPs processed: {len(gp_sources)}")
    print(f"Total patients: {len(patients_df)}")
    return tables


if __name__ == '__main__':
    try:
        run(GP_SOURCES, deprivation=RUN_DEPRIVATION, imd_version=IMD_VERSION)
    except ValueError as e:
        sys.exit(f"\nPIPELINE STOPPED:\n{e}")
