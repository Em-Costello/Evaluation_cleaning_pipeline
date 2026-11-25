'''
I think the best way for this to work is to have a separate script that takes
an input (FROM A csv OR excel ?) of a single column with postcodes or LSOAs,
converts to dataframe and then matches to LSOA and then to IMD.
The output will be a csv with the original input and the IMD fields added as
separate columns.
'''

# =============================================================================
# PLANNED LOGIC
# =============================================================================
#
# -----------------------------------------------------------------------------
# 0. CONFIG
# -----------------------------------------------------------------------------
#   - IMD_VERSION = 2025   (allowed: 2019, 2025) - this only works
#   - INPUT_FILE / INPUT_COLUMN  -> which file and which column holds the
#     postcode or LSOA (default column: 'Postcode')
#   - IMD_FIELDS -> list of IMD columns to attach. Default to a small set
#     (overall IMD score, rank, decile) with an option to add more
#     (income, employment, health, etc.) rather than all ~55 columns.
#   - OUTPUT_FILE -> where the enriched CSV is written.
#   - Later, when merged into the pipeline, these become command line
#     arguments (argparse), e.g. --deprivation --imd-version 2025.
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
#   - Final lookup structure: a pandas Series / dict  key -> lsoa21cd.
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
#   - Store this as an Input_Type column: Postcode / LSOA / Invalid / Missing.
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
#   - Keep a small config dict per version, e.g.
#       IMD_CONFIG = {
#         2025: {'file_pattern': 'IoD2025_*.csv',
#                'lsoa_col': 'LSOA code (2021)', 'lsoa_year': 2021},
#         2019: {'file_pattern': 'File_7_*IoD2019*.csv',
#                'lsoa_col': 'LSOA code (2011)', 'lsoa_year': 2011},
#       }
#   - IMPORTANT: IMD 2019 is published on 2011 LSOAs; IMD 2025 on 2021 LSOAs.
#     The current postcode lookup only gives lsoa21cd. So for 2019 we would
#     need either an older lookup with lsoa11cd, or the ONS LSOA11 -> LSOA21
#     best-fit lookup. Decision needed - possibly only support 2025 at first
#     and raise a clear error if 2019 is requested without the right lookup.
#   - Rename the long IMD headers to short pipeline-style names, prefixed
#     with the version so outputs from different versions can't be confused:
#       'Index of Multiple Deprivation (IMD) Score'  -> IMD2025_Score
#       '... (IMD) Rank (where 1 is most deprived)'  -> IMD2025_Rank
#       '... (IMD) Decile (where 1 is most ...)'     -> IMD2025_Decile
#     Do this through a mapping CSV in mapping_files/ (consistent with the
#     rest of the pipeline: mappings external, not hard-coded).
#   - Required-column check against IMD_FIELDS; hard error if missing.
#   - Index by LSOA code once (set_index) and reuse, rather than calling
#     set_index separately for every field as the current test code does.
#
# -----------------------------------------------------------------------------
# 8. LSOA -> IMD
# -----------------------------------------------------------------------------
#   - Left-merge the input onto the IMD table on LSOA (one merge attaches
#     all requested fields at once).
#   - Rows with an LSOA but no IMD match:
#       LSOA starts with W01 (Wales) or postcode maps to a Scottish / NI /
#       Welsh area -> status 'Outside England (no English IMD)'
#       otherwise  -> status 'LSOA not found in IMD file'
#
# -----------------------------------------------------------------------------
# 9. STATUS / FLAG COLUMN (flag, don't block)
# -----------------------------------------------------------------------------
#   One column, e.g. Deprivation_Match_Status, with a fixed set of values:
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
#   - Do NOT print postcodes, LSOAs or any row-level data (the current test
#     prints .head() of Postcode - remove before merging).
#   - Optional warning threshold: if match rate < e.g. 90%, print a loud
#     warning (suggests a wrong column or format issue) but still continue.
#
# -----------------------------------------------------------------------------
# 11. OUTPUT
# -----------------------------------------------------------------------------
#   - Columns: all original input columns + LSOA21CD + Input_Type +
#     Deprivation_Match_Status + chosen IMD fields. Drop helper columns
#     (cleaned key, Postcode_length).
#   - Write to CSV; once merged, also add as a sheet in the formatted
#     Excel workbook.
#
# -----------------------------------------------------------------------------
# 12. STRUCTURE FOR MERGING INTO THE PIPELINE LATER
# -----------------------------------------------------------------------------
#   Wrap each step in a function so the main pipeline can import them:
#       find_latest_lookup(folder)            -> Path
#       load_postcode_lookup(path)            -> Series (key -> lsoa21cd)
#       load_imd(version, folder, fields)     -> DataFrame indexed by LSOA
#       clean_and_classify(series)            -> DataFrame (clean, type)
#       add_deprivation(df, column, version)  -> df with IMD + status cols
#       summarise_matches(df)                 -> prints aggregate counts
#   - Load lookup + IMD ONCE outside any per-GP loop (they are large).
#   - In clean_copd_dataset.py: call add_deprivation(patients_df, 'Postcode', ...)
#     only when the --deprivation flag is set; prefix messages with [gp_id]
#     if run per GP.
#   - Input from Excel as well as CSV: choose reader by file suffix
#     (.csv -> read_csv, .xlsx -> read_excel), always dtype=str.
#
# =============================================================================

'''
Docstring for deprivation.py
'''

import sys
import re
from typing import Optional
import pandas as pd
from pathlib import Path
from datetime import date

# Read in flattened test data


# Read in the mapping file for postcode to LSOA mapping
#path = \mapping_files\deprivation_indices\PCD_OA21_LSOA21_MSOA21_LAD_FEB25_UK_LU.csv


# All paths are resolved relative to this script's folder, so the pipeline
# works no matter which folder the terminal is in.
BASE_DIR = Path(__file__).resolve().parent
TEST_DATA_DIR = BASE_DIR / 'evaluation_data/raw/'
FLAT_DATA_DIR = BASE_DIR / 'evaluation_data/flattened/'
MAPPING_DIR = BASE_DIR / 'mapping_files'
DEPRIVATION_MAPPING_DIR = MAPPING_DIR / 'deprivation_indices'

df_test_data = pd.read_csv(FLAT_DATA_DIR / 'TEST_GP00.csv')
# columns in above: NHS Number,Date of Birth,Age,Gender,Postcode,Ethnic Origin,Clinical Codes - Code Term,Clinical Codes - Date,COPD Diagnosis - Code Term,COPD Diagnosis - Date,Depression - Code Term,Depression - Date,Anxiety - Code Term,Anxiety - Date,Referral to Pulmonary Rehab - Code Term,Referral to Pulmonary Rehab - Date,Smoking Status - Code Term,Smoking Status - Date,Smoking Status - Associated Text,Review of Inhaler Technique - Code Term,Review of Inhaler Technique - Date,"Medication Issues - Name, Dosage and Quantity",Medication Issues - Date of Issue,Medication Issues - Dose,"Prednisolone - Name, Dosage and Quantity",Prednisolone - Date of Issue,Prednisolone - Dose,"Antibiotics - Name, Dosage and Quantity",Antibiotics - Date of Issue,Antibiotics - Dose,Consultations - Date,Consultations - Type of Consultation,Cardiovascular Disease / Heart Disease - Code Term,Cardiovascular Disease / Heart Disease - Date,Cardiovascular Disease / Heart Disease - Associated Text,Influenza - Code Term,Influenza - Date,Pneumococcal - Code Term,Pneumococcal - Date,COVID - Code Term,COVID - Date
# PLAN: read with dtype=str so postcodes/LSOAs are never coerced. Also note
# the flattened GP file has multiple rows per patient (event rows) - for the
# deprivation step, de-duplicate to one row per patient first (or run on
# patients_df once merged), otherwise we look up the same postcode many times.

df_PC_LSOA_mapping = pd.read_csv(DEPRIVATION_MAPPING_DIR / 'PCD_OA21_LSOA21_MSOA21_LAD_FEB25_UK_LU.csv')
#columns in above: pcd7,pcd8,pcds,dointr,doterm,usertype,oa21cd,lsoa21cd,msoa21cd,ladcd,lsoa21nm,msoa21nm,ladnm,ladnmw
# PLAN: replace hard-coded filename with find_latest_lookup() (step 1), read
# from the zip (step 2), usecols + dtype=str, lower-case the headers.

df_LSOA_IMD_mapping = pd.read_csv(DEPRIVATION_MAPPING_DIR / 'IoD2025_All_Ranks_Scores_Deciles_Population_Denominators.csv')
# columns in above: LSOA code (2021),LSOA name (2021),Local Authority District code (2024),Local Authority District name (2024),Index of Multiple Deprivation (IMD) Score,Index of Multiple Deprivation (IMD) Rank (where 1 is most deprived),Index of Multiple Deprivation (IMD) Decile (where 1 is most deprived 10% of LSOAs),Income Score (rate),Income Rank (where 1 is most deprived),Income Decile (where 1 is most deprived 10% of LSOAs),Employment Score (rate),Employment Rank (where 1 is most deprived),Employment Decile (where 1 is most deprived 10% of LSOAs),"Education, Skills and Training Score","Education, Skills and Training Rank (where 1 is most deprived)","Education, Skills and Training Decile (where 1 is most deprived 10% of LSOAs)",Health Deprivation and Disability Score,Health Deprivation and Disability Rank (where 1 is most deprived),Health Deprivation and Disability Decile (where 1 is most deprived 10% of LSOAs),Crime Score,Crime Rank (where 1 is most deprived),Crime Decile (where 1 is most deprived 10% of LSOAs),Barriers to Housing and Services Score,Barriers to Housing and Services Rank (where 1 is most deprived),Barriers to Housing and Services Decile (where 1 is most deprived 10% of LSOAs),Living Environment Score,Living Environment Rank (where 1 is most deprived),Living Environment Decile (where 1 is most deprived 10% of LSOAs),Income Deprivation Affecting Children Index (IDACI) Score (rate),Income Deprivation Affecting Children Index (IDACI) Rank (where 1 is most deprived),Income Deprivation Affecting Children Index (IDACI) Decile (where 1 is most deprived 10% of LSOAs),Income Deprivation Affecting Older People (IDAOPI) Score (rate),Income Deprivation Affecting Older People (IDAOPI) Rank (where 1 is most deprived),Income Deprivation Affecting Older People (IDAOPI) Decile (where 1 is most deprived 10% of LSOAs),Children and Young People Sub-domain Score,Children and Young People Sub-domain Rank (where 1 is most deprived),Children and Young People Sub-domain Decile (where 1 is most deprived 10% of LSO,Adult Skills Sub-domain Score,Adult Skills Sub-domain Rank (where 1 is most deprived),Adult Skills Sub-domain Decile (where 1 is most deprived 10% of LSOAs),Geographical Barriers Sub-domain Score,Geographical Barriers Sub-domain Rank (where 1 is most deprived),Geographical Barriers Sub-domain Decile (where 1 is most deprived 10% of LSOAs),Wider Barriers Sub-domain Score,Wider Barriers Sub-domain Rank (where 1 is most deprived),Wider Barriers Sub-domain Decile (where 1 is most deprived 10% of LSOAs),Indoors Sub-domain Score,Indoors Sub-domain Rank (where 1 is most deprived),Indoors Sub-domain Decile (where 1 is most deprived 10% of LSOAs),Outdoors Sub-domain Score,Outdoors Sub-domain Rank (where 1 is most deprived),Outdoors Sub-domain Decile (where 1 is most deprived 10% of LSOAs),Total population: mid 2022,Dependent Children aged 0-15: mid 2022,Older population aged 60 and over: mid 2022,Working age population 18-66 (for use with Employment Deprivation Domain): mid 2022
# PLAN: replace with load_imd(IMD_VERSION, ...) (step 7) - version-driven
# file pattern, rename to short IMD2025_* names via a mapping CSV.


#print(df_PC_LSOA_mapping.head())

print(df_test_data['Postcode'].head())
# PLAN: remove before merging - prints row-level postcodes. Replace with
# aggregate counts only (step 10).

#strip whitespace from the postcode column in the test data
df_test_data['Postcode'] = df_test_data['Postcode'].str.strip() #this gives nans and seems to give 8 character postcodes
# PLAN: .str.strip() keeps existing NaNs as NaN (it doesn't create them) and
# only removes leading/trailing spaces, so internal spacing is unchanged -
# which is why lengths vary. Replace with the normalised key in step 3/5
# (upper-case, remove ALL whitespace) in a separate working column, keeping
# the original Postcode untouched.

#count number of characters in the postcode column in the test data
df_test_data['Postcode_length'] = df_test_data['Postcode'].str.len()
# PLAN: no longer needed once matching uses the normalised key; keep only as
# a temporary diagnostic and drop from the output.

print(df_test_data['Postcode_length'].value_counts())

#find the LSOA for each postcode in the test data using the mapping file
#remove nan postcodes from the test data before mapping
df_test_data = df_test_data[df_test_data['Postcode'].notna()]
# PLAN: don't drop rows - this removes patients from the output entirely.
# Keep them and set Deprivation_Match_Status = 'Missing' (step 9).

# Map postcodes to LSOAs based on their length
for length in [7, 8]:
    df_test_data.loc[df_test_data['Postcode_length'] == length, 'LSOA'] = df_test_data.loc[df_test_data['Postcode_length'] == length, 'Postcode'].map(df_PC_LSOA_mapping.set_index(f'pcd{length}')['lsoa21cd'])
# PLAN: replace this loop with a single map on the normalised key (step 3/6).
# Current version silently leaves LSOA blank for 6-char postcodes (e.g.
# 'A1 2BC') and for any input whose spacing differs from ONS's fixed-width
# format, with no flag saying why.

print(df_test_data['LSOA'].head())
# PLAN: remove before merging (row-level output).

# match lsoa to LSOA code (2021) and IMD using other mapping file
df_test_data['IMD'] = df_test_data['LSOA'].map(df_LSOA_IMD_mapping.set_index('LSOA code (2021)')['Index of Multiple Deprivation (IMD) Score'])
df_test_data['IMD_decile'] = df_test_data['LSOA'].map(df_LSOA_IMD_mapping.set_index('LSOA code (2021)')['Index of Multiple Deprivation (IMD) Decile (where 1 is most deprived 10% of LSOAs)'])
# PLAN: replace with one left-merge on LSOA that attaches all IMD_FIELDS at
# once, with version-prefixed names, then set the status column (step 8/9).
print(df_test_data[['Postcode', 'LSOA', 'IMD', 'IMD_decile']].head())
# PLAN: remove before merging (row-level output); write CSV instead (step 11).