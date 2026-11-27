# Doccla_COPD_Evaluation

Cleaning code for the **Doccla COPD Evaluation data**. It takes the raw Excel extracts from each GP surgery, checks them, cleans them, and combines them into one set of tables (CSV files plus one formatted Excel workbook).

**Author:** Emily Costello
**Email:** e.e.costello@2019.ljmu.ac.uk

---

## Contents

- [Getting started](#getting-started)
- [Repository layout](#repository-layout)
- [1. Pre-run checks](#1-pre-run-checks)
- [2. Flatten headers](#2-flatten-headers)
- [3. Clean dataset](#3-clean-dataset)
- [4. Deprivation indices (optional)](#4-deprivation-indices-optional)
- [5. Outputs](#5-outputs)
- [6. Command reference and pipeline overview](#6-command-reference-and-pipeline-overview)

---

## Getting started

**Requirements:**
- Python 3.9 or later ([download here](https://www.python.org/downloads/) if you don't have it - tick "Add Python to PATH" during install)
- The Python packages `pandas` and `openpyxl`

**One-time setup:**
1. Download or clone this repository to your machine:
   ```
   git clone https://github.com/Em-Costello/Evaluation_cleaning_pipeline.git
   ```
2. Open a terminal (Command Prompt on Windows, Terminal on Mac) and go into the repository folder.
3. Install the required packages:
   ```
   pip install pandas openpyxl
   ```

**Running the pipeline:**
1. Put each surgery's raw Excel extract in `evaluation_data/raw/`, with its GP code in the file name, e.g. `WHC_GP01.xlsx`, `HARP_GP02.xlsx`. The code (`GP` followed by digits) becomes that surgery's `GP_ID`. The letters before it are just a reminder of which surgery it is. Two files can't share the same GP code.
2. Run:
   ```
   python run_pipeline.py
   ```
3. If anything needs fixing, the run stops **before cleaning anything** and prints one report covering every file. The report is also saved to `qa_reports/preflight_report.txt`. Raw values missing from the mapping files are listed in `qa_reports/unmapped_values.csv`. Fill in its `category` column, copy the `raw_text_lowercase,category` pairs into the mapping file named on each row, and run again. See [mapping_files/README.md](mapping_files/README.md).
4. When the checks pass, the pipeline flattens each extract, cleans and combines them, and saves the outputs to `cleaned_data/combined_data/`.

To run only the checks (e.g. when a new extract arrives): `python run_pipeline.py --check-only`.

To add a surgery, drop its file into `evaluation_data/raw/`. No code changes are needed.

To also add deprivation (IMD 2025): `python run_pipeline.py --deprivation`. The reference files it needs are included in the repository - see [4. Deprivation indices](#4-deprivation-indices-optional).

**Before running:**
- Close `copd_combined_output.xlsx` if it's open in Excel. The pipeline needs to overwrite it, and checks this at the start.
- All data stays on your local machine. Nothing is uploaded or sent anywhere by this pipeline.
- Raw data, flattened data, cleaned outputs and QA reports are all excluded from Git (see `.gitignore`), so patient data can't be committed by accident.

---

## Repository layout

| Path | What it is |
|---|---|
| `run_pipeline.py` | Runs everything in one command: checks -> flatten -> clean (-> deprivation). |
| `check_inputs.py` | Pre-run checks on the raw extracts ([section 1](#1-pre-run-checks)). |
| `flatten_headers.py` | Turns the two-row Excel header into a single row and saves a CSV ([section 2](#2-flatten-headers)). |
| `clean_copd_dataset.py` | All cleaning logic, plus the CSV and Excel output ([section 3](#3-clean-dataset)). |
| `deprivation.py` | Postcode -> LSOA -> IMD lookup ([section 4](#4-deprivation-indices-optional)). |
| `header_template.xlsx` | Reference copy of the expected raw extract layout ([section 2.1](#21-header-template-reference)). |
| `mapping_files/` | CSVs that translate raw GP values into standard categories. See [mapping_files/README.md](mapping_files/README.md). |
| `mapping_files/deprivation_indices/` | ONS postcode lookup, IMD 2025 file and `imd_field_mapping.csv` (only used with `--deprivation`). |
| `evaluation_data/raw/` | Where the raw extracts go (not tracked by Git). |
| `evaluation_data/flattened/` | Flattened CSVs, written by the pipeline. |
| `cleaned_data/combined_data/` | Final outputs. |
| `qa_reports/` | Check and deprivation reports (counts only, no identifiers). |

---

## 1. Pre-run checks

**`check_inputs.py`** (run automatically by `run_pipeline.py`)

Flattens each raw extract in memory and runs every check the cleaning step would. The difference is that it reports **all** problems across **all** files at once, rather than stopping at the first one:

- **File names**: every file must contain a GP code (e.g. `GP01`), and no two files can share one.
- **Columns**: missing, unexpected or out-of-order columns compared with `NEW_EXPECTED_COLUMNS`, with "did you mean" suggestions for renamed headers.
- **Mapping files**: every raw value with no mapping in every mapped field ([3.2](#32-clean-inconsistent-data-formats)), plus smoking statuses the parser doesn't recognise.
- **MRC codes**: any value that isn't graded 1-5.
- **Row structure**: anchor rows, whether the NHS Number is present, its checksum and duplicates, and patient-level fields appearing on detail rows ([3.3](#33-define-and-validate-the-anchor-variable-nhs-number--patient_id)). Problems are reported as line numbers in the flattened CSV. NHS numbers are never printed.
- **Dates**: values that can't be read as dates, COPD diagnosis dates before DOB, and DOBs in the future. These are warnings only: the run carries on and unreadable dates are left blank.
- **Profile**: patients, source rows, how complete the key patient fields are, and event counts per surgery. This helps spot an extract that looks wrong.

Each file is marked `OK`, `OK with warnings` or `BLOCKING ISSUES`. The pipeline only continues if no file has a blocking issue.

**Outputs:** `qa_reports/preflight_report.txt` and, if needed, `qa_reports/unmapped_values.csv`. The CSV has one row per unmapped value, with the mapping file it belongs in, the field, the number of rows, which GPs it was found in, and an empty `category` column to fill in. Neither file contains identifiers.

**Usage:**
```
python check_inputs.py                        # every .xlsx in evaluation_data/raw/
python check_inputs.py path/to/WHC_GP01.xlsx  # one file (or several, or a folder)
python run_pipeline.py --check-only           # same checks, via the main script
```

---

## 2. Flatten headers

**`flatten_headers.py`** (run automatically by `run_pipeline.py`)

The raw extracts have a two-row header with merged cells: a "group" row above a "sub-field" row. This script flattens that into a single row of clean column names and writes the result out as a CSV.

**Logic:**

- **Blank leading rows** before the header are skipped automatically.
- **Row 1 - group labels** (e.g. "COPD Diagnosis"), often merged across several columns. Each group label is filled across all the columns it spans.
- **Row 2 - sub-field labels** (e.g. "Code Term", "Date"). Any column with a blank sub-field label is treated as unused and dropped.
- **Column names**: each kept column becomes `"<Group> - <Subfield>"`, e.g. `"COPD Diagnosis - Code Term"`. Groups listed in `NO_PREFIX_GROUPS` don't get the prefix, so `"Patient Details - NHS Number"` becomes just `"NHS Number"`.
- **Dates** stored as real Excel dates are written out as `YYYY-MM-DD` text.
- **Data rows** start after the header rows and are written out using only the kept columns.

When run through `run_pipeline.py`, each flattened file is saved as `evaluation_data/flattened/<raw file name>.csv`.

**Usage (on its own):**
```
python flatten_headers.py input.xlsx output.csv
python flatten_headers.py input.xlsx output.csv --sheet "Sheet1"
python flatten_headers.py input.xlsx output.csv --header-rows 2
python flatten_headers.py input.xlsx output.csv --start-row 5
```

| Flag | Description |
|---|---|
| `--sheet <sheet_name>` | Which sheet to flatten (default: the active sheet). Also available on `run_pipeline.py` and `check_inputs.py`. |
| `--header-rows <1 or 2>` | Number of header rows (default: 2). |
| `--start-row <number>` | Row on which the data starts, if it can't be worked out automatically. |

### 2.1 Header template reference

**`header_template.xlsx`**

Shows the header layout and column order that the pipeline expects, laid out exactly as it should appear in a raw GP extract. **This is specific to the Doccla evaluation data.**

Use it to:
- Check whether a new GP surgery's extract matches the expected layout before running the pipeline on it.
- Work out why a column you expected has been dropped (usually a blank sub-field label) or has an unexpected name.
- Confirm column order and naming when requesting data from a GP surgery, so the extract arrives in a format the pipeline can handle without changes.

This file is a reference only. No script reads it.

---

## 3. Clean dataset

**`clean_copd_dataset.py`**

Normally run through `run_pipeline.py`, which builds the list of GP files automatically from `evaluation_data/raw/`. If run directly (`python clean_copd_dataset.py`), it uses the `GP_SOURCES` list at the top of the script instead, which points at already-flattened CSVs:

```python
GP_SOURCES = [
    {'gp_id': 'GP01', 'source_file': 'evaluation_data/flattened/WHC_GP01.csv'},
    {'gp_id': 'GP02', 'source_file': 'evaluation_data/flattened/HARP_GP02.csv'},
    ...
]
```

When run directly, deprivation is switched on with `RUN_DEPRIVATION = True` at the top of the script. With `run_pipeline.py`, use the `--deprivation` flag instead.

All per-surgery cleaning lives in `process_gp_file(gp_id, source_file)`, which runs once per surgery. Every error it raises starts with the GP ID (e.g. `[GP03] ...`), so a failure can be traced to the right file even when several surgeries are processed together.

The mapping files ([3.2](#32-clean-inconsistent-data-formats)) are loaded once at the start and shared across all surgeries.

### 3.1 Load and validate the data structure

Each flattened file must have exactly the columns in `NEW_EXPECTED_COLUMNS`, in the same order, or the run stops. This prevents columns being silently misaligned. Columns are then renamed to internal names using `RENAME_MAP`. The "Review of Inhaler Technique" columns are dropped, as they aren't used.

If the source file layout ever changes, `NEW_EXPECTED_COLUMNS` and `RENAME_MAP` are the two places to update.

### 3.2 Clean inconsistent data formats

#### Fields cleaned with a mapping file

GP extracts contain many ways of writing the same value. For example, "COPD" and "chronic obstructive pulmonary disease" need to become the same category. These fields are cleaned with an external mapping file:

| Field | Mapping file | Output categories |
|---|---|---|
| COPD diagnosis | `copd_diagnosis_mapping.csv` | `COPD`, `Emphysema`, `Asthma-COPD Overlap` |
| Gender | `gender_mapping.csv` | `Male`, `Female` |
| Ethnicity | `ethnicity_mapping.csv` | Gov.UK census categories, e.g. `White - British`, `Asian or Asian British - Indian`, plus `Not Stated` |
| Cardiovascular disease | `cardiovascular_disease_mapping.csv` | `Heart failure`, `Ischaemic heart disease` |
| Pulmonary rehab referral | `pulmonary_rehab_status_mapping.csv` | `Referred`, `Declined`, `Completed`, `Excluded` (see [3.8](#38-pulmonary-rehab-referral-counts)) |

Raw values are lower-cased and trimmed before lookup. For every field except pulmonary rehab, commas are removed as well. All mapped fields are listed in one place (`MAPPED_FIELDS`), which both the cleaning step and the pre-run checks use, so the two always apply the same rules.

**If any raw value isn't in its mapping file, the run stops** and lists every unmapped value across every field, rather than guessing or dropping it. To fix this, add the new values to the mapping file. See [mapping_files/README.md](mapping_files/README.md).

For cardiovascular disease, a blank cell means no recorded diagnosis and is left blank.

#### MRC code

The raw values are formatted inconsistently, but the grade number is always there. The script extracts just the number and checks it is between 1 and 5. Any other value stops the run.

#### Smoking status

Smoking status is cleaned by keyword rules (`classify_smoking_status`) rather than a mapping file, because raw entries often combine several pieces of information (e.g. status plus pack-years) rather than matching a fixed list of phrases.

- **Non-smokers**: entries containing "non-smoker" or "never smoked" become `Non-Smoker`.
- **What counts as smoking-related**: the word "smoker", rolling/hand-rolled cigarettes, cigars, or cigarette consumption / pack-years.
- **Current or former**: `Former` if the text contains an "ex-" pattern, otherwise `Current`.
- **How much**: mapped from keywords to one of five bands: `Very Heavy` (40+/day), `Heavy` (20-39/day), `Moderate` (10-19/day), `Light` (1-9/day), `Trivial` (<1/day). For example: `Former Heavy Smoker (20-39/day)`. If the amount isn't given, the result is `<Current/Former> Smoker (Amount Not Specified)` rather than a guess.
- **Anything else stops the run.** Every non-blank entry must be actively classified. If a new phrasing appears, the rules in `classify_smoking_status` need extending.

#### Depression / anxiety

Any non-blank value becomes `Yes`. Blank values are left blank rather than set to "No": the absence of a recorded diagnosis doesn't confirm the patient doesn't have the condition, so it is kept as unknown.

### 3.3 Define and validate the anchor variable (NHS Number / Patient_ID)

In the raw extract, each patient's data takes up a block of rows. The first is an **anchor row** holding the patient-level fields, identified by `Age`, `Postcode` and `DOB` all being filled. It is followed by zero or more detail rows holding only event-level data (prescriptions, consultations, etc.). The script:

1. Checks the first row of the file is an anchor row.
2. Checks the NHS Number is filled on every anchor row and only on anchor rows.
3. Validates every NHS Number against the **NHS Modulus 11 checksum**. Failures stop the run, without printing the numbers themselves.
4. Checks for duplicate NHS Numbers **within** the same file. This stops the run, as a duplicate within one surgery's extract is likely a data error.
5. Checks that no patient-level field (`PATIENT_LEVEL_COLS`) is filled on a detail row. If one is, the file doesn't match the expected structure and the run stops.
6. Copies each anchor row's NHS Number down onto its detail rows, to form `Patient_ID`.

The NHS Number is used directly as `Patient_ID`. Each row is also tagged with a separate `GP_ID` column (`GP01`, `GP02`, ...), so `Patient_ID` stays a clean, valid NHS Number throughout.

After all surgeries are combined, a `Patient_ID` that appears under more than one `GP_ID` is **not** an error. A note is printed instead, since this can be a patient who moved between surgeries. If this happens, `Patient_ID` on its own isn't a unique key, so join tables on `Patient_ID` **and** `GP_ID`.

### 3.4 Build the event tables

Once patients are anchored, the event rows are split out into separate tables, each linked back to the patient by `Patient_ID` and `GP_ID`:

| Table | Columns (besides `Patient_ID`, `GP_ID`) |
|---|---|
| Inhaler prescriptions | `Inhaler_Name_Dosage_Quantity`, `Inhaler_Dose`, `Inhaler_Issue_Date` |
| Prednisolone courses | `Prednisolone_Name_Dosage_Quantity`, `Prednisolone_Dose`, `Prednisolone_Issue_Date` |
| Antibiotic courses | `Antibiotic_Name_Dosage_Quantity`, `Antibiotic_Dose`, `Antibiotic_Issue_Date` |
| Consultations | `Consultation_Type`, `Consultation_Date` |
| Referrals (pulmonary rehab) | `Referral_to_Pulmonary_Rehab` (raw text), `Referral_to_Pulmonary_Rehab_Date` |
| Influenza vaccinations | `Influenza_Vaccine`, `Influenza_Vaccine_Date` |
| Pneumococcal vaccinations | `Pneumococcal_Vaccine`, `Pneumococcal_Vaccine_Date` |
| COVID vaccinations | `COVID_Vaccine`, `COVID_Vaccine_Date` |

A row goes into a table if any of that table's columns is filled. Together with the Patients table, this gives **nine** tables per run (ten with `--deprivation`).

### 3.5 Prescription, course and consultation counts

Four count columns are added to the Patients table: `Inhaler_Prescription_Count`, `Prednisolone_Course_Count`, `Antibiotic_Course_Count` and `Consultation_Count`. Each is the number of rows for that patient in the matching event table. Patients with no events get `0` rather than a blank, since "no prescriptions" is a known value.

### 3.6 Evidence of rescue pack

`Evidence_of_Rescue_Pack` counts distinct rescue pack events per patient. A rescue pack is a home supply of prednisolone and antibiotics that a patient keeps for treating a flare-up themselves.

Two signals are combined into one count:

- **Text mention**: the dose columns (`Inhaler_Dose`, `Prednisolone_Dose`, `Antibiotic_Dose`) are searched for "rescue pack" (any capitalisation), and the issue date on that row is taken as the event date.
- **Prescribing pattern**: prednisolone and antibiotics issued to the same patient on the same date, even with no text mention. Inhalers are left out of this signal, because they are routine maintenance medication reissued regularly regardless of flare-ups, so including them would add events that aren't rescue packs.

Each event is a (patient, date) pair. If both signals flag the same patient on the same date, it is counted **once**. Patients with no evidence get `0`.

The three issue-date columns are parsed into real dates at this point, before the main date step ([3.9](#39-parse-dates)). Otherwise the same date written in two different formats wouldn't match.

> This is a proxy indicator that depends on free-text wording at the point of prescribing and on same-day issuing, so it may under- or over-count actual rescue pack use.

### 3.7 Vaccination flags

Three columns are added to the Patients table: `Influenza_Vaccinated`, `Pneumococcal_Vaccinated` and `COVID_Vaccinated`. Each is `Yes` if the patient has at least one row in the matching vaccination table. Otherwise it is left **blank**, because no vaccination record isn't confirmation that the patient wasn't vaccinated.

### 3.8 Pulmonary rehab referral counts

A patient can have several pulmonary rehab entries. Each raw entry is mapped with `pulmonary_rehab_status_mapping.csv` to one of `Referred`, `Declined`, `Completed` or `Excluded`, and unmapped values stop the run as in [3.2](#32-clean-inconsistent-data-formats).

`Excluded` is for entries that aren't a clear referral outcome. For example:
- "pulmonary rehabilitation programme not available" is an availability issue, not something the patient did.
- "assessment for pulmonary rehabilitation completed" means the assessment was completed, not the programme.
- "pulmonary rehabilitation programme commenced" is in progress, so the outcome isn't known yet.

These are mapped to `Excluded` on purpose, rather than left out of the mapping file, so that every value is accounted for.

Three count columns are added to the Patients table: `Number_of_Times_Referred`, `Number_of_Times_Declined` and `Number_of_Times_Completed`. Each counts how many of that patient's entries mapped to that category. `Excluded` entries aren't counted, and patients with no entries get `0`.

The Referrals table keeps the original raw text, so the mapping can always be checked.

### 3.9 Parse dates

Every column in every table whose name contains `Date`, plus `DOB`, is converted to a proper date, so analysis doesn't need to handle mixed formats.

- Dates that were real Excel dates come out of `flatten_headers.py` as `YYYY-MM-DD` and are read as such.
- Text dates are read day-first (`DD/MM/YYYY`).

The two are handled separately on purpose. Reading `YYYY-MM-DD` dates as day-first swaps the day and month, and silently blanks any date where the day is above 12. Values that can't be read are left blank, and a warning gives the count per column.

### 3.10 Standardise text and sanity checks

On the Patients table:
- `MRC_Code`, `COPD_Diagnosis`, `Depression` and `Anxiety` are trimmed and put in title case.
- `Postcode` is trimmed and put in upper case.

As a sanity check (a warning, not a stop), the script prints a message for any surgery where a patient's `COPD_Diagnosis_Date` is before their `DOB`.

### 3.11 Combine and save

After every surgery is processed, each table is stacked across surgeries and saved. See [5. Outputs](#5-outputs). At the end the script prints the row and column count of each table, any cross-GP duplicate `Patient_ID` note, the number of GPs processed and the total number of patients.

---

## 4. Deprivation indices (optional)

Run with `python run_pipeline.py --deprivation` to look up each patient's Lower Layer Super Output Area (LSOA) and Index of Multiple Deprivation (IMD) from their postcode. The logic is in `deprivation.py`.

**Reference files** are stored in the repository, in `mapping_files/deprivation_indices/`, so they come down with everything else when you clone it:

| File | What it is | Example file name |
|---|---|---|
| ONS postcode lookup | Postcode -> 2021 LSOA. Kept as the `.zip` as downloaded (the pipeline reads it directly). | `PCD_OA21_LSOA21_MSOA21_LAD_FEB25_UK_LU.zip` |
| English Indices of Deprivation 2025 | LSOA -> IMD scores, ranks and deciles. | `IoD2025_All_Ranks_Scores_Deciles_Population_Denominators.csv` |
| `imd_field_mapping.csv` | Which IMD fields to add and what to call them. | - |

Both reference files are public data published under the [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/): the postcode lookup by the Office for National Statistics, and the Indices of Deprivation by the Ministry of Housing, Communities and Local Government.

**Updating the postcode lookup:** when ONS publishes a newer one, add the new `.zip` to the folder and **delete the old one in the same commit**, so the repository doesn't grow with every release. (If both are present, the pipeline picks the newest by the date in its file name, e.g. `FEB25`, so no code changes are needed either way.) The file used is printed in the run output and named on the workbook Summary sheet. Keep the lookup zipped: GitHub rejects files over 100 MB, and the extracted CSV is much larger than the zip.

**How it works:**
1. Before anything else runs, the pipeline checks the reference files exist and the IMD version is supported. A missing file stops the run straight away rather than after cleaning.
2. After all surgeries are cleaned and combined, each patient's postcode is cleaned (upper case, all spaces removed) and matched against the lookup, which is cleaned the same way. Matching doesn't depend on how the postcode is spaced.
3. A value that is already an LSOA code (e.g. `E01000001`) is used directly.
4. The LSOA is matched to the IMD file and the chosen fields are added.

**Outputs:**
- **Patients table**: gains one column, `IMD2025_Decile` (1 = most deprived 10% of areas in England, 10 = least deprived). Blank if the patient couldn't be matched; the Deprivation table says why.
- **Deprivation table**: `copd_deprivation.csv` and a **Deprivation** sheet (straight after Patients, with its own row on the Summary sheet). One row per patient, joined to Patients on `Patient_ID` + `GP_ID`. Postcode isn't repeated here, since it's already on Patients. Columns:
  - `Patient_ID`, `GP_ID`
  - `LSOA21CD`: the patient's 2021 LSOA code.
  - `Deprivation_Input_Type`: what the postcode value looked like - `Postcode`, `LSOA`, `Invalid` or `Missing`.
  - `Deprivation_Match_Status`: whether IMD fields were found, and if not, why - `Matched`, `Matched (terminated postcode)`, `Missing`, `Invalid format`, `Postcode not found in lookup`, `Outside England (no English IMD)`, `LSOA not found in IMD file`.
  - `IMD2025_Score`, `IMD2025_Rank`, `IMD2025_Decile` by default (for rank and decile, 1 = most deprived), plus any other fields switched on in `imd_field_mapping.csv`. The year prefix stops fields from different IMD versions being confused.
- **`qa_reports/deprivation_report.txt`**: match counts by status and by GP. For each kind of non-match it explains what it usually means, what to do, and how to find those patients: filter `Deprivation_Match_Status` on the Deprivation sheet, then look up the `Patient_ID` on the Patients sheet to see the postcode. Counts only, no postcodes or NHS numbers.

If the pipeline is run **without** `--deprivation`, there is no Deprivation sheet, and any `copd_deprivation.csv` or `deprivation_report.txt` from an earlier run is deleted so it can't be mistaken for current output.

**Notes:**
- Postcodes that can't be matched never stop the pipeline. They are flagged in `Deprivation_Match_Status` and their IMD fields are left blank.
- Terminated (no longer used) postcodes are still matched, because GP records often hold older addresses. They are flagged as `Matched (terminated postcode)`.
- The IMD covers England only, so Welsh, Scottish, Northern Irish, Channel Islands and Isle of Man postcodes are flagged as `Outside England`.
- The run prints counts by match status and the match rate per GP, with a warning if fewer than 90% of patients with a postcode were matched. No postcodes or other patient-level values are printed.
- To add more IMD fields (income, employment, health, crime, etc.), open `imd_field_mapping.csv` and change `include` from `No` to `Yes` for the fields you want. They are added to the Deprivation table. `Decile` must stay `Yes`, because it's the field copied onto Patients; the run stops at the start if it isn't.
- Only IMD 2025 is supported for now. IMD 2019 uses 2011 LSOAs, which the current postcode lookup doesn't provide.

**Running on its own:** `deprivation.py` also works on any CSV/Excel file with a postcode or LSOA column, which is useful for testing:
```
python deprivation.py --input <file> --column Postcode
python deprivation.py --input <file> --column Postcode --id-column none   # keep every row as-is
```
By default the input is reduced to one row per patient using the `NHS Number` column; use `--id-column none` for a plain list of postcodes. The output goes to `cleaned_data/deprivation/` (or the path given with `--output`), and the report to `qa_reports/<input name>_deprivation_report.txt`.

---

## 5. Outputs

Everything is saved to `cleaned_data/combined_data/`, in two forms:

- **CSV files**, one per table: `copd_<table_name>.csv`. Use these for further analysis in Python or R; join tables on `Patient_ID` + `GP_ID`.
- **Excel workbook**: `copd_combined_output.xlsx`. This is the main deliverable for analysts who work in Excel. It contains:
  - A **Summary** sheet first: the date it was generated, the GP surgeries included, whether deprivation was run (and the match rate if so), and a table of each data sheet with its row count and a one-line description.
  - One sheet per table, each with a bold white-on-dark-blue header row, the header frozen, filters on every column, dates shown as `DD/MM/YYYY`, and column widths fitted to the content (between 10 and 45 characters).

| Sheet | CSV | One row per |
|---|---|---|
| Patients | `copd_patients.csv` | patient |
| Deprivation (with `--deprivation`) | `copd_deprivation.csv` | patient |
| Inhaler Prescriptions | `copd_inhaler_prescriptions.csv` | inhaler prescription issued |
| Prednisolone Courses | `copd_prednisolone_courses.csv` | prednisolone course issued |
| Antibiotic Courses | `copd_antibiotic_courses.csv` | antibiotic course issued |
| Consultations | `copd_consultations.csv` | consultation |
| Referrals - Pulm Rehab | `copd_referrals_pulmonary_rehab.csv` | pulmonary rehab entry |
| Influenza Vaccinations | `copd_influenza_vaccinations.csv` | influenza vaccination record |
| Pneumococcal Vaccs | `copd_pneumococcal_vaccinations.csv` | pneumococcal vaccination record |
| COVID Vaccinations | `copd_covid_vaccinations.csv` | COVID vaccination record |

### Patients table columns

| Column | Meaning | Section |
|---|---|---|
| `Patient_ID` | NHS Number | [3.3](#33-define-and-validate-the-anchor-variable-nhs-number--patient_id) |
| `GP_ID` | Surgery code from the file name, e.g. `GP01` | [3.3](#33-define-and-validate-the-anchor-variable-nhs-number--patient_id) |
| `Postcode`, `Age`, `DOB` | As recorded; postcode in upper case | [3.10](#310-standardise-text-and-sanity-checks) |
| `Gender`, `Ethnicity` | Mapped categories | [3.2](#32-clean-inconsistent-data-formats) |
| `MRC_Code`, `MRC_Code_Date` | MRC grade 1-5 | [3.2](#32-clean-inconsistent-data-formats) |
| `COPD_Diagnosis`, `COPD_Diagnosis_Date` | Mapped category | [3.2](#32-clean-inconsistent-data-formats) |
| `Depression`, `Depression_Date`, `Anxiety`, `Anxiety_Date` | `Yes` or blank (unknown) | [3.2](#32-clean-inconsistent-data-formats) |
| `Smoking_Status`, `Smoking_Status_Date` | Current/former and amount | [3.2](#32-clean-inconsistent-data-formats) |
| `Cardiovascular_Disease`, `Cardiovascular_Disease_Date` | `Heart failure`, `Ischaemic heart disease`, or blank | [3.2](#32-clean-inconsistent-data-formats) |
| `Inhaler_Prescription_Count`, `Prednisolone_Course_Count`, `Antibiotic_Course_Count`, `Consultation_Count` | Event counts (0 if none) | [3.5](#35-prescription-course-and-consultation-counts) |
| `Evidence_of_Rescue_Pack` | Number of rescue pack events (0 if none) | [3.6](#36-evidence-of-rescue-pack) |
| `Influenza_Vaccinated`, `Pneumococcal_Vaccinated`, `COVID_Vaccinated` | `Yes` or blank | [3.7](#37-vaccination-flags) |
| `Number_of_Times_Referred`, `Number_of_Times_Declined`, `Number_of_Times_Completed` | Pulmonary rehab counts (0 if none) | [3.8](#38-pulmonary-rehab-referral-counts) |
| `IMD2025_Decile` (with `--deprivation`) | 1 = most deprived, 10 = least | [4](#4-deprivation-indices-optional) |

---

## 6. Command reference and pipeline overview

```bash
python run_pipeline.py                     # checks -> flatten -> clean, in one go
python run_pipeline.py --check-only        # checks only
python run_pipeline.py --deprivation       # also add IMD 2025 (decile on Patients + Deprivation table)
python run_pipeline.py --raw-dir <folder>  # raw extracts somewhere other than evaluation_data/raw/
python run_pipeline.py --sheet "Sheet1"    # read a specific sheet of each extract
```

The individual steps can still be run on their own:

```bash
python check_inputs.py <input_file_path>.xlsx
python flatten_headers.py <input_file_path>.xlsx <output_file_path>.csv
python clean_copd_dataset.py               # uses GP_SOURCES at the top of the script
python deprivation.py --input <input_file_path>.csv --column Postcode
```

```
Raw Excel files in evaluation_data/raw/  (one per surgery, GP code in the name)
      |
      v
run_pipeline.py
      |
      |-- (--deprivation) check reference files exist -> STOP if missing
      |-- check the output workbook isn't open in Excel
      |
      |-- STEP 1  check_inputs.py: file names, columns, mappings, MRC,
      |           row structure, dates, profile
      |             -> qa_reports/preflight_report.txt, unmapped_values.csv
      |             -> STOP here if anything is blocking
      |
      |-- STEP 2  flatten_headers.py: one flat CSV per surgery
      |             -> evaluation_data/flattened/
      |
      v
STEP 3  clean_copd_dataset.py  (process_gp_file, once per surgery)
      |
      |-- Validate columns, rename, map categorical fields
      |-- Classify smoking status, extract MRC grade
      |-- Anchor on NHS Number -> Patient_ID, tag with GP_ID
      |-- Build Patients table + 8 event tables
      |-- Add counts, Evidence_of_Rescue_Pack, vaccination flags,
      |   pulmonary rehab counts
      |-- Parse dates, standardise text, sanity checks
      |
      v
Combine across all surgeries
      |
      |-- Note any cross-GP duplicate Patient_IDs
      |-- (--deprivation) Postcode -> LSOA -> IMD: decile onto Patients,
      |       full fields -> Deprivation table, deprivation_report.txt
      |
      v
Save: 9 CSVs + 1 formatted Excel workbook (Summary + 9 sheets)
      (10 of each with --deprivation)
```