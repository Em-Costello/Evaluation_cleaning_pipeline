# Doccla_COPD_Evaluation

Cleaning code for the **DOCCLA COPD Evaluation data**.

## Data Cleaning and Preparation Methodology

**Author:** Emily Costello
**Email:** e.e.costello@2019.ljmu.ac.uk

---
## Getting Started

**Requirements:**
- Python 3.9 or later ([download here](https://www.python.org/downloads/) if you don't have it - tick "Add Python to PATH" during install)
- The following Python packages: `pandas`, `openpyxl`

**One-time setup:**
1. Download or clone this repository to your machine.
```
git clone https://github.com/Em-Costello/Evaluation_cleaning_pipeline.git
```
2. Open a terminal (Command Prompt on Windows, Terminal on Mac) and navigate into the repository folder.
3. Install the required packages:
```
pip install pandas openpyxl
```

**Running the pipeline:**
1. Put each surgery's raw Excel extract in `evaluation_data/raw/`, with its GP code in the file name - e.g. `WHC_GP01.xlsx`, `HARP_GP02.xlsx`. The code (`GP01`) becomes that surgery's `GP_ID`.
2. Run:
```
python run_pipeline.py
```
3. If anything needs fixing, the run stops before cleaning anything and prints one report covering every file (also saved to `qa_reports/preflight_report.txt`). Any raw values missing from the mapping files are listed in `qa_reports/unmapped_values.csv` - fill in the `category` column, copy the `raw_text_lowercase,category` pairs into the mapping file named on each row, and run again. See [mapping_files/README.md](mapping_files/README.md).
4. When the checks pass, the pipeline flattens each extract, cleans and combines them, and saves the outputs to `cleaned_data/combined_data/`.

To run only the checks (e.g. when a new extract arrives): `python run_pipeline.py --check-only`. To add a surgery, drop its file into `evaluation_data/raw/` - no code changes needed.

**Before running:**
- Make sure `copd_combined_output.xlsx` isn't currently open in Excel — the script needs to overwrite it, which will fail if it's open elsewhere.
- All data stays on your local machine; nothing is uploaded or sent anywhere by this pipeline.

---
## 1. Pre-run checks

**`check_inputs.py`** (run automatically by `run_pipeline.py`)

Flattens each raw extract in memory and runs every check the cleaning step would, but reports **all** problems across **all** files at once rather than stopping at the first one:

- **Columns** — missing, unexpected or out-of-order columns compared with `NEW_EXPECTED_COLUMNS`, with "did you mean" suggestions for renamed headers.
- **Row structure** — anchor rows, NHS Number presence, checksum and duplicates, and patient-level fields appearing on detail rows (reported as line numbers in the flattened CSV; NHS numbers are never printed).
- **Mapping files** — every raw value with no mapping, in every mapped field, plus smoking statuses the parser doesn't recognise.
- **Dates** — values that can't be read as dates (warning only; they're left blank), COPD diagnosis before DOB, DOB in the future.
- **Profile** — patients, source rows, completeness of key patient fields and event counts per surgery, to spot an extract that looks wrong.

Outputs `qa_reports/preflight_report.txt` and, if needed, `qa_reports/unmapped_values.csv` (same `raw_text_lowercase,category` layout as the mapping files). Neither contains identifiers.

**Usage:**
```
python check_inputs.py                       # every .xlsx in evaluation_data/raw/
python check_inputs.py path/to/WHC_GP01.xlsx
```
or 
```
python run_pipeline.py --check_only
---

## 2. Flatten Headers

**`flatten_headers.py`**

Flattens a two-row merged-cell Excel header (a "group" row sitting above a "sub-field" row) into a single row of clean column names, and writes the result out as a CSV.

**Logic:**

- **Row 1 — group labels** (e.g. "COPD Diagnosis"), often merged across several columns. Each merged group label is forward-filled across all the columns it spans.
- **Row 2 — sub-field labels** (e.g. "Code Term", "Date"). Any column whose row-2 value is blank is treated as an empty/unused column and is dropped entirely.
- **Flattened header naming** — for a kept column, the header becomes `"<Group> - <Subfield>"`, e.g. `"COPD Diagnosis - Code Term"`. Groups listed in `NO_PREFIX_GROUPS` are exempt from this prefix, since prefixing them adds noise rather than clarity (e.g. `"Patient Details - NHS Number"` becomes just `"NHS Number"`).
- **Data rows** start after the header rows (default: row 3 onward) and are written to the output CSV using only the kept columns.

**Usage:**
```
python flatten_headers.py input.xlsx output.csv
python flatten_headers.py input.xlsx output.csv --sheet "Sheet1"
python flatten_headers.py input.xlsx output.csv --header-rows 2
python flatten_headers.py input.xlsx output.csv --start-row 5
```

| Flag | Description |
|---|---|
| `--sheet <sheet_name>` | Specifies which sheet of the input file to flatten |
| `--header-rows <number>` | Specifies the number of header rows explicitly |
| `--start-row <number>` | Specifies the row on which data starts, explicitly |

---
### 2.1 Header Template Reference

**`headers_template.xlsx`**

Shows the expected header structure and column order that `flatten_headers.py` and `clean_copd_dataset.py` are built to handle — the two-row group/sub-field layout described above, laid out exactly as it should appear in a raw GP extract. **This is specific to the Doccla evaulation data**.

Use this as a reference when:
- Checking whether a new GP surgery's extract matches the expected structure before running the pipeline on it.
- Working out why `flatten_headers.py` has dropped a column you expected to keep (likely a blank sub-field label — see the logic above) or produced a header name you didn't expect.
- Confirming column order and naming if you're preparing a data request to a GP surgery, so the extract arrives in a format the pipeline can handle without modification.

This file is a reference only — it isn't read by any script in the pipeline.

---

## 3. Clean Dataset

**`clean_copd_dataset.py`**

Normally run via `run_pipeline.py`, which builds the list of GP files automatically. When run directly (`python clean_copd_dataset.py`), it uses the `GP_SOURCES` list at the top of the script, pointing at already-flattened CSVs:

```python
GP_SOURCES = [
    {'gp_id': 'GP01', 'source_file': 'evaluation_data/gp01_flattened_input.csv'},
    {'gp_id': 'GP02', 'source_file': 'evaluation_data/gp02_flattened_input.csv'},
]
```

All per-file cleaning logic lives in `process_gp_file(gp_id, source_file)`, which is run once per entry in `GP_SOURCES`. To add a new surgery, add another data file to <evaluation_data/raw/> using the naming convention 'X_GPID.xslx', where X is a combination of letters that tell you intuitively which GP it is, and a ID is a combination of digits. Ensure that this dataset matches the expected column structure. The pipeline will detect the ID, and every error raised inside `process_gp_file` is prefixed with the `gp_id`, so a failure is immediately traceable to the offending file even when several GPs are processed in one run.

The categorical mapping files (COPD diagnosis, gender, ethnicity — see 3.2) are loaded once, outside `process_gp_file`, and shared read-only across all GP files in the run.

### 3.1 Load and validate data structure

For each GP file, the script expects columns to appear in a specific order with specific names (`NEW_EXPECTED_COLUMNS`), and errors out on any mismatch. This guards against silent misalignment, since some column renaming is done by position rather than by name. Columns are then renamed to internal names via `RENAME_MAP`, and the unused columns are dropped.

### 3.2 Clean inconsistent data formats

#### COPD Diagnosis / Gender / Ethnicity

GP inputs contain many variations that refer to the same underlying value — e.g. for COPD diagnosis, "COPD" and "chronic obstructive pulmonary disease" both need to resolve to the same output. Each of these fields is cleaned using an external mapping file, loaded once and shared across all GP surgeries. The mapped fields are listed in one place (`MAPPED_FIELDS`), used by both the cleaning step and the pre-run checks. Any input not found in a mapping file causes the script to stop and list **every** unmapped value across all fields, rather than guessing or dropping the value:

- **COPD diagnosis** → `mapping_files/copd_diagnosis_mapping.csv`
- **Gender** → `mapping_files/gender_mapping.csv`
- **Ethnicity** → `mapping_files/ethnicity_mapping.csv` (maps to the Gov.UK 2022 census categories; unmapped values prompt a check)

#### MRC Code

GP inputs contain formatting variations here too, but the number grade itself is always present and consistent. The script extracts just the number grade from the raw value and validates it falls within the expected 1–5 range, erroring loudly otherwise.

#### Smoking Status

Smoking status is cleaned using rule-based keyword parsing (`classify_smoking_status`) rather than a strict lookup mapping, since raw entries often combine multiple pieces of information (status plus pack-years) rather than matching a fixed list of phrases. Any input pattern the parser doesn't recognise causes a hard error, prompting the parser to be extended, rather than silently dropping or misclassifying the value.

- **Non-smokers**: entries matching "non-smoker" or "never smoked" are classified as `Non-Smoker`.
- **Timing**: classified as `Former` if the text contains an "ex-" pattern, otherwise `Current`.
- **Heaviness band**: mapped from keywords in the text to one of five bands: `Very Heavy` (40+/day), `Heavy` (20-39/day), `Moderate` (10-19/day), `Light` (1-9/day), `Trivial` (<1/day). If a smoking-related entry doesn't specify heaviness, it's recorded as `<Timing> Smoker (Amount Not Specified)` rather than being dropped or guessed.
- **Alternative smoking habits**: entries referring to cigars, or to rolling/hand-rolled cigarettes, are also recognised as smoking-related even without the word "smoker".
- No input data is left as-is or interpreted as "No" by default — every non-blank entry is actively classified.

#### Depression / Anxiety

Any non-blank input in the depression or anxiety columns is standardised to `"Yes"`. Blank input is left as-is (`NaN`) rather than being interpreted as "No," since the absence of a recorded diagnosis doesn't confirm the patient doesn't have the condition — it's preserved as genuinely unknown.

#### Cardiovascular disease

Non-blank inputs in the cardiovascular disease column is standardised to either `Heart Failure` or `Ischaemic heart disease`. Any non-blank input *has* some form of cardiovascular disease, so can be treated as `yes`; blank cells are treated as having no diagnosis, and are left blank. 

### 3.3 Define and validate anchor variable (NHS number / Patient_ID)

Each patient's data occupies a block of rows: one **anchor row** carrying patient-level fields (`Age`, `Postcode`, `DOB` all present), followed by zero or more detail rows carrying only event-level data. The script:

1. Confirms the first row of the file is a valid anchor row.
2. Confirms `NHS_Number` presence exactly matches the anchor-row pattern (i.e. NHS Number is present on anchor rows and only anchor rows).
3. Validates every anchor row's NHS Number against the **NHS Modulus 11 checksum**, failing loudly (without printing the numbers themselves) if any fail.
4. Checks for duplicate NHS Numbers **within** the same file — this fails loudly, since a duplicate within one surgery's extract is likely a data error.
5. Forward-fills the NHS Number from each anchor row down onto its associated detail rows to form `Patient_ID`.
6. Checks that no patient-level field (`PATIENT_LEVEL_COLS`) is populated on a non-anchor row — this "off-anchor leak" check fails loudly, since it would indicate the file doesn't match the assumed anchor/detail structure.

The NHS Number is used directly as `Patient_ID`. Each patient is also tagged with a `GP_ID` column (`GP01`, `GP02`, etc., set per entry in `GP_SOURCES`) — this is a separate additive column, not a prefix on the NHS Number, so `Patient_ID` stays a clean, valid NHS Number throughout.

Once every GP file has been processed, patient tables are combined across surgeries. A `Patient_ID` appearing under more than one `GP_ID` at that point is **not** treated as an error — a printed warning is raised instead, since this can represent legitimate patient movement between surgeries rather than a data error.

### 3.4 Build events tables

Once patients are anchored, the sparse event rows attached to each patient are split out into normalised event tables, each linked back to the patient via `Patient_ID` and `GP_ID`:

- Inhaler Prescriptions
- Prednisolone Courses
- Antibiotic Courses
- Consultations
- Referrals (Pulmonary Rehab)
- Influenza Vaccinations
- Pneumococcal Vaccinations
- COVID Vaccinations

Together with the `Patients` table, this gives **nine** output tables per run.

### 3.5 Add prescription/course/vaccination/consultation counts

Count columns — `*_Vaccination_Count`, `Inhaler_Prescription_Count`, `Prednisolone_Course_Count`, `Antibiotic_Course_Count`, and `Consultation_Count` — are added to `patients_df` by grouping each corresponding event table and left-merging the counts back onto the patients table. Patients with no matching events are given a count of zero rather than `NaN`, since "zero prescriptions" is a known value while a missing `Patient_ID` match would not be.

### 3.x Pulmonary rehab referral status

Because a patient can have multiple pulmonary rehab referral entries, the raw `Referral_to_Pulmonary_Rehab` text is first standardised via an external mapping file, `mapping_files/pulmonary_rehab_status_mapping.csv`, following the same "fail loud on unmapped values" approach as the other categorical fields (3.2). Each raw value is mapped to one of: `Referred`, `Declined`, `Completed`, or `Excluded`.

`Excluded` covers raw values that don't map cleanly onto a referral outcome — for example, "pulmonary rehabilitation programme not available" (an availability issue, not a patient action), "assessment for pulmonary rehabilitation completed" (the assessment was completed, not the programme itself?), and "pulmonary rehabilitation programme commenced" (in progress, not yet resolved either way). These are mapped explicitly to `Excluded` rather than left out of the mapping file entirely.

Three count columns are added to `patients_df`: `Number_of_Times_Referred`, `Number_of_Times_Declined`, `Number_of_Times_Completed`, each counting how many of that patient's referral entries mapped to that category. Patients with no matching entries are given a count of zero.

### 3.6 Evidence of rescue pack

An `Evidence_of_Rescue_Pack` count column is added to `patients_df`, counting distinct rescue pack events per patient. A rescue pack is a home supply of prednisolone and antibiotics that a patient keeps for self-treating a flare-up.

Evidence is drawn from two signals, combined into a single count:

- **Explicit text mention**: the free-text dose columns (`Inhaler_Dose`, `Prednisolone_Dose`, `Antibiotic_Dose`) are scanned for case-insensitive mentions of the phrase "rescue pack".
- **Inferred from prescribing pattern**: prednisolone and antibiotics issued to the same patient on the same date, even where no text mention of rescue pack is present. Inhaler is excluded from this inferred signal, as it is routine maintenance medication reissued regularly regardless of exacerbation, and its inclusion would likely undercount genuine rescue packs.

Each event is identified by a (patient, date) pair. Where the same date is flagged by both signals — for example, a prednisolone/antibiotic pair issued on the same day as a text mention — it is counted once, not twice, to avoid double-counting. The combined count is merged onto `patients_df`, filling with zero where there is no evidence of either kind.

> This is a proxy indicator, not a confirmed clinical fact. It depends on free-text phrasing at the point of prescribing (for the text signal) and same-day issuance (for the inferred signal), so it may under- or over-represent actual rescue pack usage.

### 3.7 Parse dates

Every column across every table whose name contains `Date`, or which is `DOB`, is parsed with `parse_dates()`, so downstream analysis doesn't need to handle mixed date representations. Dates stored as real Excel dates come out of `flatten_headers.py` as ISO text (`YYYY-MM-DD`) and are parsed as such; text dates are parsed day-first (`DD/MM/YYYY`). The two are handled separately because `pd.to_datetime(..., dayfirst=True)` on ISO dates swaps day and month, and silently blanks any date with a day above 12. Unparseable values become `NaT`, and the script prints a warning with the count per column.

### 3.8 Standardise categorical text & sanity checks

Categorical columns on `patients_df` (`MRC_Code`, `COPD_Diagnosis`, `Depression`, `Anxiety`) are stripped and title-cased for consistent display; `Postcode` is stripped and upper-cased. As a data-quality sanity check (not a hard failure), the script prints a warning per GP if any patient's `COPD_Diagnosis_Date` falls before their `DOB`.

### 3.9 Save and summarise data

After all configured GP files are processed, each of the nine tables is concatenated across surgeries and saved two ways:

- **CSV files** — one per table, written to `cleaned_data/combined_data/copd_<table_name>.csv`. Kept alongside the Excel output for further Python analysis; tables can be joined via `Patient_ID` / `GP_ID`.
- **Combined Excel workbook** — `cleaned_data/combined_data/copd_combined_output.xlsx`, built directly with `openpyxl` (rather than pandas' Excel writer) for full formatting control. This is the analyst-facing deliverable for NHS staff who work in Excel rather than Python. The workbook contains:
  - A **Summary** sheet first, listing the GP surgeries included in the run, then a table of sheet name / row count / one-line description for each of the nine data sheets.
  - One formatted sheet per table (short, analyst-friendly sheet names, since Excel caps sheet names at 31 characters), each with: bold white-on-dark-blue header row, a frozen header row, an autofilter over the full data range, UK `DD/MM/YYYY` date formatting on all date columns, and auto-sized columns (capped between 10 and 45 characters wide).

At the end of the run, the script prints: per-table row counts (from the CSV save step), any cross-GP duplicate `Patient_ID` warning, total GPs processed, total patients, and a value count of `Cardiovascular_Disease` (temporary — pending the outstanding mapping work noted in 3.2).

### Usage

```bash
python run_pipeline.py               # checks -> flatten -> clean, in one go
python run_pipeline.py --check-only  # checks only
```

The individual steps can still be run on their own:

```bash
python check_inputs.py <input_file_path>.xlsx
python flatten_headers.py <input_file_path>.xlsx <output_file_path>.csv
python clean_copd_dataset.py
```

```
Raw Excel files in evaluation_data/raw/  (one per surgery, GP code in name)
      |
      v
run_pipeline.py
      |
      |-- check_inputs.py: columns, row structure, mappings, dates, profile
      |       -> qa_reports/preflight_report.txt, unmapped_values.csv
      |       -> STOP here if anything blocking
      |
      |-- flatten_headers.py: one flat CSV per surgery
      |       -> evaluation_data/flattened/
      |
      v
clean_copd_dataset.py  (process_gp_file, once per surgery)
      |
      |-- Validate structure, clean categorical fields
      |-- Anchor on NHS Number -> Patient_ID, tag with GP_ID
      |-- Build 8 event tables + Patients table
      |-- Add prescription/course/consultation counts
      |-- Add Evidence_of_Rescue_Pack count
      |-- Parse dates, standardise categorical text
      |
      v
Combine across all GP surgeries
      |
      |-- Warn on cross-GP duplicate Patient_IDs
      |
      v
Save: 9 CSVs + 1 formatted Excel workbook (Summary + 9 sheets)
```
