# Data-Extraction-Script

`extract_production.py` consolidates monthly production workbooks (supplied in one or more ZIP files)
into one traceable dataset. It reads **only the first worksheet** of every `.xlsx`/`.xlsm`, detects each
file's layout independently, and guarantees through cell accounting that no data is silently dropped.

## Quick start

```bash
pip install -r requirements.txt
# put the ZIP(s) into ./input, then
python extract_production.py                 # full run: inspection + extraction
python extract_production.py --inspect-only  # structure inspection only
python extract_production.py --input path/to/file.zip --output path/to/out
```

New month? Drop the new ZIP into `input/` and run the same command. All ZIPs in `input/`
(including nested ZIPs and sub-folders) are processed.

## What it produces (`output/`)

| File | Content |
|---|---|
| `production_master.csv` | One row per source row × Mill/Area × Process. Measures as columns (`Input`, `Output`, `Target`, `Yield`, `Delay`, `Output (Cum)`, …). Includes `Date, Day, Month, Year, Mill/Area, Process, Shift, RecordType, IsSummary, CountsAsProduction, SourceFile, SourceSheet, SourceRow, SourceCells, ReviewFlag, DuplicateFlag`, … |
| `production_master_long.csv` | One row per extracted value cell, with exact `SourceCell`, header path, formula, and flags. |
| `unmapped_data.xlsx` | `Unmapped` – every cell that could not be mapped (SourceFile, Sheet, Row, Column, Value, Reason, Context). `Review_Required` – mapped records that carry an assumption. |
| `extraction_report.xlsx` | Summary, Files, Files_Failed, Other_Files, ZIPs, Layout_Groups, Structure_Blocks, Structure_Columns, Errors, Duplicates, Reconciliation, Cell_Accounting, Special_Handling, Sample_Records. |
| `inspection_report.xlsx` | Structure of every first sheet: layout, header paths, layout groups and differences, special handling, sample records. |
| `extraction.log` | Run log. |

### Avoiding double counting
* `CountsAsProduction = True` marks base values only (daily rows, or period rows of a summary sheet).
* Totals / subtotals / averages / max / min / MTD rows and "Total" columns are kept with
  `IsSummary = True` and an explicit `RecordType` (e.g. `Total`, `Average`, `Total (across columns)`).
* Columns named `... (Cum)` are running totals (MTD) – never sum them.
* For daily production use: `CountsAsProduction == True and Granularity == "Daily"`.

## Layout detection (per file, never assumed)
* **vertical** – dates down a column; mills/processes/measures across (multi-row and merged headers,
  unmerged group headers, stacked sections, side-by-side tables, shift rows, long format
  `Date | Mill | Shift | Input | Output`).
* **horizontal** – dates (or day numbers 1..31) across a row; mill/measure labels down the left.
* **matrix** – no date axis (monthly summaries: `Mill | Plan | Actual | Achievement %`).

Dates may be real Excel dates, text (`01.04.2023`, `1-Apr-23`, …), Excel serial numbers, or day numbers
combined with the month/year found in the sheet dates, title, sheet name, file name or folder (the
source used is recorded; conflicts are reported). Ambiguous dates (e.g. `04/05/2023` with no period to
decide) are **not guessed** – they are flagged.

## Data-safety rules
* ZIPs are extracted into `output/_extracted/`; workbooks are opened read-only and never saved.
  SHA-256 hashes before/after the run prove the sources are unchanged.
* Every non-empty cell gets exactly one disposition: title, section, header, date, label, attribute,
  value (extracted) or unmapped. `Cell_Accounting` must show `Balanced = True` and `Unaccounted = 0`.
* Placeholders (`-`, `NIL`, `NA`, `Shut`…) are **not** converted to 0 – they go to unmapped.
* Numbers stored as text, time values, inferred measures, out-of-period dates, etc. are extracted but
  flagged (`ReviewFlag`, `ReviewReason`).
* Formulas without a cached value (file never recalculated in Excel) are reported as errors and listed
  in unmapped – open/save such files in Excel and re-run.
* `.xls`/`.xlsb` files and corrupt workbooks are listed as not processed / failed, never skipped silently.

## CRM "PRODUCTION" template (FY23-27 files)
Behaviour verified on the 42 monthly files (Jan-2023 .. Jun-2026):
* Main table: dates in column A, 3-level header (mill/segment > process > Input/Output), TOTAL row,
  header rows repeated below the table. Helper day-number columns (X, AE, AL, AS, AX) are treated as
  attributes of the date, not as data.
* `CRM04`, `CRM06`, `SPM02` are mills/lines (`Dimension = Mill/Line`). `HNT`, `TUBE`, `FULL HARD`,
  `LG BALA/ROCKMAN/TIDC`, `OEM` are a **product breakdown of the same mill output**
  (`Dimension = Product segment`, never counted as production). The breakdown is reconciled daily
  against `BOTH MILL ... o/p of both mill`.
* `CRM.. > TOTAL`, `BOTH MILL TOTAL ...` and columns proven to be calculated (e.g. `TUBE ROLLING O/P`
  = TUBE + FULL HARD outputs, `OEM ROLLING O/P`, `OEM FINISH O/P`, unlabelled copy/difference columns)
  are kept but marked `IsSummary`, so `CountsAsProduction` daily CRM04+CRM06 output equals each file's
  TOTAL row exactly.
* The calculation blocks below the table (PRODUCTION OF CRM04/CRM06, utilisation, yield helpers) are
  not daily data; they are listed in `unmapped_data.xlsx` with `NearestLabel`, `Unit` and `BlockTitle`.
* Edit `breakdown_areas` / `breakdown_reference_area` in the config if segments are renamed or added.

## Customising keywords
Mill/area names, measure synonyms (e.g. `prodn`, `o/p`, `ABP`), summary words and units live in
`DEFAULT_CONFIG` at the top of the script. Extend them without editing code:

```json
{
  "measures": [["Output", "\\b(?:net\\s*prod)\\b"]],
  "known_areas": "\\b(?:CRM\\s*-?\\d+|BAF|SPM|MY_NEW_LINE)\\b"
}
```
`python extract_production.py --config keywords.json`

## Tests
`python -m pytest -q` builds a synthetic ZIP with eight different layouts and edge cases (merged and
unmerged headers, horizontal dates, stacked sections, shift rows, summary sheet, corrupt file, `.xls`,
duplicate file, nested ZIP, formulas without cached values) and checks the extraction.
