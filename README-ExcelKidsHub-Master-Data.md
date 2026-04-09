# ExcelKidsHub Master Data Notes

This file is a working handoff note for future Codex sessions about the Excel/Google Sheets operations workbook for ExcelKidsHub.

## Main workbook files

- `ExcelKidsHub-Master-Data.xlsx`
  Current main workbook used in the workspace.
- `ExcelKidsHub-Master-Data-backup-before-schedule-sync.xlsx`
  Backup created before schedule-based batch remapping.
- `ExcelKidsHub-Master-Data-imported.xlsx`
  Workbook after first Zoho data import.
- `ExcelKidsHub-Master-Data-cleaned.xlsx`
  Imported workbook after duplicate cleanup, city improvement, and inferred batches.
- `ExcelKidsHub-Master-Data-schedule-synced.xlsx`
  Workbook copy after syncing batches with the published website schedule.
- `ExcelKidsHub-Master-Data-dashboard-fixed.xlsx`
  Copy where batch dashboard formulas were refreshed.
- `ExcelKidsHub-Master-Data-student-list-fixed.xlsx`
  Copy where batch dashboard student lists were materialized as visible text.
- `ExcelKidsHub-Master-Data-with-certificate-columns.xlsx`
  Copy with certificate-tracking columns added in `admissions`.

## Current workbook structure

The workbook design currently uses these sheets:

- `admissions`
- `payments`
- `batches`
- `expenses`
- `finance-summary`
- `batch-dashboard`

## Admissions columns

Current `admissions` design includes:

- `Admission ID`
- `Parent Name`
- `Mobile`
- `Email`
- `Address`
- `City`
- `Student Name`
- `Age`
- `Gender`
- `School`
- `Grade`
- `Level`
- `Batch Code`
- `Mode`
- `Start Date`
- `End Date`
- `Status`
- `Total Fee`
- `Discount`
- `Manual Adjustment`
- `Adjusted Fee`
- `Total Paid`
- `Pending`
- `Payment Status`
- `Admission Source`
- `Referral Type`
- `Referrer Name`
- `Created Date`

In the generator, certificate fields were later added as planned future columns:

- `Certificate Status`
- `Certificate Number`
- `Certificate Issue Date`
- `Certificate Sent Date`

These certificate columns are present in `ExcelKidsHub-Master-Data-with-certificate-columns.xlsx`, but not yet merged into the currently active `ExcelKidsHub-Master-Data.xlsx`.

## What has been done

### 1. Workbook generator created

Python generator:

- `create_excelkidshub_workbook.py`

Purpose:

- creates the workbook
- formats headers, borders, widths, alignment
- adds formulas
- adds dropdowns
- creates `batch-dashboard`

### 2. Zoho data import completed

Source file:

- `ExcelKidsHub Zoho Entries.xlsx`

Importer:

- `import_zoho_to_master.py`

Imported into workbook:

- admissions from `Form_Responses`
- payments from rows marked paid
- expenses from Zoho `Expenses`
- one explicit batch from Zoho batch columns

### 3. Cleanup pass completed

Cleanup script:

- `cleanup_imported_master.py`

What cleanup did:

- reduced admissions from 59 to 56
- removed 3 obvious duplicates
- improved some city values
- assigned inferred batch codes where missing
- added inferred batch rows

### 4. Schedule sync completed

Schedule source used:

- local file: `../excelkidshub.github.io/schedule.html`

Sync script:

- `sync_batches_from_schedule.py`

What it did:

- replaced old/inferred batch rows with website schedule batches
- mapped admissions to schedule batch codes by `Level + Mode + created month`
- updated admission batch code/start/end/status fields

Published schedule batches currently used:

- `L1-B1`
- `L1-B2`
- `L1-B3`
- `L1-B4`
- `L2-B1`
- `L2-B2`
- `L2-B3`
- `L2-B4`
- `L3-B1`
- `L3-B2`

### 5. Batch dashboard created

The `batch-dashboard` sheet is a card-based dashboard, not a table.

Each card shows:

- batch code
- level and mode
- start and end dates
- duration
- timing
- capacity
- students enrolled
- available seats
- student list

Important note:

- some Excel environments did not visibly render the `FILTER + TEXTJOIN` student-list formula
- because of that, a fallback copy was created with materialized student names

Relevant scripts:

- `refresh_batch_dashboard.py`
- `materialize_batch_dashboard_student_lists.py`

### 6. Certificate tracking design decided

Planned workflow:

- use `Certificate Status` as the manual action trigger
- when status becomes `Ready`, Google Sheets Apps Script should:
  - validate row
  - generate certificate number if missing
  - create certificate
  - send email
  - set sent date
  - change status to `Sent`

Recommended statuses:

- `Not Issued`
- `Ready`
- `Sent`

Optional future column:

- `Certificate Error`

## Known issues / caveats

### File locking / save behavior

There were repeated Windows file-lock issues while saving workbook variants.

Observed behavior:

- saving to a new filename sometimes updated the main workbook instead
- overwriting the open workbook sometimes failed with permission denied

Implication for future Codex sessions:

- always check actual file timestamps after save
- always verify the real saved workbook contents after writing
- if the workbook is open in Excel, close it before applying changes

### Batch dashboard visibility

`ExcelKidsHub-Master-Data.xlsx` may still rely on formula-based student lists depending on which copy is being used.

If student names are not visible in `batch-dashboard`, use:

- `ExcelKidsHub-Master-Data-student-list-fixed.xlsx`

That copy stores visible multi-line names directly in the card cells.

### Schedule end dates

Website schedule published:

- start date
- timing
- mode
- sessions range

But it did not publish exact end dates.

Current end dates were inferred from:

- average session count
- weekly class frequency

These are operational estimates, not exact website-provided end dates.

### City data

City was inferred from address text for some rows.

Some rows may still need manual city correction.

## Current recommended working file

If dashboard student names matter most right now:

- use `ExcelKidsHub-Master-Data-student-list-fixed.xlsx`

If staying closest to the main tracked file matters most:

- use `ExcelKidsHub-Master-Data.xlsx`

Before future edits, confirm with the user which file should be treated as the current source of truth.

## Suggested next steps

### High priority

- merge certificate columns into the final main workbook actually being used
- close the workbook in Excel before next automated update to avoid file-lock issues
- decide which workbook copy is the official source of truth

### Google Sheets migration

When moving to Google Sheets:

- upload the chosen workbook
- preserve `admissions` as the core operational sheet
- use `Certificate Status` as the manual trigger field

### Certificate automation

Recommended stack:

- Google Sheets
- Google Apps Script
- Gmail
- Google Slides certificate template

Manual certificate send flow:

1. user sets `Certificate Status = Ready`
2. Apps Script generates certificate number if blank
3. Apps Script creates certificate
4. Apps Script sends email
5. Apps Script writes `Certificate Sent Date`
6. Apps Script changes status to `Sent`

### Better batch mapping

Current admission-to-batch mapping from schedule is approximate.

Future improvement:

- assign students to exact schedule batches using a more explicit rule:
  - level
  - mode
  - admission date
  - preferred slot if available

### City insights

Potential next sheet:

- `city-summary`

Useful metrics:

- students by city
- revenue by city
- leads by city

## Related scripts in this folder

- `create_excelkidshub_workbook.py`
- `import_zoho_to_master.py`
- `cleanup_imported_master.py`
- `sync_batches_from_schedule.py`
- `refresh_batch_dashboard.py`
- `materialize_batch_dashboard_student_lists.py`
- `add_certificate_columns.py`

## Guidance for future Codex sessions

When continuing this work:

1. First identify which workbook file the user wants to continue with.
2. Check whether that file is currently open/locked before saving.
3. Verify actual workbook contents after every save.
4. If editing `batch-dashboard`, test both:
   - formula presence
   - visible rendered values
5. Prefer creating a backup copy before large workbook rewrites.
