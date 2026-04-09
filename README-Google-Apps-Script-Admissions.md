# Google Apps Script Admissions API Notes

This file is a handoff note for future Codex sessions about the Google Apps Script setup planned for ExcelKidsHub admissions form submission.

## Goal

Replace the old Zoho iframe admissions flow with:

- website form in `excelkidshub.github.io`
- Google Apps Script Web App backend
- Google Sheet storage in `ExcelKidsHub-Master-Data`

The intended submission flow is:

1. user fills admission form on website
2. website sends JSON to Google Apps Script using `fetch()`
3. Apps Script validates request
4. Apps Script writes row into `admissions` sheet
5. website shows success or error message

## Current website state

Current file:

- `../excelkidshub.github.io/admissions.html`
- `../excelkidshub.github.io/js/admissions.js`

Current status:

- custom HTML admissions form created
- website submit script created
- Apps Script Web App URL wired into frontend JS
- live testing is still needed

## Planned Google Apps Script setup

Google Sheet name:

- `ExcelKidsHub-Master-Data`

Target sheet:

- `admissions`

Apps Script project purpose:

- receive website form submissions
- write admissions into Google Sheet

## Current deployment details

The user deployed the Apps Script Web App and shared:

- Deployment ID:
  `AKfycbzFyRzQ1FRe1Vnj6NViY8tlYQQtZRLphQkgbeLHZRz_gyRDJCn2O-QXWTJCeEvAvNPMDQ`
- Web App URL:
  `https://script.google.com/macros/s/AKfycbzFyRzQ1FRe1Vnj6NViY8tlYQQtZRLphQkgbeLHZRz_gyRDJCn2O-QXWTJCeEvAvNPMDQ/exec`

This URL is currently wired in the website-side file:

- `../excelkidshub.github.io/js/admissions.js`

## Apps Script entry point

The web app entry point must be:

- `doPost(e)`

The main business logic function is:

- `admission(payload)`

Important note for future Codex:

- `admission()` is not the web endpoint by itself
- `doPost(e)` receives the request
- `doPost(e)` must parse JSON and call `admission(payload)`

## Code intended for `Code.gs`

The user asked specifically what to replace in `Code.gs` where they saw:

```javascript
function myFunction() {
  
}
```

That should be replaced with the custom admissions API code.

## Functions planned in Apps Script

Functions currently recommended:

- `doPost(e)`
- `admission(payload)`
- `getAdmissionsSheet()`
- `getNextAdmissionId(sheet)`
- `jsonResponse(data)`
- `clean(value)`
- `isValidEmail(email)`

## Expected payload from website

The website is expected to send JSON like:

```json
{
  "parentName": "Anita Sharma",
  "mobile": "9876543210",
  "email": "anita@example.com",
  "address": "Dhanori Pune",
  "city": "Pune",
  "studentName": "Aarav Sharma",
  "age": "6",
  "gender": "Male",
  "school": "ABC School",
  "grade": "1",
  "level": "Basic",
  "mode": "Online",
  "admissionSource": "Website",
  "notes": "Interested in evening batch"
}
```

## Required validation in script

Currently recommended required fields:

- `parentName`
- `mobile`
- `studentName`
- `level`
- `mode`

Optional validation:

- email format check
- numeric age check

## Expected Google Sheet column order

The Apps Script code assumes the `admissions` sheet has at least this order:

1. `Admission ID`
2. `Parent Name`
3. `Mobile`
4. `Email`
5. `Address`
6. `City`
7. `Student Name`
8. `Age`
9. `Gender`
10. `School`
11. `Grade`
12. `Level`
13. `Batch Code`
14. `Mode`
15. `Start Date`
16. `End Date`
17. `Status`
18. `Total Fee`
19. `Discount`
20. `Manual Adjustment`
21. `Adjusted Fee`
22. `Total Paid`
23. `Pending`
24. `Payment Status`
25. `Admission Source`
26. `Referral Type`
27. `Referrer Name`
28. `Created Date`

Possible later extra columns:

- `Notes`
- certificate-related columns

## Current script behavior design

When a submission comes in:

- generate next `Admission ID` like `A001`
- set `Created Date = new Date()`
- set `Status = Pending Start`
- leave financial calculated fields blank
- set `Admission Source = Website` unless explicitly sent

Default values used by current design:

- `Batch Code` = blank
- `Start Date` = blank
- `End Date` = blank
- `Status` = `Pending Start`
- `Total Fee` = blank
- `Discount` = `0`
- `Manual Adjustment` = `0`
- `Referral Type` = blank
- `Referrer Name` = blank

## Notes column caveat

In the draft Apps Script code, there was an optional section:

```javascript
if (notes) {
  sheet.getRange(lastRow, 29).setValue(notes);
}
```

This only works if there is a real `Notes` column in column 29.

If the `admissions` sheet does not include a `Notes` column after `Created Date`, this block must be removed or adapted.

## Deployment steps

Planned deployment steps in Apps Script:

1. open Google Sheet
2. `Extensions -> Apps Script`
3. paste code into `Code.gs`
4. save project
5. click `Deploy`
6. choose `New deployment`
7. choose `Web app`
8. set:
   - Execute as: `Me`
   - Who has access: `Anyone`
9. deploy
10. authorize
11. copy the Web App URL

## Response format expected

Success response:

```json
{
  "success": true,
  "message": "Admission saved successfully",
  "admissionId": "A001"
}
```

Error response:

```json
{
  "success": false,
  "message": "Mobile is required"
}
```

## Future website integration

Next website work is expected in:

- `../excelkidshub.github.io/admissions.html`
- possibly `../excelkidshub.github.io/js/admissions.js`
- maybe `../excelkidshub.github.io/css/style.css`

Website tasks still pending:

- replace Zoho iframe
- create custom HTML form
- add client-side validation
- submit with `fetch()` to Apps Script
- show success/error state

## Recommended next steps for future Codex

1. Confirm the exact current Google Sheet column order.
2. Confirm whether `Notes` exists in `admissions`.
3. Paste or adapt the Apps Script code into `Code.gs`.
4. Deploy the Apps Script Web App.
5. Test manual POST requests first.
6. Replace the website Zoho iframe with a custom form.
7. Connect the website form to the Apps Script URL.

## Related local files

- `README-ExcelKidsHub-Master-Data.md`
- `../excelkidshub.github.io/admissions.html`

## Important guidance

Before future work:

- ask which workbook / Google Sheet version is the source of truth
- confirm final `admissions` column order
- confirm whether certificate columns are already added in the live Google Sheet
