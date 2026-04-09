# Google Apps Script Admin Files

These files are meant to be created as separate files inside the same Google Apps Script project linked to your ExcelKidsHub Google Sheet.

## Files

- `00_Code.gs`
- `01_Config.gs`
- `02_Utils.gs`
- `03_AdminAuth.gs`
- `04_Admissions.gs`
- `05_Dashboard.gs`
- `06_Batches.gs`
- `07_Payments.gs`
- `08_Expenses.gs`

## Script Properties

Set these in Apps Script:

- `ADMIN_PASSWORD`
- `ADMIN_TOKEN`

Optional for the Vercel website project:

- `GOOGLE_SCRIPT_URL`

## Current actions supported

Public:

- `admission`

Admin:

- `adminLogin`
- `getDashboard`
- `getAdmissions`
- `getBatches`
- `saveBatch`
- `updateBatch`
- `assignStudentToBatch`
- `savePayment`
- `getPayments`
- `saveExpense`
- `getExpenses`

## Example admin login payload

```json
{
  "action": "adminLogin",
  "adminPassword": "your-password"
}
```

## Example admin request payload

```json
{
  "action": "getAdmissions",
  "adminToken": "your-token"
}
```

## Expected sheet names

- `admissions`
- `batches`
- `payments`
- `expenses`

## Expected headers

### admissions

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
- `Notes`

### batches

- `Batch Code`
- `Batch Name`
- `Level`
- `Mode`
- `Start Date`
- `End Date`
- `Timing`
- `Days`
- `Capacity`
- `Location`
- `Status`
- `Notes`

### payments

- `Payment ID`
- `Admission ID`
- `Student Name`
- `Batch Code`
- `Payment Date`
- `Amount`
- `Payment Mode`
- `Transaction ID`
- `Notes`

### expenses

- `Expense ID`
- `Expense Date`
- `Category`
- `Amount`
- `Payment Mode`
- `Description`
- `Vendor`

## Important note

Your current public `admission(payload)` flow is preserved.

Admin reads and writes are protected by `ADMIN_TOKEN`, so do not expose that token in the public website.
