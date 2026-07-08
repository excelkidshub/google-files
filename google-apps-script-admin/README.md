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
- `09_PaymentEmails.gs`
- `10_TeacherTraining.gs`
- `11_AdmissionCertificates.gs`

## Script Properties

Set these in Apps Script:

- `ADMIN_PASSWORD`
- `ADMIN_TOKEN`
- `RECEIPT_TEMPLATE_ID` or `RECEIPT_TEMPLATE_NAME`
- `STUDENT_CERTIFICATE_TEMPLATE_ID` or `STUDENT_CERTIFICATE_TEMPLATE_NAME`
- `TEACHER_CERTIFICATE_TEMPLATE_ID` or `TEACHER_CERTIFICATE_TEMPLATE_NAME`

Optional for the Vercel website project:

- `GOOGLE_SCRIPT_URL`

Optional but recommended:

- `RECEIPT_ARCHIVE_FOLDER_ID`
- `STUDENT_CERTIFICATE_ARCHIVE_FOLDER_ID`
- `TEACHER_CERTIFICATE_ARCHIVE_FOLDER_ID`
- `ACADEMY_NAME`
- `ACADEMY_EMAIL`
- `ACADEMY_PHONE`
- `ACADEMY_ADDRESS`
- `SENDER_NAME`

## Current actions supported

Public:

- `admission`
- `teacherTrainingAdmission`

Admin:

- `adminLogin`
- `getDashboard`
- `getAdmissions`
- `getBatches`
- `saveBatch`
- `updateBatch`
- `assignStudentToBatch`
- `savePayment`
- `sendPaymentEmail`
- `sendAdmissionCertificate`
- `sendAdmissionCertificateWhatsApp`
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
- `teacher_training_admissions`

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
- `Send Certificate`
- `Certificate Status`
- `Certificate Number`
- `Certificate Issue Date`
- `Certificate Sent Date`
- `Certificate PDF URL`
- `Certificate Image URL`
- `Send Certificate WhatsApp`
- `Certificate WhatsApp Status`
- `Certificate WhatsApp Link`
- `Certificate WhatsApp Sent Date`
- `Certificate WhatsApp Error`
- `Certificate Error`
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

Optional admissions headers for tracking email activity:

- `Last Email Sent Date`
- `Last Receipt Sent Date`
- `Full Payment Email Date`
- `Last Payment Reminder Date`

### teacher_training_admissions

- `Training ID`
- `Full Name`
- `Mobile`
- `Email`
- `City`
- `Profession`
- `Education`
- `Training Type`
- `Mode`
- `Status`
- `Total Fee`
- `Total Paid`
- `Pending`
- `Payment Status`
- `Created Date`
- `Send Receipt`
- `Receipt Status`
- `Certificate Status`
- `Certificate Number`
- `Certificate Issue Date`
- `Certificate Sent Date`
- `Certificate Error`
- `Notes`

Manual teacher-training receipt flow:

1. Set `Send Receipt = Send`
2. The `onEdit` trigger marks `Receipt Status = Pending`
3. Run `processPendingTeacherTrainingReceipts()` by time trigger or manually
4. The script sends the receipt and writes `Receipt Status = Email Sent`

Manual admissions certificate email flow:

1. Upload `Certificate_template.pptx` to Google Drive
2. Open it with Google Slides and save/convert it as a Google Slides file
3. Set `STUDENT_CERTIFICATE_TEMPLATE_ID` to that Slides file ID
4. In `admissions`, set `Send Certificate = Send`
5. Run `processPendingAdmissionCertificates()` by time trigger or manually
6. The script generates a PDF and PNG image, emails the PDF, writes Drive links, and updates only certificate email/status fields

Manual admissions certificate WhatsApp flow:

1. In `admissions`, set `Send Certificate WhatsApp = Send`
2. Run `processPendingAdmissionCertificateWhatsApps()` by time trigger or manually
3. The script generates the certificate image and writes `Certificate WhatsApp Link`
4. Open `Certificate WhatsApp Link`; WhatsApp opens with a pre-filled message containing the certificate image link
5. Send the message manually
6. After sending, set `Certificate WhatsApp Status = Sent`; the `onEdit` trigger writes `Certificate WhatsApp Sent Date`

WhatsApp note:

- Email and WhatsApp are separate triggers. WhatsApp link generation failures do not block certificate email sending.
- This is the free semi-automatic flow. It does not require WhatsApp Business API keys or Meta billing.
- Normal WhatsApp links cannot auto-attach or auto-send the image; they open a pre-filled message for you to send.

Supported student certificate placeholders:

- `{{NAME}}`
- `{{STUDENT}}`
- `{{STUDENT_NAME}}`
- `{{PARENT_NAME}}`
- `{{COURSE}}`
- `{{LEVEL}}`
- `{{MODE}}`
- `{{BATCH_CODE}}`
- `{{DATE}}`
- `{{ISSUE_DATE}}`
- `{{CERTIFICATE_NO}}`
- `{{CERTIFICATE_NUMBER}}`
- `{{ADMISSION_ID}}`
- `{{ACADEMY_NAME}}`

Manual teacher-training certificate flow:

1. Upload `Teachers_Certificate_template.pptx` to Google Drive
2. Open it with Google Slides and save/convert it as a Google Slides file
3. Set `TEACHER_CERTIFICATE_TEMPLATE_ID` to that Slides file ID
4. Set `Certificate Status = Ready`
5. Run `processReadyTeacherTrainingCertificates()` by time trigger or manually
6. The script generates the PDF, emails it, writes sent date, and changes status to `Sent`

Supported teacher certificate placeholders:

- `{{NAME}}`
- `{{FULL_NAME}}`
- `{{COURSE}}`
- `{{TRAINING_TYPE}}`
- `{{MODE}}`
- `{{DATE}}`
- `{{ISSUE_DATE}}`
- `{{CERTIFICATE_NO}}`
- `{{CERTIFICATE_NUMBER}}`
- `{{TRAINING_ID}}`
- `{{ACADEMY_NAME}}`

## Receipt template setup

The receipt email flow expects a Google Docs template in Drive.

Recommended template name:

- `ExcelKidsHub Receipt Template`

Current local source file in this repo:

- [ExcelKidsHub Receipt Template.docx](/d:/Git_ExcelKidsHub/google-files/ExcelKidsHub%20Receipt%20Template.docx)

Recommended setup:

1. Upload this `.docx` file to Google Drive
2. Open it with Google Docs
3. Save it as a Google Doc
4. Use that Google Doc as the receipt template via:
   - `RECEIPT_TEMPLATE_ID`, or
   - `RECEIPT_TEMPLATE_NAME`

If you do not set `RECEIPT_TEMPLATE_ID`, the script will search Drive by:

- `RECEIPT_TEMPLATE_NAME`

Supported placeholders inside the Google Doc template:

- `{{RECEIPT_TITLE}}`
- `{{EMAIL_TYPE_LABEL}}`
- `{{EMAIL_MESSAGE}}`
- `{{ACADEMY_NAME}}`
- `{{ACADEMY_EMAIL}}`
- `{{ACADEMY_PHONE}}`
- `{{ACADEMY_ADDRESS}}`
- `{{RECEIPT_NO}}`
- `{{PAYMENT_ID}}`
- `{{PAYMENT_DATE}}`
- `{{STUDENT_NAME}}`
- `{{PARENT_NAME}}`
- `{{ADMISSION_ID}}`
- `{{PARENT_EMAIL}}`
- `{{MOBILE}}`
- `{{LEVEL}}`
- `{{MODE}}`
- `{{BATCH_CODE}}`
- `{{PAYMENT_MODE}}`
- `{{TRANSACTION_ID}}`
- `{{AMOUNT_PAID}}`
- `{{TOTAL_FEE}}`
- `{{DISCOUNT}}`
- `{{MANUAL_ADJUSTMENT}}`
- `{{ADJUSTED_FEE}}`
- `{{TOTAL_PAID}}`
- `{{PENDING_AMOUNT}}`
- `{{PAYMENT_STATUS}}`
- `{{NOTES}}`
- `{{TODAY_DATE}}`

Your current local receipt template already includes these placeholders, which are also supported:

- `{{RECEIPT_NO}}`
- `{{DATE}}`
- `{{STUDENT}}`
- `{{AMOUNT}}`

Note:

- Your current local template has `Payment Status: PAID` as fixed text, not a dynamic placeholder.
- If you want reminder/full-payment PDFs to show dynamic status, change that line in the Google Doc template to:
  `Payment Status: {{PAYMENT_STATUS}}`

## Admin email flow now supported

From admin you can now:

- save a payment and email a receipt PDF for that transaction
- save a payment and automatically send a full-payment confirmation if balance becomes zero
- resend a receipt PDF later from payment history
- send a pending payment reminder from the payments area

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
