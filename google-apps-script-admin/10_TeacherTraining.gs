function teacherTrainingAdmission(payload) {
  try {
    var fullName = clean(payload.fullName);
    var mobile = clean(payload.mobile);
    var email = clean(payload.email);
    var trainingType = clean(payload.trainingType) || "Parent Teacher Phonics Training";
    var mode = clean(payload.mode);

    if (!fullName || !mobile || !email || !trainingType || !mode) {
      return jsonResponse({ success: false, message: "Required fields missing" });
    }

    if (!isValidEmail(email)) {
      return jsonResponse({ success: false, message: "Invalid email address" });
    }

    if (clean(payload.website)) {
      return jsonResponse({ success: false, message: "Unable to submit form right now" });
    }

    const sheet = getOrCreateTeacherTrainingSheet();
    const duplicateCheck = findDuplicateTeacherTrainingAdmission(sheet, fullName, mobile, trainingType);
    const createdDate = formatDateTimeValue(new Date());

    if (duplicateCheck.found) {
      updateTeacherTrainingAdmissionRow(sheet, duplicateCheck.rowNumber, payload);
      return jsonResponse({
        success: true,
        message: "Registration updated successfully",
        trainingId: duplicateCheck.trainingId
      });
    }

    const trainingId = getNextTeacherTrainingId(sheet);
    const rowNumber = appendObjectRow(sheet, {
      "Training ID": trainingId,
      "Full Name": fullName,
      "Mobile": mobile,
      "Email": email,
      "City": clean(payload.city),
      "Profession": clean(payload.profession),
      "Education": clean(payload.education),
      "Training Type": trainingType,
      "Mode": mode,
      "Status": DEFAULTS.admissionStatus,
      "Created Date": createdDate,
      "Payment Status": DEFAULTS.paymentStatus,
      "Notes": clean(payload.notes)
    });
    setTeacherTrainingRowFormulas(sheet, rowNumber);

    return jsonResponse({
      success: true,
      message: "Registration saved successfully",
      trainingId: trainingId
    });
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("teacherTrainingAdmission error: " + errorMsg);
    return jsonResponse({ success: false, message: "Error: " + errorMsg });
  }
}

function getOrCreateTeacherTrainingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAMES.teacherTrainingAdmissions);

  if (sheet) {
    return sheet;
  }

  sheet = spreadsheet.insertSheet(SHEET_NAMES.teacherTrainingAdmissions);
  sheet.appendRow(getTeacherTrainingHeaders());
  return sheet;
}

function getTeacherTrainingHeaders() {
  return [
    "Training ID",
    "Full Name",
    "Mobile",
    "Email",
    "City",
    "Profession",
    "Education",
    "Training Type",
    "Mode",
    "Status",
    "Total Fee",
    "Total Paid",
    "Pending",
    "Payment Status",
    "Created Date",
    "Send Receipt",
    "Receipt Status",
    "Certificate Status",
    "Certificate Number",
    "Certificate Issue Date",
    "Certificate Sent Date",
    "Certificate Error",
    "Notes"
  ];
}

function getNextTeacherTrainingId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return "T001";
  }

  var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  var maxNumber = 0;

  idValues.forEach(function(id) {
    var text = String(id || "").trim();
    var match = text.match(/^T(\d+)$/i);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNumber) {
        maxNumber = num;
      }
    }
  });

  return "T" + String(maxNumber + 1).padStart(3, "0");
}

function findDuplicateTeacherTrainingAdmission(sheet, fullName, mobile, trainingType) {
  const rows = getSheetObjects(sheet);
  const normalizedName = clean(fullName).toLowerCase();
  const normalizedMobile = clean(mobile);
  const normalizedTrainingType = clean(trainingType).toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    if (
      clean(rows[i]["Training ID"]) &&
      clean(rows[i]["Full Name"]).toLowerCase() === normalizedName &&
      clean(rows[i]["Mobile"]) === normalizedMobile &&
      clean(rows[i]["Training Type"]).toLowerCase() === normalizedTrainingType
    ) {
      return {
        found: true,
        rowNumber: rows[i]._rowNumber,
        trainingId: clean(rows[i]["Training ID"])
      };
    }
  }

  return { found: false };
}

function updateTeacherTrainingAdmissionRow(sheet, rowNumber, payload) {
  const headerMap = getHeaderMap(sheet);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Full Name", clean(payload.fullName));
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Mobile", clean(payload.mobile));
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Email", clean(payload.email));
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "City", clean(payload.city));
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Profession", clean(payload.profession));
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Education", clean(payload.education));
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Training Type", clean(payload.trainingType) || "Parent Teacher Phonics Training");
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Mode", clean(payload.mode));
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Notes", clean(payload.notes));
  setTeacherTrainingRowFormulas(sheet, rowNumber);
}

function setTeacherTrainingRowFormulas(sheet, rowNumber) {
  const headerMap = getHeaderMap(sheet);

  if (headerMap["Pending"]) {
    sheet.getRange(rowNumber, headerMap["Pending"]).setFormulaR1C1('=IF(RC[-2]="","",RC[-2]-RC[-1])');
  }

  if (headerMap["Payment Status"]) {
    sheet.getRange(rowNumber, headerMap["Payment Status"]).setFormulaR1C1('=IF(RC[-3]="","",IF(RC[-2]=0,"Not Started",IF(RC[-2]<RC[-3],"Partial","Completed")))');
  }
}

function handleTeacherTrainingEdit(sheet, range) {
  const headerMap = getHeaderMap(sheet);
  const editedColumn = range.getColumn();
  const rowNumber = range.getRow();

  if (rowNumber <= 1) {
    return;
  }

  if (headerMap["Send Receipt"] && editedColumn === headerMap["Send Receipt"]) {
    if (clean(range.getValue()) === "Send") {
      setCellIfHeaderExists(sheet, rowNumber, headerMap, "Receipt Status", "Pending");
    }
    return;
  }

  if (headerMap["Certificate Status"] && editedColumn === headerMap["Certificate Status"]) {
    if (clean(range.getValue()) === "Ready") {
      setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Error", "");
    }
  }
}

function processPendingTeacherTrainingReceipts() {
  const sheet = getSheet(SHEET_NAMES.teacherTrainingAdmissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);

  rows.forEach(function(row) {
    if (clean(row["Receipt Status"]) !== "Pending") {
      return;
    }

    try {
      if (!clean(row["Email"]) || !isValidEmail(clean(row["Email"]))) {
        setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Receipt Status", "Email Failed - Invalid Email");
        return;
      }

      sendTeacherTrainingReceiptInternal(row);
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Receipt Status", "Email Sent");
    } catch (error) {
      const errorMsg = error && error.message ? error.message : "Unknown error";
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Receipt Status", "Email Failed - " + errorMsg.substring(0, 60));
    }
  });
}

function sendTeacherTrainingReceiptInternal(row) {
  const academyName = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName;
  const academyEmail = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyEmail) || DEFAULTS.academyEmail;
  const academyPhone = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyPhone) || DEFAULTS.academyPhone;
  const senderName = getScriptProperty(SCRIPT_PROPERTY_KEYS.senderName) || DEFAULTS.senderName;
  const fullName = clean(row["Full Name"]);
  const email = clean(row["Email"]);
  const paymentStatus = clean(row["Payment Status"]) || DEFAULTS.paymentStatus;
  const totalPaid = toNumber(row["Total Paid"], 0);
  const pending = toNumber(row["Pending"], 0);
  const templateFile = getReceiptTemplateFile();
  const receiptTitle = "Training Payment Receipt";
  const templateValues = {
    "{{RECEIPT_TITLE}}": receiptTitle,
    "{{EMAIL_TYPE_LABEL}}": receiptTitle,
    "{{EMAIL_MESSAGE}}": "Thank you for your payment. Please find the receipt attached.",
    "{{ACADEMY_NAME}}": academyName,
    "{{ACADEMY_EMAIL}}": academyEmail,
    "{{ACADEMY_PHONE}}": academyPhone,
    "{{ACADEMY_ADDRESS}}": getScriptProperty(SCRIPT_PROPERTY_KEYS.academyAddress) || DEFAULTS.academyAddress,
    "{{RECEIPT_NO}}": clean(row["Training ID"]),
    "{{PAYMENT_ID}}": clean(row["Training ID"]),
    "{{DATE}}": formatDisplayDate(new Date()),
    "{{PAYMENT_DATE}}": formatDisplayDate(new Date()),
    "{{STUDENT}}": fullName,
    "{{STUDENT_NAME}}": fullName,
    "{{PARENT_NAME}}": fullName,
    "{{ADMISSION_ID}}": clean(row["Training ID"]),
    "{{PARENT_EMAIL}}": email,
    "{{MOBILE}}": clean(row["Mobile"]),
    "{{LEVEL}}": clean(row["Training Type"]),
    "{{MODE}}": clean(row["Mode"]),
    "{{BATCH_CODE}}": "",
    "{{PAYMENT_MODE}}": "",
    "{{TRANSACTION_ID}}": "",
    "{{AMOUNT}}": Utilities.formatString("%.0f", totalPaid),
    "{{AMOUNT_PAID}}": formatMoney(totalPaid),
    "{{TOTAL_FEE}}": formatMoney(toNumber(row["Total Fee"], 0)),
    "{{DISCOUNT}}": formatMoney(0),
    "{{MANUAL_ADJUSTMENT}}": formatMoney(0),
    "{{ADJUSTED_FEE}}": formatMoney(toNumber(row["Total Fee"], 0)),
    "{{TOTAL_PAID}}": formatMoney(totalPaid),
    "{{PENDING_AMOUNT}}": formatMoney(pending),
    "{{PAYMENT_STATUS}}": paymentStatus,
    "{{NOTES}}": clean(row["Notes"]),
    "{{TODAY_DATE}}": formatDisplayDate(new Date())
  };
  const pdfBlob = buildReceiptPdfBlob(templateFile, templateValues);
  const subject = "Training payment receipt for " + fullName + " | " + academyName;
  const plainTextBody =
    "Dear " + fullName + ",\n\n" +
    "Thank you for your payment. Please find the receipt attached.\n\n" +
    "Training ID: " + clean(row["Training ID"]) + "\n" +
    "Payment Status: " + paymentStatus + "\n" +
    "Total Paid: " + formatMoney(totalPaid) + "\n" +
    "Pending: " + formatMoney(pending) + "\n\n" +
    "Regards,\n" + academyName + "\n" + academyPhone + "\n" + academyEmail;
  const htmlBody =
    "<p>Dear " + sanitizeHtmlText(fullName) + ",</p>" +
    "<p>Thank you for your payment. Please find the receipt attached.</p>" +
    "<p><strong>Training ID:</strong> " + sanitizeHtmlText(clean(row["Training ID"])) + "<br>" +
    "<strong>Payment Status:</strong> " + sanitizeHtmlText(paymentStatus) + "<br>" +
    "<strong>Total Paid:</strong> " + sanitizeHtmlText(formatMoney(totalPaid)) + "<br>" +
    "<strong>Pending:</strong> " + sanitizeHtmlText(formatMoney(pending)) + "</p>" +
    "<p>Regards,<br>" + sanitizeHtmlText(academyName) + "<br>" + sanitizeHtmlText(academyPhone) + "<br>" + sanitizeHtmlText(academyEmail) + "</p>";

  GmailApp.sendEmail(email, subject, plainTextBody, {
    htmlBody: htmlBody,
    attachments: [pdfBlob],
    name: senderName
  });
}

function processReadyTeacherTrainingCertificates() {
  const sheet = getSheet(SHEET_NAMES.teacherTrainingAdmissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);

  rows.forEach(function(row) {
    if (clean(row["Certificate Status"]) !== "Ready") {
      return;
    }

    try {
      if (!clean(row["Email"]) || !isValidEmail(clean(row["Email"]))) {
        setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", "Invalid Email");
        return;
      }

      let certificateNumber = clean(row["Certificate Number"]);
      if (!certificateNumber) {
        certificateNumber = buildTeacherCertificateNumber(row);
        setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Number", certificateNumber);
      }

      const issueDate = row["Certificate Issue Date"] || new Date();
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Issue Date", issueDate);
      sendTeacherTrainingCertificateInternal(row, certificateNumber, issueDate);
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Sent Date", new Date());
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Sent");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", "");
    } catch (error) {
      const errorMsg = error && error.message ? error.message : "Unknown error";
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", errorMsg.substring(0, 100));
    }
  });
}

function buildTeacherCertificateNumber(row) {
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  return "EKH-TC-" + year + "-" + clean(row["Training ID"]);
}

function getTeacherCertificateTemplateFile() {
  const templateId = getScriptProperty(SCRIPT_PROPERTY_KEYS.teacherCertificateTemplateId);
  if (templateId) {
    return DriveApp.getFileById(templateId);
  }

  const templateName = getScriptProperty(SCRIPT_PROPERTY_KEYS.teacherCertificateTemplateName) || DEFAULTS.teacherCertificateTemplateName;
  const matches = DriveApp.getFilesByName(templateName);

  if (matches.hasNext()) {
    return matches.next();
  }

  throw new Error("Teacher certificate template not found. Set TEACHER_CERTIFICATE_TEMPLATE_ID or upload a Google Slides file named '" + templateName + "'");
}

function sendTeacherTrainingCertificateInternal(row, certificateNumber, issueDate) {
  const academyName = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName;
  const academyEmail = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyEmail) || DEFAULTS.academyEmail;
  const academyPhone = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyPhone) || DEFAULTS.academyPhone;
  const senderName = getScriptProperty(SCRIPT_PROPERTY_KEYS.senderName) || DEFAULTS.senderName;
  const fullName = clean(row["Full Name"]);
  const templateFile = getTeacherCertificateTemplateFile();
  const pdfBlob = buildTeacherCertificatePdfBlob(templateFile, row, certificateNumber, issueDate);
  const subject = "Certificate for " + clean(row["Training Type"]) + " | " + academyName;
  const plainTextBody =
    "Dear " + fullName + ",\n\n" +
    "Congratulations on completing " + clean(row["Training Type"]) + ". Your certificate is attached.\n\n" +
    "Certificate Number: " + certificateNumber + "\n\n" +
    "Regards,\n" + academyName + "\n" + academyPhone + "\n" + academyEmail;
  const htmlBody =
    "<p>Dear " + sanitizeHtmlText(fullName) + ",</p>" +
    "<p>Congratulations on completing " + sanitizeHtmlText(clean(row["Training Type"])) + ". Your certificate is attached.</p>" +
    "<p><strong>Certificate Number:</strong> " + sanitizeHtmlText(certificateNumber) + "</p>" +
    "<p>Regards,<br>" + sanitizeHtmlText(academyName) + "<br>" + sanitizeHtmlText(academyPhone) + "<br>" + sanitizeHtmlText(academyEmail) + "</p>";

  GmailApp.sendEmail(clean(row["Email"]), subject, plainTextBody, {
    htmlBody: htmlBody,
    attachments: [pdfBlob],
    name: senderName
  });
}

function buildTeacherCertificatePdfBlob(templateFile, row, certificateNumber, issueDate) {
  const parentFolder = templateFile.getParents().hasNext()
    ? templateFile.getParents().next()
    : DriveApp.getRootFolder();
  const fullName = clean(row["Full Name"]);
  const copyName = "Teacher Certificate - " + fullName + " - " + clean(row["Training ID"]);
  const workingCopy = templateFile.makeCopy(copyName, parentFolder);
  const presentation = SlidesApp.openById(workingCopy.getId());
  const issueDateText = formatDisplayDate(issueDate);
  const values = {
    "{{NAME}}": fullName,
    "{{FULL_NAME}}": fullName,
    "{{COURSE}}": clean(row["Training Type"]),
    "{{TRAINING_TYPE}}": clean(row["Training Type"]),
    "{{MODE}}": clean(row["Mode"]),
    "{{DATE}}": issueDateText,
    "{{ISSUE_DATE}}": issueDateText,
    "{{CERTIFICATE_NO}}": certificateNumber,
    "{{CERTIFICATE_NUMBER}}": certificateNumber,
    "{{TRAINING_ID}}": clean(row["Training ID"]),
    "{{ACADEMY_NAME}}": getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName
  };

  Object.keys(values).forEach(function(key) {
    presentation.replaceAllText(key, String(values[key] || ""));
  });

  presentation.saveAndClose();

  const pdfBlob = workingCopy.getAs(MimeType.PDF).setName(copyName + ".pdf");
  const archiveFolderId = getScriptProperty(SCRIPT_PROPERTY_KEYS.teacherCertificateArchiveFolderId);

  if (archiveFolderId) {
    DriveApp.getFolderById(archiveFolderId).createFile(pdfBlob.copyBlob());
  }

  workingCopy.setTrashed(true);
  return pdfBlob;
}
