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
      "Send Certificate": "",
      "Certificate Status": "",
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
    "Send Certificate",
    "Certificate Status",
    "Certificate Number",
    "Date of Cert",
    "Certificate Issue Date",
    "Certificate Sent Date",
    "Certificate PDF URL",
    "Certificate Image URL",
    "Certificate Error",
    "Send Certificate WhatsApp",
    "Certificate WhatsApp Status",
    "Certificate WhatsApp Link",
    "Certificate WhatsApp Sent Date",
    "Certificate WhatsApp Error",
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

  if (headerMap["Send Certificate"] && editedColumn === headerMap["Send Certificate"]) {
    if (clean(range.getValue()) === "Send") {
      setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Status", "Pending");
    }
    return;
  }

  if (headerMap["Send Certificate WhatsApp"] && editedColumn === headerMap["Send Certificate WhatsApp"]) {
    if (clean(range.getValue()) === "Send") {
      setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate WhatsApp Status", "Pending");
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
  const academyName = getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName;
  const academyEmail = getScriptProperty("ACADEMY_EMAIL") || DEFAULTS.academyEmail;
  const academyPhone = getScriptProperty("ACADEMY_PHONE") || DEFAULTS.academyPhone;
  const senderName = getScriptProperty("SENDER_NAME") || DEFAULTS.senderName;
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
    "{{ACADEMY_ADDRESS}}": getScriptProperty("ACADEMY_ADDRESS") || DEFAULTS.academyAddress,
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

function sendTeacherTrainingCertificate(payload) {
  authorizeAdmin(payload);

  const result = sendTeacherTrainingCertificateForTrainingId(clean(payload.trainingId));

  return jsonResponse({
    success: true,
    message: result.message
  });
}

function sendTeacherTrainingCertificateWhatsApp(payload) {
  authorizeAdmin(payload);

  const result = sendTeacherTrainingCertificateWhatsAppForTrainingId(clean(payload.trainingId));

  return jsonResponse({
    success: true,
    message: result.message
  });
}

function processPendingTeacherTrainingCertificates() {
  const sheet = getSheet(SHEET_NAMES.teacherTrainingAdmissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);

  Logger.log("Total rows to process: " + rows.length);
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  rows.forEach(function(row) {
    const sendCertificate = clean(row["Send Certificate"]);
    const certificateStatus = clean(row["Certificate Status"]);
    const trainingId = clean(row["Training ID"]);
    const email = clean(row["Email"]);
    const fullName = clean(row["Full Name"]);

    Logger.log("Row " + row._rowNumber + " - Training ID: " + trainingId + ", Send Certificate: " + sendCertificate + ", Certificate Status: " + certificateStatus + ", Email: " + email);

    // Only process if status is "Pending" or "Ready"
    if (certificateStatus !== "Pending" && certificateStatus !== "Ready") {
      Logger.log("Row " + row._rowNumber + " - Skipping (status is not Pending or Ready, current status: '" + certificateStatus + "')");
      skippedCount++;
      return;
    }

    // Skip if already sent to prevent duplicate emails
    if (certificateStatus && certificateStatus.indexOf("Certificate Sent") === 0) {
      Logger.log("Row " + row._rowNumber + " - Skipping (already sent)");
      skippedCount++;
      return;
    }

    // Skip if already failed to prevent repeated errors
    if (certificateStatus === "Failed") {
      Logger.log("Row " + row._rowNumber + " - Skipping (already failed)");
      skippedCount++;
      return;
    }

    // Skip if missing critical data
    if (!trainingId) {
      Logger.log("Row " + row._rowNumber + " - Skipping (missing Training ID)");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", "Missing Training ID");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Failed");
      errorCount++;
      return;
    }

    if (!fullName) {
      Logger.log("Row " + row._rowNumber + " - Skipping (missing Full Name)");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", "Missing Full Name");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Failed");
      errorCount++;
      return;
    }

    if (!email || !isValidEmail(email)) {
      Logger.log("Row " + row._rowNumber + " - Skipping (invalid or missing email: '" + email + "')");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", "Invalid or missing email");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Failed");
      errorCount++;
      return;
    }

    try {
      Logger.log("Row " + row._rowNumber + " - Processing certificate for training " + trainingId);
      const result = sendTeacherTrainingCertificateInternal(row);
      writeTeacherTrainingCertificateResult(sheet, row._rowNumber, headerMap, result);
      processedCount++;
      Logger.log("Row " + row._rowNumber + " - Certificate sent successfully");
    } catch (error) {
      const errorMsg = error && error.message ? error.message : "Unknown error";
      Logger.log("Row " + row._rowNumber + " - ERROR: " + errorMsg);
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", errorMsg.substring(0, 100));
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Failed");
      errorCount++;
    }
  });

  Logger.log("Summary - Processed: " + processedCount + ", Skipped: " + skippedCount + ", Errors: " + errorCount + " out of " + rows.length + " total rows");
}

function processPendingTeacherTrainingCertificateWhatsApps() {
  const sheet = getSheet(SHEET_NAMES.teacherTrainingAdmissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);

  rows.forEach(function(row) {
    const sendWhatsApp = clean(row["Send Certificate WhatsApp"]);
    const whatsAppStatus = clean(row["Certificate WhatsApp Status"]);

    if (sendWhatsApp !== "Send" && whatsAppStatus !== "Pending") {
      return;
    }

    try {
      const result = sendTeacherTrainingCertificateWhatsAppInternal(row);
      writeTeacherTrainingCertificateWhatsAppResult(sheet, row._rowNumber, headerMap, result);
    } catch (error) {
      const errorMsg = error && error.message ? error.message : "Unknown error";
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate WhatsApp Error", errorMsg.substring(0, 250));
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate WhatsApp Status", "Failed");
    }
  });
}

function sendTeacherTrainingCertificateForTrainingId(trainingId) {
  if (!trainingId) {
    throw new Error("Training ID is required");
  }

  const sheet = getSheet(SHEET_NAMES.teacherTrainingAdmissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);
  const row = rows.find(function(item) {
    return clean(item["Training ID"]) === trainingId;
  });

  if (!row) {
    throw new Error("Training admission not found");
  }

  const result = sendTeacherTrainingCertificateInternal(row);
  writeTeacherTrainingCertificateResult(sheet, row._rowNumber, headerMap, result);
  return {
    message: "Certificate sent for " + clean(row["Full Name"])
  };
}

function sendTeacherTrainingCertificateWhatsAppForTrainingId(trainingId) {
  if (!trainingId) {
    throw new Error("Training ID is required");
  }

  const sheet = getSheet(SHEET_NAMES.teacherTrainingAdmissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);
  const row = rows.find(function(item) {
    return clean(item["Training ID"]) === trainingId;
  });

  if (!row) { 
    throw new Error("Training admission not found");
  }

  const result = sendTeacherTrainingCertificateWhatsAppInternal(row);
  writeTeacherTrainingCertificateWhatsAppResult(sheet, row._rowNumber, headerMap, result);
  return {
    message: "Certificate WhatsApp link created for " + clean(row["Full Name"])
  };
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
      const result = sendTeacherTrainingCertificateInternal(row, certificateNumber, issueDate);
      writeTeacherTrainingCertificateResult(sheet, row._rowNumber, headerMap, result);
    } catch (error) {
      const errorMsg = error && error.message ? error.message : "Unknown error";
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", errorMsg.substring(0, 100));
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Failed");
    }
  });
}

function buildTeacherCertificateNumber(row) {
  // Use Certificate Number column if available, otherwise use Training ID
  const certNumber = clean(row["Certificate Number"]);
  if (certNumber) {
    return certNumber;
  }
  return clean(row["Training ID"]);
}

function getTeacherCertificateTemplateFile() {
  const templateId = getScriptProperty("TEACHER_CERTIFICATE_TEMPLATE_ID");
  if (templateId) {
    return DriveApp.getFileById(templateId);
  }

  const templateName = getScriptProperty("TEACHER_CERTIFICATE_TEMPLATE_NAME") || DEFAULTS.teacherCertificateTemplateName;
  const matches = DriveApp.getFilesByName(templateName);

  if (matches.hasNext()) {
    return matches.next();
  }

  throw new Error("Teacher certificate template not found. Set TEACHER_CERTIFICATE_TEMPLATE_ID or upload a Google Slides file named '" + templateName + "'");
}

function sendTeacherTrainingCertificateInternal(row, certificateNumber, issueDate) {
  Logger.log("=== Starting sendTeacherTrainingCertificateInternal ===");
  const email = clean(row["Email"]);
  const fullName = clean(row["Full Name"]);
  const trainingId = clean(row["Training ID"]);
  
  Logger.log("Email: '" + email + "'");
  Logger.log("Full Name: '" + fullName + "'");
  Logger.log("Training ID: '" + trainingId + "'");

  if (!email || !isValidEmail(email)) {
    throw new Error("Email is missing or invalid: '" + email + "'");
  }

  if (!fullName) {
    throw new Error("Full name is missing");
  }

  if (!trainingId) {
    throw new Error("Training ID is missing");
  }

  Logger.log("Building certificate number");
  const certNumber = certificateNumber || buildTeacherCertificateNumber(row);
  Logger.log("Certificate Number: '" + certNumber + "'");
  
  // Use Date of Cert column if available, otherwise use issueDate parameter or current date
  const dateOfCert = row["Date of Cert"];
  const certIssueDate = dateOfCert || issueDate || new Date();
  Logger.log("Issue Date: '" + certIssueDate + "'");
  
  Logger.log("Getting template file");
  let templateFile;
  try {
    templateFile = getTeacherCertificateTemplateFile();
    Logger.log("Template file obtained: " + templateFile.getName());
  } catch (templateError) {
    throw new Error("Failed to get certificate template: " + (templateError.message || "Unknown error"));
  }
  
  Logger.log("Calling buildTeacherCertificateFiles");
  const certificateFiles = buildTeacherCertificateFiles(templateFile, row, certNumber, certIssueDate);
  Logger.log("Certificate files built successfully");

  if (!certificateFiles.pdfFile) {
    throw new Error("PDF file was not created successfully");
  }

  Logger.log("Sending certificate email");
  sendTeacherTrainingCertificateEmail(row, certNumber, certificateFiles.pdfFile, certificateFiles.pdfUrl);
  Logger.log("Certificate email sent successfully");

  return {
    certificateNumber: certNumber,
    issueDate: certIssueDate,
    pdfUrl: certificateFiles.pdfFile.getUrl(),
    imageUrl: certificateFiles.imageFile ? certificateFiles.imageFile.getUrl() : ""
  };
}

function sendTeacherTrainingCertificateEmail(row, certificateNumber, pdfFile, pdfUrl) {
  const academyName = getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName;
  const academyEmail = getScriptProperty("ACADEMY_EMAIL") || DEFAULTS.academyEmail;
  const academyPhone = getScriptProperty("ACADEMY_PHONE") || DEFAULTS.academyPhone;
  const senderName = getScriptProperty("SENDER_NAME") || DEFAULTS.senderName;
  const fullName = clean(row["Full Name"]);
  const trainingType = clean(row["Training Type"]);
  const subject = "Certificate for " + trainingType + " | " + academyName;
  const plainTextBody =
    "Dear " + fullName + ",\n\n" +
    "Congratulations on completing " + trainingType + ". Your certificate is attached.\n\n" +
    "Certificate Number: " + certificateNumber + "\n\n" +
    "Regards,\n" + academyName + "\n" + academyPhone + "\n" + academyEmail;
  const htmlBody =
    "<p>Dear " + sanitizeHtmlText(fullName) + ",</p>" +
    "<p>Congratulations on completing " + sanitizeHtmlText(trainingType) + ". Your certificate is attached.</p>" +
    "<p><strong>Certificate Number:</strong> " + sanitizeHtmlText(certificateNumber) + "</p>" +
    "<p>Regards,<br>" + sanitizeHtmlText(academyName) + "<br>" + sanitizeHtmlText(academyPhone) + "<br>" + sanitizeHtmlText(academyEmail) + "</p>";

  GmailApp.sendEmail(clean(row["Email"]), subject, plainTextBody, {
    htmlBody: htmlBody,
    attachments: [pdfFile.getBlob()],
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
    "{{ACADEMY_NAME}}": getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName
  };

  Object.keys(values).forEach(function(key) {
    presentation.replaceAllText(key, String(values[key] || ""));
  });

  presentation.saveAndClose();

  const pdfBlob = workingCopy.getAs(MimeType.PDF).setName(copyName + ".pdf");
  const archiveFolderId = getScriptProperty("TEACHER_CERTIFICATE_ARCHIVE_FOLDER_ID");

  if (archiveFolderId) {
    DriveApp.getFolderById(archiveFolderId).createFile(pdfBlob.copyBlob());
  }

  workingCopy.setTrashed(true);
  return pdfBlob;
}

function sendTeacherTrainingCertificateWhatsAppInternal(row) {
  const certificateNumber = clean(row["Certificate Number"]) || buildTeacherCertificateNumber(row);
  // Use Date of Cert column if available, otherwise use Certificate Issue Date or current date
  const dateOfCert = row["Date of Cert"];
  const issueDate = dateOfCert || row["Certificate Issue Date"] || new Date();
  const templateFile = getTeacherCertificateTemplateFile();
  
  // Build certificate files similar to admission certificates
  const certificateFiles = buildTeacherCertificateFiles(templateFile, row, certificateNumber, issueDate);

  if (!certificateFiles.imageFile) {
    throw new Error("Certificate image could not be generated");
  }

  const whatsAppLink = buildTeacherTrainingCertificateWhatsAppLink(row, certificateNumber, certificateFiles.imageFile.getUrl());

  return {
    certificateNumber: certificateNumber,
    issueDate: issueDate,
    pdfUrl: certificateFiles.pdfFile.getUrl(),
    imageUrl: certificateFiles.imageFile.getUrl(),
    whatsAppStatus: whatsAppLink ? "Link Ready" : "Mobile Missing",
    whatsAppLink: whatsAppLink,
    whatsAppError: whatsAppLink ? "" : "WhatsApp mobile number missing or invalid"
  };
}

function buildTeacherCertificateFiles(templateFile, row, certificateNumber, issueDate) {
  Logger.log("=== Starting buildTeacherCertificateFiles ===");
  
  let parentFolder;
  try {
    parentFolder = templateFile.getParents().hasNext()
      ? templateFile.getParents().next()
      : DriveApp.getRootFolder();
    Logger.log("Parent folder: " + parentFolder.getName());
  } catch (folderError) {
    Logger.log("Error getting parent folder: " + folderError.message + ", using root folder");
    parentFolder = DriveApp.getRootFolder();
  }
  
  let outputFolder = parentFolder;
  const archiveFolderId = getScriptProperty("TEACHER_CERTIFICATE_ARCHIVE_FOLDER_ID");
  if (archiveFolderId) {
    try {
      outputFolder = DriveApp.getFolderById(archiveFolderId);
      Logger.log("Using archive folder: " + outputFolder.getName());
    } catch (archiveError) {
      Logger.log("Error accessing archive folder: " + archiveError.message + ", using parent folder");
      outputFolder = parentFolder;
    }
  }
  
  const fullName = clean(row["Full Name"]);
  const trainingId = clean(row["Training ID"]);
  
  Logger.log("Full Name: '" + fullName + "'");
  Logger.log("Training ID: '" + trainingId + "'");
  
  let copyName = "Teacher Certificate - " + fullName;
  let workingCopy;
  
  try {
    Logger.log("Attempting to create copy with full name: '" + copyName + "'");
    workingCopy = templateFile.makeCopy(copyName, parentFolder);
    Logger.log("Successfully created copy with full name");
  } catch (nameError) {
    Logger.log("Failed to create copy with full name: " + nameError.message);
    copyName = "Teacher Certificate - " + trainingId;
    Logger.log("Attempting fallback with Training ID: '" + copyName + "'");
    try {
      workingCopy = templateFile.makeCopy(copyName, parentFolder);
      Logger.log("Successfully created copy with Training ID");
    } catch (fallbackError) {
      Logger.log("Failed to create copy with Training ID: " + fallbackError.message);
      throw new Error("Unable to create certificate copy. Full name error: " + nameError.message + ", Fallback error: " + fallbackError.message);
    }
  }
  
  Logger.log("Opening presentation with ID: " + workingCopy.getId());
  const presentation = SlidesApp.openById(workingCopy.getId());
  const firstSlide = presentation.getSlides()[0];
  const firstSlideObjectId = firstSlide.getObjectId();
  const issueDateText = formatDisplayDate(issueDate);
  const values = {
    "{{NAME}}": fullName,
    "{{FULL_NAME}}": fullName,
    "{{STUDENT}}": fullName,
    "{{COURSE}}": clean(row["Training Type"]),
    "{{TRAINING_TYPE}}": clean(row["Training Type"]),
    "{{MODE}}": clean(row["Mode"]),
    "{{DATE}}": issueDateText,
    "{{ISSUE_DATE}}": issueDateText,
    "{{CERT_ID}}": certificateNumber,
    "{{CERTIFICATE_NO}}": certificateNumber,
    "{{CERTIFICATE_NUMBER}}": certificateNumber,
    "{{TRAINING_ID}}": trainingId,
    "{{ACADEMY_NAME}}": getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName
  };

  Object.keys(values).forEach(function(key) {
    presentation.replaceAllText(key, String(values[key] || ""));
  });

  Logger.log("Saving and closing presentation");
  presentation.saveAndClose();
  Logger.log("Presentation saved and closed");

  // Use Drive API export method
  Logger.log("Creating PDF file with name: " + copyName + ".pdf");
  const exportUrl = "https://www.googleapis.com/drive/v3/files/" + workingCopy.getId() + "/export?mimeType=application/pdf";
  Logger.log("Export URL: " + exportUrl);
  
  let response;
  try {
    response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    Logger.log("PDF export response code: " + response.getResponseCode());
    
    if (response.getResponseCode() !== 200) {
      throw new Error("PDF export failed with status " + response.getResponseCode() + ": " + response.getContentText());
    }
  } catch (fetchError) {
    Logger.log("PDF export fetch error: " + (fetchError.message || "Unknown error"));
    throw new Error("Failed to export PDF: " + (fetchError.message || "Unknown error"));
  }
  
  const pdfBlob = response.getBlob().setName(copyName + ".pdf");
  Logger.log("PDF blob created, size: " + pdfBlob.getBytes().length);
  
  if (pdfBlob.getBytes().length === 0) {
    throw new Error("PDF blob is empty - export may have failed");
  }
  
  let pdfFile;
  try {
    pdfFile = outputFolder.createFile(pdfBlob.copyBlob());
    Logger.log("PDF file created in Drive");
  } catch (createError) {
    throw new Error("Failed to create PDF file in Drive: " + (createError.message || "Unknown error"));
  }
  
  try {
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log("PDF file sharing set");
  } catch (sharingError) {
    Logger.log("Warning: Failed to set PDF sharing: " + sharingError.message);
  }

  let imageFile = null;
  try {
    Logger.log("Creating PNG file with name: " + copyName + ".png");
    const imageExportUrl = "https://docs.google.com/presentation/d/" + workingCopy.getId() + "/export/png?pageid=" + encodeURIComponent(firstSlideObjectId);
    const imageResponse = UrlFetchApp.fetch(imageExportUrl, {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    
    if (imageResponse.getResponseCode() !== 200) {
      Logger.log("PNG export failed with status " + imageResponse.getResponseCode() + " - will continue without image");
    } else {
      const imageBlob = imageResponse.getBlob().setName(copyName + ".png");
      if (imageBlob.getBytes().length > 0) {
        imageFile = outputFolder.createFile(imageBlob.copyBlob());
        try {
          imageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (sharingError) {
          Logger.log("Warning: Failed to set PNG sharing: " + sharingError.message);
        }
        Logger.log("PNG file created successfully");
      } else {
        Logger.log("PNG blob is empty - skipping image creation");
      }
    }
  } catch (error) {
    Logger.log("Certificate image export failed: " + (error && error.message ? error.message : "Unknown error") + " - will continue without image");
  }

  workingCopy.setTrashed(true);

  return {
    pdfFile: pdfFile,
    imageFile: imageFile,
    pdfUrl: pdfFile.getUrl()
  };
}

function buildTeacherTrainingCertificateWhatsAppLink(row, certificateNumber, certificateUrl) {
  const mobile = normalizeIndianMobileForWhatsApp(clean(row["Mobile"]));
  if (!mobile) {
    return "";
  }

  const academyName = getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName;
  const academyPhone = getScriptProperty("ACADEMY_PHONE") || DEFAULTS.academyPhone;
  const fullName = clean(row["Full Name"]);
  const trainingType = clean(row["Training Type"]);
  const message =
    "Dear " + fullName + ",\n\n" +
    "Congratulations on completing " + trainingType + ". Please find the certificate image here:\n" +
    certificateUrl + "\n\n" +
    "Certificate Number: " + certificateNumber + "\n\n" +
    "Regards,\n" + academyName + "\n" + academyPhone;

  return "https://wa.me/" + mobile + "?text=" + encodeURIComponent(message);
}

function writeTeacherTrainingCertificateResult(sheet, rowNumber, headerMap, result) {
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Number", result.certificateNumber);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Issue Date", result.issueDate);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Sent Date", new Date());
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate PDF URL", result.pdfUrl);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Image URL", result.imageUrl);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Error", "");
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Status", "Certificate Sent - " + result.certificateNumber);
  // Don't clear Send Certificate column - keep it as "Send"
}

function writeTeacherTrainingCertificateWhatsAppResult(sheet, rowNumber, headerMap, result) {
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Number", result.certificateNumber);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Issue Date", result.issueDate);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate PDF URL", result.pdfUrl);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Image URL", result.imageUrl);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate WhatsApp Status", result.whatsAppStatus);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate WhatsApp Link", result.whatsAppLink);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate WhatsApp Sent Date", "");
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate WhatsApp Error", result.whatsAppError);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Send Certificate WhatsApp", "");
}
