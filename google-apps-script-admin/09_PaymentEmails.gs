function sendPaymentEmail(payload) {
  authorizeAdmin(payload);

  const result = sendPaymentEmailInternal({
    admissionId: clean(payload.admissionId),
    paymentId: clean(payload.paymentId),
    emailType: clean(payload.emailType)
  });

  return jsonResponse({
    success: true,
    message: result.message
  });
}

function sendPaymentEmailInternal(options) {
  const emailType = normalizePaymentEmailType(options.emailType);
  const admissionsSheet = getSheet(SHEET_NAMES.admissions);
  const paymentsSheet = getSheet(SHEET_NAMES.payments);
  const context = getPaymentEmailContext(admissionsSheet, paymentsSheet, options.admissionId, options.paymentId);
  const email = clean(context.admission["Email"]);

  if (!email) {
    throw new Error("Parent email is missing for this admission");
  }

  if (!isValidEmail(email)) {
    throw new Error("Parent email is invalid");
  }

  const templateFile = getReceiptTemplateFile();
  const paymentMeta = buildPaymentEmailMeta(emailType, context);
  const pdfBlob = buildReceiptPdfBlob(templateFile, paymentMeta.templateValues);

  GmailApp.sendEmail(email, paymentMeta.subject, paymentMeta.plainTextBody, {
    htmlBody: paymentMeta.htmlBody,
    attachments: [pdfBlob],
    name: paymentMeta.senderName
  });

  recordPaymentEmailLog(admissionsSheet, context.admission._rowNumber, emailType);

  return {
    message: paymentMeta.successMessage
  };
}

function normalizePaymentEmailType(value) {
  const normalized = clean(value).toLowerCase();
  if (normalized === "full-payment" || normalized === "pending-reminder" || normalized === "receipt") {
    return normalized;
  }
  return "receipt";
}

function getPaymentEmailContext(admissionsSheet, paymentsSheet, admissionId, paymentId) {
  const admissions = getSheetObjects(admissionsSheet);
  const payments = getSheetObjects(paymentsSheet);
  let paymentRow = null;
  let admissionRow = null;
  let resolvedAdmissionId = clean(admissionId);

  if (paymentId) {
    paymentRow = payments.find(function(item) {
      return clean(item["Payment ID"]) === paymentId;
    }) || null;

    if (!paymentRow) {
      throw new Error("Payment not found");
    }

    resolvedAdmissionId = clean(paymentRow["Admission ID"]);
  }

  if (!resolvedAdmissionId) {
    throw new Error("Admission ID is required");
  }

  admissionRow = admissions.find(function(item) {
    return clean(item["Admission ID"]) === resolvedAdmissionId;
  }) || null;

  if (!admissionRow) {
    throw new Error("Admission not found");
  }

  if (!paymentRow) {
    paymentRow = payments
      .filter(function(item) {
        return clean(item["Admission ID"]) === resolvedAdmissionId;
      })
      .sort(function(left, right) {
        return right._rowNumber - left._rowNumber;
      })[0] || null;
  }

  const headerMap = getHeaderMap(admissionsSheet);
  const financials = {
    totalFee: toNumber(admissionRow["Total Fee"], 0),
    discount: toNumber(admissionRow["Discount"], 0),
    manualAdjustment: toNumber(admissionRow["Manual Adjustment"], 0),
    adjustedFee: toNumber(admissionRow["Adjusted Fee"], 0),
    totalPaid: toNumber(admissionRow["Total Paid"], 0),
    pending: toNumber(admissionRow["Pending"], 0),
    paymentStatus: clean(admissionRow["Payment Status"])
  };

  if (!financials.adjustedFee && (financials.totalFee || financials.discount || financials.manualAdjustment)) {
    const recalculated = recalculateAdmissionFinancials(admissionsSheet, admissionRow._rowNumber);
    financials.adjustedFee = recalculated.adjustedFee;
    financials.totalPaid = recalculated.totalPaid;
    financials.pending = recalculated.pending;
    financials.paymentStatus = recalculated.paymentStatus;
    admissionRow["Adjusted Fee"] = recalculated.adjustedFee;
    admissionRow["Total Paid"] = recalculated.totalPaid;
    admissionRow["Pending"] = recalculated.pending;
    admissionRow["Payment Status"] = recalculated.paymentStatus;
  }

  if (!financials.paymentStatus) {
    financials.paymentStatus = clean(admissionRow["Payment Status"]) || DEFAULTS.paymentStatus;
  }

  return {
    admission: admissionRow,
    payment: paymentRow,
    financials: financials,
    admissionHeaderMap: headerMap
  };
}

function getReceiptTemplateFile() {
  const templateId = getScriptProperty(SCRIPT_PROPERTY_KEYS.receiptTemplateId) || DEFAULTS.receiptTemplateId;
  if (templateId) {
    return DriveApp.getFileById(templateId);
  }

  const templateName = getScriptProperty(SCRIPT_PROPERTY_KEYS.receiptTemplateName) || DEFAULTS.receiptTemplateName;
  const matches = DriveApp.getFilesByName(templateName);

  if (matches.hasNext()) {
    return matches.next();
  }

  throw new Error("Receipt template not found. Set RECEIPT_TEMPLATE_ID or upload a file named '" + templateName + "'");
}

function buildReceiptPdfBlob(templateFile, templateValues) {
  const parentFolder = templateFile.getParents().hasNext()
    ? templateFile.getParents().next()
    : DriveApp.getRootFolder();

  const copyName = templateValues["{{RECEIPT_TITLE}}"] + " - " + templateValues["{{STUDENT_NAME}}"];
  const workingCopy = templateFile.makeCopy(copyName, parentFolder);
  const document = DocumentApp.openById(workingCopy.getId());
  const body = document.getBody();

  Object.keys(templateValues).forEach(function(key) {
    body.replaceText(escapeReplaceTextPattern(key), String(templateValues[key] || ""));
  });

  document.saveAndClose();

  const pdfBlob = workingCopy.getAs(MimeType.PDF).setName(copyName + ".pdf");
  const archiveFolderId = getScriptProperty(SCRIPT_PROPERTY_KEYS.receiptArchiveFolderId);

  if (archiveFolderId) {
    DriveApp.getFolderById(archiveFolderId).createFile(pdfBlob.copyBlob());
  }

  workingCopy.setTrashed(true);
  return pdfBlob;
}

function buildPaymentEmailMeta(emailType, context) {
  const academyName = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName;
  const academyEmail = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyEmail) || DEFAULTS.academyEmail;
  const academyPhone = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyPhone) || DEFAULTS.academyPhone;
  const academyAddress = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyAddress) || DEFAULTS.academyAddress;
  const senderName = getScriptProperty(SCRIPT_PROPERTY_KEYS.senderName) || DEFAULTS.senderName;
  const admission = context.admission;
  const payment = context.payment || {};
  const financials = context.financials;
  const studentName = clean(admission["Student Name"]);
  const parentName = clean(admission["Parent Name"]);
  const paymentDate = formatDisplayDate(payment["Payment Date"] || new Date());
  // Use payment amount if available, otherwise use total paid from admission
  const paymentAmount = toNumber(payment["Amount"], 0) || financials.totalPaid;
  const paymentId = clean(payment["Payment ID"]) || "Pending";
  const emailTypeLabel = emailType === "full-payment"
    ? "Full Payment Confirmation"
    : emailType === "pending-reminder"
      ? "Pending Payment Reminder"
      : "Payment Receipt";

  const summaryLine = emailType === "full-payment"
    ? "Your child's fee payment is now fully completed."
    : emailType === "pending-reminder"
      ? "This is a gentle reminder that fee payment is still pending."
      : "Thank you for your payment. Please find the receipt attached.";

  const subject = emailType === "full-payment"
    ? "Full payment completed for " + studentName + " | " + academyName
    : emailType === "pending-reminder"
      ? "Pending fee reminder for " + studentName + " | " + academyName
      : "Payment receipt for " + studentName + " | " + academyName;

  const htmlBody =
    "<p>Dear " + sanitizeHtmlText(parentName || "Parent") + ",</p>" +
    "<p>" + sanitizeHtmlText(summaryLine) + "</p>" +
    "<p><strong>Student:</strong> " + sanitizeHtmlText(studentName) + "<br>" +
    "<strong>Admission ID:</strong> " + sanitizeHtmlText(clean(admission["Admission ID"])) + "<br>" +
    "<strong>Payment Status:</strong> " + sanitizeHtmlText(financials.paymentStatus || DEFAULTS.paymentStatus) + "<br>" +
    "<strong>Total Paid:</strong> " + sanitizeHtmlText(formatMoney(financials.totalPaid)) + "<br>" +
    "<strong>Pending:</strong> " + sanitizeHtmlText(formatMoney(financials.pending)) + "</p>" +
    "<p>The PDF is attached for your records.</p>" +
    "<p>Regards,<br>" + sanitizeHtmlText(academyName) + "<br>" + sanitizeHtmlText(academyPhone) + "<br>" + sanitizeHtmlText(academyEmail) + "</p>";

  return {
    senderName: senderName,
    subject: subject,
    plainTextBody:
      "Dear " + (parentName || "Parent") + ",\n\n" +
      summaryLine + "\n\n" +
      "Student: " + studentName + "\n" +
      "Admission ID: " + clean(admission["Admission ID"]) + "\n" +
      "Payment Status: " + (financials.paymentStatus || DEFAULTS.paymentStatus) + "\n" +
      "Total Paid: " + formatMoney(financials.totalPaid) + "\n" +
      "Pending: " + formatMoney(financials.pending) + "\n\n" +
      "Regards,\n" + academyName + "\n" + academyPhone + "\n" + academyEmail,
    htmlBody: htmlBody,
    successMessage: emailTypeLabel + " sent successfully to " + clean(admission["Email"]),
    templateValues: {
      "{{RECEIPT_TITLE}}": emailTypeLabel,
      "{{EMAIL_TYPE_LABEL}}": emailTypeLabel,
      "{{EMAIL_MESSAGE}}": summaryLine,
      "{{ACADEMY_NAME}}": academyName,
      "{{ACADEMY_EMAIL}}": academyEmail,
      "{{ACADEMY_PHONE}}": academyPhone,
      "{{ACADEMY_ADDRESS}}": academyAddress,
      "{{RECEIPT_NO}}": clean(admission["Admission ID"]),
      "{{PAYMENT_ID}}": paymentId,
      "{{DATE}}": paymentDate,
      "{{PAYMENT_DATE}}": paymentDate,
      "{{STUDENT}}": studentName,
      "{{STUDENT_NAME}}": studentName,
      "{{PARENT_NAME}}": parentName,
      "{{ADMISSION_ID}}": clean(admission["Admission ID"]),
      "{{PARENT_EMAIL}}": clean(admission["Email"]),
      "{{MOBILE}}": clean(admission["Mobile"]),
      "{{LEVEL}}": clean(admission["Level"]),
      "{{MODE}}": clean(admission["Mode"]),
      "{{BATCH_CODE}}": clean(admission["Batch Code"]),
      "{{PAYMENT_MODE}}": clean(payment["Payment Mode"]),
      "{{TRANSACTION_ID}}": clean(payment["Transaction ID"]),
      "{{AMOUNT}}": Utilities.formatString("%.0f", paymentAmount),
      "{{AMOUNT_PAID}}": formatMoney(paymentAmount),
      "{{TOTAL_FEE}}": formatMoney(financials.totalFee),
      "{{DISCOUNT}}": formatMoney(financials.discount),
      "{{MANUAL_ADJUSTMENT}}": formatMoney(financials.manualAdjustment),
      "{{ADJUSTED_FEE}}": formatMoney(financials.adjustedFee),
      "{{TOTAL_PAID}}": formatMoney(financials.totalPaid),
      "{{PENDING_AMOUNT}}": formatMoney(financials.pending),
      "{{PAYMENT_STATUS}}": financials.paymentStatus || DEFAULTS.paymentStatus,
      "{{NOTES}}": clean(payment["Notes"]) || clean(admission["Notes"]),
      "{{TODAY_DATE}}": formatDisplayDate(new Date())
    }
  };
}

function recordPaymentEmailLog(admissionsSheet, rowNumber, emailType) {
  const headerMap = getHeaderMap(admissionsSheet);
  const now = new Date();

  setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Last Email Sent Date", now);

  if (emailType === "receipt") {
    setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Last Receipt Sent Date", now);
  }

  if (emailType === "full-payment") {
    setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Full Payment Email Date", now);
  }

  if (emailType === "pending-reminder") {
    setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Last Payment Reminder Date", now);
  }
}

function formatMoney(value) {
  const amount = toNumber(value, 0);
  return "INR " + Utilities.formatString("%.0f", amount);
}

function formatDisplayDate(value) {
  if (!value) {
    return "";
  }

  const timezone = Session.getScriptTimeZone();

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timezone, "dd MMM yyyy");
  }

  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, timezone, "dd MMM yyyy");
  }

  return clean(value);
}

function sanitizeHtmlText(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeReplaceTextPattern(value) {
  return String(value || "").replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
}
