function sendAdmissionCertificate(payload) {
  authorizeAdmin(payload);

  const result = sendAdmissionCertificateForAdmissionId(clean(payload.admissionId));

  return jsonResponse({
    success: true,
    message: result.message
  });
}

function sendAdmissionCertificateWhatsApp(payload) {
  authorizeAdmin(payload);

  const result = sendAdmissionCertificateWhatsAppForAdmissionId(clean(payload.admissionId));

  return jsonResponse({
    success: true,
    message: result.message
  });
}

function processPendingAdmissionCertificates() {
  const sheet = getSheet(SHEET_NAMES.admissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);

  rows.forEach(function(row) {
    const sendCertificate = clean(row["Send Certificate"]);
    const certificateStatus = clean(row["Certificate Status"]);

    if (sendCertificate !== "Send" && certificateStatus !== "Pending" && certificateStatus !== "Ready") {
      return;
    }

    try {
      const result = sendAdmissionCertificateInternal(row);
      writeAdmissionCertificateResult(sheet, row._rowNumber, headerMap, result);
    } catch (error) {
      const errorMsg = error && error.message ? error.message : "Unknown error";
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", errorMsg.substring(0, 100));
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Failed");
    }
  });
}

function processPendingAdmissionCertificateWhatsApps() {
  const sheet = getSheet(SHEET_NAMES.admissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);

  rows.forEach(function(row) {
    const sendWhatsApp = clean(row["Send Certificate WhatsApp"]);
    const whatsAppStatus = clean(row["Certificate WhatsApp Status"]);

    if (sendWhatsApp !== "Send" && whatsAppStatus !== "Pending") {
      return;
    }

    try {
      const result = sendAdmissionCertificateWhatsAppInternal(row);
      writeAdmissionCertificateWhatsAppResult(sheet, row._rowNumber, headerMap, result);
    } catch (error) {
      const errorMsg = error && error.message ? error.message : "Unknown error";
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate WhatsApp Error", errorMsg.substring(0, 250));
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate WhatsApp Status", "Failed");
    }
  });
}

function sendAdmissionCertificateForAdmissionId(admissionId) {
  if (!admissionId) {
    throw new Error("Admission ID is required");
  }

  const sheet = getSheet(SHEET_NAMES.admissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);
  const row = rows.find(function(item) {
    return clean(item["Admission ID"]) === admissionId;
  });

  if (!row) {
    throw new Error("Admission not found");
  }

  const result = sendAdmissionCertificateInternal(row);
  writeAdmissionCertificateResult(sheet, row._rowNumber, headerMap, result);
  return {
    message: "Certificate sent for " + clean(row["Student Name"])
  };
}

function sendAdmissionCertificateWhatsAppForAdmissionId(admissionId) {
  if (!admissionId) {
    throw new Error("Admission ID is required");
  }

  const sheet = getSheet(SHEET_NAMES.admissions);
  const rows = getSheetObjects(sheet);
  const headerMap = getHeaderMap(sheet);
  const row = rows.find(function(item) {
    return clean(item["Admission ID"]) === admissionId;
  });

  if (!row) {
    throw new Error("Admission not found");
  }

  const result = sendAdmissionCertificateWhatsAppInternal(row);
  writeAdmissionCertificateWhatsAppResult(sheet, row._rowNumber, headerMap, result);
  return {
    message: "Certificate WhatsApp link created for " + clean(row["Student Name"])
  };
}

function sendAdmissionCertificateInternal(row) {
  const email = clean(row["Email"]);

  if (!email || !isValidEmail(email)) {
    throw new Error("Parent email is missing or invalid");
  }

  const certificateNumber = clean(row["Certificate Number"]) || buildAdmissionCertificateNumber(row);
  const issueDate = row["Certificate Issue Date"] || new Date();
  const templateFile = getAdmissionCertificateTemplateFile();
  const certificateFiles = buildAdmissionCertificateFiles(templateFile, row, certificateNumber, issueDate);

  sendAdmissionCertificateEmail(row, certificateNumber, certificateFiles.pdfFile);

  return {
    certificateNumber: certificateNumber,
    issueDate: issueDate,
    pdfUrl: certificateFiles.pdfFile.getUrl(),
    imageUrl: certificateFiles.imageFile ? certificateFiles.imageFile.getUrl() : ""
  };
}

function sendAdmissionCertificateWhatsAppInternal(row) {
  const certificateNumber = clean(row["Certificate Number"]) || buildAdmissionCertificateNumber(row);
  const issueDate = row["Certificate Issue Date"] || new Date();
  const templateFile = getAdmissionCertificateTemplateFile();
  const certificateFiles = buildAdmissionCertificateFiles(templateFile, row, certificateNumber, issueDate);

  if (!certificateFiles.imageFile) {
    throw new Error("Certificate image could not be generated");
  }

  const whatsAppLink = buildAdmissionCertificateWhatsAppLink(row, certificateNumber, certificateFiles.imageFile.getUrl());

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

function writeAdmissionCertificateResult(sheet, rowNumber, headerMap, result) {
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Number", result.certificateNumber);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Issue Date", result.issueDate);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Sent Date", new Date());
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate PDF URL", result.pdfUrl);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Image URL", result.imageUrl);
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Error", "");
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Status", "Sent");
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Send Certificate", "");
}

function writeAdmissionCertificateWhatsAppResult(sheet, rowNumber, headerMap, result) {
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

function buildAdmissionCertificateNumber(row) {
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy");
  return "EKH-ST-" + year + "-" + clean(row["Admission ID"]);
}

function getAdmissionCertificateTemplateFile() {
  const templateId = getScriptProperty(SCRIPT_PROPERTY_KEYS.studentCertificateTemplateId);
  if (templateId) {
    return DriveApp.getFileById(templateId);
  }

  const templateName = getScriptProperty(SCRIPT_PROPERTY_KEYS.studentCertificateTemplateName) || DEFAULTS.studentCertificateTemplateName;
  const matches = DriveApp.getFilesByName(templateName);

  if (matches.hasNext()) {
    return matches.next();
  }

  throw new Error("Student certificate template not found. Set STUDENT_CERTIFICATE_TEMPLATE_ID or upload a Google Slides file named '" + templateName + "'");
}

function buildAdmissionCertificateFiles(templateFile, row, certificateNumber, issueDate) {
  const parentFolder = templateFile.getParents().hasNext()
    ? templateFile.getParents().next()
    : DriveApp.getRootFolder();
  const archiveFolderId = getScriptProperty(SCRIPT_PROPERTY_KEYS.studentCertificateArchiveFolderId);
  const outputFolder = archiveFolderId ? DriveApp.getFolderById(archiveFolderId) : parentFolder;
  const studentName = clean(row["Student Name"]);
  const copyName = "Student Certificate - " + studentName + " - " + clean(row["Admission ID"]);
  const workingCopy = templateFile.makeCopy(copyName, parentFolder);
  const presentation = SlidesApp.openById(workingCopy.getId());
  const firstSlide = presentation.getSlides()[0];
  const firstSlideObjectId = firstSlide.getObjectId();
  const issueDateText = formatDisplayDate(issueDate);
  const values = {
    "{{NAME}}": studentName,
    "{{STUDENT}}": studentName,
    "{{STUDENT_NAME}}": studentName,
    "{{PARENT_NAME}}": clean(row["Parent Name"]),
    "{{COURSE}}": clean(row["Level"]),
    "{{LEVEL}}": clean(row["Level"]),
    "{{MODE}}": clean(row["Mode"]),
    "{{BATCH_CODE}}": clean(row["Batch Code"]),
    "{{DATE}}": issueDateText,
    "{{ISSUE_DATE}}": issueDateText,
    "{{CERTIFICATE_NO}}": certificateNumber,
    "{{CERTIFICATE_NUMBER}}": certificateNumber,
    "{{ADMISSION_ID}}": clean(row["Admission ID"]),
    "{{ACADEMY_NAME}}": getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName
  };

  Object.keys(values).forEach(function(key) {
    presentation.replaceAllText(key, String(values[key] || ""));
  });

  presentation.saveAndClose();

  const pdfFile = outputFolder.createFile(workingCopy.getAs(MimeType.PDF).setName(copyName + ".pdf"));
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  let imageFile = null;
  try {
    const imageBlob = exportSlideAsPngBlob(workingCopy.getId(), firstSlideObjectId).setName(copyName + ".png");
    imageFile = outputFolder.createFile(imageBlob);
    imageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    Logger.log("Certificate image export failed: " + (error && error.message ? error.message : "Unknown error"));
  }

  workingCopy.setTrashed(true);

  return {
    pdfFile: pdfFile,
    imageFile: imageFile
  };
}

function exportSlideAsPngBlob(presentationId, slideObjectId) {
  const url = "https://docs.google.com/presentation/d/" + encodeURIComponent(presentationId) +
    "/export/png?id=" + encodeURIComponent(presentationId) +
    "&pageid=" + encodeURIComponent(slideObjectId);
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("PNG export failed with status " + response.getResponseCode());
  }

  return response.getBlob();
}

function sendAdmissionCertificateEmail(row, certificateNumber, pdfFile) {
  const academyName = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName;
  const academyEmail = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyEmail) || DEFAULTS.academyEmail;
  const academyPhone = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyPhone) || DEFAULTS.academyPhone;
  const senderName = getScriptProperty(SCRIPT_PROPERTY_KEYS.senderName) || DEFAULTS.senderName;
  const parentName = clean(row["Parent Name"]) || "Parent";
  const studentName = clean(row["Student Name"]);
  const level = clean(row["Level"]);
  const subject = "Certificate for " + studentName + " | " + academyName;
  const plainTextBody =
    "Dear " + parentName + ",\n\n" +
    "Congratulations to " + studentName + " on completing " + level + ". The certificate is attached.\n\n" +
    "Certificate Number: " + certificateNumber + "\n\n" +
    "Regards,\n" + academyName + "\n" + academyPhone + "\n" + academyEmail;
  const htmlBody =
    "<p>Dear " + sanitizeHtmlText(parentName) + ",</p>" +
    "<p>Congratulations to " + sanitizeHtmlText(studentName) + " on completing " + sanitizeHtmlText(level) + ". The certificate is attached.</p>" +
    "<p><strong>Certificate Number:</strong> " + sanitizeHtmlText(certificateNumber) + "</p>" +
    "<p>Regards,<br>" + sanitizeHtmlText(academyName) + "<br>" + sanitizeHtmlText(academyPhone) + "<br>" + sanitizeHtmlText(academyEmail) + "</p>";

  GmailApp.sendEmail(clean(row["Email"]), subject, plainTextBody, {
    htmlBody: htmlBody,
    attachments: [pdfFile.getBlob()],
    name: senderName
  });
}

function buildAdmissionCertificateWhatsAppLink(row, certificateNumber, certificateUrl) {
  const mobile = normalizeIndianMobileForWhatsApp(clean(row["Mobile"]));
  if (!mobile) {
    return "";
  }

  const academyName = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName;
  const academyPhone = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyPhone) || DEFAULTS.academyPhone;
  const studentName = clean(row["Student Name"]);
  const level = clean(row["Level"]);
  const message =
    "Dear Parent,\n\n" +
    "Congratulations to " + studentName + " on completing " + level + ". Please find the certificate image here:\n" +
    certificateUrl + "\n\n" +
    "Certificate Number: " + certificateNumber + "\n\n" +
    "Regards,\n" + academyName + "\n" + academyPhone;

  return "https://wa.me/" + mobile + "?text=" + encodeURIComponent(message);
}

function normalizeIndianMobileForWhatsApp(value) {
  const digits = clean(value).replace(/\D/g, "");

  if (digits.length === 10) {
    return "91" + digits;
  }

  if (digits.length === 12 && digits.indexOf("91") === 0) {
    return digits;
  }

  return digits.length >= 11 ? digits : "";
}
