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

  Logger.log("Total rows to process: " + rows.length);
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  rows.forEach(function(row) {
    const sendCertificate = clean(row["Send Certificate"]);
    const certificateStatus = clean(row["Certificate Status"]);
    const admissionId = clean(row["Admission ID"]);
    const email = clean(row["Email"]);
    const studentName = clean(row["Student Name"]);

    Logger.log("Row " + row._rowNumber + " - Admission ID: " + admissionId + ", Send Certificate: " + sendCertificate + ", Certificate Status: " + certificateStatus + ", Email: " + email);

    // Only process if status is "Pending" or "Ready"
    // Do NOT process based on "Send Certificate" column value alone
    // This ensures onEdit must set status to "Pending" first
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
    if (!admissionId) {
      Logger.log("Row " + row._rowNumber + " - Skipping (missing Admission ID)");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", "Missing Admission ID");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Status", "Failed");
      errorCount++;
      return;
    }

    if (!studentName) {
      Logger.log("Row " + row._rowNumber + " - Skipping (missing Student Name)");
      setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Certificate Error", "Missing Student Name");
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
      Logger.log("Row " + row._rowNumber + " - Processing certificate for admission " + admissionId);
      const result = sendAdmissionCertificateInternal(row);
      writeAdmissionCertificateResult(sheet, row._rowNumber, headerMap, result);
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
  Logger.log("=== Starting sendAdmissionCertificateInternal ===");
  const email = clean(row["Email"]);
  const studentName = clean(row["Student Name"]);
  const admissionId = clean(row["Admission ID"]);
  
  Logger.log("Email: '" + email + "'");
  Logger.log("Student Name: '" + studentName + "'");
  Logger.log("Admission ID: '" + admissionId + "'");

  if (!email || !isValidEmail(email)) {
    throw new Error("Parent email is missing or invalid: '" + email + "'");
  }

  if (!studentName) {
    throw new Error("Student name is missing");
  }

  if (!admissionId) {
    throw new Error("Admission ID is missing");
  }

  Logger.log("Building certificate number");
  const certificateNumber = clean(row["Certificate Number"]) || buildAdmissionCertificateNumber(row);
  Logger.log("Certificate Number: '" + certificateNumber + "'");
  
  const issueDate = row["Certificate Issue Date"] || new Date();
  Logger.log("Issue Date: '" + issueDate + "'");
  
  Logger.log("Getting template file");
  let templateFile;
  try {
    templateFile = getAdmissionCertificateTemplateFile();
    Logger.log("Template file obtained: " + templateFile.getName());
  } catch (templateError) {
    throw new Error("Failed to get certificate template: " + (templateError.message || "Unknown error"));
  }
  
  Logger.log("Calling buildAdmissionCertificateFiles");
  const certificateFiles = buildAdmissionCertificateFiles(templateFile, row, certificateNumber, issueDate);
  Logger.log("Certificate files built successfully");

  if (!certificateFiles.pdfFile) {
    throw new Error("PDF file was not created successfully");
  }

  Logger.log("Sending certificate email");
  sendAdmissionCertificateEmail(row, certificateNumber, certificateFiles.pdfFile, certificateFiles.pdfUrl);
  Logger.log("Certificate email sent successfully");

  return {
    certificateNumber: certificateNumber,
    issueDate: issueDate,
    pdfUrl: certificateFiles.pdfFile.getUrl(),
    imageUrl: certificateFiles.imageFile ? certificateFiles.imageFile.getUrl() : ""
  };
  Logger.log("Certificate files built successfully");

  if (!certificateFiles.pdfFile) {
    throw new Error("PDF file was not created successfully");
  }

  Logger.log("Sending certificate email");
  sendAdmissionCertificateEmail(row, certificateNumber, certificateFiles.pdfFile, certificateFiles.pdfUrl);
  Logger.log("Certificate email sent successfully");

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
  setCellIfHeaderExists(sheet, rowNumber, headerMap, "Certificate Status", "Certificate Sent - " + result.certificateNumber);
  // Don't clear Send Certificate column - keep it as "Send"
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
  Logger.log("=== Starting getAdmissionCertificateTemplateFile ===");
  const templateId = getScriptProperty("STUDENT_CERTIFICATE_TEMPLATE_ID");
  Logger.log("Template ID from script property: '" + templateId + "'");
  
  // Try script property ID first
  if (templateId) {
    Logger.log("Attempting to get file by script property ID: " + templateId);
    try {
      const file = DriveApp.getFileById(templateId);
      Logger.log("Template file found by script property ID: " + file.getName());
      return file;
    } catch (idError) {
      Logger.log("Failed to get file by script property ID: " + idError.message);
      // Continue to fallback methods
    }
  }
  
  // Fallback to hardcoded ID
  const fallbackTemplateId = "1_lBMJqch9ptN9VCgk4jWIuizYWoGPhpO_Ek3YeFcY1c";
  Logger.log("Attempting fallback with hardcoded ID: " + fallbackTemplateId);
  try {
    const file = DriveApp.getFileById(fallbackTemplateId);
    Logger.log("Template file found by fallback ID: " + file.getName());
    return file;
  } catch (fallbackError) {
    Logger.log("Failed to get file by fallback ID: " + fallbackError.message);
    // Continue to name-based search
  }

  // Try name-based search as last resort
  const templateName = getScriptProperty("STUDENT_CERTIFICATE_TEMPLATE_NAME") || DEFAULTS.studentCertificateTemplateName;
  Logger.log("Template name to search: '" + templateName + "'");
  
  Logger.log("Searching for files by name");
  const matches = DriveApp.getFilesByName(templateName);

  if (matches.hasNext()) {
    const file = matches.next();
    Logger.log("Template file found by name: " + file.getName());
    return file;
  }

  Logger.log("Template file not found by any method");
  throw new Error("Student certificate template not found. Tried: 1) Script Property ID (" + templateId + "), 2) Fallback ID (" + fallbackTemplateId + "), 3) Name search (" + templateName + "). Please check file permissions and IDs.");
}

function buildAdmissionCertificateFiles(templateFile, row, certificateNumber, issueDate) {
  Logger.log("=== Starting buildAdmissionCertificateFiles ===");
  
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
  const archiveFolderId = getScriptProperty("STUDENT_CERTIFICATE_ARCHIVE_FOLDER_ID");
  if (archiveFolderId) {
    try {
      outputFolder = DriveApp.getFolderById(archiveFolderId);
      Logger.log("Using archive folder: " + outputFolder.getName());
    } catch (archiveError) {
      Logger.log("Error accessing archive folder: " + archiveError.message + ", using parent folder");
      outputFolder = parentFolder;
    }
  }
  
  const studentName = clean(row["Student Name"]);
  const admissionId = clean(row["Admission ID"]);
  
  Logger.log("Student Name: '" + studentName + "'");
  Logger.log("Admission ID: '" + admissionId + "'");
  
  // Use the same naming pattern as the working reference script
  // Fallback to admission ID if student name causes issues
  let copyName = "Certificate - " + studentName;
  let workingCopy;
  
  try {
    Logger.log("Attempting to create copy with student name: '" + copyName + "'");
    workingCopy = templateFile.makeCopy(copyName, parentFolder);
    Logger.log("Successfully created copy with student name");
  } catch (nameError) {
    Logger.log("Failed to create copy with student name: " + nameError.message);
    copyName = "Certificate - " + admissionId;
    Logger.log("Attempting fallback with Admission ID: '" + copyName + "'");
    try {
      workingCopy = templateFile.makeCopy(copyName, parentFolder);
      Logger.log("Successfully created copy with Admission ID");
    } catch (fallbackError) {
      Logger.log("Failed to create copy with Admission ID: " + fallbackError.message);
      throw new Error("Unable to create certificate copy. Student name error: " + nameError.message + ", Fallback error: " + fallbackError.message);
    }
  }
  
  Logger.log("Opening presentation with ID: " + workingCopy.getId());
  const presentation = SlidesApp.openById(workingCopy.getId());
  const firstSlide = presentation.getSlides()[0];
  const firstSlideObjectId = firstSlide.getObjectId();
  const issueDateText = formatDisplayDate(issueDate);
  const values = {
    "{{NAME}}": studentName,
    "{{STUDENT}}": studentName,
    "{{STUDENT_NAME}}": studentName,
    "{{STUDENT_ID}}": admissionId,
    "{{PARENT_NAME}}": clean(row["Parent Name"]),
    "{{COURSE}}": clean(row["Level"]),
    "{{LEVEL}}": clean(row["Level"]),
    "{{MODE}}": clean(row["Mode"]),
    "{{BATCH_CODE}}": clean(row["Batch Code"]),
    "{{DATE}}": issueDateText,
    "{{ISSUE_DATE}}": issueDateText,
    "{{CERTIFICATE_NO}}": certificateNumber,
    "{{CERTIFICATE_NUMBER}}": certificateNumber,
    "{{ADMISSION_ID}}": admissionId,
    "{{ACADEMY_NAME}}": getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName
  };

  Object.keys(values).forEach(function(key) {
    presentation.replaceAllText(key, String(values[key] || ""));
  });

  Logger.log("Saving and closing presentation");
  presentation.saveAndClose();
  Logger.log("Presentation saved and closed");

  // Use Drive API export method like the working reference script
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
    // Continue anyway, file is still created
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

function sendAdmissionCertificateEmail(row, certificateNumber, pdfFile, pdfUrl) {
  const academyName = getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName;
  const academyEmail = getScriptProperty("ACADEMY_EMAIL") || DEFAULTS.academyEmail;
  const academyPhone = getScriptProperty("ACADEMY_PHONE") || DEFAULTS.academyPhone;
  const senderName = getScriptProperty("SENDER_NAME") || DEFAULTS.senderName;
  const parentName = clean(row["Parent Name"]) || "Parent";
  const studentName = clean(row["Student Name"]);
  const level = clean(row["Level"]);
  const subject = "Certificate for " + studentName + " | " + academyName;
  
  const plainTextBody =
    "Dear " + parentName + ",\n\n" +
    "Congratulations to " + studentName + " on completing the Jolly Phonics Workshop! The certificate is attached.\n\n" +
    "Level: " + level + "\n" +
    "Certificate Number: " + certificateNumber + "\n\n" +
    "Warm regards,\n" + academyName + "\n" + academyPhone + "\n" + academyEmail;
  
  const htmlBody =
    '<div style="font-family: Comic Sans MS, Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #fff5e6;">' +
    '<div style="background: linear-gradient(135deg, #ff6b6b 0%, #feca57 50%, #48dbfb 100%); padding: 30px; border-radius: 20px 20px 0 0; text-align: center; border: 4px solid #ff9ff3;">' +
    '<img src="https://excelkidshub.in/images/logo.svg" alt="ExcelKidsHub Logo" style="width: 150px; height: auto; margin-bottom: 15px;">' +
    '<h1 style="margin: 0; color: #ffffff; font-size: 28px; text-shadow: 2px 2px 4px rgba(0,0,0,0.2);">Congratulations!</h1>' +
    '<p style="margin: 10px 0 0; color: #ffffff; font-size: 18px; font-weight: bold;">Jolly Phonics Workshop Certificate</p>' +
    '</div>' +
    '<div style="padding: 30px; background-color: #ffffff; border: 4px solid #ff9ff3; border-top: none; border-radius: 0 0 20px 20px;">' +
    '<p style="font-size: 18px; line-height: 1.6; color: #2d3436;">Dear <strong style="color: #ff6b6b;">' + sanitizeHtmlText(parentName) + '</strong>,</p>' +
    '<p style="font-size: 18px; line-height: 1.6; color: #2d3436;">We are delighted to share that <strong style="color: #feca57;">' + sanitizeHtmlText(studentName) + '</strong> has successfully completed the Jolly Phonics Workshop!</p>' +
    '<p style="font-size: 18px; line-height: 1.6; color: #2d3436;">Well done, <strong style="color: #48dbfb; font-weight: bold;">' + sanitizeHtmlText(studentName) + '</strong>! This is a wonderful milestone in the phonics learning journey, and we are proud to celebrate it with you.</p>' +
    '<p style="font-size: 18px; line-height: 1.6; color: #2d3436;">The certificate is attached to this email.</p>' +
    '<div style="background-color: #fff9c4; padding: 20px; border-radius: 15px; margin: 25px 0; border: 3px dashed #feca57;">' +
    '<p style="margin: 0; font-size: 16px; color: #2d3436;"><strong>Course Details:</strong></p>' +
    '<p style="margin: 10px 0 0; font-size: 15px; color: #2d3436;">Level: <span style="color: #ff6b6b; font-weight: bold;">' + sanitizeHtmlText(level) + '</span><br>Certificate Number: <span style="color: #48dbfb; font-weight: bold;">' + sanitizeHtmlText(certificateNumber) + '</span></p>' +
    '</div>' +
    '<div style="background-color: #e8f4f8; padding: 20px; border-radius: 15px; margin: 25px 0; border: 3px solid #48dbfb; text-align: center;">' +
    '<p style="margin: 0; font-size: 16px; color: #2d3436; font-weight: bold;">Share the Achievement</p>' +
    '<p style="margin: 10px 0 0; font-size: 14px; color: #2d3436;">This is a proud moment for <strong style="color: #48dbfb; font-weight: bold;">' + sanitizeHtmlText(studentName) + '</strong>. Click below to view the certificate and share the achievement with family and friends.</p>' +
    '<a href="' + sanitizeHtmlText(pdfUrl) + '" style="display: inline-block; margin-top: 15px; padding: 12px 30px; background: linear-gradient(135deg, #ff6b6b 0%, #feca57 100%); color: #ffffff; text-decoration: none; font-weight: bold; border-radius: 25px; font-size: 16px;">View & Share Certificate</a>' +
    '</div>' +
    '<p style="font-size: 18px; line-height: 1.6; color: #2d3436; text-align: center;">Keep learning, keep growing, and keep shining!</p>' +
    '<div style="text-align: center; margin: 30px 0; padding: 20px; background-color: #e8f4f8; border-radius: 15px; border: 3px solid #48dbfb;">' +
    '<p style="margin: 0 0 15px; font-size: 18px; color: #2d3436; font-weight: bold;">ExcelKidsHub Phonics Academy</p>' +
    '<p style="margin: 0 0 10px; font-size: 15px; color: #2d3436;">' + sanitizeHtmlText(academyPhone) + ' | ' + sanitizeHtmlText(academyEmail) + '</p>' +
    '<p style="margin: 0 0 15px; font-size: 14px; color: #2d3436;">Visit us: <a href="https://excelkidshub.in" style="color: #48dbfb; text-decoration: none; font-weight: bold;">excelkidshub.in</a></p>' +
    '<p style="margin: 15px 0 0; font-size: 14px; color: #2d3436;">Follow us: <a href="https://www.facebook.com/excelkidshubphonics/" style="color: #1877f2; text-decoration: none; font-weight: bold;">Facebook</a> | <a href="https://www.instagram.com/excelkidshub/" style="color: #e4405f; text-decoration: none; font-weight: bold;">Instagram</a></p>' +
    '<p style="margin: 15px 0 0; font-size: 16px; color: #ff6b6b; font-weight: bold;">Making Phonics Fun!</p>' +
    '</div>' +
    '<div style="text-align: center; font-size: 12px; color: #636e72; margin-top: 20px; padding-top: 15px; border-top: 2px solid #ffeaa7;">' +
    '<p style="margin: 0;">© 2026 ' + sanitizeHtmlText(academyName) + '. All rights reserved.</p>' +
    '</div>' +
    '</div></div>';

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

  const academyName = getScriptProperty("ACADEMY_NAME") || DEFAULTS.academyName;
  const academyPhone = getScriptProperty("ACADEMY_PHONE") || DEFAULTS.academyPhone;
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

function cleanupOldCertificateFiles() {
  Logger.log("=== Starting cleanup of old certificate files ===");
  const archiveFolderId = getScriptProperty("STUDENT_CERTIFICATE_ARCHIVE_FOLDER_ID");
  
  if (!archiveFolderId) {
    Logger.log("No archive folder ID set, skipping cleanup");
    return;
  }
  
  const archiveFolder = DriveApp.getFolderById(archiveFolderId);
  const files = archiveFolder.getFiles();
  const daysToKeep = 30; // Delete files older than 30 days
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
  
  let deletedCount = 0;
  
  while (files.hasNext()) {
    const file = files.next();
    const createdDate = file.getDateCreated();
    
    if (createdDate < cutoffDate) {
      Logger.log("Deleting old file: " + file.getName() + " (created: " + createdDate + ")");
      file.setTrashed(true);
      deletedCount++;
    }
  }
  
  Logger.log("Cleanup complete. Deleted " + deletedCount + " old certificate files.");
}

function setupCleanupTrigger() {
  // Delete any existing cleanup triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "cleanupOldCertificateFiles") {
      ScriptApp.deleteTrigger(trigger);
      Logger.log("Deleted existing cleanup trigger");
    }
  });
  
  // Create new trigger to run daily at 2 AM
  ScriptApp.newTrigger("cleanupOldCertificateFiles")
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  
  Logger.log("Created new cleanup trigger to run daily at 2 AM");
}
