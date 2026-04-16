function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, message: "No data received" });
    }

    const payload = JSON.parse(e.postData.contents);
    const action = clean(payload.action);

    if (!action || action === "admission") {
      return admission(payload);
    }

    if (action === "adminLogin") {
      return adminLogin(payload);
    }

    if (action === "getDashboard") {
      return getDashboard(payload);
    }

    if (action === "getAdmissions") {
      return getAdmissions(payload);
    }

    if (action === "updateStudent") {
      return updateStudent(payload);
    }

    if (action === "getBatches") {
      return getBatches(payload);
    }

    if (action === "saveBatch") {
      return saveBatch(payload);
    }

    if (action === "updateBatch") {
      return updateBatch(payload);
    }

    if (action === "assignStudentToBatch") {
      return assignStudentToBatch(payload);
    }

    if (action === "savePayment") {
      return savePayment(payload);
    }

    if (action === "getPayments") {
      return getPayments(payload);
    }

    if (action === "sendPaymentEmail") {
      return sendPaymentEmail(payload);
    }

    if (action === "saveExpense") {
      return saveExpense(payload);
    }

    if (action === "getExpenses") {
      return getExpenses(payload);
    }

    return jsonResponse({ success: false, message: "Invalid action" });
  } catch (error) {
    return jsonResponse({
      success: false,
      message: error && error.message ? error.message : "Something went wrong"
    });
  }
}

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    // Log every edit attempt for debugging
    Logger.log("onEdit triggered - Sheet: " + sheet.getName() + ", Column: " + range.getColumn() + ", Value: " + range.getValue());
    
    // Only process changes in "admissions" sheet
    if (sheet.getName() !== "admissions") {
      Logger.log("Not admissions sheet, skipping");
      return;
    }

    // Get column headers
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let sendReceiptColIndex = -1;
    let receiptStatusColIndex = -1;
    let admissionIdColIndex = -1;

    headerRow.forEach(function(header, index) {
      const cleanHeader = clean(header);
      if (cleanHeader === "Send Receipt") sendReceiptColIndex = index + 1;
      if (cleanHeader === "Receipt Status") receiptStatusColIndex = index + 1;
      if (cleanHeader === "Admission ID") admissionIdColIndex = index + 1;
    });

    Logger.log("Found columns - Send Receipt: " + sendReceiptColIndex + ", Receipt Status: " + receiptStatusColIndex + ", Admission ID: " + admissionIdColIndex);

    // Check if this edit is in the "Send Receipt" column
    if (range.getColumn() === sendReceiptColIndex) {
      const cellValue = clean(range.getValue());
      Logger.log("Edit in Send Receipt column. Value: '" + cellValue + "'");
      
      if (cellValue === "Send") {
        const rowNumber = range.getRow();
        const admissionId = clean(sheet.getRange(rowNumber, admissionIdColIndex).getValue());

        Logger.log("Send Receipt = 'Send' - Admission: " + admissionId + ", Row: " + rowNumber);

        if (!admissionId) {
          Logger.log("ERROR: Admission ID is empty");
          if (receiptStatusColIndex > 0) {
            sheet.getRange(rowNumber, receiptStatusColIndex).setValue("Email Failed - No Admission ID");
          }
          return;
        }

        if (receiptStatusColIndex <= 0) {
          Logger.log("ERROR: Receipt Status column not found");
          return;
        }

        // Mark as pending so the batch process can handle it
        Logger.log("Marking row " + rowNumber + " as pending for email processing");
        sheet.getRange(rowNumber, receiptStatusColIndex).setValue("Pending");
      }
    }
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("onEdit error: " + errorMsg);
  }
}

function processPendingEmailsFromMaster() {
  try {
    const sheet = getSheet(SHEET_NAMES.admissions);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let receiptStatusColIndex = -1;
    let admissionIdColIndex = -1;
    let emailColIndex = -1;
    let paymentStatusColIndex = -1;

    headerRow.forEach(function(header, index) {
      const cleanHeader = clean(header);
      if (cleanHeader === "Receipt Status") receiptStatusColIndex = index + 1;
      if (cleanHeader === "Admission ID") admissionIdColIndex = index + 1;
      if (cleanHeader === "Email") emailColIndex = index + 1;
      if (cleanHeader === "Payment Status") paymentStatusColIndex = index + 1;
    });

    if (receiptStatusColIndex <= 0) {
      Logger.log("Receipt Status column not found");
      return;
    }

    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const receiptStatus = clean(data[i][receiptStatusColIndex - 1]);
      const admissionId = clean(data[i][admissionIdColIndex - 1]);
      const email = clean(data[i][emailColIndex - 1]);
      const paymentStatus = paymentStatusColIndex > 0 ? clean(data[i][paymentStatusColIndex - 1]) : "";

      if (receiptStatus === "Pending" && admissionId && email) {
        const rowNumber = i + 1;
        Logger.log("Processing pending email for admission " + admissionId);

        try {
          // Check if payment is completed
          if (paymentStatus !== "Completed") {
            Logger.log("Payment not completed for admission " + admissionId + ". Status: " + paymentStatus);
            sheet.getRange(rowNumber, receiptStatusColIndex).setValue("Not Sent - Payment not completed");
            continue;
          }

          if (!isValidEmail(email)) {
            Logger.log("ERROR: Invalid email " + email);
            sheet.getRange(rowNumber, receiptStatusColIndex).setValue("Email Failed - Invalid Email");
            continue;
          }

          // Send payment email with template (same as UI)
          Logger.log("Attempting to send receipt email for admission " + admissionId);
          sendPaymentEmailInternal({
            admissionId: admissionId,
            paymentId: null,
            emailType: "receipt"
          });

          Logger.log("Email sent successfully for admission " + admissionId);
          sheet.getRange(rowNumber, receiptStatusColIndex).setValue("Email Sent");
        } catch (emailError) {
          const errorMsg = emailError && emailError.message ? emailError.message : "Unknown error";
          Logger.log("ERROR sending email for admission " + admissionId + ": " + errorMsg);
          // Store first 60 chars of error for display
          const shortError = errorMsg.substring(0, 60);
          sheet.getRange(rowNumber, receiptStatusColIndex).setValue("Email Failed - " + shortError);
        }
      }
    }
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("processPendingEmailsFromMaster error: " + errorMsg);
  }
}

function testReceiptTemplate() {
  try {
    Logger.log("Testing Receipt Template...");
    
    const templateId = getScriptProperty(SCRIPT_PROPERTY_KEYS.receiptTemplateId) || DEFAULTS.receiptTemplateId;
    Logger.log("Template ID from configuration: " + templateId);
    
    const templateFile = getReceiptTemplateFile();
    Logger.log("Template file found: " + templateFile.getName());
    Logger.log("Template file size: " + templateFile.getSize() + " bytes");
    Logger.log("Template file owner: " + templateFile.getOwner().getEmail());
    Logger.log("Template file access: OK");
    Logger.log("✅ Receipt template is accessible!");
    
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("❌ ERROR: " + errorMsg);
    Logger.log("This means the template file cannot be accessed. Please:");
    Logger.log("1. Check if the template file exists in Google Drive");
    Logger.log("2. Check if you have permission to access it");
    Logger.log("3. Check the template ID in Script Properties");
  }
}

function processAdmissionQueue() {
  try {
    Logger.log("=== Processing Admission Queue ===");
    
    const queueSheet = getSheet("admission_queue");
    if (!queueSheet) {
      Logger.log("Queue sheet not found - nothing to process");
      return;
    }
    
    const data = queueSheet.getDataRange().getValues();
    const admissionsSheet = getSheet(SHEET_NAMES.admissions);
    let processedCount = 0;
    let rowsToDelete = [];
    
    // Start from row 2 (skip header)
    for (let i = 1; i < data.length; i++) {
      const timestamp = data[i][0];
      const payloadJson = data[i][1];
      const status = clean(data[i][2]);
      
      if (status && status !== "Pending") {
        continue; // Skip already processed rows
      }
      
      try {
        const payload = JSON.parse(payloadJson);
        Logger.log("Processing queued admission: " + payload.studentName + " - " + payload.mobile);
        
        // Process the admission
        processQueuedAdmission(payload, admissionsSheet);
        
        // Mark as processed
        queueSheet.getRange(i + 1, 3).setValue("Processed");
        processedCount++;
        rowsToDelete.push(i + 1);
        
      } catch (error) {
        const errorMsg = error && error.message ? error.message : "Unknown error";
        Logger.log("ERROR processing queue row " + (i+1) + ": " + errorMsg);
        queueSheet.getRange(i + 1, 3).setValue("Error: " + errorMsg.substring(0, 30));
      }
    }
    
    // Clean up processed rows (delete from bottom to top to avoid index shift)
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      queueSheet.deleteRow(rowsToDelete[i]);
    }
    
    Logger.log("=== Completed: processed " + processedCount + " admissions from queue ===");
    
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("processAdmissionQueue ERROR: " + errorMsg);
  }
}

function processQueuedAdmission(payload, admissionsSheet) {
  var parentName = clean(payload.parentName);
  var mobile = clean(payload.mobile);
  var email = clean(payload.email);
  var address = clean(payload.address);
  var city = clean(payload.city);
  var studentName = clean(payload.studentName);
  var age = clean(payload.age);
  var gender = clean(payload.gender);
  var school = clean(payload.school);
  var grade = clean(payload.grade);
  var level = clean(payload.level);
  var mode = clean(payload.mode);
  var admissionSource = clean(payload.admissionSource) || "Website";

  // Validate required fields
  if (!parentName || !mobile || !studentName || !level || !mode) {
    throw new Error("Required fields missing");
  }

  if (email && !isValidEmail(email)) {
    throw new Error("Invalid email address");
  }

  // Check for duplicate (Student Name + Mobile + Level)
  const duplicateCheck = findQueuedDuplicateAdmission(admissionsSheet, studentName, mobile, level);
  var createdDate = formatDateTimeValue(new Date());
  var totalFee = getFeeByLevel(level);
  var isUpdate = false;

  if (duplicateCheck.found) {
    // UPDATE existing record
    const admissionId = duplicateCheck.admissionId;
    const rowNumber = duplicateCheck.rowNumber;
    isUpdate = true;
    
    const headerRow = admissionsSheet.getRange(1, 1, 1, admissionsSheet.getLastColumn()).getValues()[0];
    const headerMap = {};
    for (let i = 0; i < headerRow.length; i++) {
      headerMap[clean(headerRow[i])] = i + 1;
    }
    
    // Update fields
    if (headerMap["Parent Name"]) admissionsSheet.getRange(rowNumber, headerMap["Parent Name"]).setValue(parentName);
    if (headerMap["Mobile"]) admissionsSheet.getRange(rowNumber, headerMap["Mobile"]).setValue(mobile);
    if (headerMap["Email"]) admissionsSheet.getRange(rowNumber, headerMap["Email"]).setValue(email);
    if (headerMap["Address"]) admissionsSheet.getRange(rowNumber, headerMap["Address"]).setValue(address);
    if (headerMap["City"]) admissionsSheet.getRange(rowNumber, headerMap["City"]).setValue(city);
    if (headerMap["Student Name"]) admissionsSheet.getRange(rowNumber, headerMap["Student Name"]).setValue(studentName);
    if (headerMap["Age"]) admissionsSheet.getRange(rowNumber, headerMap["Age"]).setValue(age);
    if (headerMap["Gender"]) admissionsSheet.getRange(rowNumber, headerMap["Gender"]).setValue(gender);
    if (headerMap["School"]) admissionsSheet.getRange(rowNumber, headerMap["School"]).setValue(school);
    if (headerMap["Grade"]) admissionsSheet.getRange(rowNumber, headerMap["Grade"]).setValue(grade);
    if (headerMap["Level"]) admissionsSheet.getRange(rowNumber, headerMap["Level"]).setValue(level);
    if (headerMap["Mode"]) admissionsSheet.getRange(rowNumber, headerMap["Mode"]).setValue(mode);
    if (headerMap["Total Fee"]) admissionsSheet.getRange(rowNumber, headerMap["Total Fee"]).setValue(totalFee);
    if (headerMap["Notification Status"]) admissionsSheet.getRange(rowNumber, headerMap["Notification Status"]).setValue("Pending-Update");
    
    Logger.log("Updated existing admission: " + admissionId);
    
  } else {
    // CREATE new record
    const admissionId = getNextAdmissionId(admissionsSheet);

    var row = [
      admissionId,
      parentName,
      mobile,
      email,
      address,
      city,
      studentName,
      age,
      gender,
      school,
      grade,
      level,
      "",
      mode,
      "",
      "",
      "Pending Start",
      totalFee,
      0,
      0,
      "",
      "",
      "",
      "",
      admissionSource,
      "",
      "",
      createdDate,
      "", // Certificate Status
      "", // Certificate Number
      "", // Certificate Issue Date
      "", // Certificate Sent Date
      "", // Send Receipt
      "", // Receipt Status
      "Pending-New" // Notification Status
    ];

    // Find the first empty row
    var data = admissionsSheet.getDataRange().getValues();
    var insertRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || clean(data[i][0]) === "") {
        insertRow = i + 1;
        break;
      }
    }

    if (insertRow === -1) {
      admissionsSheet.appendRow(row);
    } else {
      admissionsSheet.getRange(insertRow, 1, 1, row.length).setValues([row]);
    }
    
    Logger.log("Created new admission: " + admissionId);
  }
}

function findQueuedDuplicateAdmission(sheet, studentName, mobile, level) {
  const data = sheet.getDataRange().getValues();
  const headerRow = data[0];
  
  let studentNameColIndex = -1;
  let mobileColIndex = -1;
  let levelColIndex = -1;
  let admissionIdColIndex = -1;
  
  for (let i = 0; i < headerRow.length; i++) {
    const cleanHeader = clean(headerRow[i]);
    if (cleanHeader === "Student Name") studentNameColIndex = i;
    if (cleanHeader === "Mobile") mobileColIndex = i;
    if (cleanHeader === "Level") levelColIndex = i;
    if (cleanHeader === "Admission ID") admissionIdColIndex = i;
  }
  
  // Search for matching record
  for (let i = 1; i < data.length; i++) {
    const existingStudentName = clean(data[i][studentNameColIndex]);
    const existingMobile = clean(data[i][mobileColIndex]);
    const existingLevel = clean(data[i][levelColIndex]);
    const existingAdmissionId = clean(data[i][admissionIdColIndex]);
    
    if (existingAdmissionId && 
        existingStudentName === clean(studentName) && 
        existingMobile === clean(mobile) && 
        existingLevel === clean(level)) {
      return {
        found: true,
        rowNumber: i + 1,
        admissionId: existingAdmissionId,
        rowIndex: i
      };
    }
  }
  
  return { found: false };
}

function testAdmissionNotificationSetup() {
  try {
    Logger.log("=== Testing Admission Notification Setup ===");
    
    const sheet = getSheet(SHEET_NAMES.admissions);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    Logger.log("Total columns: " + headerRow.length);
    
    let notificationStatusColIndex = -1;
    let admissionIdColIndex = -1;
    let studentNameColIndex = -1;
    
    for (let i = 0; i < headerRow.length; i++) {
      const cleanHeader = clean(headerRow[i]);
      
      if (cleanHeader === "Notification Status") {
        notificationStatusColIndex = i + 1;
        Logger.log("✅ Found Notification Status column at index " + notificationStatusColIndex);
      }
      if (cleanHeader === "Admission ID") {
        admissionIdColIndex = i + 1;
        Logger.log("✅ Found Admission ID column at index " + admissionIdColIndex);
      }
      if (cleanHeader === "Student Name") {
        studentNameColIndex = i + 1;
        Logger.log("✅ Found Student Name column at index " + studentNameColIndex);
      }
    }
    
    if (notificationStatusColIndex <= 0) {
      Logger.log("❌ ERROR: Notification Status column NOT found!");
      Logger.log("Headers found: " + headerRow.join(" | "));
      return;
    }
    
    // Check data rows
    const data = sheet.getDataRange().getValues();
    Logger.log("\nScanning first 30 data rows for 'Pending' status:");
    
    let pendingRows = [];
    for (let i = 1; i < data.length && i < 31; i++) {
      const notifStatus = clean(data[i][notificationStatusColIndex - 1]);
      const admissionId = clean(data[i][admissionIdColIndex - 1]);
      const studentName = clean(data[i][studentNameColIndex - 1]);
      
      if (notifStatus === "Pending") {
        Logger.log("Row " + (i+1) + ": " + admissionId + " - " + studentName + " [Status: Pending]");
        pendingRows.push({
          rowNumber: i + 1,
          admissionId: admissionId,
          studentName: studentName
        });
      }
    }
    
    if (pendingRows.length === 0) {
      Logger.log("❌ No rows with Pending status found");
    } else {
      Logger.log("✅ Found " + pendingRows.length + " rows with Pending status");
    }
    
    Logger.log("=== Test Complete ===");
    
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("❌ Test error: " + errorMsg);
  }
}

function processNewAdmissionNotifications() {
  try {
    Logger.log("=== Starting processNewAdmissionNotifications ===");
    const sheet = getSheet(SHEET_NAMES.admissions);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let admissionIdColIndex = -1;
    let studentNameColIndex = -1;
    let parentNameColIndex = -1;
    let emailColIndex = -1;
    let mobileColIndex = -1;
    let notificationStatusColIndex = -1;
    let createdDateColIndex = -1;

    headerRow.forEach(function(header, index) {
      const cleanHeader = clean(header);
      if (cleanHeader === "Admission ID") admissionIdColIndex = index + 1;
      if (cleanHeader === "Student Name") studentNameColIndex = index + 1;
      if (cleanHeader === "Parent Name") parentNameColIndex = index + 1;
      if (cleanHeader === "Email") emailColIndex = index + 1;
      if (cleanHeader === "Mobile") mobileColIndex = index + 1;
      if (cleanHeader === "Notification Status") notificationStatusColIndex = index + 1;
      if (cleanHeader === "Created Date") createdDateColIndex = index + 1;
    });

    Logger.log("Columns - Admission ID: " + admissionIdColIndex + ", Notification Status: " + notificationStatusColIndex);

    if (notificationStatusColIndex <= 0) {
      Logger.log("ERROR: Notification Status column not found");
      return;
    }

    const data = sheet.getDataRange().getValues();
    Logger.log("Total rows in sheet: " + data.length);
    let processedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const admissionId = clean(data[i][admissionIdColIndex - 1]);
      const studentName = clean(data[i][studentNameColIndex - 1]);
      const parentName = clean(data[i][parentNameColIndex - 1]);
      const email = clean(data[i][emailColIndex - 1]);
      const mobile = mobileColIndex > 0 ? clean(data[i][mobileColIndex - 1]) : "";
      const notificationStatus = clean(data[i][notificationStatusColIndex - 1]);

      // Process any row marked as Pending, Pending-New, or Pending-Update
      if (admissionId && (notificationStatus === "Pending" || notificationStatus === "Pending-New" || notificationStatus === "Pending-Update")) {
        const rowNumber = i + 1;
        const isUpdate = notificationStatus === "Pending-Update";
        
        Logger.log("Row " + rowNumber + ": Found " + (isUpdate ? "Update" : "New") + " notification for " + admissionId);
        
        try {
          if (!studentName) {
            Logger.log("Skipping - Student Name is empty");
            sheet.getRange(rowNumber, notificationStatusColIndex).setValue("Failed - No Student Name");
            continue;
          }

          if (!email) {
            Logger.log("Skipping - Email is empty");
            sheet.getRange(rowNumber, notificationStatusColIndex).setValue("Failed - No Email");
            continue;
          }

          Logger.log("Sending " + (isUpdate ? "update" : "new") + " notification for admission " + admissionId);
          
          // Build complete student record
          const studentRecord = buildAdmissionRecordForEmail(data[i], headerRow);
          
          // Send notification email
          sendAdmissionNotificationEmail(admissionId, studentName, parentName, email, mobile, studentRecord, isUpdate);
          
          Logger.log("Email sent - updating status to Sent");
          sheet.getRange(rowNumber, notificationStatusColIndex).setValue("Sent");
          processedCount++;
          
        } catch (emailError) {
          const errorMsg = emailError && emailError.message ? emailError.message : "Unknown error";
          Logger.log("ERROR: " + errorMsg);
          sheet.getRange(rowNumber, notificationStatusColIndex).setValue("Failed - " + errorMsg.substring(0, 40));
        }
      }
    }
    
    Logger.log("=== Completed: processed " + processedCount + " notifications ===");
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("processNewAdmissionNotifications ERROR: " + errorMsg);
  }
}

function buildAdmissionRecordForEmail(rowData, headers) {
  let record = "";
  for (let i = 0; i < headers.length; i++) {
    const header = clean(headers[i]);
    const value = clean(rowData[i]);
    if (header && value) {
      record += header + ": " + value + "\n";
    }
  }
  return record;
}

function sendAdmissionNotificationEmail(admissionId, studentName, parentName, parentEmail, mobile, studentRecord, isUpdate) {
  try {
    const adminEmail = getScriptProperty(SCRIPT_PROPERTY_KEYS.adminEmail) || "excelkidshub.edu@gmail.com";
    const academyName = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName;
    const academyPhone = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyPhone) || DEFAULTS.academyPhone;
    
    Logger.log("Sending " + (isUpdate ? "update" : "new registration") + " notification email from " + Session.getActiveUser().getEmail() + " to " + adminEmail);
    
    const typeLabel = isUpdate ? "UPDATED" : "NEW";
    const subject = typeLabel + " Registration: " + studentName + " (ID: " + admissionId + ", Mobile: " + mobile + ")";
    
    const messageIntro = isUpdate 
      ? "Student Registration UPDATED"
      : "New Student Registration Received";
    
    const plainTextBody =
      messageIntro + "\n\n" +
      "Student Name: " + studentName + "\n" +
      "Admission ID: " + admissionId + "\n" +
      "Parent Name: " + parentName + "\n" +
      "Parent Mobile: " + mobile + "\n" +
      "Parent Email: " + parentEmail + "\n\n" +
      (isUpdate ? "This is an UPDATE to the existing record.\n\n" : "") +
      "Complete Student Record:\n" +
      "====================================\n" +
      studentRecord + "\n" +
      "====================================\n\n" +
      "Updated at: " + formatDisplayDate(new Date()) + "\n" +
      "Academy: " + academyName + "\n" +
      "Contact: " + academyPhone;
    
    const headerColor = isUpdate ? "#ff6b6b" : "#1a5490";
    const htmlBody =
      "<div style='font-family:Arial,sans-serif; color:#333;'>" +
      "<h2 style='color:" + headerColor + ";'>" + messageIntro + "</h2>" +
      "<p style='background:" + (isUpdate ? "#ffe0e0" : "#e8f1f8") + "; padding:10px; border-left:4px solid " + headerColor + ";'>" +
      "<b>Status:</b> " + (isUpdate ? "UPDATED (ID: " + admissionId + ")" : "NEW (ID: " + admissionId + ")") +
      "</p>" +
      "<p><b>Student Name:</b> " + sanitizeHtmlText(studentName) + "</p>" +
      "<p><b>Parent Name:</b> " + sanitizeHtmlText(parentName) + "</p>" +
      "<p><b>Parent Mobile:</b> " + sanitizeHtmlText(mobile) + "</p>" +
      "<p><b>Parent Email:</b> " + sanitizeHtmlText(parentEmail) + "</p>" +
      "<hr>" +
      "<h3>Complete Student Record:</h3>" +
      "<pre style='background:#f5f5f5; padding:10px; border-radius:3px; overflow-x:auto;'>" + sanitizeHtmlText(studentRecord) + "</pre>" +
      "<hr>" +
      "<p style='color:#666; font-size:0.9em;'>" +
      "Updated at: " + formatDisplayDate(new Date()) + "<br>" +
      "Academy: " + sanitizeHtmlText(academyName) + "<br>" +
      "Contact: " + sanitizeHtmlText(academyPhone) +
      "</p>" +
      "</div>";
    
    GmailApp.sendEmail(adminEmail, subject, plainTextBody, {
      htmlBody: htmlBody
    });
    
    Logger.log("Email sent successfully to " + adminEmail);
  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    Logger.log("ERROR in sendAdmissionNotificationEmail: " + errorMsg);
    throw new Error("Failed to send email: " + errorMsg);
  }
}

