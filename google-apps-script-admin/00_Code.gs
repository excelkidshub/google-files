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

      // Process any row marked as Pending (regardless of date - for testing)
      if (admissionId && notificationStatus === "Pending") {
        const rowNumber = i + 1;
        
        Logger.log("Row " + rowNumber + ": Found Pending notification for " + admissionId);
        
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

          Logger.log("Sending notification for admission " + admissionId);
          
          // Build complete student record
          const studentRecord = buildAdmissionRecordForEmail(data[i], headerRow);
          
          // Send notification email
          sendAdmissionNotificationEmail(admissionId, studentName, parentName, email, mobile, studentRecord);
          
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

function sendAdmissionNotificationEmail(admissionId, studentName, parentName, parentEmail, mobile, studentRecord) {
  try {
    const adminEmail = getScriptProperty(SCRIPT_PROPERTY_KEYS.adminEmail) || "excelkidshub.edu@gmail.com";
    const academyName = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyName) || DEFAULTS.academyName;
    const academyPhone = getScriptProperty(SCRIPT_PROPERTY_KEYS.academyPhone) || DEFAULTS.academyPhone;
    
    Logger.log("Sending notification email from " + Session.getActiveUser().getEmail() + " to " + adminEmail);
    
    const subject = "New Registration: " + studentName + " (Parent: " + parentName + ", Mobile: " + mobile + ")";
    
    const plainTextBody =
      "New Student Registration Received\n\n" +
      "Student Name: " + studentName + "\n" +
      "Parent Name: " + parentName + "\n" +
      "Parent Mobile: " + mobile + "\n" +
      "Parent Email: " + parentEmail + "\n" +
      "Admission ID: " + admissionId + "\n\n" +
      "Complete Student Record:\n" +
      "====================================\n" +
      studentRecord + "\n" +
      "====================================\n\n" +
      "Registered at: " + formatDisplayDate(new Date()) + "\n" +
      "Academy: " + academyName + "\n" +
      "Contact: " + academyPhone;
    
    const htmlBody =
      "<div style='font-family:Arial,sans-serif; color:#333;'>" +
      "<h2 style='color:#1a5490;'>New Student Registration Received</h2>" +
      "<p><b>Student Name:</b> " + sanitizeHtmlText(studentName) + "</p>" +
      "<p><b>Parent Name:</b> " + sanitizeHtmlText(parentName) + "</p>" +
      "<p><b>Parent Mobile:</b> " + sanitizeHtmlText(mobile) + "</p>" +
      "<p><b>Parent Email:</b> " + sanitizeHtmlText(parentEmail) + "</p>" +
      "<p><b>Admission ID:</b> " + sanitizeHtmlText(admissionId) + "</p>" +
      "<hr>" +
      "<h3>Complete Student Record:</h3>" +
      "<pre style='background:#f5f5f5; padding:10px; border-radius:3px; overflow-x:auto;'>" + sanitizeHtmlText(studentRecord) + "</pre>" +
      "<hr>" +
      "<p style='color:#666; font-size:0.9em;'>" +
      "Registered at: " + formatDisplayDate(new Date()) + "<br>" +
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

