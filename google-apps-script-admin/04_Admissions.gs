function getFeeByLevel(level) {
  const normalizedLevel = clean(level).toLowerCase();
  
  if (normalizedLevel === "basic") {
    return 4500;
  } else if (normalizedLevel === "advanced" || normalizedLevel === "proficient") {
    return 6500;
  }
  
  return 0;
}

function admission(payload) {
  // ⚡ FAST RESPONSE - Queue for background processing
  try {
    // Basic validation only - quick checks
    var parentName = clean(payload.parentName);
    var mobile = clean(payload.mobile);
    var studentName = clean(payload.studentName);
    var level = clean(payload.level);
    var mode = clean(payload.mode);

    if (!parentName || !mobile || !studentName || !level || !mode) {
      return jsonResponse({ success: false, message: "Required fields missing" });
    }

    // Get or create the queue sheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var queueSheet = null;
    
    try {
      queueSheet = spreadsheet.getSheetByName("admission_queue");
    } catch (e) {
      queueSheet = null;
    }
    
    if (!queueSheet) {
      Logger.log("Creating admission_queue sheet...");
      queueSheet = spreadsheet.insertSheet("admission_queue");
      queueSheet.appendRow(["Timestamp", "Payload", "Status"]);
      Logger.log("✅ admission_queue sheet created");
    }

    // Store the entire payload as JSON for background processing
    var timestamp = new Date();
    var payloadJson = JSON.stringify(payload);
    
    queueSheet.appendRow([
      formatDateTimeValue(timestamp),
      payloadJson,
      "Pending"
    ]);

    Logger.log("✅ Admission queued for processing: " + studentName + " - " + mobile + " - " + level);

    // ✅ Return immediately to user - NO WAITING
    return jsonResponse({
      success: true,
      message: "Registration received! Processing in background...",
      processingInBackground: true
    });

  } catch (error) {
    const errorMsg = error && error.message ? error.message : "Unknown error";
    const errorStack = error && error.stack ? error.stack : "No stack trace";
    Logger.log("❌ Admission queue error: " + errorMsg);
    Logger.log("Stack: " + errorStack);
    
    return jsonResponse({
      success: false,
      message: "Error: " + errorMsg
    });
  }
}


function getAdmissions(payload) {
  authorizeAdmin(payload);

  const sheet = getSheet(SHEET_NAMES.admissions);
  const admissions = getSheetObjects(sheet).map(function(item) {
    return {
      admissionId: clean(item["Admission ID"]),
      parentName: clean(item["Parent Name"]),
      mobile: clean(item["Mobile"]),
      email: clean(item["Email"]),
      address: clean(item["Address"]),
      city: clean(item["City"]),
      studentName: clean(item["Student Name"]),
      age: clean(item["Age"]),
      gender: clean(item["Gender"]),
      school: clean(item["School"]),
      grade: clean(item["Grade"]),
      level: clean(item["Level"]),
      batchCode: clean(item["Batch Code"]),
      mode: clean(item["Mode"]),
      startDate: formatDateValue(item["Start Date"]),
      endDate: formatDateValue(item["End Date"]),
      status: clean(item["Status"]),
      totalFee: toNumber(item["Total Fee"], 0),
      discount: toNumber(item["Discount"], 0),
      manualAdjustment: toNumber(item["Manual Adjustment"], 0),
      adjustedFee: toNumber(item["Adjusted Fee"], 0),
      totalPaid: toNumber(item["Total Paid"], 0),
      pending: toNumber(item["Pending"], 0),
      paymentStatus: clean(item["Payment Status"]),
      admissionSource: clean(item["Admission Source"]),
      referralType: clean(item["Referral Type"]),
      referrerName: clean(item["Referrer Name"]),
      createdDate: item["Created Date"] ? new Date(item["Created Date"]).toISOString() : "",
      notes: clean(item["Notes"])
    };
  });

  sortByDateDesc(admissions, "createdDate");

  return jsonResponse({
    success: true,
    data: admissions
  });
}

function getNextAdmissionId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return "A001";
  }

  var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  var maxNumber = 0;

  idValues.forEach(function(id) {
    var text = String(id || "").trim();
    var match = text.match(/^A(\d+)$/i);
    if (match) {
      var num = parseInt(match[1], 10);
      if (num > maxNumber) {
        maxNumber = num;
      }
    }
  });

  return "A" + String(maxNumber + 1).padStart(3, "0");
}

function updateStudent(payload) {
  var sheet = getSheet(SHEET_NAMES.admissions);
  var admissionId = clean(payload.admissionId);

  if (!admissionId) {
    return jsonResponse({ success: false, message: "Admission ID is required" });
  }

  // Find the row with this admission ID
  var data = sheet.getDataRange().getValues();
  var headerRow = data[0];
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == admissionId) {
      rowIndex = i + 1; // Sheets use 1-based indexing
      break;
    }
  }

  if (rowIndex === -1) {
    return jsonResponse({ success: false, message: "Student not found" });
  }

  // Update fields
  var updates = {};
  
  if (payload.parentName !== undefined) updates["Parent Name"] = clean(payload.parentName);
  if (payload.mobile !== undefined) updates["Mobile"] = clean(payload.mobile);
  if (payload.email !== undefined) updates["Email"] = clean(payload.email);
  if (payload.address !== undefined) updates["Address"] = clean(payload.address);
  if (payload.city !== undefined) updates["City"] = clean(payload.city);
  if (payload.studentName !== undefined) updates["Student Name"] = clean(payload.studentName);
  if (payload.age !== undefined) updates["Age"] = clean(payload.age);
  if (payload.gender !== undefined) updates["Gender"] = clean(payload.gender);
  if (payload.school !== undefined) updates["School"] = clean(payload.school);
  if (payload.grade !== undefined) updates["Grade"] = clean(payload.grade);
  if (payload.level !== undefined) updates["Level"] = clean(payload.level);
  if (payload.mode !== undefined) updates["Mode"] = clean(payload.mode);
  if (payload.batchCode !== undefined) updates["Batch Code"] = clean(payload.batchCode);
  if (payload.startDate !== undefined) updates["Start Date"] = clean(payload.startDate);
  if (payload.endDate !== undefined) updates["End Date"] = clean(payload.endDate);
  if (payload.status !== undefined) updates["Status"] = clean(payload.status);
  if (payload.totalFee !== undefined) updates["Total Fee"] = Number(payload.totalFee) || 0;
  if (payload.discount !== undefined) updates["Discount"] = Number(payload.discount) || 0;
  if (payload.manualAdjustment !== undefined) updates["Manual Adjustment"] = Number(payload.manualAdjustment) || 0;
  if (payload.totalPaid !== undefined) updates["Total Paid"] = Number(payload.totalPaid) || 0;
  if (payload.notes !== undefined) updates["Notes"] = clean(payload.notes);

  // Apply updates
  for (var col = 0; col < headerRow.length; col++) {
    var header = headerRow[col];
    if (updates[header] !== undefined) {
      sheet.getRange(rowIndex, col + 1).setValue(updates[header]);
    }
  }

  return jsonResponse({
    success: true,
    message: "Student updated successfully"
  });
}
