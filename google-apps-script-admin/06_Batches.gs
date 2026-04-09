function getBatches(payload) {
  authorizeAdmin(payload);

  const sheet = getSheet(SHEET_NAMES.batches);
  const batches = getSheetObjects(sheet).map(function(item) {
    return {
      batchCode: clean(item["Batch Code"]),
      batchName: clean(item["Batch Name"]),
      level: clean(item["Level"]),
      mode: clean(item["Mode"]),
      startDate: formatDateValue(item["Start Date"]),
      endDate: formatDateValue(item["End Date"]),
      timing: clean(item["Timing"]),
      days: clean(item["Days"]),
      capacity: toNumber(item["Capacity"], 0),
      location: clean(item["Location"]),
      status: clean(item["Status"]),
      notes: clean(item["Notes"])
    };
  });

  return jsonResponse({ success: true, data: batches });
}

function saveBatch(payload) {
  authorizeAdmin(payload);

  const sheet = getSheet(SHEET_NAMES.batches);
  const batchCode = clean(payload.batchCode);

  if (!batchCode) {
    return jsonResponse({ success: false, message: "Batch Code is required" });
  }

  const existing = getSheetObjects(sheet).some(function(item) {
    return clean(item["Batch Code"]) === batchCode;
  });

  if (existing) {
    return jsonResponse({ success: false, message: "Batch Code already exists" });
  }

  appendObjectRow(sheet, {
    "Batch Code": batchCode,
    "Batch Name": clean(payload.batchName),
    "Level": clean(payload.level),
    "Mode": clean(payload.mode),
    "Start Date": clean(payload.startDate),
    "End Date": clean(payload.endDate),
    "Timing": clean(payload.timing),
    "Days": clean(payload.days),
    "Capacity": toNumber(payload.capacity, 0),
    "Location": clean(payload.location),
    "Status": clean(payload.status) || DEFAULTS.batchStatus,
    "Notes": clean(payload.notes)
  });

  return jsonResponse({ success: true, message: "Batch saved successfully" });
}

function updateBatch(payload) {
  authorizeAdmin(payload);

  const sheet = getSheet(SHEET_NAMES.batches);
  const batchCode = clean(payload.batchCode);
  const headerMap = getHeaderMap(sheet);
  const rows = getSheetObjects(sheet);
  const row = rows.find(function(item) {
    return clean(item["Batch Code"]) === batchCode;
  });

  if (!row) {
    return jsonResponse({ success: false, message: "Batch not found" });
  }

  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Batch Name", clean(payload.batchName));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Level", clean(payload.level));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Mode", clean(payload.mode));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Start Date", clean(payload.startDate));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "End Date", clean(payload.endDate));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Timing", clean(payload.timing));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Days", clean(payload.days));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Capacity", toNumber(payload.capacity, 0));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Location", clean(payload.location));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Status", clean(payload.status));
  setCellIfHeaderExists(sheet, row._rowNumber, headerMap, "Notes", clean(payload.notes));

  return jsonResponse({ success: true, message: "Batch updated successfully" });
}

function assignStudentToBatch(payload) {
  authorizeAdmin(payload);

  const admissionsSheet = getSheet(SHEET_NAMES.admissions);
  const admissionId = clean(payload.admissionId);
  const batchCode = clean(payload.batchCode);
  const startDate = clean(payload.startDate);
  const endDate = clean(payload.endDate);

  if (!admissionId) {
    return jsonResponse({ success: false, message: "Admission ID is required" });
  }
  if (!batchCode) {
    return jsonResponse({ success: false, message: "Batch Code is required" });
  }

  const rows = getSheetObjects(admissionsSheet);
  const row = rows.find(function(item) {
    return clean(item["Admission ID"]) === admissionId;
  });

  if (!row) {
    return jsonResponse({ success: false, message: "Admission not found" });
  }

  const headerMap = getHeaderMap(admissionsSheet);
  setCellIfHeaderExists(admissionsSheet, row._rowNumber, headerMap, "Batch Code", batchCode);
  if (startDate) {
    setCellIfHeaderExists(admissionsSheet, row._rowNumber, headerMap, "Start Date", startDate);
  }
  if (endDate) {
    setCellIfHeaderExists(admissionsSheet, row._rowNumber, headerMap, "End Date", endDate);
  }
  setCellIfHeaderExists(admissionsSheet, row._rowNumber, headerMap, "Status", "Active");

  return jsonResponse({ success: true, message: "Student assigned to batch successfully" });
}
