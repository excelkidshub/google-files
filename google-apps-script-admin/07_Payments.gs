function savePayment(payload) {
  authorizeAdmin(payload);

  const admissionsSheet = getSheet(SHEET_NAMES.admissions);
  const paymentsSheet = getSheet(SHEET_NAMES.payments);
  const admissionId = clean(payload.admissionId);
  const amount = toNumber(payload.amount, 0);

  if (!admissionId) {
    return jsonResponse({ success: false, message: "Admission ID is required" });
  }
  if (amount <= 0) {
    return jsonResponse({ success: false, message: "Amount must be greater than zero" });
  }

  const admissionRow = getSheetObjects(admissionsSheet).find(function(item) {
    return clean(item["Admission ID"]) === admissionId;
  });

  if (!admissionRow) {
    return jsonResponse({ success: false, message: "Admission not found" });
  }

  appendObjectRow(paymentsSheet, {
    "Payment ID": getNextPrefixedId(paymentsSheet, "Payment ID", "P"),
    "Admission ID": admissionId,
    "Student Name": clean(admissionRow["Student Name"]),
    "Batch Code": clean(admissionRow["Batch Code"]),
    "Payment Date": clean(payload.paymentDate) || new Date(),
    "Amount": amount,
    "Payment Mode": clean(payload.paymentMode),
    "Transaction ID": clean(payload.transactionId),
    "Notes": clean(payload.notes)
  });

  recalculateAdmissionFinancials(admissionsSheet, admissionRow._rowNumber);

  return jsonResponse({ success: true, message: "Payment saved successfully" });
}

function getPayments(payload) {
  authorizeAdmin(payload);

  var paymentsSheet = getSheet(SHEET_NAMES.payments);
  var payments = getSheetObjects(paymentsSheet).map(function(item) {
    return {
      paymentId: clean(item["Payment ID"]),
      admissionId: clean(item["Admission ID"]),
      studentName: clean(item["Student Name"]),
      batchCode: clean(item["Batch Code"]),
      paymentDate: formatDateValue(item["Payment Date"]),
      amount: toNumber(item["Amount"], 0),
      paymentMode: clean(item["Payment Mode"]),
      transactionId: clean(item["Transaction ID"]),
      notes: clean(item["Notes"])
    };
  });

  return jsonResponse({ success: true, data: payments });
}

function recalculateAdmissionFinancials(admissionsSheet, rowNumber) {
  const headerMap = getHeaderMap(admissionsSheet);
  const admissionId = clean(admissionsSheet.getRange(rowNumber, headerMap["Admission ID"]).getValue());
  const totalFee = toNumber(admissionsSheet.getRange(rowNumber, headerMap["Total Fee"]).getValue(), 0);
  const discount = toNumber(admissionsSheet.getRange(rowNumber, headerMap["Discount"]).getValue(), 0);
  const manualAdjustment = toNumber(admissionsSheet.getRange(rowNumber, headerMap["Manual Adjustment"]).getValue(), 0);

  const payments = getSheetObjects(getSheet(SHEET_NAMES.payments)).filter(function(item) {
    return clean(item["Admission ID"]) === admissionId;
  });

  const totalPaid = payments.reduce(function(sum, item) {
    return sum + toNumber(item["Amount"], 0);
  }, 0);

  const adjustedFee = Math.max(0, totalFee - discount + manualAdjustment);
  const pending = Math.max(0, adjustedFee - totalPaid);
  const paymentStatus = pending === 0 && adjustedFee > 0 ? "Paid" : totalPaid > 0 ? "Partial" : "Pending";

  setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Adjusted Fee", adjustedFee);
  setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Total Paid", totalPaid);
  setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Pending", pending);
  setCellIfHeaderExists(admissionsSheet, rowNumber, headerMap, "Payment Status", paymentStatus);
}

function getNextPrefixedId(sheet, headerName, prefix) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return prefix + "001";
  }

  const headerMap = requireHeaders(sheet, [headerName]);
  const values = sheet.getRange(2, headerMap[headerName], lastRow - 1, 1).getValues().flat();
  let maxNumber = 0;

  values.forEach(function(value) {
    const text = clean(value);
    const match = text.match(new RegExp("^" + prefix + "(\\d+)$", "i"));
    if (match) {
      const number = parseInt(match[1], 10);
      if (number > maxNumber) {
        maxNumber = number;
      }
    }
  });

  return prefix + String(maxNumber + 1).padStart(3, "0");
}
