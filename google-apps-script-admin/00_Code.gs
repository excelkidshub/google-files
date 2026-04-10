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
