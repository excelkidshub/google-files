function saveExpense(payload) {
  authorizeAdmin(payload);

  const expensesSheet = getSheet(SHEET_NAMES.expenses);
  const amount = toNumber(payload.amount, 0);

  if (amount <= 0) {
    return jsonResponse({ success: false, message: "Amount must be greater than zero" });
  }

  appendObjectRow(expensesSheet, {
    "Expense ID": getNextPrefixedId(expensesSheet, "Expense ID", "E"),
    "Expense Date": clean(payload.expenseDate) || new Date(),
    "Category": clean(payload.category),
    "Amount": amount,
    "Payment Mode": clean(payload.paymentMode),
    "Description": clean(payload.description),
    "Vendor": clean(payload.vendor)
  });

  return jsonResponse({ success: true, message: "Expense saved successfully" });
}

function getExpenses(payload) {
  authorizeAdmin(payload);

  var expensesSheet = getSheet(SHEET_NAMES.expenses);
  var expenses = getSheetObjects(expensesSheet).map(function(item) {
    return {
      expenseId: clean(item["Expense ID"]),
      expenseDate: formatDateValue(item["Expense Date"]),
      category: clean(item["Category"]),
      amount: toNumber(item["Amount"], 0),
      paymentMode: clean(item["Payment Mode"]),
      description: clean(item["Description"]),
      vendor: clean(item["Vendor"])
    };
  });

  return jsonResponse({ success: true, data: expenses });
}
