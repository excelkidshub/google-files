function getDashboard(payload) {
  authorizeAdmin(payload);

  const admissions = getSheetObjects(getSheet(SHEET_NAMES.admissions));
  const batches = getSheetObjects(getSheet(SHEET_NAMES.batches));
  const payments = getSheetObjects(getSheet(SHEET_NAMES.payments));
  const expenses = getSheetObjects(getSheet(SHEET_NAMES.expenses));

  const totalRevenue = payments.reduce(function(sum, item) {
    return sum + toNumber(item["Amount"], 0);
  }, 0);

  const totalExpenses = expenses.reduce(function(sum, item) {
    return sum + toNumber(item["Amount"], 0);
  }, 0);

  const pendingFees = admissions.reduce(function(sum, item) {
    return sum + toNumber(item["Pending"], 0);
  }, 0);

  return jsonResponse({
    success: true,
    data: {
      totalStudents: admissions.length,
      activeStudents: admissions.filter(function(item) { return clean(item["Status"]) === "Active"; }).length,
      pendingStudents: admissions.filter(function(item) { return clean(item["Status"]) === "Pending Start"; }).length,
      activeBatches: batches.filter(function(item) { return clean(item["Status"]) === "Active"; }).length,
      totalRevenue: totalRevenue,
      pendingFees: pendingFees,
      totalExpenses: totalExpenses
    }
  });
}
