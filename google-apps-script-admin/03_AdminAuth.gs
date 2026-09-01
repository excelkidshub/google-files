function adminLogin(payload) {
  const providedPassword = clean(payload.adminPassword);
  const expectedPassword = getScriptProperty("ADMIN_PASSWORD");
  const adminToken = getScriptProperty("ADMIN_TOKEN");

  if (!expectedPassword) {
    throw new Error("ADMIN_PASSWORD is not configured in Script Properties");
  }

  if (!adminToken) {
    throw new Error("ADMIN_TOKEN is not configured in Script Properties");
  }

  if (!providedPassword || providedPassword !== expectedPassword) {
    return jsonResponse({ success: false, message: "Invalid admin password" });
  }

  return jsonResponse({
    success: true,
    message: "Login successful",
    adminToken: adminToken
  });
}

function authorizeAdmin(payload) {
  const adminToken = clean(payload.adminToken);
  const expectedToken = getScriptProperty("ADMIN_TOKEN");

  if (!expectedToken) {
    throw new Error("ADMIN_TOKEN is not configured in Script Properties");
  }

  if (!adminToken || adminToken !== expectedToken) {
    throw new Error("Unauthorized");
  }
}
