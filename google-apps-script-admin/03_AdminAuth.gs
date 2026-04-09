function adminLogin(payload) {
  const providedPassword = clean(payload.adminPassword);
  const expectedPassword = getScriptProperty(SCRIPT_PROPERTY_KEYS.adminPassword);
  const adminToken = getScriptProperty(SCRIPT_PROPERTY_KEYS.adminToken);

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
  const expectedToken = getScriptProperty(SCRIPT_PROPERTY_KEYS.adminToken);

  if (!expectedToken) {
    throw new Error("ADMIN_TOKEN is not configured in Script Properties");
  }

  if (!adminToken || adminToken !== expectedToken) {
    throw new Error("Unauthorized");
  }
}
