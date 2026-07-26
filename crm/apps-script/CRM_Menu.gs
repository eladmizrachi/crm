// ============================================================
// CRM_Menu.gs — Menu Setup ONLY
// This is the ONLY file that should contain onOpen().
// Make sure no other file in the project has an onOpen() function.
// ============================================================

function onOpen() {
  const ui          = SpreadsheetApp.getUi();
  const isAdmin = isCurrentUserAdmin_();

  ui.createMenu("🔍 CRM")
    .addItem("Setup / Reset Dashboard", "setupCRM")
    .addSeparator()
    .addItem("➕ New Customer", "openNewCustomerDialog")
    .addItem("👤 New Contact", "openNewContactDialog")
    .addItem("🛒 New Purchase Order", "openNewPurchaseOrderDialog")
    .addSeparator()
    .addItem("🔍 Search Customer", "openSearchDialog")
    .addItem("🔎 Search Purchase Order", "openSearchPODialog")
    .addItem("🗑️ Clear Results", "clearResults")
    .addToUi();

  if (isAdmin) {
    ui.createMenu("⚙️ CRM Admin")
      .addItem("🔧 CRM Settings",          "openCRMSettingsDialog")
      .addSeparator()
      .addItem("🔒 Protect Data Sheets",   "protectDataSheets")
      .addItem("🔓 Remove Protection",     "unprotectDataSheets")
      .addSeparator()
      .addItem("📊 Monthly Billing Report","openMonthlyReportDialog")
      .addToUi();
  }

  ui.createMenu("✏️ CRM Update")
    .addItem("✏️ Update Customer", "openUpdateCustomerDialog")
    .addItem("✏️ Update Contact", "openUpdateContactDialog")
    .addItem("✏️ Update Purchase Order", "openUpdatePODialog")
    .addToUi();
}
