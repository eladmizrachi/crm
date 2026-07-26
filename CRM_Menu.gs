// ============================================================
// CRM_Menu.gs — Menu Setup ONLY
// This is the ONLY file that should contain onOpen().
// Make sure no other file in the project has an onOpen() function.
// ============================================================

function onOpen() {
  const ui          = SpreadsheetApp.getUi();
  const currentUser = Session.getActiveUser().getEmail().toLowerCase();
  const isAdmin     = SHEET_EDITORS.map(function(e) { return e.toLowerCase(); })
                        .indexOf(currentUser) !== -1;

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
      .addItem("🔒 Protect Data Sheets", "protectDataSheets")
      .addItem("🔓 Remove Protection",   "unprotectDataSheets")
      .addToUi();
  }

  ui.createMenu("✏️ CRM Update")
    .addItem("✏️ Update Customer", "openUpdateCustomerDialog")
    .addItem("✏️ Update Contact", "openUpdateContactDialog")
    .addItem("✏️ Update Purchase Order", "openUpdatePODialog")
    .addToUi();
}
