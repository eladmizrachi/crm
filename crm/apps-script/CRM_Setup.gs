// ============================================================
// CRM_Setup.gs — Sheet Builders & Utilities
// ============================================================

// ── SETTINGS are now managed via CRM Admin → CRM Settings dialog ──
// Defaults are defined in CRM_Config.gs → CONFIG_DEFAULTS_

// Derives manager email from their name: "David Morali" → "david@teamiff.com"
function managerNameToEmail(managerName) {
  var first = String(managerName || '').trim().split(' ')[0].toLowerCase();
  return first ? first + '@teamiff.com' : '';
}

// Looks up the Customer Manager name for an org from the Customers sheet.
function getManagerNameForOrg(orgName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Customers");
    if (!sh) return '';
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return '';
    const n      = h => String(h).trim().replace(/\s+/g,'').toLowerCase();
    const hdrs   = data[0].map(n);
    const orgIdx = hdrs.findIndex(h => h === 'organization');
    const mgrIdx = hdrs.findIndex(h => h === 'customermanager');
    if (orgIdx === -1 || mgrIdx === -1) return '';
    const row = data.slice(1).find(
      r => String(r[orgIdx]||'').trim().toLowerCase() === String(orgName||'').trim().toLowerCase()
    );
    return row ? String(row[mgrIdx]||'').trim() : '';
  } catch(e) { return ''; }
}

// Builds the final recipient list: manager email + always-recipients, deduped.
// When send_email_enabled is false, returns only the test email.
function buildToList(managerEmail) {
  if (getCRMSetting_('send_email_enabled') === 'false') {
    return getCRMSetting_('test_email_recipient');
  }
  const parts = getCRMSetting_('notification_always_to')
    .split(',').map(function(e){ return e.trim(); }).filter(Boolean);
  if (managerEmail && parts.indexOf(managerEmail) === -1) parts.unshift(managerEmail);
  return parts.join(',');
}


// ── INITIAL SETUP ──────────────────────────────────────────
function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSearchSheet(ss);
  createDetailsSheet(ss);
  createRecurringSheet(ss);
  createContactsSheet(ss);
  createOrdersSheet(ss);
  createTicketsSheet(ss);

  // Move Search to first position
  ss.setActiveSheet(ss.getSheetByName(OUT.SEARCH));
  ss.moveActiveSheet(1);

  SpreadsheetApp.getUi().alert(
    "✅ CRM Dashboard Ready!\n\n" +
    "Use the CRM menu → 'Search Customer' to get started."
  );
}


// ── PROTECT DATA SHEETS ────────────────────────────────────
// Uses warning-only protection: users see a warning when editing directly in Sheets,
// but CRM dialogs (which run as the current user via google.script.run) are never blocked.
// Strict editor-list protection would prevent non-admin users from using the CRM dialogs.
function protectDataSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToProtect = [
    SRC.CUSTOMERS,      // "Customers"
    SRC.CONTACTS,       // "Contacts "
    SRC.ORDERS,         // "Purchase Orders"
    BILLING_SHEET_NAME, // "Billing"
  ];

  const results = [];

  sheetsToProtect.forEach(function(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) { results.push('⚠️ Sheet not found: ' + sheetName); return; }

    // Remove any existing protections
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });

    // Warning-only: warns on direct UI edits, never blocks scripts
    const protection = sh.protect();
    protection.setDescription('CRM protected — use CRM dialogs to make changes');
    protection.setWarningOnly(true);

    // Hide from regular users
    sh.hideSheet();

    results.push('✅ Protected & hidden: ' + sheetName.trim());
  });

  SpreadsheetApp.getUi().alert(
    '🔒 Sheet Protection Applied\n\n' +
    results.join('\n') + '\n\n' +
    'Users will see a warning when trying to edit these sheets directly.\n' +
    'All CRM dialogs remain fully functional for every user.'
  );
}


// ── REMOVE PROTECTION (if needed) ──────────────────────────
function unprotectDataSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToUnprotect = [SRC.CUSTOMERS, SRC.CONTACTS, SRC.ORDERS, BILLING_SHEET_NAME];
  const results = [];

  sheetsToUnprotect.forEach(function(sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });
    try { sh.showSheet(); } catch(e) {}
    results.push('🔓 Unprotected & visible: ' + sheetName.trim());
  });

  SpreadsheetApp.getUi().alert(results.join('\n') || 'No protections found.');
}


// ── BUILD SEARCH SHEET ─────────────────────────────────────
function createSearchSheet(ss) {
  let sh = ss.getSheetByName(OUT.SEARCH);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(OUT.SEARCH);

  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 30);

  sh.getRange("A1:E1").merge()
    .setValue("🔍 CRM Customer Search")
    .setBackground("#1a237e").setFontColor("#FFFFFF")
    .setFontSize(18).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 50);

  sh.getRange("A2:E2").merge()
    .setValue("Search by Customer Name, Contact Name or Phone Number")
    .setBackground("#5c6bc0").setFontColor("#FFFFFF")
    .setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 28);

  sh.getRange("A3:E3").merge().setBackground("#f3f4ff");
  sh.setRowHeight(3, 10);

  sh.getRange("B4").setValue("🔎  Search Name:")
    .setBackground("#f3f4ff").setFontSize(12).setFontWeight("bold")
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sh.getRange("C4").setValue("")
    .setBackground("#FFFFFF").setFontSize(13)
    .setBorder(true, true, true, true, false, false, "#5c6bc0", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(4, 36);
  sh.getRange("A4").setBackground("#f3f4ff");
  sh.getRange("D4:E4").setBackground("#f3f4ff");

  sh.getRange("A5:E5").merge().setBackground("#f3f4ff");
  sh.setRowHeight(5, 10);

  sh.getRange("B6:D6").merge()
    .setValue("💡 Type a name or phone number above, then use menu: CRM → Search Customer")
    .setBackground("#f3f4ff").setFontColor("#5c6bc0").setFontSize(10).setHorizontalAlignment("center");
  sh.getRange("A6").setBackground("#f3f4ff");
  sh.getRange("E6").setBackground("#f3f4ff");

  sh.getRange("A7:E7").merge().setBackground("#f3f4ff");
  sh.setRowHeight(7, 15);

  sh.getRange("A8:E8").setBackground("#3949ab");
  sh.getRange("B8").setValue("MATCHING CUSTOMERS")
    .setFontColor("#FFFFFF").setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center");

  sh.getRange("B9").setValue("Organization").setBackground("#e8eaf6").setFontWeight("bold");
  sh.getRange("C9").setValue("Industry").setBackground("#e8eaf6").setFontWeight("bold");
  sh.getRange("D9").setValue("Customer Manager").setBackground("#e8eaf6").setFontWeight("bold");
  sh.getRange("A9:E9").setBorder(false, false, true, false, false, false, "#9fa8da", SpreadsheetApp.BorderStyle.SOLID);

  for (let r = 10; r <= 14; r++) {
    sh.getRange(`B${r}:D${r}`).setBackground(r % 2 === 0 ? "#f5f5ff" : "#FFFFFF");
    sh.getRange(`A${r}`).setBackground("#FFFFFF");
    sh.getRange(`E${r}`).setBackground("#FFFFFF");
  }
  sh.getRange("B10").setValue("← Enter a name above and search").setFontColor("#9e9e9e").setFontStyle("italic");
}


// ── BUILD DETAILS SHEET ────────────────────────────────────
function createDetailsSheet(ss) {
  let sh = ss.getSheetByName(OUT.DETAILS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(OUT.DETAILS);

  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 280);
  sh.setColumnWidth(5, 30);

  titleRow(sh, "📋 Customer Details", 1);
  sh.getRange("A2:E2").merge().setValue("Customer information will appear here after a search")
    .setBackground("#5c6bc0").setFontColor("#FFFFFF")
    .setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("A3:E30").setBackground("#FAFAFA");
}


// ── BUILD RECURRING ACTIVITY SHEET ────────────────────────
function createRecurringSheet(ss) {
  let sh = ss.getSheetByName(OUT.RECURRING);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(OUT.RECURRING);

  sh.setColumnWidth(1, 20);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 120);
  sh.setColumnWidth(5, 130);
  sh.setColumnWidth(6, 160);
  sh.setColumnWidth(7, 20);

  titleRow(sh, "📅 Recurring Activity", 1);
  sh.getRange("A2:G2").merge().setValue("Recurring activity will appear here after a search")
    .setBackground("#5c6bc0").setFontColor("#FFFFFF")
    .setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("A3:G30").setBackground("#FAFAFA");
}


// ── BUILD CONTACTS SHEET ───────────────────────────────────
function createContactsSheet(ss) {
  let sh = ss.getSheetByName(OUT.CONTACTS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(OUT.CONTACTS);

  sh.setColumnWidth(1, 30);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 140);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 140);
  sh.setColumnWidth(7, 30);

  titleRow(sh, "👤 Customer Contacts", 1);
  sh.getRange("A2:G2").merge().setValue("Contacts will appear here after a search")
    .setBackground("#5c6bc0").setFontColor("#FFFFFF")
    .setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("A3:G30").setBackground("#FAFAFA");
}


// ── BUILD ORDERS SHEET ─────────────────────────────────────
function createOrdersSheet(ss) {
  let sh = ss.getSheetByName(OUT.ORDERS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(OUT.ORDERS);

  for (let c = 1; c <= 13; c++) sh.setColumnWidth(c, c === 1 || c === 13 ? 20 : 140);

  titleRow(sh, "🛒 Purchase Orders", 1);
  sh.getRange("A2:M2").merge().setValue("Purchase orders will appear here after a search")
    .setBackground("#5c6bc0").setFontColor("#FFFFFF")
    .setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("A3:M30").setBackground("#FAFAFA");
}


// ── BUILD TICKETS SHEET ────────────────────────────────────
function createTicketsSheet(ss) {
  let sh = ss.getSheetByName(OUT.TICKETS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(OUT.TICKETS);

  for (let c = 1; c <= 8; c++) sh.setColumnWidth(c, c === 1 || c === 8 ? 20 : 170);

  titleRow(sh, "🎫 Service Tickets", 1);
  sh.getRange("A2:H2").merge().setValue("Service tickets will appear here after a search")
    .setBackground("#5c6bc0").setFontColor("#FFFFFF")
    .setFontSize(11).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.getRange("A3:H30").setBackground("#FAFAFA"); 
}


// ── HELPER: title row ──────────────────────────────────────
function titleRow(sh, text, row) {
  const lastCol = sh.getMaxColumns();
  const endCol  = String.fromCharCode(64 + lastCol);
  sh.getRange(`A${row}:${endCol}${row}`).merge()
    .setValue(text)
    .setBackground("#1a237e").setFontColor("#FFFFFF")
    .setFontSize(16).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(row, 44);
}


// ── UTILITY: read sheet into array of objects ───────────────
function sheetToObjects(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}
