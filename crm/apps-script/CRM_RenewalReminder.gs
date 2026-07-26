// ============================================================
// CRM_RenewalReminder.gs — Renewal Date Email Reminders
// Depends on: CRM_Setup.gs (managerNameToEmail, getManagerNameForOrg, buildToList)
//             CRM_Search.gs (SRC, readSheet, fmtDate)
//
// Set up a daily time-based trigger on sendRenewalReminders()
// via Apps Script → Triggers → Add Trigger.
// ============================================================


// ── MAIN: check all POs and email those renewing in 7 days ─
function sendRenewalReminders() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const orders  = readSheet(ss, SRC.ORDERS);
  const today   = new Date();
  today.setHours(0, 0, 0, 0);

  const target = new Date(today);
  target.setDate(target.getDate() + 7);

  var sent = 0;
  var errors = [];

  orders.forEach(function(o) {
    var rawDate = o["Renwal Date"];
    if (!rawDate) return;

    var renewalDate = parseRenewalDate(rawDate);
    if (!renewalDate) return;

    renewalDate.setHours(0, 0, 0, 0);
    if (renewalDate.getTime() !== target.getTime()) return;

    // Build recipient list
    var org         = String(o["Organization"] || "").trim();
    var mgrName     = getManagerNameForOrg(org);
    var mgrEmail    = managerNameToEmail(mgrName);
    var toList      = buildToList(mgrEmail);

    var project     = String(o["Project"] || "").trim() || "—";
    var poNumber    = String(o["PO number"] || "").trim() || "—";
    var billingType = String(o["Billing Type"] || "").trim() || "—";
    var dateStr     = fmtDate(renewalDate);

    var subject = "🔔 Renewal Reminder: " + org + " — " + project + " (PO " + poNumber + ")";
    var body    = buildEmailBody(org, project, poNumber, billingType, dateStr, mgrName);

    try {
      MailApp.sendEmail({
        to:      toList,
        subject: subject,
        htmlBody: body,
      });
      sent++;
    } catch (e) {
      errors.push(org + ": " + e.message);
    }
  });

  // Log result (visible in Apps Script → Executions)
  Logger.log("Renewal reminders sent: " + sent + (errors.length ? " | Errors: " + errors.join("; ") : ""));
}


// ── MANUAL TEST: run from editor to preview what would fire ─
function previewRenewalReminders() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const orders  = readSheet(ss, SRC.ORDERS);
  const today   = new Date();
  today.setHours(0, 0, 0, 0);

  const target = new Date(today);
  target.setDate(target.getDate() + 7);

  var lines = ["Checking renewal dates for: " + Utilities.formatDate(target, Session.getScriptTimeZone(), "dd/MM/yyyy"), ""];

  orders.forEach(function(o) {
    var rawDate = o["Renwal Date"];
    if (!rawDate) return;

    var renewalDate = parseRenewalDate(rawDate);
    if (!renewalDate) return;

    renewalDate.setHours(0, 0, 0, 0);
    if (renewalDate.getTime() !== target.getTime()) return;

    var org      = String(o["Organization"] || "").trim();
    var mgrName  = getManagerNameForOrg(org);
    var mgrEmail = managerNameToEmail(mgrName);
    var toList   = buildToList(mgrEmail);

    lines.push(
      "MATCH: " + org +
      " | Project: " + (o["Project"] || "—") +
      " | PO: " + (o["PO number"] || "—") +
      " | Renewal: " + fmtDate(renewalDate) +
      " | To: " + toList
    );
  });

  if (lines.length === 2) lines.push("No renewals due in 7 days.");

  SpreadsheetApp.getUi().alert(lines.join("\n"));
}


// ── HELPERS ─────────────────────────────────────────────────

// Accepts a Date object, a dd/mm/yyyy string, or yyyy-mm-dd string
function parseRenewalDate(val) {
  if (!val) return null;
  if (val instanceof Date) return new Date(val);

  var s = String(val).trim();

  // dd/mm/yyyy
  var m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m1) return new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));

  // yyyy-mm-dd
  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));

  return null;
}


function buildEmailBody(org, project, poNumber, billingType, dateStr, mgrName) {
  return '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">'
    + '<div style="background:#1a237e;padding:20px 24px;border-radius:8px 8px 0 0">'
    +   '<h2 style="color:#fff;margin:0;font-size:18px">🔔 PO Renewal Reminder</h2>'
    +   '<p style="color:#c5cae9;margin:4px 0 0;font-size:13px">Renewal due in 7 days</p>'
    + '</div>'
    + '<div style="background:#f3f4ff;padding:20px 24px;border:1px solid #c5cae9;border-top:none;border-radius:0 0 8px 8px">'
    +   '<table style="width:100%;border-collapse:collapse;font-size:14px">'
    +     row2("Organization",   org)
    +     row2("Project",        project)
    +     row2("PO Number",      poNumber)
    +     row2("Billing Type",   billingType)
    +     row2("Renewal Date",   '<strong style="color:#c62828">' + dateStr + '</strong>')
    +     row2("Account Manager", mgrName || "—")
    +   '</table>'
    +   '<p style="margin:18px 0 0;font-size:12px;color:#666">'
    +     'This is an automated reminder from the TeamIFF CRM system.'
    +   '</p>'
    + '</div>'
  + '</div>';
}

function row2(label, value) {
  return '<tr>'
    + '<td style="padding:7px 10px;font-weight:700;color:#5c6bc0;width:38%;background:#e8eaf6;border:1px solid #c5cae9">' + label + '</td>'
    + '<td style="padding:7px 10px;color:#212121;background:#fff;border:1px solid #c5cae9">' + value + '</td>'
    + '</tr>';
}
