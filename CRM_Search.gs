// ============================================================
// CRM_Search.gs — Customer Search & Results Population
// ============================================================


// ── SOURCE SHEET NAMES ─────────────────────────────────────
const SRC = {
  CUSTOMERS:  "Customers",
  CONTACTS:   "Contacts ",
  ORDERS:     "Purchase Orders",
  TICKETS:    "Service Ticket ",
};

// ── OUTPUT SHEET NAMES ─────────────────────────────────────
const OUT = {
  SEARCH:    "🔍 Search",
  DETAILS:   "📋 Customer Details",
  CONTACTS:  "👤 Contacts",
  ORDERS:    "🛒 Purchase Orders",
  TICKETS:   "🎫 Service Tickets",
};


// ── UTILITIES ──────────────────────────────────────────────
function norm(val) {
  return String(val || "").trim().toLowerCase();
}

function fmtDate(val) {
  if (!val || val === "") return "—";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(val).trim() || "—";
}

function fmtCurrency(val, currency) {
  if (val === null || val === undefined || val === "" || isNaN(Number(val))) return val || "—";
  var sym = (String(currency || "").trim().toUpperCase() === "USD") ? "$" : "₪";
  return sym + Number(val).toLocaleString();
}

function fmtVal(val) {
  if (val === null || val === undefined || val === "") return "—";
  if (val instanceof Date) return fmtDate(val);
  return String(val).trim() || "—";
}

// Read a sheet into an array of objects keyed by trimmed header name
function readSheet(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[String(h).trim()] = row[i]; });
    return obj;
  });
}


// ── SEARCH DIALOG ──────────────────────────────────────────
function openSearchDialog() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(OUT.SEARCH);
  const pre = sh ? sh.getRange("C4").getValue().toString().trim() : "";

  // Load all customer org names for the dropdown
  const orgs = readSheet(ss, SRC.CUSTOMERS)
    .map(r => String(r["Organization"] || "").trim())
    .filter(v => v !== "")
    .sort();

  function buildOptions(arr) {
    return arr.map(v => {
      const esc = v.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
      return `<option value="${esc}">${esc}</option>`;
    }).join('');
  }

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; background: #f3f4ff; margin:0; }
        h2   { color: #1a237e; margin-bottom: 4px; font-size:16px; }
        p    { color: #555; font-size: 12px; margin-top: 0 0 10px; }

        .field-label {
          font-size: 11px; font-weight: 700; color: #5c6bc0;
          text-transform: uppercase; letter-spacing: .05em;
          margin: 12px 0 4px;
        }

        select, input[type=text] {
          width: 100%; padding: 9px 12px; font-size: 13px;
          border: 2px solid #9fa8da; border-radius: 6px;
          box-sizing: border-box; outline: none;
          background: #fff; color: #212121;
        }
        select:focus, input[type=text]:focus { border-color: #3949ab; }

        .divider {
          display: flex; align-items: center; gap: 8px;
          margin: 12px 0; color: #9fa8da; font-size: 11px; font-weight: 600;
        }
        .divider::before, .divider::after {
          content: ''; flex: 1; border-top: 1px solid #c5cae9;
        }

        .btn-row { display: flex; gap: 10px; margin-top: 14px; }
        button {
          flex: 1; padding: 10px; font-size: 13px; border: none;
          border-radius: 6px; cursor: pointer; font-weight: 700;
        }
        #btnSearch { background: #1a237e; color: #fff; }
        #btnSearch:hover { background: #3949ab; }
        #btnClear  { background: #e8eaf6; color: #1a237e; }
        #btnClear:hover { background: #c5cae9; }

        #status { margin-top: 12px; font-size: 13px; min-height: 20px; text-align: center; }
        #status.err { color: #c62828; }
        #status.ok  { color: #2e7d32; }
      </style>
    </head>
    <body>
      <h2>🔍 CRM Customer Search</h2>
      <p>Pick a customer from the dropdown, or search by name / contact / phone.</p>

      <div class="field-label">Select Customer</div>
      <select id="dropdown" onchange="onDropdownChange()">
        <option value="">— All Customers —</option>
        ${buildOptions(orgs)}
      </select>

      <div class="divider">or free-text search</div>

      <input type="text" id="query" placeholder="e.g. UPS, Omri, 052-1234567…" value="${pre}" />

      <div class="btn-row">
        <button id="btnSearch" onclick="doSearch()">Search</button>
        <button id="btnClear"  onclick="doClear()">Clear</button>
      </div>
      <div id="status"></div>

      <script>
        // When dropdown changes, copy value to text box and search immediately
        function onDropdownChange() {
          var val = document.getElementById('dropdown').value;
          if (val) {
            document.getElementById('query').value = val;
            doSearch();
          }
        }

        document.getElementById('query').addEventListener('keydown', function(e) {
          if (e.key === 'Enter') doSearch();
        });

        // Sync text box → reset dropdown if user types manually
        document.getElementById('query').addEventListener('input', function() {
          document.getElementById('dropdown').value = '';
        });

        function doSearch() {
          var q = document.getElementById('query').value.trim();
          if (!q) { setStatus('⚠️ Please enter a search term or select a customer.', ''); return; }
          setStatus('⏳ Searching…', '');
          google.script.run
            .withSuccessHandler(function(msg) { setStatus(msg, msg.startsWith('✅') ? 'ok' : 'err'); })
            .withFailureHandler(function(e)   { setStatus('❌ ' + e.message, 'err'); })
            .runSearch(q);
        }

        function doClear() {
          document.getElementById('query').value = '';
          document.getElementById('dropdown').value = '';
          setStatus('', '');
          google.script.run.clearResults();
        }

        function setStatus(msg, cls) {
          var el = document.getElementById('status');
          el.innerText = msg;
          el.className = cls;
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(440).setHeight(360).setTitle("Search CRM");

  SpreadsheetApp.getUi().showModalDialog(html, "Search CRM");
}


// ── MAIN SEARCH ────────────────────────────────────────────
function runSearch(query) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const q  = norm(query);
  if (!q) return "⚠️ Empty search.";

  // Write query into Search sheet C4
  const searchSh = ss.getSheetByName(OUT.SEARCH);
  if (searchSh) searchSh.getRange("C4").setValue(query);

  // Load all source data
  const customers = readSheet(ss, SRC.CUSTOMERS);
  const contacts  = readSheet(ss, SRC.CONTACTS);
  const orders    = readSheet(ss, SRC.ORDERS);
  const tickets   = readSheet(ss, SRC.TICKETS);

  // Find matching organizations
  const matchedOrgs = new Set();

  customers.forEach(r => {
    if (norm(r["Organization"]).includes(q)) matchedOrgs.add(norm(r["Organization"]));
  });

  contacts.forEach(r => {
    const full = norm(r["First Name"]) + " " + norm(r["Last Name"]);
    if (
      norm(r["First Name"]).includes(q)   ||
      norm(r["Last Name"]).includes(q)    ||
      full.includes(q)                    ||
      norm(r["Phone Number"]).includes(q) ||
      norm(r["Email"]).includes(q)        ||
      norm(r["Organization"]).includes(q)
    ) {
      matchedOrgs.add(norm(r["Organization"]));
    }
  });

  if (matchedOrgs.size === 0) {
    clearResultAreas(ss);
    if (searchSh) {
      searchSh.getRange("B10:D50").clearContent().clearFormat();
      searchSh.getRange("B10")
        .setValue('❌ No results for "' + query + '"')
        .setFontColor("#c62828").setFontStyle("italic");
    }
    return '❌ No results found for "' + query + '"';
  }

  // Filter all data to matched orgs
  const inOrg = r => matchedOrgs.has(norm(r["Organization"]));

  const filtCustomers = customers.filter(inOrg);
  const filtContacts  = contacts.filter(inOrg);
  const filtOrders    = orders.filter(inOrg);
  const filtTickets   = tickets.filter(inOrg);

  // Populate all output sheets
  populateSearchResults(ss, filtCustomers, query);
  populateDetails(ss, filtCustomers);
  populateContacts(ss, filtContacts);
  populateOrders(ss, filtOrders);
  populateTickets(ss, filtTickets);

  // Switch to Search tab
  if (searchSh) ss.setActiveSheet(searchSh);

  return "✅ Found " + matchedOrgs.size + ' customer(s) matching "' + query + '"';
}


// ── POPULATE: SEARCH RESULTS TAB ───────────────────────────
function populateSearchResults(ss, customers, query) {
  const sh = ss.getSheetByName(OUT.SEARCH);
  if (!sh) return;

  sh.getRange("B10:D50").clearContent().clearFormat();

  if (customers.length === 0) {
    sh.getRange("B10").setValue("No customers found").setFontColor("#9e9e9e").setFontStyle("italic");
    return;
  }

  customers.forEach((c, i) => {
    const row = 10 + i;
    const bg  = i % 2 === 0 ? "#e8f5e9" : "#FFFFFF";
    sh.getRange(row, 2).setValue(fmtVal(c["Organization"])).setBackground(bg).setFontWeight("bold");
    sh.getRange(row, 3).setValue(fmtVal(c["Industry"])).setBackground(bg);
    sh.getRange(row, 4).setValue(fmtVal(c["Customer Manager"])).setBackground(bg);
  });

  sh.getRange("B10:D" + (9 + customers.length))
    .setBorder(true, true, true, true, true, true, "#9fa8da", SpreadsheetApp.BorderStyle.SOLID_THIN);
}


// ── POPULATE: CUSTOMER DETAILS TAB ─────────────────────────
function populateDetails(ss, customers) {
  const sh = ss.getSheetByName(OUT.DETAILS);
  if (!sh) return;

  sh.getRange("A3:E100").clearContent().clearFormat().setBackground("#FAFAFA");

  const HEADER_BG = "#3949ab", HEADER_FG = "#FFFFFF";
  const LABEL_BG  = "#e8eaf6", VALUE_BG  = "#FFFFFF";

  let row = 3;

  customers.forEach((c, idx) => {
    const org = fmtVal(c["Organization"]);

    // Customer block title
    sh.getRange(row, 1, 1, 5).merge()
      .setValue("  " + (idx + 1) + ".  " + org)
      .setBackground(HEADER_BG).setFontColor(HEADER_FG)
      .setFontSize(13).setFontWeight("bold").setVerticalAlignment("middle");
    sh.setRowHeight(row, 32);
    row++;

    [
      ["Organization",     c["Organization"]],
      ["Industry",         c["Industry"]],
      ["Main Activity",    c["Main activity type"]],
      ["Customer Manager", c["Customer Manager"]],
    ].forEach(([label, value]) => {
      sh.getRange(row, 2).setValue(label).setBackground(LABEL_BG).setFontWeight("bold").setFontSize(11);
      sh.getRange(row, 3, 1, 2).merge().setValue(fmtVal(value)).setBackground(VALUE_BG).setFontSize(11);
      sh.getRange(row, 1).setBackground(LABEL_BG);
      sh.getRange(row, 5).setBackground(VALUE_BG);
      row++;
    });

    // Spacer
    sh.getRange(row, 1, 1, 5).setBackground("#FAFAFA");
    row++;
  });
}


// ── POPULATE: CONTACTS TAB ─────────────────────────────────
function populateContacts(ss, contacts) {
  const sh = ss.getSheetByName(OUT.CONTACTS);
  if (!sh) return;

  sh.getRange("A3:J100").clearContent().clearFormat().setBackground("#FAFAFA");

  const LABEL_BG = "#e8eaf6", BORDER = "#9fa8da", ROW_ALT = "#f5f5ff";

  const headers = ["First Name", "Last Name", "Position", "Organization", "Phone Number", "Date of Birth", "Email", "Address"];
  headers.forEach((h, i) => {
    sh.getRange(3, i + 2)
      .setValue(h).setBackground(LABEL_BG).setFontWeight("bold").setFontSize(11)
      .setBorder(false, false, true, false, false, false, BORDER, SpreadsheetApp.BorderStyle.SOLID);
  });

  if (contacts.length === 0) {
    sh.getRange("B4").setValue("No contacts found").setFontColor("#9e9e9e").setFontStyle("italic");
    return;
  }

  contacts.forEach((c, i) => {
    const row = 4 + i;
    const bg  = i % 2 === 0 ? "#FFFFFF" : ROW_ALT;
    sh.getRange(row, 2).setValue(fmtVal(c["First Name"])).setBackground(bg);
    sh.getRange(row, 3).setValue(fmtVal(c["Last Name"])).setBackground(bg);
    sh.getRange(row, 4).setValue(fmtVal(c["Postion"])).setBackground(bg);
    sh.getRange(row, 5).setValue(fmtVal(c["Organization"])).setBackground(bg).setFontWeight("bold");
    sh.getRange(row, 6).setValue(fmtVal(c["Phone Number"])).setBackground(bg);
    sh.getRange(row, 7).setValue(fmtDate(c["Date of Birth"])).setBackground(bg);
    sh.getRange(row, 8).setValue(fmtVal(c["Email"])).setBackground(bg);
    sh.getRange(row, 9).setValue(fmtVal(c["Address"])).setBackground(bg);
    sh.getRange(row, 1).setBackground(bg);
    sh.getRange(row, 10).setBackground(bg);
  });

  sh.getRange(3, 2, contacts.length + 1, 8)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID_THIN);
}


// ── POPULATE: PURCHASE ORDERS TAB ──────────────────────────
// HOW TO ADD A NEW COLUMN:
//   1. Add the display label to the `headers` array below.
//   2. Add the matching o["Exact Sheet Header"] value to the vals array.
// ─────────────────────────────────────────────────────────────
function populateOrders(ss, orders) {
  const sh = ss.getSheetByName(OUT.ORDERS);
  if (!sh) return;

  sh.getRange("A3:R100").clearContent().clearFormat().setBackground("#FAFAFA");

  const LABEL_BG = "#e8eaf6", BORDER = "#9fa8da", ROW_ALT = "#f5f5ff";

  // ── Step 1: Add/remove display labels here ──────────────
  const headers = [
    "Date",           // col B  ← o["PO Date"]
    "Organization",   // col C  ← o["Organization"]
    "Project",        // col D  ← o["Project"]
    "Description",    // col E  ← o["Project description"]
    "Project Name",   // col F  ← o["Project Name"]
    "PO Number",      // col G  ← o["PO number"]
    "Amount",         // col G  ← o["Amount"]
    "Recurring Amt",  // col H  ← o["Recurring Amount"]
    "Milestones",     // col I  ← o["Milestones"]
    "Hours",          // col J  ← o["Hours"]
    "Price/Hour",     // col K  ← o["Price per hour"]
    "Total",          // col L  ← o["Total Amount"]
    "Customer",       // col M  ← o["Customer"]
    "Proposal Link",  // col N  ← o["Proposal link"]
    "Billing Type",   // col O  ← o["Billing Type"]
    "Renewal Date",   // col P  ← o["Renwal Date"]
    "Commbox ARR",    // col Q  ← o["Commbox ARR"]
  ];

  headers.forEach((h, i) => {
    sh.getRange(3, i + 2)
      .setValue(h).setBackground(LABEL_BG).setFontWeight("bold").setFontSize(10)
      .setBorder(false, false, true, false, false, false, BORDER, SpreadsheetApp.BorderStyle.SOLID);
  });

  if (orders.length === 0) {
    sh.getRange("B4").setValue("No purchase orders found").setFontColor("#9e9e9e").setFontStyle("italic");
    return;
  }

  orders.forEach((o, i) => {
    const row = 4 + i;
    const bg  = i % 2 === 0 ? "#FFFFFF" : ROW_ALT;
    const cur = String(o["Currency"] || "").trim();

    // ── Step 2: Add/remove row values here (must match headers order) ──
    const vals = [
      fmtDate(o["PO Date"]),                    // "Date"
      fmtVal(o["Organization"]),                // "Organization"
      fmtVal(o["Project"]),                     // "Project"
      fmtVal(o["Project description"]),         // "Description"
      fmtVal(o["Project Name"]),                // "Project Name"
      fmtVal(o["PO number"]),                   // "PO Number"
      fmtCurrency(o["Amount"], cur),            // "Amount"
      fmtCurrency(o["Recurring Amount"], cur),  // "Recurring Amt"
      fmtVal(o["Milestones"]),                  // "Milestones"
      fmtVal(o["Hours"]),                       // "Hours"
      fmtCurrency(o["Price per hour"], cur),    // "Price/Hour"
      fmtCurrency(o["Total Amount"]),             // "Total" (always NIS)
      fmtVal(o["Customer"]),                    // "Customer"
      fmtVal(o["Proposal link"]),               // "Proposal Link"
      fmtVal(o["Billing Type"]),                // "Billing Type"
      fmtDate(o["Renwal Date"]),                // "Renewal Date"
      fmtCurrency(o["Commbox ARR"], cur),       // "Commbox ARR"
    ];

    vals.forEach((val, j) => {
      sh.getRange(row, j + 2).setValue(val).setBackground(bg);
    });
    sh.getRange(row, 1).setBackground(bg);
    sh.getRange(row, headers.length + 2).setBackground(bg);
  });

  sh.getRange(3, 2, orders.length + 1, headers.length)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID_THIN);
}


// ── POPULATE: SERVICE TICKETS TAB ──────────────────────────
function populateTickets(ss, tickets) {
  const sh = ss.getSheetByName(OUT.TICKETS);
  if (!sh) return;

  sh.getRange("A3:H100").clearContent().clearFormat().setBackground("#FAFAFA");

  const LABEL_BG = "#e8eaf6", BORDER = "#9fa8da", ROW_ALT = "#f5f5ff";

  const headers = ["Organization", "Ticket Date", "End Date", "Topic", "Status", "Contact"];
  headers.forEach((h, i) => {
    sh.getRange(3, i + 2)
      .setValue(h).setBackground(LABEL_BG).setFontWeight("bold").setFontSize(11)
      .setBorder(false, false, true, false, false, false, BORDER, SpreadsheetApp.BorderStyle.SOLID);
  });

  if (tickets.length === 0) {
    sh.getRange("B4").setValue("No service tickets found").setFontColor("#9e9e9e").setFontStyle("italic");
    return;
  }

  tickets.forEach((t, i) => {
    const row = 4 + i;
    const bg  = i % 2 === 0 ? "#FFFFFF" : ROW_ALT;
    [
      fmtVal(t["Organization"]),
      fmtDate(t["Ticket date"]),
      fmtDate(t["Ticket End date"]),
      fmtVal(t["Ticke Topic"]),
      fmtVal(t["Status"]),
      fmtVal(t["Ticke contact"]),
    ].forEach((val, j) => {
      sh.getRange(row, j + 2).setValue(val).setBackground(bg);
    });
    sh.getRange(row, 1).setBackground(bg);
    sh.getRange(row, 8).setBackground(bg);
  });

  sh.getRange(3, 2, tickets.length + 1, 6)
    .setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID_THIN);
}


// ── CLEAR RESULTS ──────────────────────────────────────────
function clearResults() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearResultAreas(ss);
  const sh = ss.getSheetByName(OUT.SEARCH);
  if (sh) {
    sh.getRange("C4").clearContent();
    sh.getRange("B10:D50").clearContent().clearFormat();
    sh.getRange("B10").setValue("← Enter a name above and search").setFontColor("#9e9e9e").setFontStyle("italic");
  }
}

function clearResultAreas(ss) {
  [OUT.DETAILS, OUT.CONTACTS, OUT.ORDERS, OUT.TICKETS].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.getRange("A3:Z100").clearContent().clearFormat().setBackground("#FAFAFA");
  });
}
