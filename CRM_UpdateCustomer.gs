// ============================================================
// CRM_UpdateCustomer.gs — Edit Existing Customer Details
// Depends on: CRM_Search.gs (for SRC, norm, runSearch)
//             CRM_NewCustomer.gs (for getCodeTableValues, buildRow, escapeHtml)
// ============================================================


// ── GET CURRENT DATA FOR A CUSTOMER ────────────────────────
function getCustomerForEdit(orgName) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sh      = ss.getSheetByName(SRC.CUSTOMERS);
  if (!sh) return null;

  const data    = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const orgIdx      = headers.findIndex(h => h.replace(/\s+/g, "") === "Organization");
  const indIdx      = headers.findIndex(h => h.replace(/\s+/g, "") === "Industry");
  const actIdx      = headers.findIndex(h => h.replace(/\s+/g, "") === "Mainactivitytype");
  const mgrIdx      = headers.findIndex(h => h.replace(/\s+/g, "") === "CustomerManager");

  const row = data.slice(1).find(r => norm(String(r[orgIdx] || "")) === norm(orgName));
  if (!row) return null;

  return {
    orgName:  String(row[orgIdx]  || "").trim(),
    industry: String(row[indIdx]  || "").trim(),
    activity: String(row[actIdx]  || "").trim(),
    manager:  String(row[mgrIdx]  || "").trim(),
  };
}


// ── GET ALL ORGANIZATION NAMES ──────────────────────────────
function getCustomersForUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SRC.CUSTOMERS);
  if (!sh) return [];

  const data    = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const orgIdx  = headers.findIndex(h => String(h).trim().replace(/\s+/g, "") === "Organization");

  return data.slice(1)
    .map(row => String(row[orgIdx] || "").trim())
    .filter(v => v !== "")
    .sort();
}


// ── OPEN UPDATE CUSTOMER DIALOG ─────────────────────────────
function openUpdateCustomerDialog() {
  const orgs  = getCustomersForUpdate();
  const codes = getCodeTableValues();

  // Pre-select the last searched customer if it exactly matches an org
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const searchSh = ss.getSheetByName(OUT.SEARCH);
  const lastQuery = searchSh ? searchSh.getRange("C4").getValue().toString().trim() : "";
  const preselect = orgs.find(o => norm(o) === norm(lastQuery)) || "";

  function buildOptions(arr, selected) {
    return arr.map(v => {
      const sel = norm(v) === norm(selected) ? ' selected' : '';
      return `<option value="${escapeHtml(v)}"${sel}>${escapeHtml(v)}</option>`;
    }).join('');
  }

  function buildCodeOptions(arr, selected) {
    return arr.map(v => {
      const sel = norm(v) === norm(selected) ? ' selected' : '';
      return `<option value="${escapeHtml(v)}"${sel}>${escapeHtml(v)}</option>`;
    }).join('');
  }

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        *, *::before, *::after { box-sizing: border-box; }
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          background: #f3f4ff;
          margin: 0; padding: 16px 20px 20px;
          font-size: 13px; color: #1a237e;
        }
        h2 { margin: 0 0 4px; font-size: 17px; color: #1a237e; }
        p.sub { margin: 0 0 14px; color: #555; font-size: 12px; }

        .section-title {
          font-size: 11px; font-weight: 700;
          letter-spacing: .06em; text-transform: uppercase;
          color: #5c6bc0; margin: 14px 0 6px;
          border-bottom: 1px solid #c5cae9; padding-bottom: 3px;
        }
        .field { margin-bottom: 10px; }
        label { display: block; font-weight: 600; margin-bottom: 3px; color: #1a237e; }
        label .req { color: #c62828; margin-left: 2px; }

        input[type=text], select {
          width: 100%; padding: 8px 10px; font-size: 13px;
          border: 1.5px solid #9fa8da; border-radius: 6px;
          background: #fff; color: #212121; outline: none;
          transition: border-color .15s;
        }
        input[type=text]:focus, select:focus { border-color: #3949ab; }
        input.error, select.error { border-color: #c62828 !important; }

        #editFields {
          background: #eef0fb; border: 1px solid #c5cae9;
          border-radius: 8px; padding: 12px 14px; margin-top: 8px;
        }
        #editFields .field { margin-bottom: 8px; }
        #editFields .field:last-child { margin-bottom: 0; }

        #loadingMsg {
          color: #5c6bc0; font-style: italic; font-size: 12px;
          padding: 10px 0; display: none;
        }

        .btn-row { display: flex; gap: 10px; margin-top: 18px; }
        button {
          flex: 1; padding: 10px; font-size: 13px;
          font-weight: 700; border: none; border-radius: 6px; cursor: pointer;
        }
        #btnSave         { background: #1a237e; color: #fff; }
        #btnSave:hover   { background: #3949ab; }
        #btnCancel       { background: #e8eaf6; color: #1a237e; }
        #btnCancel:hover { background: #c5cae9; }
        #btnSave:disabled { background: #9fa8da; cursor: default; }

        #status {
          margin-top: 12px; font-size: 12px;
          min-height: 18px; text-align: center;
        }
        #status.ok  { color: #2e7d32; }
        #status.err { color: #c62828; }
      </style>
    </head>
    <body>
      <h2>✏️ Update Customer</h2>
      <p class="sub">Select a customer to edit their details.</p>

      <div class="section-title">Select Customer</div>
      <div class="field">
        <label>Organization <span class="req">*</span></label>
        <select id="orgSelector" onchange="onOrgSelected()">
          <option value="">— Select Customer —</option>
          ${buildOptions(orgs, preselect)}
        </select>
      </div>

      <div id="loadingMsg">⏳ Loading customer data…</div>

      <div id="editFields" style="display:none">
        <div class="section-title" style="margin-top:0">Edit Details</div>

        <div class="field">
          <label>Organization Name <span class="req">*</span></label>
          <input type="text" id="orgName" />
        </div>
        <div class="field">
          <label>Industry <span class="req">*</span></label>
          <select id="industry">
            <option value="">— Select —</option>
            ${buildCodeOptions(codes.industries, "")}
          </select>
        </div>
        <div class="field">
          <label>Main Activity Type <span class="req">*</span></label>
          <select id="activity">
            <option value="">— Select —</option>
            ${buildCodeOptions(codes.activities, "")}
          </select>
        </div>
        <div class="field">
          <label>Customer Manager <span class="req">*</span></label>
          <select id="manager">
            <option value="">— Select —</option>
            ${buildCodeOptions(codes.managers, "")}
          </select>
        </div>
      </div>

      <div class="btn-row">
        <button id="btnSave" onclick="doSave()" disabled>💾 Save Changes</button>
        <button id="btnCancel" onclick="google.script.host.close()">Cancel</button>
      </div>
      <div id="status"></div>

      <script>
        var originalOrgName = '';

        function setStatus(msg, type) {
          var el = document.getElementById('status');
          el.innerText = msg;
          el.className = type || '';
        }

        function setOption(selectId, value) {
          var sel = document.getElementById(selectId);
          for (var i = 0; i < sel.options.length; i++) {
            if (sel.options[i].value === value) { sel.selectedIndex = i; return; }
          }
          // Value not in list — add a temporary option so it's visible
          var opt = document.createElement('option');
          opt.value = value; opt.text = value; opt.selected = true;
          sel.insertBefore(opt, sel.options[1]);
        }

        function onOrgSelected() {
          var org = document.getElementById('orgSelector').value;
          if (!org) {
            document.getElementById('editFields').style.display = 'none';
            document.getElementById('btnSave').disabled = true;
            originalOrgName = '';
            return;
          }
          document.getElementById('loadingMsg').style.display = 'block';
          document.getElementById('editFields').style.display = 'none';
          document.getElementById('btnSave').disabled = true;
          setStatus('', '');

          google.script.run
            .withSuccessHandler(function(data) {
              document.getElementById('loadingMsg').style.display = 'none';
              if (!data) { setStatus('❌ Could not load customer data.', 'err'); return; }

              originalOrgName = data.orgName;
              document.getElementById('orgName').value = data.orgName;
              setOption('industry', data.industry);
              setOption('activity', data.activity);
              setOption('manager',  data.manager);

              // Clear errors
              ['orgName','industry','activity','manager'].forEach(function(id) {
                document.getElementById(id).classList.remove('error');
              });

              document.getElementById('editFields').style.display = 'block';
              document.getElementById('btnSave').disabled = false;
            })
            .withFailureHandler(function(e) {
              document.getElementById('loadingMsg').style.display = 'none';
              setStatus('❌ Error: ' + e.message, 'err');
            })
            .getCustomerForEdit(org);
        }

        function doSave() {
          var orgName  = document.getElementById('orgName').value.trim();
          var industry = document.getElementById('industry').value;
          var activity = document.getElementById('activity').value;
          var manager  = document.getElementById('manager').value;

          // Clear previous errors
          ['orgName','industry','activity','manager'].forEach(function(id) {
            document.getElementById(id).classList.remove('error');
          });

          var valid = true;
          if (!orgName)  { document.getElementById('orgName').classList.add('error');   valid = false; }
          if (!industry) { document.getElementById('industry').classList.add('error'); valid = false; }
          if (!activity) { document.getElementById('activity').classList.add('error'); valid = false; }
          if (!manager)  { document.getElementById('manager').classList.add('error');  valid = false; }
          if (!valid) { setStatus('⚠️ Please fill in all required fields.', 'err'); return; }

          setStatus('⏳ Saving…', '');
          document.getElementById('btnSave').disabled = true;

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.ok) {
                setStatus('✅ ' + result.msg, 'ok');
                // Update selector label if org was renamed
                var sel = document.getElementById('orgSelector');
                sel.options[sel.selectedIndex].value = orgName;
                sel.options[sel.selectedIndex].text  = orgName;
                originalOrgName = orgName;
              } else {
                setStatus('❌ ' + result.msg, 'err');
              }
              document.getElementById('btnSave').disabled = false;
            })
            .withFailureHandler(function(e) {
              setStatus('❌ Error: ' + e.message, 'err');
              document.getElementById('btnSave').disabled = false;
            })
            .updateCustomer({
              originalOrgName: originalOrgName,
              orgName:         orgName,
              industry:        industry,
              activity:        activity,
              manager:         manager,
            });
        }

        // Auto-load if a customer was pre-selected
        window.onload = function() {
          if (document.getElementById('orgSelector').value) onOrgSelected();
        };
      </script>
    </body>
    </html>
  `)
  .setWidth(460).setHeight(560).setTitle("Update Customer");

  SpreadsheetApp.getUi().showModalDialog(html, "Update Customer");
}


// ── SAVE CUSTOMER UPDATE ────────────────────────────────────
function updateCustomer(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const sh = ss.getSheetByName(SRC.CUSTOMERS);
    if (!sh) return { ok: false, msg: 'Sheet "' + SRC.CUSTOMERS + '" not found.' };

    const allRows = sh.getDataRange().getValues();
    const headers = allRows[0];

    const orgColIdx = headers.findIndex(h => String(h).trim().replace(/\s+/g, "") === "Organization");
    const rowIdx    = allRows.slice(1).findIndex(r => norm(String(r[orgColIdx] || "")) === norm(data.originalOrgName));
    if (rowIdx === -1) return { ok: false, msg: '"' + data.originalOrgName + '" not found in Customers.' };

    // Check for duplicate if org name was changed
    if (norm(data.orgName) !== norm(data.originalOrgName)) {
      const duplicate = allRows.slice(1).find(r => norm(String(r[orgColIdx] || "")) === norm(data.orgName));
      if (duplicate) return { ok: false, msg: '"' + data.orgName + '" already exists.' };
    }

    // Build the updated row (preserve any columns we don't edit)
    const sheetRow = rowIdx + 2; // +1 for header, +1 for 1-based index
    const currentRow = allRows[rowIdx + 1].slice(); // copy

    const fieldMap = {
      "Organization ":       data.orgName,
      "Industry ":           data.industry,
      "Main activity type ": data.activity,
      "Customer Manager ":   data.manager,
    };

    headers.forEach((h, i) => {
      const key   = String(h).trim();
      const match = Object.keys(fieldMap).find(k => k.trim() === key);
      if (match !== undefined) currentRow[i] = fieldMap[match];
    });

    sh.getRange(sheetRow, 1, 1, currentRow.length).setValues([currentRow]);

    // If org was renamed, also update the organization column in Contacts, Orders, Tickets
    if (norm(data.orgName) !== norm(data.originalOrgName)) {
      _renameOrgInSheet(ss, SRC.CONTACTS,        data.originalOrgName, data.orgName);
      _renameOrgInSheet(ss, SRC.ORDERS,          data.originalOrgName, data.orgName);
      _renameOrgInSheet(ss, SRC.TICKETS,         data.originalOrgName, data.orgName);
      _renameOrgInSheet(ss, SHEETS.SRC_RECURRING, data.originalOrgName, data.orgName);
    }

    runSearch(data.orgName);
    return { ok: true, msg: '"' + data.orgName + '" updated successfully!' };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}


// ── HELPER: rename org in another sheet's Organization column ─
function _renameOrgInSheet(ss, sheetName, oldName, newName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return;

  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  const orgIdx  = headers.findIndex(h => String(h).trim().replace(/\s+/g, "") === "Organization");
  if (orgIdx === -1) return;

  data.slice(1).forEach((row, i) => {
    if (norm(String(row[orgIdx] || "")) === norm(oldName)) {
      sh.getRange(i + 2, orgIdx + 1).setValue(newName);
    }
  });
}
