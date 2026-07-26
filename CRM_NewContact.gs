// ============================================================
// CRM_NewContact.gs — New Contact Creation
// Depends on: CRM_Setup.gs (for SHEETS, norm, runSearch)
// Add this file to the same Apps Script project.
// ============================================================


// ── LOAD EXISTING ORGANIZATIONS FROM CUSTOMERS SHEET ───────
function getOrganizations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SRC.CUSTOMERS);
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const orgIdx  = headers.findIndex(h => String(h).trim().replace(/\s+/g,"") === "Organization");

  return data.slice(1)
    .map(row => String(row[orgIdx] || "").trim())
    .filter(v => v !== "")
    .sort();
}


// ── OPEN NEW CONTACT DIALOG ────────────────────────────────
function openNewContactDialog() {
  const orgs = getOrganizations();

  function buildOptions(arr) {
    return arr.map(v => `<option value="${escapeHtmlContact(v)}">${escapeHtmlContact(v)}</option>`).join('');
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

        .row2 { display: flex; gap: 10px; }
        .row2 .field { flex: 1; }

        .field { margin-bottom: 10px; }
        label { display: block; font-weight: 600; margin-bottom: 3px; color: #1a237e; }
        label .req { color: #c62828; margin-left: 2px; }
        label .opt { color: #888; font-weight: 400; font-size: 11px; margin-left: 4px; }

        input[type=text], input[type=date], select {
          width: 100%; padding: 8px 10px; font-size: 13px;
          border: 1.5px solid #9fa8da; border-radius: 6px;
          background: #fff; color: #212121; outline: none;
          transition: border-color .15s;
        }
        input:focus, select:focus { border-color: #3949ab; }
        input.error, select.error { border-color: #c62828 !important; }

        .btn-row { display: flex; gap: 10px; margin-top: 18px; }
        button {
          flex: 1; padding: 10px; font-size: 13px;
          font-weight: 700; border: none; border-radius: 6px; cursor: pointer;
        }
        #btnSave         { background: #1a237e; color: #fff; }
        #btnSave:hover   { background: #3949ab; }
        #btnCancel       { background: #e8eaf6; color: #1a237e; }
        #btnCancel:hover { background: #c5cae9; }

        #status {
          margin-top: 12px; font-size: 12px;
          min-height: 18px; text-align: center;
        }
        #status.ok  { color: #2e7d32; }
        #status.err { color: #c62828; }
      </style>
    </head>
    <body>
      <h2>👤 New Contact</h2>
      <p class="sub">Add a new contact and link them to an existing customer.</p>

      <!-- Organization -->
      <div class="section-title">Organization</div>
      <div class="field">
        <label>Organization <span class="req">*</span></label>
        <select id="organization">
          <option value="">— Select Customer —</option>
          ${buildOptions(orgs)}
        </select>
      </div>

      <!-- Personal Info -->
      <div class="section-title">Personal Information</div>
      <div class="row2">
        <div class="field">
          <label>First Name <span class="req">*</span></label>
          <input type="text" id="firstName" placeholder="e.g. John" />
        </div>
        <div class="field">
          <label>Last Name <span class="req">*</span></label>
          <input type="text" id="lastName" placeholder="e.g. Smith" />
        </div>
      </div>

      <div class="row2">
        <div class="field">
          <label>Position <span class="opt">(optional)</span></label>
          <input type="text" id="position" placeholder="e.g. CTO" />
        </div>
        <div class="field">
          <label>Phone Number <span class="opt">(optional)</span></label>
          <input type="text" id="phone" placeholder="e.g. 052-1234567" />
        </div>
      </div>

      <div class="field">
        <label>Email <span class="req">*</span></label>
        <input type="text" id="email" placeholder="e.g. john@company.com" />
      </div>

      <div class="field">
        <label>Date of Birth <span class="opt">(optional)</span></label>
        <input type="date" id="dob" />
      </div>

      <div class="field">
        <label>Address <span class="opt">(optional)</span></label>
        <input type="text" id="address" placeholder="e.g. 123 Main St, Tel Aviv" />
      </div>

      <!-- Buttons -->
      <div class="btn-row">
        <button id="btnSave" onclick="doSave()">💾 Save Contact</button>
        <button id="btnCancel" onclick="google.script.host.close()">Cancel</button>
      </div>
      <div id="status"></div>

      <script>
        function setStatus(msg, type) {
          var el = document.getElementById('status');
          el.innerText = msg;
          el.className = type || '';
        }
        function clearErrors() {
          ['organization','firstName','lastName','email'].forEach(function(id) {
            document.getElementById(id).classList.remove('error');
          });
        }
        function doSave() {
          clearErrors();
          var organization = document.getElementById('organization').value;
          var firstName    = document.getElementById('firstName').value.trim();
          var lastName     = document.getElementById('lastName').value.trim();
          var position     = document.getElementById('position').value.trim();
          var phone        = document.getElementById('phone').value.trim();
          var email        = document.getElementById('email').value.trim();
          var dob          = document.getElementById('dob').value;
          var address      = document.getElementById('address').value.trim();

          var valid = true;
          if (!organization) { document.getElementById('organization').classList.add('error'); valid = false; }
          if (!firstName)    { document.getElementById('firstName').classList.add('error');    valid = false; }
          if (!lastName)     { document.getElementById('lastName').classList.add('error');     valid = false; }
          if (!email)        { document.getElementById('email').classList.add('error');        valid = false; }

          if (!valid) { setStatus('⚠️ Please fill in all required fields.', 'err'); return; }

          setStatus('⏳ Saving…');
          document.getElementById('btnSave').disabled = true;

          google.script.run
            .withSuccessHandler(function(result) {
              if (result.ok) {
                setStatus('✅ ' + result.msg, 'ok');
                // Reset form
                document.getElementById('organization').selectedIndex = 0;
                ['firstName','lastName','position','phone','email','dob','address'].forEach(function(id) {
                  document.getElementById(id).value = '';
                });
              } else {
                setStatus('❌ ' + result.msg, 'err');
              }
              document.getElementById('btnSave').disabled = false;
            })
            .withFailureHandler(function(e) {
              setStatus('❌ Error: ' + e.message, 'err');
              document.getElementById('btnSave').disabled = false;
            })
            .saveNewContact({
              organization: organization,
              firstName:    firstName,
              lastName:     lastName,
              position:     position,
              phone:        phone,
              email:        email,
              dob:          dob,
              address:      address,
            });
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(460).setHeight(570).setTitle("New Contact");

  SpreadsheetApp.getUi().showModalDialog(html, "New Contact");
}


// ── SAVE NEW CONTACT ───────────────────────────────────────
function saveNewContact(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const contactSheet = ss.getSheetByName(SRC.CONTACTS);
    if (!contactSheet) return { ok: false, msg: 'Sheet "' + SRC.CONTACTS + '" not found.' };

    // Read headers from the Contacts sheet
    const headers = contactSheet.getRange(1, 1, 1, contactSheet.getLastColumn()).getValues()[0];

    // Format date as dd/MM/yyyy if provided
    let dobFormatted = "";
    if (data.dob) {
      const parts = data.dob.split("-"); // input date is yyyy-mm-dd
      if (parts.length === 3) {
        dobFormatted = parts[2] + "/" + parts[1] + "/" + parts[0];
      }
    }

    // Build the row and write it — clear any data validation first so
    // values like Hebrew text aren't rejected by sheet-level rules.
    const newRowIndex = contactSheet.getLastRow() + 1;
    const rowValues = buildContactRow(headers, {
      "First Name":    data.firstName,
      "Last Name ":    data.lastName,
      "Date of Birth": dobFormatted,
      "Postion ":      data.position,   // note: typo matches the actual sheet header
      "Organization ": data.organization,
      "Phone Number":  data.phone,
      "Email":         data.email,
      "Address":       data.address || "",
    });
    const newRange = contactSheet.getRange(newRowIndex, 1, 1, rowValues.length);
    newRange.clearDataValidations();
    newRange.setValues([rowValues]);

    // Auto-search the organization so the new contact appears immediately
    runSearch(data.organization);

    return { ok: true, msg: data.firstName + " " + data.lastName + " saved successfully!" };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}


// ── HELPER: build a row array aligned to sheet headers ─────
function buildContactRow(headers, fieldMap) {
  return headers.map(h => {
    const key = String(h).trim();
    if (fieldMap.hasOwnProperty(key)) return fieldMap[key];
    // Fallback: match ignoring trailing spaces
    const found = Object.keys(fieldMap).find(k => k.trim() === key);
    return found !== undefined ? fieldMap[found] : "";
  });
}


// ── HELPER: escape HTML special characters ─────────────────
function escapeHtmlContact(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
