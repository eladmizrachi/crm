// ============================================================
// CRM_UpdateContact.gs — Search & Edit Contacts
// Depends on: CRM_NewContact.gs (for getOrganizations)
//             CRM_Setup.gs (for SHEETS)
// ============================================================

const CONTACT_SHEET = "Contacts ";

// ── SERVER: search contacts ──────────────────────────────────
function searchContacts(query, searchBy) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONTACT_SHEET);
    if (!sh) return { ok: false, msg: 'Sheet "' + CONTACT_SHEET.trim() + '" not found.', results: [] };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, results: [] };

    const headers = data[0].map(function(h) { return String(h).trim(); });

    // Locate relevant column indexes
    function col(name) {
      return headers.findIndex(function(h) { return h.toLowerCase().replace(/\s+/g,'') === name.toLowerCase().replace(/\s+/g,''); });
    }
    const emailIdx = col('email');
    const phoneIdx = col('phonenumber');
    const orgIdx   = col('organization');

    const q = query.trim().toLowerCase();
    const results = [];

    data.slice(1).forEach(function(row, i) {
      var match = false;
      if (searchBy === 'email') {
        match = String(row[emailIdx]||'').toLowerCase().indexOf(q) !== -1;
      } else if (searchBy === 'phone') {
        match = String(row[phoneIdx]||'').toLowerCase().indexOf(q) !== -1;
      } else if (searchBy === 'organization') {
        match = String(row[orgIdx]||'').trim().toLowerCase().indexOf(q) !== -1;
      }

      if (match) {
        var obj = { _rowIndex: i + 2 };
        headers.forEach(function(h, j) {
          var v = row[j];
          if (v instanceof Date) {
            obj[h] = ('0'+v.getDate()).slice(-2)+'/'+('0'+(v.getMonth()+1)).slice(-2)+'/'+v.getFullYear();
          } else {
            obj[h] = (v !== null && v !== undefined) ? String(v) : '';
          }
        });
        results.push(obj);
      }
    });

    return { ok: true, results: results, orgs: getOrganizations() };

  } catch(e) {
    return { ok: false, msg: e.message, results: [] };
  }
}


// ── SERVER: update a contact row ─────────────────────────────
function updateContact(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONTACT_SHEET);
    if (!sh) return { ok: false, msg: 'Sheet "' + CONTACT_SHEET.trim() + '" not found.' };

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });

    // yyyy-mm-dd → dd/mm/yyyy
    var dob = payload.dob || '';
    if (dob.indexOf('-') !== -1) {
      var p = dob.split('-');
      if (p.length === 3) dob = p[2]+'/'+p[1]+'/'+p[0];
    }

    // Field map — keys match trimmed sheet headers
    const fieldMap = {
      'First Name':    payload.firstName,
      'Last Name':     payload.lastName,
      'Date of Birth': dob,
      'Postion':       payload.position,   // typo in sheet header
      'Organization':  payload.organization,
      'Phone Number':  payload.phone,
      'Email':         payload.email,
      'Address':       payload.address || '',
    };

    const row = sh.getRange(payload._rowIndex, 1, 1, lastCol).getValues()[0];
    headers.forEach(function(h, i) {
      const key = Object.keys(fieldMap).find(function(k) { return k.trim().toLowerCase() === h.trim().toLowerCase(); });
      if (key !== undefined) row[i] = fieldMap[key];
    });
    sh.getRange(payload._rowIndex, 1, 1, lastCol).setValues([row]);

    return { ok: true, msg: payload.firstName + ' ' + payload.lastName + ' updated successfully!' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}


// ── DIALOG ───────────────────────────────────────────────────
function openUpdateContactDialog() {
  const html = HtmlService.createHtmlOutput(UPDATE_CONTACT_HTML)
    .setWidth(520)
    .setHeight(620)
    .setTitle('Update Contact');
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Contact');
}


const UPDATE_CONTACT_HTML = `<!DOCTYPE html>
<html>
<head>
<style>
*,*::before,*::after{box-sizing:border-box}
body{font-family:'Google Sans',Arial,sans-serif;background:#f3f4ff;margin:0;padding:14px 18px 18px;font-size:13px;color:#1a237e}
h2{margin:0 0 3px;font-size:16px}
p.sub{margin:0 0 12px;color:#555;font-size:12px}

.sr{display:flex;flex-direction:column;gap:6px;margin-bottom:8px}
.sr select{width:100%;padding:8px 10px;font-size:13px;border:2px solid #9fa8da;border-radius:6px;background:#fff;color:#1a237e;outline:none}
.sr select:focus{border-color:#3949ab}
.sr-row{display:flex;gap:8px}
.sr-row input{flex:1;padding:8px 12px;font-size:13px;border:2px solid #9fa8da;border-radius:6px;outline:none;color:#212121}
.sr-row input:focus{border-color:#3949ab}
.sr-row button{padding:8px 18px;font-size:13px;font-weight:700;background:#1a237e;color:#fff;border:none;border-radius:6px;cursor:pointer;white-space:nowrap}
.sr-row button:hover{background:#3949ab}
.sr-row button:disabled{background:#9fa8da;cursor:default}

#gst{font-size:12px;min-height:16px;text-align:center;margin-bottom:8px}
#gst.ok{color:#2e7d32}#gst.err{color:#c62828}

.card{background:#fff;border:1.5px solid #c5cae9;border-radius:8px;margin-bottom:12px;overflow:hidden}
.ch{background:#1a237e;color:#fff;padding:9px 14px;font-size:13px;font-weight:700}
.cb{padding:12px 14px}

.sec{font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#5c6bc0;margin:10px 0 5px;border-bottom:1px solid #e8eaf6;padding-bottom:2px}
.sec:first-child{margin-top:0}
.r2{display:flex;gap:10px}
.r2 .fi{flex:1}
.fi{margin-bottom:7px}
label{display:block;font-size:11px;font-weight:600;color:#5c6bc0;margin-bottom:2px}
.req{color:#c62828;margin-left:2px}
.opt{color:#888;font-weight:400;font-size:10px;margin-left:2px}

input[type=text],input[type=date],input[type=email],select{
  width:100%;padding:7px 9px;font-size:12px;border:1.5px solid #9fa8da;
  border-radius:5px;background:#fff;color:#212121;outline:none}
input:focus,select:focus{border-color:#3949ab}
.ef{border-color:#c62828!important}

.btn-upd{width:100%;padding:8px;font-size:13px;font-weight:700;background:#1a237e;
  color:#fff;border:none;border-radius:6px;cursor:pointer;margin-top:6px}
.btn-upd:hover{background:#3949ab}
.btn-upd:disabled{background:#9fa8da;cursor:default}
.cst{font-size:11px;min-height:14px;text-align:center;margin-top:4px}
.cst.ok{color:#2e7d32}.cst.err{color:#c62828}

.no-res{color:#9e9e9e;font-style:italic;text-align:center;padding:16px 0;font-size:12px}
</style>
</head>
<body>
<h2>✏️ Update Contact</h2>
<p class="sub">Search by name, phone number, or organization.</p>

<div class="sr">
  <select id="searchBy" onchange="updatePlaceholder()">
    <option value="email">Search by Email</option>
    <option value="phone">Search by Phone</option>
    <option value="organization">Search by Organization</option>
  </select>
  <div class="sr-row">
    <input type="text" id="qi" placeholder="Enter email…" />
    <button id="bs" onclick="go()">Search</button>
  </div>
</div>
<div id="gst"></div>
<div id="res"></div>

<script>
var _orgs    = [];
var _cards   = [];   // { _rowIndex, ... } per result

document.getElementById('qi').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') go();
});

var PLACEHOLDERS = { email: 'Enter email…', phone: 'Enter phone number…', organization: 'Enter organization name…' };
function updatePlaceholder() {
  var by = document.getElementById('searchBy').value;
  document.getElementById('qi').placeholder = PLACEHOLDERS[by] || 'Search…';
}

function gst(m, c) { var e = document.getElementById('gst'); e.innerText = m; e.className = c || ''; }

function esc(v) {
  return String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// dd/mm/yyyy → yyyy-mm-dd for date input
function toInp(v) {
  if (!v) return '';
  var p = String(v).split('/');
  return p.length === 3 ? p[2]+'-'+p[1]+'-'+p[0] : '';
}

function buildOrgSelect(id, selected) {
  var s = '<select id="' + id + '">';
  s += '<option value="">— Select —</option>';
  _orgs.forEach(function(v) {
    s += '<option value="' + esc(v) + '"' + (v === selected ? ' selected' : '') + '>' + esc(v) + '</option>';
  });
  return s + '</select>';
}

function fi(lbl, html, req, opt) {
  var mark = req ? '<span class="req">*</span>' : (opt ? '<span class="opt">(optional)</span>' : '');
  return '<div class="fi"><label>' + lbl + mark + '</label>' + html + '</div>';
}

function inp(id, val, placeholder) {
  return '<input type="text" id="' + id + '" value="' + esc(val||'') + '" placeholder="' + esc(placeholder||'') + '" />';
}

// ── Search ──────────────────────────────────────────────────
function go() {
  var q  = document.getElementById('qi').value.trim();
  var by = document.getElementById('searchBy').value;
  document.getElementById('res').innerHTML = '';
  _cards = [];
  if (!q) { gst('Please enter a search term.', 'err'); return; }
  gst('Searching…', '');
  document.getElementById('bs').disabled = true;

  google.script.run
    .withSuccessHandler(function(r) {
      document.getElementById('bs').disabled = false;
      if (!r.ok) { gst(r.msg, 'err'); return; }
      _orgs = r.orgs || [];
      if (r.results.length === 0) { gst('No contacts found.', 'err'); return; }
      gst('Found ' + r.results.length + ' contact(s).', 'ok');
      var html = '';
      r.results.forEach(function(o, idx) {
        _cards.push(o);
        html += buildCard(o, idx);
      });
      document.getElementById('res').innerHTML = html;
    })
    .withFailureHandler(function(e) {
      document.getElementById('bs').disabled = false;
      gst('Error: ' + e.message, 'err');
    })
    .searchContacts(q, by);
}

// ── Card builder ─────────────────────────────────────────────
function buildCard(o, idx) {
  var c   = 'c' + idx;
  var org = o['Organization'] || '';

  var h = '<div class="card" id="card' + idx + '">';
  h += '<div class="ch" id="' + c + '_hdr">'
     + esc(o['First Name']||'') + ' ' + esc(o['Last Name']||'')
     + (org ? ' &mdash; ' + esc(org) : '') + '</div>';
  h += '<div class="cb">';

  h += '<div class="sec">Personal Information</div>';
  h += '<div class="r2">';
  h += fi('First Name', inp(c+'_fn', o['First Name']||'', 'e.g. John'), true);
  h += fi('Last Name',  inp(c+'_ln', o['Last Name']||'',  'e.g. Smith'), true);
  h += '</div>';
  h += '<div class="r2">';
  h += fi('Position', inp(c+'_pos', o['Postion']||'', 'e.g. CTO'), false, true);
  h += fi('Phone Number', inp(c+'_phone', o['Phone Number']||'', 'e.g. 052-1234567'), false, true);
  h += '</div>';
  h += fi('Email', inp(c+'_email', o['Email']||'', 'e.g. john@company.com'), true);
  h += fi('Date of Birth', '<input type="date" id="' + c + '_dob" value="' + toInp(o['Date of Birth']||'') + '" />', false, true);
  h += fi('Address', inp(c+'_address', o['Address']||'', 'e.g. 123 Main St, Tel Aviv'), false, true);

  h += '<div class="sec">Organization</div>';
  h += fi('Organization', buildOrgSelect(c+'_org', org), true);

  h += '<button class="btn-upd" id="' + c + '_btn" onclick="doUpdate(' + idx + ')">💾 Update Contact</button>';
  h += '<div class="cst" id="' + c + '_st"></div>';
  h += '</div></div>';
  return h;
}

// ── Update ───────────────────────────────────────────────────
function doUpdate(idx) {
  var c  = 'c' + idx;
  var st = document.getElementById(c + '_st');
  st.innerText = ''; st.className = 'cst';

  var fn    = document.getElementById(c+'_fn').value.trim();
  var ln    = document.getElementById(c+'_ln').value.trim();
  var email = document.getElementById(c+'_email').value.trim();
  var org   = document.getElementById(c+'_org').value;

  // Clear errors
  [c+'_fn', c+'_ln', c+'_email', c+'_org'].forEach(function(id) {
    document.getElementById(id).classList.remove('ef');
  });

  var ok = true;
  function err(id) { document.getElementById(id).classList.add('ef'); ok = false; }
  if (!fn)    err(c+'_fn');
  if (!ln)    err(c+'_ln');
  if (!email) err(c+'_email');
  if (!org)   err(c+'_org');
  if (!ok) { st.innerText = 'Please fill in all required fields.'; st.className = 'cst err'; return; }

  st.innerText = 'Saving…'; st.className = 'cst';
  document.getElementById(c+'_btn').disabled = true;

  google.script.run
    .withSuccessHandler(function(r) {
      document.getElementById(c+'_btn').disabled = false;
      st.innerText = r.ok ? '✅ ' + r.msg : '❌ ' + r.msg;
      st.className = 'cst ' + (r.ok ? 'ok' : 'err');
      if (r.ok) {
        document.getElementById(c+'_hdr').innerText =
          fn + ' ' + ln + (org ? ' — ' + org : '');
      }
    })
    .withFailureHandler(function(e) {
      document.getElementById(c+'_btn').disabled = false;
      st.innerText = '❌ ' + e.message; st.className = 'cst err';
    })
    .updateContact({
      _rowIndex:    _cards[idx]._rowIndex,
      firstName:    fn,
      lastName:     ln,
      position:     document.getElementById(c+'_pos').value.trim(),
      phone:        document.getElementById(c+'_phone').value.trim(),
      email:        email,
      dob:          document.getElementById(c+'_dob').value,
      organization: org,
      address:      document.getElementById(c+'_address').value.trim(),
    });
}
</script>
</body>
</html>`;
