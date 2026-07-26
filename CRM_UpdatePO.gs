// ============================================================
// CRM_UpdatePO.gs — Search & Edit Purchase Orders
// Depends on: CRM_NewPurchaseOrder.gs (getPOOrganizations, getPOActivities, getUsdToNisRate)
// ============================================================

// SERVER: find PO by number
function findPOByNumber(poNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Purchase Orders");
    if (!sh) return { ok: false, msg: 'Sheet "Purchase Orders" not found.', data: null };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: false, msg: 'No data in sheet.', data: null };

    const headers = data[0].map(function(h) { return String(h).trim(); });

    function col(name) {
      return headers.findIndex(function(h) {
        return h.toLowerCase().replace(/\s+/g, '') === name.toLowerCase().replace(/\s+/g, '');
      });
    }
    const poIdx = col('ponumber');
    if (poIdx === -1) return { ok: false, msg: '"PO number" column not found.', data: null };

    const q = String(poNumber).trim();
    let rowIndex = -1;
    let found = null;

    data.slice(1).forEach(function(row, i) {
      if (!found && String(row[poIdx] || '').trim() === q) {
        rowIndex = i + 2;
        found = row;
      }
    });

    if (!found) return { ok: false, msg: 'No PO found for "' + q + '".', data: null };

    function fmtDate(v) {
      if (v instanceof Date) {
        return ('0'+v.getDate()).slice(-2) + '/' + ('0'+(v.getMonth()+1)).slice(-2) + '/' + v.getFullYear();
      }
      return String(v || '');
    }

    const result = {};
    headers.forEach(function(h, i) {
      const v = found[i];
      result[h] = (v instanceof Date) ? fmtDate(v) : (v !== null && v !== undefined ? String(v) : '');
    });

    return {
      ok: true,
      data: result,
      rowIndex: rowIndex,
      orgs: getPOOrganizations(),
      activities: getPOActivities()
    };
  } catch(e) {
    return { ok: false, msg: e.message, data: null };
  }
}


// SERVER: save updated PO
function saveUpdatedPO(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName("Purchase Orders");
    if (!sh) return { ok: false, msg: 'Sheet "Purchase Orders" not found.' };

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });

    function fmtDate(v) {
      if (!v || v.indexOf('-') === -1) return v || '';
      const p = v.split('-');
      return p.length === 3 ? p[2]+'/'+p[1]+'/'+p[0] : v;
    }

    const amount          = parseFloat(payload.amount)          || 0;
    const recurring       = parseFloat(payload.recurringAmount) || 0;
    const recurringPeriod = parseInt(payload.recurringPeriod)   || 0;
    const hours           = parseFloat(payload.hours)           || 0;
    const pph             = parseFloat(payload.pricePerHour)    || 0;
    const bt              = payload.billingType || '';
    const exchangeRate    = parseFloat(payload.exchangeRate)    || 1;

    let totalInCurrency = 0;
    if      (bt === 'Fixed Project') totalInCurrency = amount;
    else if (bt === 'Recurring')     totalInCurrency = recurring * recurringPeriod;
    else if (bt === 'Hourly')        totalInCurrency = hours * pph;
    const total = totalInCurrency * exchangeRate;

    const fieldMap = {
      'PO Date':             fmtDate(payload.poDate),
      'Organization':        payload.organization,
      'Project':             payload.project,
      'Project description': payload.description,
      'PO number':           payload.poNumber,
      'Amount':              amount || '',
      'Recurring Amount':    recurring || '',
      'Milestones':          payload.milestones,
      'Hours':               hours || '',
      'Price per hour':      pph || '',
      'Total Amount':        total || '',
      'Customer':            payload.customer,
      'Proposal link':       payload.proposalLink,
      'Billing Type':        payload.billingType,
      'Renwal Date':         fmtDate(payload.renewalDate),
      'Commbox ARR':         payload.commboxARR || '',
      'Recurring Period':    recurringPeriod || '',
      'Currency':            payload.currency || 'NIS',
      'Exchange Rate':       exchangeRate
    };

    const row = sh.getRange(payload._rowIndex, 1, 1, lastCol).getValues()[0];
    headers.forEach(function(h, i) {
      const key = Object.keys(fieldMap).find(function(k) { return k.trim().toLowerCase() === h.trim().toLowerCase(); });
      if (key !== undefined) row[i] = fieldMap[key];
    });
    sh.getRange(payload._rowIndex, 1, 1, lastCol).setValues([row]);

    return { ok: true, msg: 'PO #' + payload.poNumber + ' updated successfully!' };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}


// DIALOG
function openUpdatePODialog() {
  const html = HtmlService.createHtmlOutput(UPDATE_PO_HTML)
    .setWidth(540)
    .setHeight(700)
    .setTitle('Update Purchase Order');
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Purchase Order');
}


const UPDATE_PO_HTML = `<!DOCTYPE html>
<html>
<head>
<style>
*,*::before,*::after{box-sizing:border-box}
body{font-family:'Google Sans',Arial,sans-serif;background:#f3f4ff;margin:0;padding:14px 18px 18px;font-size:13px;color:#1a237e}
h2{margin:0 0 3px;font-size:16px}
p.sub{margin:0 0 12px;color:#555;font-size:12px}

.sr{display:flex;gap:8px;margin-bottom:8px}
.sr input{flex:1;padding:8px 12px;font-size:13px;border:2px solid #9fa8da;border-radius:6px;outline:none;color:#212121}
.sr input:focus{border-color:#3949ab}
.sr button{padding:8px 18px;font-size:13px;font-weight:700;background:#1a237e;color:#fff;border:none;border-radius:6px;cursor:pointer;white-space:nowrap}
.sr button:hover{background:#3949ab}
.sr button:disabled{background:#9fa8da;cursor:default}

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

input[type=text],input[type=number],input[type=date],input[type=url],select,textarea{
  width:100%;padding:7px 9px;font-size:12px;border:1.5px solid #9fa8da;
  border-radius:5px;background:#fff;color:#212121;outline:none}
input:focus,select:focus,textarea:focus{border-color:#3949ab}
.ef{border-color:#c62828!important}
textarea{resize:vertical;min-height:44px;font-family:inherit}

.cur-badge{display:inline-block;padding:3px 10px;font-size:11px;font-weight:700;border-radius:4px;background:#e8eaf6;color:#3949ab;margin-bottom:6px}

.tb{background:#e8f5e9;border:1.5px solid #81c784;border-radius:5px;padding:8px 12px;margin-bottom:8px}
.tb-row{display:flex;justify-content:space-between;align-items:center}
.tl{font-weight:700;color:#2e7d32;font-size:12px}
.tv{font-weight:700;color:#2e7d32;font-size:15px}
.tc{font-size:10px;color:#5c6bc0;margin-top:2px}

.btn-upd{width:100%;padding:8px;font-size:13px;font-weight:700;background:#1a237e;
  color:#fff;border:none;border-radius:6px;cursor:pointer;margin-top:6px}
.btn-upd:hover{background:#3949ab}
.btn-upd:disabled{background:#9fa8da;cursor:default}
.cst{font-size:11px;min-height:14px;text-align:center;margin-top:4px}
.cst.ok{color:#2e7d32}.cst.err{color:#c62828}
</style>
</head>
<body>
<h2>Update Purchase Order</h2>
<p class="sub">Enter a PO number to load and edit the order.</p>

<div class="sr">
  <input type="text" id="qi" placeholder="Enter PO number..." />
  <button id="bs" onclick="go()">Find</button>
</div>
<div id="gst"></div>
<div id="res"></div>

<script>
var _po     = null;   // current PO data object
var _curCur = 'NIS';
var _curRate = 1;

document.getElementById('qi').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') go();
});

function gst(m, c) { var e = document.getElementById('gst'); e.innerText = m; e.className = c || ''; }

function esc(v) {
  return String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function toInp(v) {
  if (!v) return '';
  var p = String(v).split('/');
  return p.length === 3 ? p[2]+'-'+p[1]+'-'+p[0] : '';
}

function fmtNIS(v) {
  var n = parseFloat(v) || 0;
  return '\u20AA' + n.toLocaleString('he-IL', {minimumFractionDigits:2, maximumFractionDigits:2});
}

function buildSelect(id, arr, selected) {
  var s = '<select id="' + id + '"><option value="">-- Select --</option>';
  arr.forEach(function(v) {
    s += '<option value="' + esc(v) + '"' + (v === selected ? ' selected' : '') + '>' + esc(v) + '</option>';
  });
  return s + '</select>';
}

function fi(lbl, html, req, opt) {
  var mark = req ? '<span class="req">*</span>' : (opt ? '<span class="opt">(optional)</span>' : '');
  return '<div class="fi"><label>' + lbl + mark + '</label>' + html + '</div>';
}

// ── Search ───────────────────────────────────────────────────
function go() {
  var q = document.getElementById('qi').value.trim();
  document.getElementById('res').innerHTML = '';
  _po = null;
  if (!q) { gst('Please enter a PO number.', 'err'); return; }
  gst('Searching\u2026', '');
  document.getElementById('bs').disabled = true;

  google.script.run
    .withSuccessHandler(function(r) {
      document.getElementById('bs').disabled = false;
      if (!r.ok) { gst(r.msg, 'err'); return; }
      _po = r;
      _curCur  = r.data['Currency'] || 'NIS';
      _curRate = parseFloat(r.data['Exchange Rate']) || 1;
      gst('PO found. Edit the fields below and save.', 'ok');
      document.getElementById('res').innerHTML = buildCard(r.data, r.rowIndex, r.orgs || [], r.activities || []);
      calcTotal();
    })
    .withFailureHandler(function(e) {
      document.getElementById('bs').disabled = false;
      gst('Error: ' + e.message, 'err');
    })
    .findPOByNumber(q);
}

// ── Card builder ─────────────────────────────────────────────
function buildCard(o, rowIndex, orgs, acts) {
  var curBt   = o['Billing Type'] || '';
  var curRp   = o['Recurring Period'] || '';
  var sym = _curCur === 'USD' ? '$' : '\u20AA';

  var btSel = '<select id="f_billing" onchange="onBtChange()">'
    + '<option value="">-- Select --</option>'
    + ['Hourly','Fixed Project','Recurring'].map(function(v) {
        return '<option value="' + v + '"' + (v === curBt ? ' selected' : '') + '>' + v + '</option>';
      }).join('')
    + '</select>';

  var rpOpts = '<option value="">-- Select --</option>';
  for (var m = 1; m <= 12; m++) {
    rpOpts += '<option value="' + m + '"' + (String(m) === String(curRp) ? ' selected' : '') + '>' + m + '</option>';
  }

  var h = '<div class="card">';
  h += '<div class="ch">PO #' + esc(o['PO number']||'') + ' &mdash; ' + esc(o['Organization']||'') + '</div>';
  h += '<div class="cb">';

  h += '<div class="sec">Order Details</div>';
  h += '<div class="r2">';
  h += fi('PO Number', '<input type="text" id="f_poNum" value="' + esc(o['PO number']||'') + '" />', true);
  h += fi('PO Date',   '<input type="date" id="f_poDate" value="' + toInp(o['PO Date']||'') + '" />', false, true);
  h += '</div>';
  h += '<div class="r2">';
  h += fi('Organization', buildSelect('f_org',     orgs, o['Organization']||''), true);
  h += fi('Project',      buildSelect('f_project', acts, o['Project']||''),      true);
  h += '</div>';
  h += fi('Project Description', '<textarea id="f_desc">' + esc(o['Project description']||'') + '</textarea>', false, true);
  h += '<div class="r2">';
  h += fi('Customer',     '<input type="text" id="f_customer" value="' + esc(o['Customer']||'') + '" />', false, true);
  h += fi('Billing Type', btSel, true);
  h += '</div>';
  h += fi('Proposal Link', '<input type="url" id="f_proposal" value="' + esc(o['Proposal link']||'') + '" />', false, true);

  h += '<div class="sec">Financial Details</div>';
  h += '<div class="cur-badge">' + esc(_curCur) + (_curCur === 'USD' && _curRate > 1 ? ' &mdash; 1 USD = ' + _curRate.toFixed(4) + ' NIS' : '') + '</div>';
  h += '<div class="r2">';
  h += fi('Amount (' + sym + ')',    '<input type="number" id="f_amount"    value="' + esc(o['Amount']||'')           + '" min="0" step="0.01" oninput="calcTotal()" />', false, true);
  h += fi('Recurring (' + sym + ')', '<input type="number" id="f_recurring" value="' + esc(o['Recurring Amount']||'') + '" min="0" step="0.01" oninput="calcTotal()" />', false, true);
  h += '</div>';
  h += '<div id="f_recPeriodRow" style="display:' + (curBt === 'Recurring' ? 'block' : 'none') + '">';
  h += fi('Recurring Period (months)', '<select id="f_recPeriod" onchange="calcTotal()">' + rpOpts + '</select>', false, true);
  h += '</div>';
  h += '<div class="r2">';
  h += fi('Hours',                                                '<input type="number" id="f_hours" value="' + esc(o['Hours']||'')          + '" min="0" step="0.5"  oninput="calcTotal()" />', false, true);
  h += fi('Price/Hour (' + sym + ')',      '<input type="number" id="f_pph"   value="' + esc(o['Price per hour']||'') + '" min="0" step="0.01" oninput="calcTotal()" />', false, true);
  h += '</div>';
  h += fi('Milestones',   '<input type="text" id="f_milestones" value="' + esc(o['Milestones']||'')   + '" />', false, true);
  h += fi('Renewal Date', '<input type="date" id="f_renewal"    value="' + toInp(o['Renwal Date']||'') + '" />', false, true);
  h += fi('Commbox ARR (' + sym + ')', '<input type="number" id="f_commboxARR" value="' + esc(o['Commbox ARR']||'') + '" min="0" step="0.01" />', false, true);

  h += '<div class="tb"><div class="tb-row"><span class="tl">Total Amount (NIS)</span><span class="tv" id="f_total">\u20AA0.00</span></div>';
  h += '<div class="tc" id="f_conv"></div></div>';

  h += '<button class="btn-upd" id="btnSave" onclick="doSave()">Save Changes</button>';
  h += '<div class="cst" id="cst"></div>';
  h += '</div></div>';
  return h;
}

// ── Billing type toggle ───────────────────────────────────────
function onBtChange() {
  var bt = document.getElementById('f_billing').value;
  var rp = document.getElementById('f_recPeriodRow');
  if (rp) rp.style.display = (bt === 'Recurring') ? 'block' : 'none';
  calcTotal();
}

// ── Total calculation ─────────────────────────────────────────
function calcTotal() {
  var bEl = document.getElementById('f_billing'); if (!bEl) return;
  var bt  = bEl.value;
  var amt = parseFloat((document.getElementById('f_amount')    || {value:'0'}).value) || 0;
  var rec = parseFloat((document.getElementById('f_recurring') || {value:'0'}).value) || 0;
  var rp  = parseInt( (document.getElementById('f_recPeriod') || {value:'0'}).value)  || 0;
  var hrs = parseFloat((document.getElementById('f_hours')     || {value:'0'}).value) || 0;
  var pph = parseFloat((document.getElementById('f_pph')       || {value:'0'}).value) || 0;
  var totalInCur = 0;
  if      (bt === 'Fixed Project') totalInCur = amt;
  else if (bt === 'Recurring')     totalInCur = rec * rp;
  else if (bt === 'Hourly')        totalInCur = hrs * pph;
  var totalNIS = totalInCur * _curRate;
  var totEl = document.getElementById('f_total'); if (totEl) totEl.innerText = fmtNIS(totalNIS);
  var convEl = document.getElementById('f_conv');
  if (convEl) {
    if (_curCur === 'USD' && totalInCur > 0 && _curRate > 1) {
      convEl.innerText = '(from $' + totalInCur.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}) + ' at ' + _curRate.toFixed(4) + ')';
    } else { convEl.innerText = ''; }
  }
}


// ── Save ─────────────────────────────────────────────────────
function doSave() {
  var st = document.getElementById('cst');
  st.innerText = ''; st.className = 'cst';

  var poNum   = document.getElementById('f_poNum').value.trim();
  var org     = document.getElementById('f_org').value;
  var project = document.getElementById('f_project').value;
  var billing = document.getElementById('f_billing').value;

  ['f_poNum','f_org','f_project','f_billing'].forEach(function(id) {
    document.getElementById(id).classList.remove('ef');
  });

  var ok = true;
  function err(id) { document.getElementById(id).classList.add('ef'); ok = false; }
  if (!poNum)   err('f_poNum');
  if (!org)     err('f_org');
  if (!project) err('f_project');
  if (!billing) err('f_billing');
  if (!ok) { st.innerText = 'Please fill in all required fields.'; st.className = 'cst err'; return; }

  st.innerText = 'Saving\u2026'; st.className = 'cst';
  document.getElementById('btnSave').disabled = true;

  var recPeriodEl = document.getElementById('f_recPeriod');

  google.script.run
    .withSuccessHandler(function(r) {
      document.getElementById('btnSave').disabled = false;
      st.innerText = r.ok ? '\u2705 ' + r.msg : '\u274c ' + r.msg;
      st.className = 'cst ' + (r.ok ? 'ok' : 'err');
    })
    .withFailureHandler(function(e) {
      document.getElementById('btnSave').disabled = false;
      st.innerText = '\u274c ' + e.message; st.className = 'cst err';
    })
    .saveUpdatedPO({
      _rowIndex:       _po.rowIndex,
      poNumber:        poNum,
      organization:    org,
      project:         project,
      description:     document.getElementById('f_desc').value,
      poDate:          document.getElementById('f_poDate').value,
      customer:        document.getElementById('f_customer').value.trim(),
      billingType:     billing,
      proposalLink:    document.getElementById('f_proposal').value.trim(),
      amount:          document.getElementById('f_amount').value,
      recurringAmount: document.getElementById('f_recurring').value,
      hours:           document.getElementById('f_hours').value,
      pricePerHour:    document.getElementById('f_pph').value,
      milestones:      document.getElementById('f_milestones').value.trim(),
      renewalDate:     document.getElementById('f_renewal').value,
      commboxARR:      document.getElementById('f_commboxARR').value,
      recurringPeriod: recPeriodEl ? recPeriodEl.value : '',
      currency:        _curCur,
      exchangeRate:    _curCur === 'USD' ? _curRate : 1
    });
}
</script>
</body>
</html>`;
