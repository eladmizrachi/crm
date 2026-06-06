// ============================================================
// CRM_SearchPO.gs — Search Purchase Orders
//   Search only (by PO Number OR Customer / Org)
// ============================================================

const PO_SHEET = "Purchase Orders";

// ── SERVER: search POs ───────────────────────────────────────
function searchPurchaseOrders(mode, term) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(PO_SHEET);
    if (!sh) return { ok: false, msg: 'Sheet "Purchase Orders" not found.' };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: false, msg: 'No data found in the sheet.' };

    const headers = data[0].map(function(h) { return String(h).trim(); });
    const q = String(term).trim().toLowerCase();
    if (!q) return { ok: false, msg: 'Please enter a search term.' };

    function col(name) {
      return headers.findIndex(function(h) {
        return h.toLowerCase().replace(/\s+/g, '') === name;
      });
    }
    const iOrg   = col('organization');
    const iCust  = col('customer');
    const iPO    = col('ponumber');
    const iProj  = col('project');
    const iDate  = col('podate');
    const iTotal = col('totalamount');
    const iBT    = col('billingtype');
    const iCur   = col('currency');

    function fmtDate(v) {
      if (v instanceof Date) {
        return ('0'+v.getDate()).slice(-2) + '/' + ('0'+(v.getMonth()+1)).slice(-2) + '/' + v.getFullYear();
      }
      return String(v || '');
    }

    const matches = [];
    data.slice(1).forEach(function(row, i) {
      var match = false;
      if (mode === 'po') {
        match = String(row[iPO] || '').trim().toLowerCase() === q;
      } else {
        var org  = String(row[iOrg]  || '').toLowerCase();
        var cust = String(row[iCust] || '').toLowerCase();
        match = org.indexOf(q) !== -1 || cust.indexOf(q) !== -1;
      }
      if (match) {
        matches.push({
          rowIndex:     i + 2,
          poNumber:     String(row[iPO]   || ''),
          organization: String(row[iOrg]  || ''),
          customer:     String(row[iCust] || ''),
          project:      String(row[iProj] || ''),
          poDate:       fmtDate(row[iDate]),
          totalAmount:  String(row[iTotal]|| ''),
          billingType:  String(row[iBT]   || ''),
          currency:     String(row[iCur]  || 'NIS'),
        });
      }
    });

    if (!matches.length) return { ok: false, msg: 'No orders found for "' + term + '".' };
    return { ok: true, matches: matches };

  } catch(e) {
    return { ok: false, msg: 'Error: ' + e.message };
  }
}


// ── DIALOG ───────────────────────────────────────────────────
function openSearchPODialog() {
  const html = HtmlService.createHtmlOutput(SEARCH_PO_HTML)
    .setWidth(540)
    .setHeight(600)
    .setTitle('Search Purchase Orders');
  SpreadsheetApp.getUi().showModalDialog(html, 'Search Purchase Orders');
}


// ── HTML ─────────────────────────────────────────────────────
const SEARCH_PO_HTML = `<!DOCTYPE html>
<html>
<head>
<style>
*,*::before,*::after { box-sizing: border-box; }
body {
  font-family: 'Google Sans', Arial, sans-serif;
  background: #f3f4ff;
  margin: 0;
  padding: 16px 20px 20px;
  font-size: 13px;
  color: #1a237e;
}
h2 { margin: 0 0 4px; font-size: 17px; }
p.sub { margin: 0 0 14px; color: #666; font-size: 12px; }

/* Mode toggle */
.mode-row {
  display: flex;
  border: 2px solid #9fa8da;
  border-radius: 7px;
  overflow: hidden;
  margin-bottom: 12px;
}
.mode-btn {
  flex: 1;
  padding: 9px;
  font-size: 12px;
  font-weight: 700;
  border: none;
  cursor: pointer;
  background: #fff;
  color: #555;
  transition: all .15s;
}
.mode-btn.active { background: #1a237e; color: #fff; }

/* Search bar */
.sr { display: flex; gap: 8px; margin-bottom: 8px; }
.sr input {
  flex: 1;
  padding: 9px 13px;
  font-size: 14px;
  border: 2px solid #9fa8da;
  border-radius: 6px;
  outline: none;
  color: #212121;
}
.sr input:focus { border-color: #3949ab; }
.sr button {
  padding: 9px 20px;
  font-size: 13px;
  font-weight: 700;
  background: #1a237e;
  color: #fff;
  border: none;
  border-radius: 6px;
  cursor: pointer;
}
.sr button:hover { background: #3949ab; }
.sr button:disabled { background: #9fa8da; cursor: default; }

/* Status */
#status {
  font-size: 12px;
  min-height: 18px;
  text-align: center;
  margin-bottom: 10px;
}
#status.ok  { color: #2e7d32; }
#status.err { color: #c62828; }

/* Results table */
.tbl { width: 100%; border-collapse: collapse; background: #fff; border-radius: 7px; overflow: hidden; border: 1.5px solid #c5cae9; }
.tbl thead tr { background: #1a237e; color: #fff; }
.tbl th { padding: 9px 10px; text-align: left; font-size: 11px; font-weight: 700; letter-spacing: .04em; }
.tbl tbody tr { border-bottom: 1px solid #e8eaf6; transition: background .12s; }
.tbl tbody tr:last-child { border-bottom: none; }
.tbl tbody tr:hover { background: #f0f2ff; }
.tbl td { padding: 8px 10px; font-size: 12px; color: #212121; }
.tbl td.po  { font-weight: 700; color: #1a237e; }
.tbl td.amt { font-weight: 700; color: #2e7d32; text-align: right; }
.tbl td.bt  { font-size: 11px; color: #666; }
.badge {
  display: inline-block;
  font-size: 10px;
  font-weight: 700;
  padding: 2px 7px;
  border-radius: 10px;
  background: #e8eaf6;
  color: #3949ab;
}
.badge.usd { background: #e8f5e9; color: #2e7d32; }
</style>
</head>
<body>

<h2>🔎 Search Purchase Orders</h2>
<p class="sub">Find orders by PO number or by customer / organization.</p>

<div class="mode-row">
  <button class="mode-btn active" id="mPO"   onclick="setMode('po')">By PO Number</button>
  <button class="mode-btn"        id="mCust" onclick="setMode('customer')">By Customer / Org</button>
</div>

<div class="sr">
  <input type="text" id="srch" placeholder="Enter PO number…" />
  <button id="btn" onclick="go()">Search</button>
</div>

<div id="status"></div>
<div id="results"></div>

<script>
var _mode = 'po';

function setMode(m) {
  _mode = m;
  document.getElementById('mPO').className   = 'mode-btn' + (m === 'po'       ? ' active' : '');
  document.getElementById('mCust').className = 'mode-btn' + (m === 'customer' ? ' active' : '');
  document.getElementById('srch').placeholder = m === 'po'
    ? 'Enter PO number\u2026'
    : 'Enter customer or organization name\u2026';
  document.getElementById('srch').value = '';
  setStatus('', '');
  document.getElementById('results').innerHTML = '';
  document.getElementById('srch').focus();
}

document.getElementById('srch').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') go();
});

function setStatus(msg, cls) {
  var el = document.getElementById('status');
  el.innerText = msg;
  el.className = cls || '';
}

function fmtNIS(v) {
  var n = parseFloat(v) || 0;
  if (n === 0) return '';
  return '\u20AA' + n.toLocaleString('he-IL', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function esc(v) {
  return String(v)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}

function go() {
  var q = document.getElementById('srch').value.trim();
  if (!q) { setStatus('Please enter a search term.', 'err'); return; }

  document.getElementById('results').innerHTML = '';
  setStatus('Searching\u2026', '');
  document.getElementById('btn').disabled = true;

  google.script.run
    .withSuccessHandler(function(r) {
      document.getElementById('btn').disabled = false;
      if (!r.ok) { setStatus(r.msg, 'err'); return; }
      setStatus(r.matches.length + ' result(s) found.', 'ok');
      renderTable(r.matches);
    })
    .withFailureHandler(function(e) {
      document.getElementById('btn').disabled = false;
      setStatus('Error: ' + (e.message || e), 'err');
    })
    .searchPurchaseOrders(_mode, q);
}

function renderTable(rows) {
  var h = '<table class="tbl">';
  h += '<thead><tr>';
  h += '<th>PO #</th><th>Organization</th><th>Project</th><th>Date</th><th>Type</th><th style="text-align:right">Total</th>';
  h += '</tr></thead><tbody>';

  rows.forEach(function(m) {
    h += '<tr>';
    h += '<td class="po">' + esc(m.poNumber) + '</td>';
    h += '<td>' + esc(m.organization) + (m.customer && m.customer !== m.organization ? '<br><span style="color:#888;font-size:11px">' + esc(m.customer) + '</span>' : '') + '</td>';
    h += '<td>' + esc(m.project) + '</td>';
    h += '<td>' + esc(m.poDate) + '</td>';
    h += '<td class="bt"><span class="badge">' + esc(m.billingType || '\u2014') + '</span></td>';
    var amtStr = fmtNIS(m.totalAmount);
    h += '<td class="amt">' + (amtStr ? esc(amtStr) : '\u2014') + (m.currency === 'USD' ? ' <span class="badge usd">USD</span>' : '') + '</td>';
    h += '</tr>';
  });

  h += '</tbody></table>';
  document.getElementById('results').innerHTML = h;
}
</script>
</body>
</html>`;
