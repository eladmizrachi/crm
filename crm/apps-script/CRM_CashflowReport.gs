// ============================================================
// CRM_CashflowReport.gs — Expected Cashflow Report
// Shows projected monthly income based on billing records + payment terms.
// ============================================================

function openCashflowReportDialog() {
  const now        = new Date();
  const initYear   = now.getFullYear();
  const initMonth  = now.getMonth() + 1;

  // Default range: current month → 6 months ahead
  const toDate     = new Date(now.getFullYear(), now.getMonth() + 6, 1);
  const toYear     = toDate.getFullYear();
  const toMonth    = toDate.getMonth() + 1;

  const html = HtmlService.createHtmlOutput(
    buildCashflowHtml_(initYear, initMonth, toYear, toMonth)
  )
    .setTitle('Cashflow Report')
    .setWidth(960)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Expected Cashflow Report');
}


// ── Server-side data fetch ─────────────────────────────────

function getCashflowData(fromYear, fromMonth, toYear, toMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(BILLING_SHEET_NAME);
    if (!sh) return { ok: true, months: [], grandTotal: 0 };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, months: [], grandTotal: 0 };

    const hdrs = data[0].map(function(h) { return String(h).trim(); });
    function findCol(name) {
      const lc = name.toLowerCase();
      return hdrs.findIndex(function(h) { return h.toLowerCase() === lc; });
    }

    const colMonth        = findCol('Month');
    const colPaying       = findCol('Paying Customer');
    const colAmt          = findCol('Amount');
    const colPT           = findCol('Payment terms');
    const colCust         = findCol('Customer');
    const colProj         = findCol('Project type');
    const colBillingDesc  = findCol('Billing description');
    const colPO           = findCol('PO Number');

    const fromKey = String(fromYear) + '-' + String(fromMonth).padStart(2, '0');
    const toKey   = String(toYear)   + '-' + String(toMonth).padStart(2, '0');

    // Map: expectedMonthKey → { payingCustomer → { total, records[] } }
    const monthMap = {};

    data.slice(1).forEach(function(row) {
      const rawMonth = row[colMonth];
      const billingMonth = normalizeCashflowMonth_(rawMonth);
      if (!billingMonth) return;

      const amount       = parseFloat(row[colAmt]) || 0;
      if (amount === 0) return;

      const paymentTerms = parseInt(row[colPT]) || 0;
      const payingCust   = colPaying !== -1 ? String(row[colPaying] || '') : '';
      const customer     = colCust   !== -1 ? String(row[colCust]   || '') : '';
      const project      = colProj   !== -1 ? String(row[colProj]   || '') : '';
      const billingDesc  = colBillingDesc !== -1 ? String(row[colBillingDesc] || '') : '';
      const poNumber     = colPO     !== -1 ? String(row[colPO]     || '') : '';

      // Expected payment date: 1st of billing month + paymentTerms days
      const parts = billingMonth.split('-');
      const baseDate = new Date(+parts[0], +parts[1] - 1, 1);
      baseDate.setDate(baseDate.getDate() + paymentTerms);

      const expYear  = baseDate.getFullYear();
      const expMonth = baseDate.getMonth() + 1;
      const expKey   = String(expYear) + '-' + String(expMonth).padStart(2, '0');

      // Filter to requested range
      if (expKey < fromKey || expKey > toKey) return;

      if (!monthMap[expKey]) monthMap[expKey] = {};
      const custKey = payingCust || customer || '(Unknown)';
      if (!monthMap[expKey][custKey]) monthMap[expKey][custKey] = { total: 0, records: [] };

      monthMap[expKey][custKey].total += amount;
      monthMap[expKey][custKey].records.push({
        customer:    customer,
        paying:      payingCust,
        amount:      amount,
        project:     project,
        billingDesc: billingDesc,
        poNumber:    poNumber,
        billingMonth: billingMonth,
        paymentTerms: paymentTerms,
      });
    });

    // Sort months and build response
    const sortedKeys = Object.keys(monthMap).sort();
    let grandTotal = 0;

    const months = sortedKeys.map(function(key) {
      const custMap   = monthMap[key];
      const custKeys  = Object.keys(custMap).sort();
      let monthTotal  = 0;

      const customers = custKeys.map(function(ck) {
        monthTotal += custMap[ck].total;
        return {
          name:    ck,
          total:   Math.round(custMap[ck].total * 100) / 100,
          records: custMap[ck].records,
        };
      });

      monthTotal = Math.round(monthTotal * 100) / 100;
      grandTotal += monthTotal;

      return {
        key:      key,
        label:    monthLabel_(key),
        total:    monthTotal,
        customers: customers,
      };
    });

    return { ok: true, months: months, grandTotal: Math.round(grandTotal * 100) / 100 };
  } catch(e) {
    return { ok: false, error: e.message, months: [], grandTotal: 0 };
  }
}


// ── Helpers ────────────────────────────────────────────────

function normalizeCashflowMonth_(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth() + 1).padStart(2, '0');
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{1,2})$/);
  if (m) return m[1] + '-' + m[2].padStart(2, '0');
  return '';
}

function monthLabel_(key) {
  const parts = key.split('-');
  const names = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
  return names[+parts[1] - 1] + ' ' + parts[0];
}


// ── HTML builder ───────────────────────────────────────────

function buildCashflowHtml_(fromYear, fromMonth, toYear, toMonth) {
  const years = (function() {
    let opts = '';
    for (let y = 2024; y <= fromYear + 2; y++) {
      opts += '<option value="' + y + '">' + y + '</option>';
    }
    return opts;
  })();

  const monthNames = ['January','February','March','April','May','June',
                      'July','August','September','October','November','December'];

  function monthOpts(selectedVal) {
    return monthNames.map(function(n, i) {
      const v = i + 1;
      return '<option value="' + v + '"' + (v === selectedVal ? ' selected' : '') + '>' + n + '</option>';
    }).join('');
  }

  function yearOpts(selectedVal) {
    let opts = '';
    for (let y = 2024; y <= fromYear + 2; y++) {
      opts += '<option value="' + y + '"' + (y === selectedVal ? ' selected' : '') + '>' + y + '</option>';
    }
    return opts;
  }

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f7fa; color: #333; }

  .header {
    background: #1b5e20; color: #fff;
    padding: 16px 20px; font-size: 17px; font-weight: bold;
    display: flex; align-items: center; gap: 10px;
  }

  .controls {
    display: flex; align-items: center; gap: 10px;
    padding: 12px 20px; background: #fff;
    border-bottom: 1px solid #e0e0e0; flex-wrap: wrap;
  }
  .controls label { font-weight: 600; font-size: 13px; color: #555; }
  .controls select {
    padding: 6px 10px; border: 1px solid #bbb; border-radius: 6px;
    font-size: 13px; background: #fff; cursor: pointer;
  }
  .controls select:focus { outline: none; border-color: #2e7d32; }
  .btn-load {
    padding: 8px 20px; background: #2e7d32; color: #fff;
    border: none; border-radius: 6px; font-size: 13px;
    font-weight: 600; cursor: pointer;
  }
  .btn-load:hover { background: #1b5e20; }
  .range-sep { color: #888; font-size: 13px; }

  /* Grand total bar */
  .grand-bar {
    display: flex; gap: 16px; padding: 10px 20px;
    background: #e8f5e9; border-bottom: 1px solid #a5d6a7;
    font-size: 13px; flex-wrap: wrap; align-items: center;
  }
  .grand-chip {
    background: #fff; border: 1px solid #a5d6a7;
    border-radius: 20px; padding: 4px 14px;
    font-weight: 600; color: #2e7d32;
  }
  .grand-chip span { color: #1b5e20; }
  .grand-total-chip {
    background: #1b5e20 !important; color: #fff !important;
    border-color: #1b5e20 !important;
  }
  .grand-total-chip span { color: #b9f6ca !important; }

  .content { padding: 16px 20px; }

  /* Month card */
  .month-card {
    background: #fff; border: 1px solid #c8e6c9;
    border-radius: 8px; margin-bottom: 16px; overflow: hidden;
  }
  .month-header {
    background: #2e7d32; color: #fff;
    padding: 10px 16px;
    display: flex; justify-content: space-between; align-items: center;
  }
  .month-header .month-name { font-size: 15px; font-weight: 700; }
  .month-header .month-total { font-size: 15px; font-weight: 700; color: #b9f6ca; }

  /* Customer section inside a month */
  .cust-section { border-top: 1px solid #e8f5e9; }
  .cust-header {
    background: #f1f8e9;
    padding: 7px 16px;
    display: flex; justify-content: space-between; align-items: center;
    cursor: pointer; user-select: none;
  }
  .cust-header:hover { background: #dcedc8; }
  .cust-name { font-weight: 700; font-size: 13px; color: #33691e; }
  .cust-total { font-weight: 700; font-size: 13px; color: #2e7d32; }
  .cust-toggle { font-size: 11px; color: #66bb6a; margin-left: 8px; }

  /* Records table inside each customer */
  .records-wrap { display: none; overflow-x: auto; padding: 0 16px 10px; }
  .records-wrap.open { display: block; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 8px; }
  thead tr { background: #388e3c; color: #fff; }
  thead th { padding: 7px 10px; text-align: left; font-weight: 600; white-space: nowrap; }
  tbody tr:nth-child(even) { background: #f9fbe7; }
  tbody tr:hover { background: #f1f8e9; }
  tbody td { padding: 7px 10px; border-bottom: 1px solid #e8f5e9; vertical-align: top; }
  td.amount { text-align: right; font-weight: 600; color: #2e7d32; white-space: nowrap; }

  .empty { padding: 40px; text-align: center; color: #9e9e9e; font-size: 14px; }
  .error-msg { padding: 16px 20px; color: #c62828; font-size: 13px; }
  .loading { padding: 40px; text-align: center; color: #388e3c; font-size: 14px; }
</style>
</head>
<body>

<div class="header">💰 Expected Cashflow Report</div>

<div class="controls">
  <label>From:</label>
  <select id="fromMonth">${monthOpts(fromMonth)}</select>
  <select id="fromYear">${yearOpts(fromYear)}</select>
  <span class="range-sep">→</span>
  <label>To:</label>
  <select id="toMonth">${monthOpts(toMonth)}</select>
  <select id="toYear">${yearOpts(toYear)}</select>
  <button class="btn-load" onclick="loadReport()">Load Report</button>
</div>

<div id="grandBar" class="grand-bar" style="display:none"></div>
<div id="content"><div class="loading">Select a date range and click Load Report.</div></div>

<script>
  function loadReport() {
    var fromMonth = parseInt(document.getElementById('fromMonth').value, 10);
    var fromYear  = parseInt(document.getElementById('fromYear').value, 10);
    var toMonth   = parseInt(document.getElementById('toMonth').value, 10);
    var toYear    = parseInt(document.getElementById('toYear').value, 10);

    document.getElementById('grandBar').style.display = 'none';
    document.getElementById('content').innerHTML = '<div class="loading">Loading…</div>';

    google.script.run
      .withSuccessHandler(renderReport)
      .withFailureHandler(function(err) {
        document.getElementById('content').innerHTML =
          '<div class="error-msg">⚠ Error: ' + err.message + '</div>';
      })
      .getCashflowData(fromYear, fromMonth, toYear, toMonth);
  }

  function renderReport(data) {
    if (!data.ok) {
      document.getElementById('content').innerHTML =
        '<div class="error-msg">⚠ ' + data.error + '</div>';
      return;
    }

    var months = data.months;

    if (months.length === 0) {
      document.getElementById('grandBar').style.display = 'none';
      document.getElementById('content').innerHTML =
        '<div class="empty">No expected payments found for this period.</div>';
      return;
    }

    // Grand total bar
    var barHtml = '<div class="grand-chip">Months with income: <span>' + months.length + '</span></div>';
    barHtml += '<div class="grand-chip total grand-total-chip">Total expected: <span>' + fmt(data.grandTotal) + '</span></div>';
    document.getElementById('grandBar').innerHTML = barHtml;
    document.getElementById('grandBar').style.display = 'flex';

    // Month cards
    var html = '';
    months.forEach(function(mo) {
      html += '<div class="month-card">' +
        '<div class="month-header">' +
          '<span class="month-name">' + esc(mo.label) + '</span>' +
          '<span class="month-total">₪ ' + fmt(mo.total) + '</span>' +
        '</div>';

      mo.customers.forEach(function(cust, ci) {
        var uid = mo.key.replace('-','') + '_' + ci;
        html += '<div class="cust-section">' +
          '<div class="cust-header" onclick="toggleRecords(\'' + uid + '\')">' +
            '<span class="cust-name">' + esc(cust.name) + '</span>' +
            '<span>' +
              '<span class="cust-total">₪ ' + fmt(cust.total) + '</span>' +
              '<span class="cust-toggle" id="tog_' + uid + '">▼ show</span>' +
            '</span>' +
          '</div>' +
          '<div class="records-wrap" id="rec_' + uid + '">' +
            '<table><thead><tr>' +
              '<th>PO #</th><th>Customer</th><th>Project type</th>' +
              '<th>Billing description</th><th>Billing month</th><th>Payment terms</th><th>Amount (₪)</th>' +
            '</tr></thead><tbody>';

        cust.records.forEach(function(r) {
          html += '<tr>' +
            '<td>' + esc(r.poNumber) + '</td>' +
            '<td>' + esc(r.customer) + '</td>' +
            '<td>' + esc(r.project) + '</td>' +
            '<td>' + esc(r.billingDesc) + '</td>' +
            '<td>' + esc(r.billingMonth) + '</td>' +
            '<td>' + (r.paymentTerms ? r.paymentTerms + ' days' : '') + '</td>' +
            '<td class="amount">₪ ' + fmt(r.amount) + '</td>' +
          '</tr>';
        });

        html += '</tbody></table></div></div>';
      });

      html += '</div>'; // month-card
    });

    document.getElementById('content').innerHTML = html;
  }

  function toggleRecords(uid) {
    var wrap = document.getElementById('rec_' + uid);
    var tog  = document.getElementById('tog_' + uid);
    if (!wrap) return;
    var open = wrap.classList.toggle('open');
    tog.innerHTML = open ? '▲ hide' : '▼ show';
  }

  function fmt(n) {
    return Number(n).toLocaleString('he-IL', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }
  function esc(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  // Auto-load on open
  loadReport();
</script>
</body>
</html>`;
}
