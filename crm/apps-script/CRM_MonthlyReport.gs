// ============================================================
// CRM_MonthlyReport.gs — Monthly Billing Closing Report
// ============================================================

function openMonthlyReportDialog() {
  const now       = new Date();
  const initYear  = now.getFullYear();
  const initMonth = now.getMonth() + 1; // 1-based

  const html = HtmlService.createHtmlOutput(buildMonthlyReportHtml_(initYear, initMonth))
    .setTitle('Monthly Billing Report')
    .setWidth(900)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'Monthly Billing Report');
}


// ── Month cell normalizer (handles Date objects and string variants) ──

function normalizeMonthCell_(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth() + 1).padStart(2, '0');
  }
  const s = String(val).trim();
  // already YYYY-MM
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  // YYYY-M (no zero-pad)
  const m = s.match(/^(\d{4})-(\d{1,2})$/);
  if (m) return m[1] + '-' + m[2].padStart(2, '0');
  return s;
}


// ── Server-side data fetch ────────────────────────────────

function getMonthlyBillingRecords(year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(BILLING_SHEET_NAME);
    if (!sh) return { ok: true, rows: [], total: 0 };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, rows: [], total: 0 };

    const hdrs      = data[0].map(function(h) { return String(h).trim(); });
    const colYear   = hdrs.indexOf('Year');
    const colMonth  = hdrs.indexOf('Month');
    const colCust   = hdrs.indexOf('Customer');
    const colType   = hdrs.indexOf('Billing Type');
    const colPaying = hdrs.indexOf('Paying Customer');
    const colAmt    = hdrs.indexOf('Amount');
    const colMsName = hdrs.indexOf('Milestone name');
    const colMsDesc = hdrs.indexOf('Milestone description');
    const colPO          = hdrs.indexOf('PO Number');
    const colPT          = hdrs.indexOf('Payment terms');
    const colProj        = hdrs.indexOf('Project type');
    const colHoursReport  = hdrs.indexOf('Hours report');
    const colInvoiceType  = hdrs.indexOf('Initiate invoice type');
    const colBillingPeriod = hdrs.indexOf('Billing period');
    const colBillingDesc   = hdrs.indexOf('Billing description');

    const targetMonth = String(year) + '-' + String(month).padStart(2, '0');

    const rows = [];
    let total  = 0;

    data.slice(1).forEach(function(row) {
      const rowMonth = normalizeMonthCell_(row[colMonth]);
      if (rowMonth !== targetMonth) return;

      const amount = parseFloat(row[colAmt]) || 0;
      total += amount;
      rows.push({
        year:         String(row[colYear]   || ''),
        month:        rowMonth,
        customer:     String(row[colCust]   || ''),
        type:         String(row[colType]   || ''),
        paying:       colPaying  !== -1 ? String(row[colPaying]  || '') : '',
        amount:       amount,
        msName:       colMsName  !== -1 ? String(row[colMsName]  || '') : '',
        msDesc:       colMsDesc  !== -1 ? String(row[colMsDesc]  || '') : '',
        poNumber:     colPO          !== -1 ? String(row[colPO]          || '') : '',
        paymentTerms: colPT          !== -1 ? String(row[colPT]          || '') : '',
        project:      colProj        !== -1 ? String(row[colProj]        || '') : '',
        hoursReport:    colHoursReport  !== -1 ? String(row[colHoursReport]  || '') : '',
        invoiceType:    colInvoiceType  !== -1 ? String(row[colInvoiceType]  || '') : '',
        billingPeriod:  colBillingPeriod !== -1 ? String(row[colBillingPeriod] || '') : '',
        billingDesc:    colBillingDesc   !== -1 ? String(row[colBillingDesc]   || '') : '',
      });
    });

    return { ok: true, rows: rows, total: total };
  } catch(e) {
    return { ok: false, error: e.message, rows: [], total: 0 };
  }
}


// ── HTML builder ──────────────────────────────────────────

function buildMonthlyReportHtml_(initYear, initMonth) {
  const yearOptions = (function() {
    let opts = '';
    for (let y = 2022; y <= initYear + 1; y++) {
      opts += '<option value="' + y + '"' + (y === initYear ? ' selected' : '') + '>' + y + '</option>';
    }
    return opts;
  })();

  const monthOptions = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December'
  ].map(function(name, i) {
    const val = i + 1;
    return '<option value="' + val + '"' + (val === initMonth ? ' selected' : '') + '>' + name + '</option>';
  }).join('');

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f7fa; color: #333; }

  .header {
    background: #1a237e; color: #fff;
    padding: 16px 20px; font-size: 17px; font-weight: bold;
    display: flex; align-items: center; gap: 10px;
  }

  .controls {
    display: flex; align-items: center; gap: 12px;
    padding: 14px 20px; background: #fff;
    border-bottom: 1px solid #e0e0e0; flex-wrap: wrap;
  }
  .controls label { font-weight: 600; font-size: 13px; color: #555; }
  .controls select {
    padding: 7px 10px; border: 1px solid #bbb; border-radius: 6px;
    font-size: 14px; background: #fff; cursor: pointer;
  }
  .controls select:focus { outline: none; border-color: #3949ab; }
  .btn-load {
    padding: 8px 20px; background: #3949ab; color: #fff;
    border: none; border-radius: 6px; font-size: 14px;
    font-weight: 600; cursor: pointer;
  }
  .btn-load:hover { background: #1a237e; }

  .summary-bar {
    display: flex; gap: 20px; padding: 12px 20px;
    background: #e8eaf6; border-bottom: 1px solid #c5cae9;
    font-size: 13px; flex-wrap: wrap;
  }
  .summary-bar .chip {
    background: #fff; border: 1px solid #9fa8da;
    border-radius: 20px; padding: 4px 14px;
    font-weight: 600; color: #3949ab;
  }
  .summary-bar .chip span { color: #1a237e; }
  .total-chip {
    background: #1a237e !important; color: #fff !important;
    border-color: #1a237e !important;
  }
  .total-chip span { color: #ffd740 !important; }

  .table-wrap { padding: 16px 20px; overflow-x: auto; }

  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  thead tr { background: #3949ab; color: #fff; }
  thead th { padding: 10px 12px; text-align: left; font-weight: 600; white-space: nowrap; }
  tbody tr:nth-child(even) { background: #f5f5ff; }
  tbody tr:hover { background: #e8eaf6; }
  tbody td { padding: 9px 12px; border-bottom: 1px solid #e8eaf6; vertical-align: top; }
  td.amount { text-align: right; font-weight: 600; color: #1a237e; white-space: nowrap; }
  td.type { white-space: nowrap; }

  .badge {
    display: inline-block; padding: 2px 8px; border-radius: 12px;
    font-size: 11px; font-weight: 700; text-transform: uppercase;
  }
  .badge-recurring  { background: #e3f2fd; color: #1565c0; }
  .badge-upfront    { background: #e8f5e9; color: #2e7d32; }
  .badge-milestones { background: #fff3e0; color: #e65100; }

  .empty { padding: 40px 20px; text-align: center; color: #9e9e9e; font-size: 14px; }
  .error-msg { padding: 16px 20px; color: #c62828; font-size: 13px; }
  .loading { padding: 40px 20px; text-align: center; color: #5c6bc0; font-size: 14px; }
</style>
</head>
<body>

<div class="header">📊 Monthly Billing Closing Report</div>

<div class="controls">
  <label>Month:</label>
  <select id="selMonth">${monthOptions}</select>
  <label>Year:</label>
  <select id="selYear">${yearOptions}</select>
  <button class="btn-load" onclick="loadReport()">Load Report</button>
</div>

<div id="summaryBar" class="summary-bar" style="display:none"></div>
<div id="content"><div class="loading">Select a month and year, then click Load Report.</div></div>

<script>
  function loadReport() {
    const month = parseInt(document.getElementById('selMonth').value, 10);
    const year  = parseInt(document.getElementById('selYear').value, 10);
    document.getElementById('summaryBar').style.display = 'none';
    document.getElementById('content').innerHTML = '<div class="loading">Loading...</div>';
    google.script.run
      .withSuccessHandler(renderReport)
      .withFailureHandler(function(err) {
        document.getElementById('content').innerHTML =
          '<div class="error-msg">⚠ Error: ' + err.message + '</div>';
      })
      .getMonthlyBillingRecords(year, month);
  }

  function renderReport(data) {
    if (!data.ok) {
      document.getElementById('content').innerHTML =
        '<div class="error-msg">⚠ ' + data.error + '</div>';
      return;
    }

    const rows = data.rows;
    const total = data.total;

    // Summary bar
    const byType = {};
    rows.forEach(function(r) {
      byType[r.type] = (byType[r.type] || 0) + r.amount;
    });
    let chips = '<div class="chip">Records: <span>' + rows.length + '</span></div>';
    Object.keys(byType).sort().forEach(function(t) {
      chips += '<div class="chip">' + cap(t) + ': <span>' + fmt(byType[t]) + '</span></div>';
    });
    chips += '<div class="chip total-chip">Total: <span>' + fmt(total) + '</span></div>';
    document.getElementById('summaryBar').innerHTML = chips;
    document.getElementById('summaryBar').style.display = 'flex';

    const tableHeader = '<div class="table-wrap"><table><thead><tr>' +
      '<th>#</th><th>PO #</th><th>Customer</th><th>Project type</th><th>Billing Type</th>' +
      '<th>Billing period</th><th>Paying Customer</th><th>Payment terms</th><th>Amount (&#8362;)</th>' +
      '<th>Hours report</th><th>Invoice type</th>' +
      '<th>Milestone</th><th>Description</th><th>Billing description</th>' +
      '</tr></thead><tbody>';

    if (rows.length === 0) {
      document.getElementById('content').innerHTML =
        tableHeader + '<tr><td colspan="14" style="text-align:center;color:#9e9e9e;padding:30px">No billing records found for this period.</td></tr></tbody></table></div>';
      return;
    }

    let html = tableHeader;

    rows.forEach(function(r, i) {
      const badge = '<span class="badge badge-' + r.type + '">' + cap(r.type) + '</span>';
      const ptLabel = r.paymentTerms ? esc(r.paymentTerms) + ' days' : '';
      html += '<tr>' +
        '<td>' + (i + 1) + '</td>' +
        '<td>' + esc(r.poNumber) + '</td>' +
        '<td>' + esc(r.customer) + '</td>' +
        '<td>' + esc(r.project) + '</td>' +
        '<td class="type">' + badge + '</td>' +
        '<td>' + esc(r.billingPeriod) + '</td>' +
        '<td>' + esc(r.paying) + '</td>' +
        '<td>' + ptLabel + '</td>' +
        '<td class="amount">&#8362; ' + fmt(r.amount) + '</td>' +
        '<td>' + esc(r.hoursReport) + '</td>' +
        '<td>' + esc(r.invoiceType) + '</td>' +
        '<td>' + esc(r.msName) + '</td>' +
        '<td>' + esc(r.msDesc) + '</td>' +
        '<td>' + esc(r.billingDesc) + '</td>' +
        '</tr>';
    });

    html += '</tbody></table></div>';
    document.getElementById('content').innerHTML = html;
  }

  function fmt(n) {
    return Number(n).toLocaleString('he-IL', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }
  function cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : ''; }
  function esc(s) {
    return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }

  // Auto-load on open
  loadReport();
</script>
</body>
</html>`;
}
