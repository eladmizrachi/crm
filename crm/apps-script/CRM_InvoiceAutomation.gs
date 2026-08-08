// ============================================================
// CRM_InvoiceAutomation.gs — Automatic invoice creation via GreenInvoice API
// Handles recurring invoices only. Run manually or via cron.
// Auth token reused from CRM_Billing.gs (getGreenInvoiceToken_).
// Tracking columns expected in Billing sheet: Invoice ID, Invoice Status,
// Invoice Created At, Invoice number  (looked up by header name).
// ============================================================

const GI_DOCUMENTS_URL = 'https://api.greeninvoice.co.il/api/v1/documents';


// ── Preview: show what invoices WOULD be created ───────────
function openInvoicePreviewDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BILLING_SHEET_NAME);
  if (!sh) {
    SpreadsheetApp.getUi().alert('Billing sheet not found.');
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data in Billing sheet.');
    return;
  }

  var now             = new Date();
  var currentMonthStr = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  var data = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  var hdrs = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  function col(name) { return hdrs.findIndex(function(h) { return h === name.toLowerCase(); }); }

  var idxType        = col('billing type');
  var idxMonth       = col('month');
  var idxPaying      = col('paying customer');
  var idxPayingId    = col('paying customer id');
  var idxAmount      = col('amount');
  var idxPO          = col('po number');
  var idxPayTerms    = col('payment terms');
  var idxInvoiceType = col('initiate invoice type');
  var idxBillingDesc = col('billing description');
  var idxInvoiceId   = col('invoice id');

  var token = null;
  try { token = getGreenInvoiceToken_(); } catch(e) {}

  var previews = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[idxType]).toLowerCase().trim() !== 'recurring') continue;
    if (idxInvoiceId !== -1 && row[idxInvoiceId] && String(row[idxInvoiceId]).trim() !== '') continue;

    var rowMonth = normalizeMonthCell_(row[idxMonth]);
    if (rowMonth !== currentMonthStr) continue;

    var paying       = String(row[idxPaying]      || '').trim();
    var payingId     = String(row[idxPayingId]    || '').trim();
    var amount       = parseFloat(row[idxAmount]) || 0;
    var poNumber     = String(row[idxPO]          || '').trim();
    var paymentTerms = parseInt(row[idxPayTerms]) || 0;
    var invoiceType  = String(row[idxInvoiceType] || '').trim();
    var billingDesc  = String(row[idxBillingDesc] || '').trim();

    if (!payingId || amount <= 0) continue;

    var emails = [];
    if (token) {
      var cd = gi_getClientDetails_(token, payingId);
      emails = cd.emails;
    }

    var isProforma  = invoiceType.toLowerCase().indexOf('proforma')  !== -1 ||
                      invoiceType.toLowerCase().indexOf('performan') !== -1;
    var rowDueDate  = gi_dueDateFromNow_(paymentTerms);
    var description = billingDesc + (poNumber ? ' ' + poNumber : '');

    previews.push({
      paying:      paying,
      description: description,
      itemDesc:    billingDesc,
      amount:      amount,
      docType:     isProforma ? 'Proforma' : 'Tax Invoice',
      email:       emails.join(', ') || '—',
      dueDate:     rowDueDate
    });
  }

  var monthLabel = gi_monthLabel_(currentMonthStr);

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:16px;color:#212121;margin:0}' +
    'h2{margin:0 0 4px;color:#1a237e}' +
    '.subtitle{color:#666;margin-bottom:14px;font-size:12px}' +
    '.summary{background:#e8eaf6;border-radius:8px;padding:10px 16px;margin-bottom:14px;font-size:13px}' +
    '.summary b{color:#1a237e}' +
    'table{width:100%;border-collapse:collapse;font-size:12px}' +
    'thead tr{background:#1a237e;color:#fff}' +
    'th{padding:9px 10px;text-align:left;white-space:nowrap}' +
    'tbody tr:nth-child(even){background:#f0f4ff}' +
    'td{padding:8px 10px;border-bottom:1px solid #e0e0e0;vertical-align:top}' +
    'td.amt{text-align:right;font-weight:600;color:#1a237e}' +
    '.tax{background:#e8f5e9;color:#2e7d32;border-radius:10px;padding:2px 8px;font-size:11px;white-space:nowrap}' +
    '.pro{background:#fff3e0;color:#e65100;border-radius:10px;padding:2px 8px;font-size:11px;white-space:nowrap}' +
    '.none{color:#9e9e9e;text-align:center;padding:40px}' +
    '</style></head><body>' +
    '<h2>Invoice Preview — ' + monthLabel + '</h2>' +
    '<div class="subtitle">Recurring invoices that will be created when you run Create Invoices. Verify before proceeding.</div>';

  if (previews.length === 0) {
    html += '<div class="none">No eligible recurring invoices found for ' + monthLabel + '.</div>';
  } else {
    var totalAmt = 0;
    for (var k = 0; k < previews.length; k++) totalAmt += previews[k].amount;
    html += '<div class="summary">Will create <b>' + previews.length + ' invoice' + (previews.length !== 1 ? 's' : '') + '</b>' +
      ' &nbsp;·&nbsp; Total net amount: <b>&#8362; ' + gi_fmt_(totalAmt) + '</b></div>';

    html += '<table><thead><tr>' +
      '<th>#</th><th>Paying Customer</th><th>Description</th>' +
      '<th>Item Description</th><th>Amount (net)</th>' +
      '<th>Type</th><th>Email</th><th>Due Date</th>' +
      '</tr></thead><tbody>';

    for (var j = 0; j < previews.length; j++) {
      var p    = previews[j];
      var badge = p.docType === 'Tax Invoice'
        ? '<span class="tax">Tax Invoice</span>'
        : '<span class="pro">Proforma</span>';
      html += '<tr>' +
        '<td>' + (j + 1) + '</td>' +
        '<td>' + gi_esc_(p.paying)      + '</td>' +
        '<td>' + gi_esc_(p.description) + '</td>' +
        '<td>' + gi_esc_(p.itemDesc)    + '</td>' +
        '<td class="amt">&#8362; ' + gi_fmt_(p.amount) + '</td>' +
        '<td>' + badge + '</td>' +
        '<td>' + gi_esc_(p.email)   + '</td>' +
        '<td>' + gi_esc_(p.dueDate) + '</td>' +
        '</tr>';
    }
    html += '</tbody></table>';
  }

  html += '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1050).setHeight(520),
    'Invoice Preview — ' + monthLabel
  );
}


// ── Create invoices for all eligible rows in Billing sheet ─
// Eligibility: Billing Type = 'recurring', current month, Invoice ID empty.
// Run manually or via cron on the billing date.
function createInvoicesFromBilling() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BILLING_SHEET_NAME);
  if (!sh) {
    Logger.log('ERROR: Billing sheet not found.');
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows in Billing sheet.');
    return;
  }

  var data = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  var hdrs = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  function col(name) { return hdrs.findIndex(function(h) { return h === name.toLowerCase(); }); }

  var idxType        = col('billing type');
  var idxMonth       = col('month');
  var idxPaying      = col('paying customer');
  var idxPayingId    = col('paying customer id');
  var idxAmount      = col('amount');
  var idxPO          = col('po number');
  var idxPayTerms    = col('payment terms');
  var idxInvoiceType = col('initiate invoice type');
  var idxBillingDesc = col('billing description');
  var idxInvoiceId   = col('invoice id');
  var idxStatus      = col('invoice status');
  var idxCreatedAt   = col('invoice created at');
  var idxInvoiceNum  = col('invoice number');

  var now             = new Date();
  var currentMonthStr = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  var token;
  try {
    token = getGreenInvoiceToken_();
  } catch(e) {
    Logger.log('ERROR getting GreenInvoice token: ' + e.message);
    return;
  }

  var created = 0, failed = 0, skipped = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    if (String(row[idxType]).toLowerCase().trim() !== 'recurring') { skipped++; continue; }
    if (idxInvoiceId !== -1 && row[idxInvoiceId] && String(row[idxInvoiceId]).trim() !== '') { skipped++; continue; }

    var rowMonth = normalizeMonthCell_(row[idxMonth]);
    if (rowMonth !== currentMonthStr) { skipped++; continue; }

    var paying       = String(row[idxPaying]      || '').trim();
    var payingId     = String(row[idxPayingId]    || '').trim();
    var amount       = parseFloat(row[idxAmount]) || 0;
    var poNumber     = String(row[idxPO]          || '').trim();
    var paymentTerms = parseInt(row[idxPayTerms]) || 0;
    var invoiceType  = String(row[idxInvoiceType] || '').trim();
    var billingDesc  = String(row[idxBillingDesc] || '').trim();

    if (!payingId || !rowMonth || amount <= 0) {
      Logger.log('Row ' + (i + 1) + ': skipped — missing paying customer id, month, or amount.');
      skipped++;
      continue;
    }

    var result = gi_createDocument_(token, {
      monthStr:     rowMonth,
      payingId:     payingId,
      payingName:   paying,
      amount:       amount,
      poNumber:     poNumber,
      paymentTerms: paymentTerms,
      invoiceType:  invoiceType,
      billingDesc:  billingDesc
    });

    var rowNum = i + 1;
    if (result.ok) {
      var invoiceNum = gi_getDocumentNumber_(token, result.invoiceId);
      if (idxInvoiceId  !== -1) sh.getRange(rowNum, idxInvoiceId  + 1).setValue(result.invoiceId);
      if (idxInvoiceNum !== -1) sh.getRange(rowNum, idxInvoiceNum + 1).setValue(invoiceNum);
      if (idxStatus     !== -1) sh.getRange(rowNum, idxStatus     + 1).setValue('Created');
      if (idxCreatedAt  !== -1) sh.getRange(rowNum, idxCreatedAt  + 1).setValue(gi_nowStr_());
      Logger.log('Row ' + rowNum + ': created — id: ' + result.invoiceId + ', number: ' + invoiceNum);
      created++;
    } else {
      if (idxStatus    !== -1) sh.getRange(rowNum, idxStatus    + 1).setValue('Error: ' + result.error);
      if (idxCreatedAt !== -1) sh.getRange(rowNum, idxCreatedAt + 1).setValue(gi_nowStr_());
      Logger.log('Row ' + rowNum + ': FAILED — ' + result.error);
      failed++;
    }
  }

  Logger.log('Done — Created: ' + created + ', Failed: ' + failed + ', Skipped: ' + skipped);
}


// ── Fetch client details (email) from GreenInvoice ─────────
function gi_getClientDetails_(token, clientId) {
  try {
    var res = UrlFetchApp.fetch('https://api.greeninvoice.co.il/api/v1/clients/' + clientId, {
      method:             'get',
      headers:            { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { emails: [] };
    var parsed = JSON.parse(res.getContentText());
    var emails = [];
    if (Array.isArray(parsed.emails)) {
      emails = parsed.emails.filter(function(e) { return e && String(e).trim(); });
    } else if (parsed.email && String(parsed.email).trim()) {
      emails = [String(parsed.email).trim()];
    }
    return { emails: emails };
  } catch(e) {
    Logger.log('GI client details error: ' + e.message);
    return { emails: [] };
  }
}


// ── Build and POST a single document to GreenInvoice ───────
function gi_createDocument_(token, d) {
  try {
    var invoiceDate   = gi_todayStr_();
    var dueDate       = gi_dueDateFromNow_(d.paymentTerms);
    var clientDetails = gi_getClientDetails_(token, d.payingId);

    var docType = (d.invoiceType.toLowerCase().indexOf('proforma')  !== -1 ||
                   d.invoiceType.toLowerCase().indexOf('performan') !== -1) ? 300 : 305;

    var description = d.billingDesc + (d.poNumber ? ' ' + d.poNumber : '');
    var clientObj   = { id: d.payingId, name: d.payingName, add: false, self: false };
    if (clientDetails.emails.length > 0) clientObj.emails = clientDetails.emails;

    var payload = {
      type:        docType,
      date:        invoiceDate,
      dueDate:     dueDate,
      lang:        'he',
      currency:    'ILS',
      vatType:     0,
      signed:      true,
      rounding:    false,
      description: description,
      client:      clientObj,
      income: [{
        catalogNum:   '',
        description:  d.billingDesc,
        quantity:     1,
        price:        d.amount,
        currency:     'ILS',
        currencyRate: 1,
        vatType:      0
      }],
      payment: [{
        date:         dueDate,
        type:         3,
        price:        d.amount,
        currency:     'ILS',
        currencyRate: 1
      }]
    };

    Logger.log('GI create payload: ' + JSON.stringify(payload));

    var res = UrlFetchApp.fetch(GI_DOCUMENTS_URL, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify(payload),
      headers:            { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    Logger.log('GI response code: ' + res.getResponseCode());
    Logger.log('GI response body: ' + res.getContentText());

    var parsed = JSON.parse(res.getContentText());
    var code   = res.getResponseCode();
    if ((code !== 200 && code !== 201) || !parsed.id) {
      return { ok: false, error: parsed.errorMessage || parsed.error || ('HTTP ' + code) };
    }
    return { ok: true, invoiceId: parsed.id };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}


// ── Fetch invoice number from created document ─────────────
function gi_getDocumentNumber_(token, documentId) {
  try {
    var res = UrlFetchApp.fetch(GI_DOCUMENTS_URL + '/' + documentId, {
      method:             'get',
      headers:            { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    var parsed = JSON.parse(res.getContentText());
    return parsed.number || '';
  } catch(e) {
    Logger.log('GI get document error: ' + e.message);
    return '';
  }
}


// ── Date / formatting helpers ──────────────────────────────

function gi_todayStr_() {
  var now = new Date();
  return now.getFullYear() + '-' +
    String(now.getMonth() + 1).padStart(2, '0') + '-' +
    String(now.getDate()).padStart(2, '0');
}

function gi_dueDateFromNow_(paymentTermsDays) {
  var now         = new Date();
  var lastOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  lastOfMonth.setDate(lastOfMonth.getDate() + (paymentTermsDays || 0));
  return lastOfMonth.getFullYear() + '-' +
    String(lastOfMonth.getMonth() + 1).padStart(2, '0') + '-' +
    String(lastOfMonth.getDate()).padStart(2, '0');
}

function gi_nowStr_() {
  var now = new Date();
  return now.getFullYear() + '-' +
    String(now.getMonth() + 1).padStart(2, '0') + '-' +
    String(now.getDate()).padStart(2, '0') + ' ' +
    String(now.getHours()).padStart(2, '0') + ':' +
    String(now.getMinutes()).padStart(2, '0');
}

function gi_monthLabel_(monthStr) {
  var parts = String(monthStr).split('-');
  if (parts.length < 2) return monthStr;
  var names = ['January','February','March','April','May','June',
               'July','August','September','October','November','December'];
  return names[parseInt(parts[1]) - 1] + ' ' + parts[0];
}

function gi_fmt_(n) {
  return Number(n).toLocaleString('he-IL', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function gi_esc_(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
