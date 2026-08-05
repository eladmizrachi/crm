// ============================================================
// CRM_InvoiceAutomation.gs — Automatic invoice creation via GreenInvoice API
// Handles recurring invoices only. Run manually or via cron.
// Auth token reused from CRM_Billing.gs (getGreenInvoiceToken_).
// ============================================================

const GI_DOCUMENTS_URL   = 'https://api.greeninvoice.co.il/api/v1/documents';
const BILLING_TEST_SHEET = 'Billing test';

// ── Setup "Billing test" sheet ─────────────────────────────
// Copies headers from "Billing" sheet and adds invoice tracking columns T/U/V.
// Safe to run multiple times.
function setupBillingTestSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const billing = ss.getSheetByName(BILLING_SHEET_NAME);
  if (!billing) {
    SpreadsheetApp.getUi().alert('Billing sheet not found.');
    return;
  }

  let testSheet = ss.getSheetByName(BILLING_TEST_SHEET);
  if (!testSheet) {
    testSheet = ss.insertSheet(BILLING_TEST_SHEET);
  } else {
    testSheet.clearContents();
  }

  const lastCol = billing.getLastColumn();
  const headers = billing.getRange(1, 1, 1, lastCol).getValues()[0];
  const fullHeaders = headers.concat(['Invoice ID', 'Invoice Status', 'Invoice Created At']);

  testSheet.getRange(1, 1, 1, fullHeaders.length).setValues([fullHeaders]);
  testSheet.getRange(1, 1, 1, fullHeaders.length)
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  SpreadsheetApp.getUi().alert('"Billing test" sheet is ready (' + fullHeaders.length + ' columns).');
}


// ── Create invoices for all eligible rows ──────────────────
// Eligibility: Billing Type = 'recurring' AND Invoice ID column is empty.
// Run manually to test; attach to cron for automation.
function createInvoicesFromBillingTest() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName(BILLING_TEST_SHEET);
  if (!testSheet) {
    Logger.log('ERROR: "Billing test" sheet not found. Run setupBillingTestSheet() first.');
    return;
  }

  const lastRow = testSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows in Billing test sheet.');
    return;
  }

  const lastCol = testSheet.getLastColumn();
  const data    = testSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const hdrs    = data[0].map(function(h) { return String(h).trim().toLowerCase(); });

  function col(name) {
    return hdrs.findIndex(function(h) { return h === name.toLowerCase(); });
  }

  const idxType        = col('billing type');
  const idxMonth       = col('month');
  const idxPaying      = col('paying customer');
  const idxPayingId    = col('paying customer id');
  const idxAmount      = col('amount');
  const idxPO          = col('po number');
  const idxPayTerms    = col('payment terms');
  const idxInvoiceType = col('initiate invoice type');
  const idxBillingDesc = col('billing description');
  const idxInvoiceId   = col('invoice id');
  const idxStatus      = col('invoice status');
  const idxCreatedAt   = col('invoice created at');

  // Only process rows matching the current month (cron runs on billing date)
  var now            = new Date();
  var currentMonthStr = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  let token;
  try {
    token = getGreenInvoiceToken_();
  } catch(e) {
    Logger.log('ERROR getting GreenInvoice token: ' + e.message);
    return;
  }

  var created = 0, failed = 0, skipped = 0;

  for (var i = 1; i < data.length; i++) {
    const row = data[i];

    if (String(row[idxType]).toLowerCase().trim() !== 'recurring') { skipped++; continue; }
    if (idxInvoiceId !== -1 && row[idxInvoiceId] && String(row[idxInvoiceId]).trim() !== '') { skipped++; continue; }

    // Skip rows not belonging to the current month.
    // normalizeMonthCell_ handles both Date objects and 'YYYY-MM' strings from the sheet.
    const rowMonth = normalizeMonthCell_(row[idxMonth]);
    if (rowMonth !== currentMonthStr) { skipped++; continue; }

    const monthStr     = String(row[idxMonth]       || '').trim();
    const paying       = String(row[idxPaying]      || '').trim();
    const payingId     = String(row[idxPayingId]    || '').trim();
    const amount       = parseFloat(row[idxAmount]) || 0;
    const poNumber     = String(row[idxPO]          || '').trim();
    const paymentTerms = parseInt(row[idxPayTerms]) || 0;
    const invoiceType  = String(row[idxInvoiceType] || '').trim();
    const billingDesc  = String(row[idxBillingDesc] || '').trim();

    if (!payingId || !monthStr || amount <= 0) {
      Logger.log('Row ' + (i + 1) + ': skipped — missing paying customer id, month, or amount.');
      skipped++;
      continue;
    }

    const result = gi_createDocument_(token, {
      monthStr:     monthStr,
      payingId:     payingId,
      payingName:   paying,
      amount:       amount,
      poNumber:     poNumber,
      paymentTerms: paymentTerms,
      invoiceType:  invoiceType,
      billingDesc:  billingDesc
    });

    const rowNum = i + 1;
    if (result.ok) {
      if (idxInvoiceId !== -1) testSheet.getRange(rowNum, idxInvoiceId + 1).setValue(result.invoiceId);
      if (idxStatus    !== -1) testSheet.getRange(rowNum, idxStatus    + 1).setValue('Created');
      if (idxCreatedAt !== -1) testSheet.getRange(rowNum, idxCreatedAt + 1).setValue(gi_nowStr_());
      Logger.log('Row ' + rowNum + ': invoice created — ' + result.invoiceId);
      created++;
    } else {
      if (idxStatus    !== -1) testSheet.getRange(rowNum, idxStatus    + 1).setValue('Error: ' + result.error);
      if (idxCreatedAt !== -1) testSheet.getRange(rowNum, idxCreatedAt + 1).setValue(gi_nowStr_());
      Logger.log('Row ' + rowNum + ': FAILED — ' + result.error);
      failed++;
    }
  }

  Logger.log('Done — Created: ' + created + ', Failed: ' + failed + ', Skipped: ' + skipped);
}


// ── Build and POST a single document to GreenInvoice ───────
function gi_createDocument_(token, d) {
  try {
    var invoiceDate = gi_invoiceDateFromMonth_(d.monthStr);
    var dueDate     = gi_dueDateFromMonth_(d.monthStr, d.paymentTerms);

    // 305 = Tax Invoice, 300 = Proforma Invoice
    var docType = (d.invoiceType.toLowerCase().indexOf('proforma')  !== -1 ||
                   d.invoiceType.toLowerCase().indexOf('performan') !== -1) ? 300 : 305;

    // Top-level description: billing desc + bare PO number
    var description = d.billingDesc + (d.poNumber ? ' ' + d.poNumber : '');

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
      client: {
        id:   d.payingId,
        name: d.payingName,
        add:  false,
        self: false
      },
      income: [
        {
          catalogNum:   '',
          description:  d.billingDesc,
          quantity:     1,
          price:        d.amount,
          currency:     'ILS',
          currencyRate: 1,
          vatType:      1
        }
      ],
      payment: [
        {
          date:         dueDate,
          type:         3,
          price:        d.amount,
          currency:     'ILS',
          currencyRate: 1
        }
      ]
    };

    Logger.log('GI create payload: ' + JSON.stringify(payload));

    var res = UrlFetchApp.fetch(GI_DOCUMENTS_URL, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify(payload),
      headers:            { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    Logger.log('GI response code: '  + res.getResponseCode());
    Logger.log('GI response body: '  + res.getContentText());

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


// ── Date helpers ───────────────────────────────────────────

// Invoice date = first day of billing month (YYYY-MM-01)
function gi_invoiceDateFromMonth_(monthStr) {
  var parts = String(monthStr).split('-');
  if (parts.length < 2) return '';
  return parts[0] + '-' + String(parts[1]).padStart(2, '0') + '-01';
}

// Due date = last day of billing month + paymentTermsDays
function gi_dueDateFromMonth_(monthStr, paymentTermsDays) {
  var parts = String(monthStr).split('-');
  if (parts.length < 2) return '';
  var lastDay = new Date(+parts[0], +parts[1], 0); // day 0 of next month = last day of billing month
  lastDay.setDate(lastDay.getDate() + (paymentTermsDays || 0));
  var y = lastDay.getFullYear();
  var m = String(lastDay.getMonth() + 1).padStart(2, '0');
  var d = String(lastDay.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

function gi_nowStr_() {
  var now = new Date();
  return now.getFullYear() + '-' +
    String(now.getMonth() + 1).padStart(2, '0') + '-' +
    String(now.getDate()).padStart(2, '0') + ' ' +
    String(now.getHours()).padStart(2, '0') + ':' +
    String(now.getMinutes()).padStart(2, '0');
}
