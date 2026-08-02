// ============================================================
// CRM_Billing.gs — GreenInvoice integration & Billing sheet
// ============================================================

const GI_TOKEN_URL      = 'https://api.morning.co/idp/v1/oauth/token';
const GI_CLIENTS_URL    = 'https://api.greeninvoice.co.il/api/v1/clients/search';
const GI_CLIENT_ID_KEY  = '36df7665-3902-4f33-86b9-9f07da77dd4e';
const GI_CLIENT_SECRET_KEY = 'pHyFg>s+lrg1N{@hS1ol6P^:9A^V>f]p';
const BILLING_SHEET_NAME = 'Billing';


// ── GreenInvoice Auth ──────────────────────────────────────

function getGreenInvoiceToken_() {
  const payload = {
    grant_type: 'client_credentials',
    client_id: GI_CLIENT_ID_KEY,
    client_secret: GI_CLIENT_SECRET_KEY
  };
  Logger.log('GI token request URL: ' + GI_TOKEN_URL);
  Logger.log('GI token request body: ' + JSON.stringify(payload));

  const res = UrlFetchApp.fetch(GI_TOKEN_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log('GI token response code: ' + res.getResponseCode());
  Logger.log('GI token response body: ' + res.getContentText());

  const parsed = JSON.parse(res.getContentText());
  Logger.log('GI token parsed keys: ' + Object.keys(parsed).join(', '));

  const token = parsed.access_token || parsed.token || parsed.id_token || parsed.accessToken;
  if (!token) {
    throw new Error('GI token error: ' + (parsed.error_description || parsed.error || 'no token returned. Keys: ' + Object.keys(parsed).join(', ')));
  }
  return token;
}


// ── DEBUG: run this directly from the editor to test the token ──
function debugGreenInvoiceToken() {
  try {
    const token = getGreenInvoiceToken_();
    Logger.log('SUCCESS — token: ' + token.substring(0, 40) + '...');
  } catch(e) {
    Logger.log('FAILED — ' + e.message);
  }
}


// ── GreenInvoice Client List ───────────────────────────────

function getGreenInvoiceClients() {
  try {
    const token = getGreenInvoiceToken_();
    Logger.log('GI clients request URL: ' + GI_CLIENTS_URL);

    const res = UrlFetchApp.fetch(GI_CLIENTS_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ active: true }),
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    Logger.log('GI clients response code: ' + res.getResponseCode());
    Logger.log('GI clients response body: ' + res.getContentText());

    const data = JSON.parse(res.getContentText());
    Logger.log('GI clients parsed keys: ' + Object.keys(data).join(', '));

    const list = data.items || data.clients || data.data || (Array.isArray(data) ? data : []);
    Logger.log('GI clients list length: ' + list.length);
    if (list.length > 0) Logger.log('GI first client sample: ' + JSON.stringify(list[0]));

    const clients = list
      .filter(function(c) { return c.id && c.name; })
      .map(function(c)    {
        return {
          id:           String(c.id),
          name:         String(c.name),
          paymentTerms: String(c.paymentTerms || c.paymentDays || c.payment_terms || c.net || '')
        };
      })
      .sort(function(a, b){ return a.name.localeCompare(b.name, 'he'); });
    return { ok: true, clients: clients };
  } catch(e) {
    Logger.log('GI clients error: ' + e.message);
    return { ok: false, error: e.message, clients: [] };
  }
}


// ── DEBUG: run this directly from the editor to test client list ──
function debugGreenInvoiceClients() {
  const result = getGreenInvoiceClients();
  Logger.log('Result ok: ' + result.ok);
  Logger.log('Client count: ' + result.clients.length);
  if (!result.ok) Logger.log('Error: ' + result.error);
  result.clients.slice(0, 5).forEach(function(c) {
    Logger.log('  Client: ' + c.name + ' | id: ' + c.id);
  });
}


// ── Required billing column headers in order ──────────────
const BILLING_HEADERS = [
  'Year', 'Month', 'Customer', 'Billing Type', 'Paying Customer',
  'Paying Customer Id', 'Amount', 'Milestone name', 'Milestone description',
  'PO Number', 'Payment terms', 'Project type', 'Hours report', 'Initiate invoice type',
  'Billing period', 'Billing description'
];
// Col Q (17) = Price per hour, Col R (18) = Number of hours, Col S (19) = Milestone rate.
// Written at fixed positions; NOT managed by ensureBillingHeaders_ to avoid
// misalignment when the sheet already has data beyond column P.

// Ensures all required headers exist in row 1; appends any that are missing.
function ensureBillingHeaders_(sh) {
  const lastCol      = sh.getLastColumn();
  const existing     = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); })
    : [];
  const existingLow  = existing.map(function(h) { return h.toLowerCase(); });
  BILLING_HEADERS.forEach(function(hdr) {
    if (existingLow.indexOf(hdr.toLowerCase()) === -1) {
      const newCol = sh.getLastColumn() + 1;
      sh.getRange(1, newCol).setValue(hdr);
      existing.push(hdr);
      existingLow.push(hdr.toLowerCase());
    }
  });
}


// ── Save Billing Records to Sheet ─────────────────────────

function saveBillingRecords(d) {
  Logger.log('saveBillingRecords — billingType: ' + d.billingType + ', isMilestones: ' + d.isMilestones);
  Logger.log('saveBillingRecords — billMilestones: ' + JSON.stringify(d.billMilestones));
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(BILLING_SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(BILLING_SHEET_NAME);
    }
    // Always ensure all headers are present (adds missing cols to existing sheets)
    ensureBillingHeaders_(sh);
    const rows = buildBillingRows_(d);
    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 19).setValues(rows);
    }
    return { ok: true, count: rows.length };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}


// ── Build row arrays for each billing type ─────────────────

function buildBillingRows_(d) {
  const rows         = [];
  const org          = d.organization         || '';
  const customer     = d.billPayingCustomer   || '';
  const customerId   = d.billPayingCustomerId || '';
  const poNumber     = d.poNumber             || '';
  const paymentTerms = d.billPaymentTerms     || '';
  const project      = d.project              || '';
  const hoursReport  = d.billHoursReport      || '';
  const invoiceType  = d.billInvoiceType      || '';
  const pricePerHour = d.pricePerHour         || '';
  const numHours     = d.hours                || '';

  // Milestones — billing period/description = project type
  if (d.isMilestones) {
    const msList = Array.isArray(d.billMilestones) ? d.billMilestones : [];
    Logger.log('buildBillingRows_ milestones count: ' + msList.length);
    msList.forEach(function(ms) {
      const year = ms.month ? ms.month.split('-')[0] : '';
      rows.push([year, ms.month || '', org, 'milestones', customer, customerId,
        parseFloat(ms.amount) || 0, ms.name || '', ms.description || '',
        poNumber, paymentTerms, project, hoursReport, invoiceType, '', project,
        ms.pricePerHour || '', ms.hours || '', ms.percentage || '']);
    });
    return rows;
  }

  const rawBt = (d.billingType || '').toLowerCase().replace(/\s+/g, '');
  const type  = rawBt === 'recurring' ? 'recurring' : 'upfront';

  if (type === 'recurring') {
    const start = parseBillingDate_(d.startDate || d.billStartDate);
    const end   = parseBillingDate_(d.renewalDate);
    if (!start || !end) return rows;

    const baseAmount = parseFloat(d.recurringAmount) || 0;
    const period     = d.billPeriod || 'Monthly';
    const step       = period === 'Quarterly' ? 3 : period === 'Yearly' ? 12 : 1;

    const startDay          = start.getDate();
    const endDay            = end.getDate();
    const lastDayOfEndMonth = daysInMonth_(end.getFullYear(), end.getMonth());

    // Exact cycle: end+1 day falls on the same day-of-month as start.
    // e.g. start=Jul 15, end=Jul 14 → end+1=Jul 15 → exactly 12 complete months.
    // In this case every record is full amount; no proration for first or last period.
    const dayAfterEnd  = new Date(end.getFullYear(), end.getMonth(), end.getDate() + 1);
    const isExactCycle = dayAfterEnd.getDate() === startDay;

    let cur     = new Date(start.getFullYear(), start.getMonth(), 1);
    const endMo = new Date(end.getFullYear(), end.getMonth(), 1);
    let isFirst = true;

    // For exact cycles the end month is the start of the next cycle, not a billing month.
    while (isExactCycle ? cur < endMo : cur <= endMo) {
      const yr       = String(cur.getFullYear());
      const monthStr = billingMonthStr_(cur);

      // Is this the last billing period?
      const nextStart = new Date(cur.getFullYear(), cur.getMonth() + step, 1);
      const isLast    = nextStart > endMo;

      let rowAmount, billingDesc;

      if (period === 'Monthly') {
        const dInMonth = daysInMonth_(cur.getFullYear(), cur.getMonth());
        // Prorate only when NOT an exact cycle
        const fromDay  = (!isExactCycle && isFirst && startDay > 1) ? startDay : 1;
        const toDay    = (!isExactCycle && isLast && endDay < lastDayOfEndMonth) ? endDay : dInMonth;
        rowAmount   = Math.round((toDay - fromDay + 1) / dInMonth * baseAmount * 100) / 100;
        billingDesc = project + ' - ' + monthName_(cur) + ' ' + yr;

      } else {
        // Quarterly / Yearly: sum up prorated months across the period
        rowAmount = 0;
        for (var m = 0; m < step; m++) {
          const mo = new Date(cur.getFullYear(), cur.getMonth() + m, 1);
          if (mo > endMo) break;
          const dInMo   = daysInMonth_(mo.getFullYear(), mo.getMonth());
          // Prorate only when NOT an exact cycle
          const fromDay = (!isExactCycle && isFirst && m === 0 && startDay > 1) ? startDay : 1;
          const isEndMo = (mo.getFullYear() === end.getFullYear() && mo.getMonth() === end.getMonth());
          const toDay   = (!isExactCycle && isLast && isEndMo && endDay < lastDayOfEndMonth) ? endDay : dInMo;
          rowAmount += (toDay - fromDay + 1) / dInMo * baseAmount;
        }
        rowAmount = Math.round(rowAmount * 100) / 100;

        // Description: show month range, capped at renewal month
        const periodEndMo = new Date(cur.getFullYear(), cur.getMonth() + step - 1, 1);
        const capEnd = periodEndMo <= endMo ? periodEndMo : endMo;
        billingDesc = project + ' - ' + monthName_(cur) + ' ' + yr
          + ' - ' + monthName_(capEnd) + ' ' + capEnd.getFullYear();
      }

      rows.push([yr, monthStr, org, 'recurring', customer, customerId,
        rowAmount, '', '', poNumber, paymentTerms, project, hoursReport, invoiceType,
        period, billingDesc, '', '', '']);

      cur.setMonth(cur.getMonth() + step);
      isFirst = false;
    }

  } else {
    const amount    = parseFloat(d.totalAmount) || parseFloat(d.amount) || 0;
    const year      = d.billMonth ? d.billMonth.split('-')[0] : '';
    const billingDesc = project;
    rows.push([year, d.billMonth || '', org, 'upfront', customer, customerId,
      amount, '', '', poNumber, paymentTerms, project, hoursReport, invoiceType,
      '', billingDesc, pricePerHour, numHours, '']);
  }

  return rows;
}


// ── Date / calendar helpers ────────────────────────────────

function daysInMonth_(year, month) {
  return new Date(year, month + 1, 0).getDate();
}

function monthName_(date) {
  return ['January','February','March','April','May','June',
          'July','August','September','October','November','December'][date.getMonth()];
}


// ── Date helpers ───────────────────────────────────────────

function parseBillingDate_(str) {
  if (!str) return null;
  str = String(str).trim();
  // dd/mm/yyyy
  let m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  // yyyy-mm-dd
  m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function billingMonthStr_(date) {
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
}
