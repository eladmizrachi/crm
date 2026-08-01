// ============================================================
// CRM_CashflowReport.gs — Expected Cashflow Report
// Shows projected monthly income based on billing records + payment terms.
// ============================================================

function openCashflowReportDialog() {
  const now       = new Date();
  const initYear  = now.getFullYear();
  const initMonth = now.getMonth() + 1;

  // Default range: 2 months back → 6 months ahead (covers short-term records)
  const fromDate  = new Date(now.getFullYear(), now.getMonth() - 2, 1);
  const toDate    = new Date(now.getFullYear(), now.getMonth() + 6, 1);

  const html = HtmlService.createHtmlOutput(
    buildCashflowHtml_(
      fromDate.getFullYear(), fromDate.getMonth() + 1,
      toDate.getFullYear(),   toDate.getMonth() + 1
    )
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
    const sh = ss.getSheetByName('Billing');
    if (!sh) return { ok: true, months: [], grandTotal: 0 };

    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, months: [], grandTotal: 0 };

    // Case-insensitive column finder
    const hdrs = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var colMonth       = hdrs.indexOf('month');
    var colPaying      = hdrs.indexOf('paying customer');
    var colAmt         = hdrs.indexOf('amount');
    var colPT          = hdrs.indexOf('payment terms');
    var colCust        = hdrs.indexOf('customer');
    var colProj        = hdrs.indexOf('project type');
    var colBillingDesc = hdrs.indexOf('billing description');
    var colPO          = hdrs.indexOf('po number');

    var fromKey = String(fromYear) + '-' + String(fromMonth).padStart(2, '0');
    var toKey   = String(toYear)   + '-' + String(toMonth).padStart(2, '0');

    // monthMap: expectedMonthKey → { payingCustomer → { total, records[] } }
    var monthMap = {};

    data.slice(1).forEach(function(row) {
      var billingMonth = cashflowNormalizeMonth_(row[colMonth]);
      if (!billingMonth) return;

      var amount = parseFloat(row[colAmt]) || 0;
      if (amount === 0) return;

      var paymentTerms = parseInt(row[colPT]) || 0;
      var payingCust   = colPaying      >= 0 ? String(row[colPaying]      || '') : '';
      var customer     = colCust        >= 0 ? String(row[colCust]        || '') : '';
      var project      = colProj        >= 0 ? String(row[colProj]        || '') : '';
      var billingDesc  = colBillingDesc >= 0 ? String(row[colBillingDesc] || '') : '';
      var poNumber     = colPO          >= 0 ? String(row[colPO]          || '') : '';

      // Expected payment date: 1st of billing month + paymentTerms days
      var parts    = billingMonth.split('-');
      var baseDate = new Date(+parts[0], +parts[1] - 1, 1);
      baseDate.setDate(baseDate.getDate() + paymentTerms);

      var expYear  = baseDate.getFullYear();
      var expMonth = baseDate.getMonth() + 1;
      var expKey   = String(expYear) + '-' + String(expMonth).padStart(2, '0');

      if (expKey < fromKey || expKey > toKey) return;

      var custKey = payingCust || customer || '(Unknown)';
      if (!monthMap[expKey]) monthMap[expKey] = {};
      if (!monthMap[expKey][custKey]) monthMap[expKey][custKey] = { total: 0, records: [] };

      monthMap[expKey][custKey].total += amount;
      monthMap[expKey][custKey].records.push({
        customer:     customer,
        paying:       payingCust,
        amount:       amount,
        project:      project,
        billingDesc:  billingDesc,
        poNumber:     poNumber,
        billingMonth: billingMonth,
        paymentTerms: paymentTerms,
      });
    });

    var sortedKeys = Object.keys(monthMap).sort();
    var grandTotal = 0;

    var months = sortedKeys.map(function(key) {
      var custMap  = monthMap[key];
      var custKeys = Object.keys(custMap).sort();
      var monthTotal = 0;

      var customers = custKeys.map(function(ck) {
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
        key:       key,
        label:     cashflowMonthLabel_(key),
        total:     monthTotal,
        customers: customers,
      };
    });

    return { ok: true, months: months, grandTotal: Math.round(grandTotal * 100) / 100 };

  } catch(e) {
    Logger.log('getCashflowData error: ' + e.message);
    return { ok: false, error: e.message, months: [], grandTotal: 0 };
  }
}


// ── Helpers ────────────────────────────────────────────────

function cashflowNormalizeMonth_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + String(val.getMonth() + 1).padStart(2, '0');
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{4})-(\d{1,2})$/);
  if (m) return m[1] + '-' + m[2].padStart(2, '0');
  return '';
}

function cashflowMonthLabel_(key) {
  var parts = key.split('-');
  var names = ['January','February','March','April','May','June',
               'July','August','September','October','November','December'];
  return names[+parts[1] - 1] + ' ' + parts[0];
}


// ── HTML builder ───────────────────────────────────────────

function buildCashflowHtml_(fromYear, fromMonth, toYear, toMonth) {
  var monthNames = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];

  var monthOpts = monthNames.map(function(n, i) {
    return { val: i + 1, name: n };
  });

  function buildMonthSel(id, selected) {
    return '<select id="' + id + '">' +
      monthOpts.map(function(o) {
        return '<option value="' + o.val + '"' + (o.val === selected ? ' selected' : '') + '>' + o.name + '</option>';
      }).join('') +
    '</select>';
  }

  function buildYearSel(id, selected) {
    var opts = '';
    for (var y = 2024; y <= fromYear + 3; y++) {
      opts += '<option value="' + y + '"' + (y === selected ? ' selected' : '') + '>' + y + '</option>';
    }
    return '<select id="' + id + '">' + opts + '</select>';
  }

  return '<!DOCTYPE html>\n' +
'<html>\n' +
'<head>\n' +
'<meta charset="utf-8"/>\n' +
'<style>\n' +
'* { box-sizing: border-box; margin: 0; padding: 0; }\n' +
'body { font-family: \'Segoe UI\', Arial, sans-serif; background: #f5f7fa; color: #333; }\n' +
'.header { background: #1b5e20; color: #fff; padding: 16px 20px; font-size: 17px; font-weight: bold; }\n' +
'.controls { display: flex; align-items: center; gap: 10px; padding: 12px 20px; background: #fff; border-bottom: 1px solid #e0e0e0; flex-wrap: wrap; }\n' +
'.controls label { font-weight: 600; font-size: 13px; color: #555; }\n' +
'.controls select { padding: 6px 10px; border: 1px solid #bbb; border-radius: 6px; font-size: 13px; background: #fff; cursor: pointer; }\n' +
'.btn-load { padding: 8px 20px; background: #2e7d32; color: #fff; border: none; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; }\n' +
'.btn-load:hover { background: #1b5e20; }\n' +
'.grand-bar { display: flex; gap: 16px; padding: 10px 20px; background: #e8f5e9; border-bottom: 1px solid #a5d6a7; font-size: 13px; flex-wrap: wrap; align-items: center; }\n' +
'.grand-chip { background: #fff; border: 1px solid #a5d6a7; border-radius: 20px; padding: 4px 14px; font-weight: 600; color: #2e7d32; }\n' +
'.grand-chip span { color: #1b5e20; }\n' +
'.gtotal { background: #1b5e20 !important; color: #fff !important; border-color: #1b5e20 !important; }\n' +
'.gtotal span { color: #b9f6ca !important; }\n' +
'.content-area { padding: 16px 20px; }\n' +
'.month-card { background: #fff; border: 1px solid #c8e6c9; border-radius: 8px; margin-bottom: 16px; overflow: hidden; }\n' +
'.month-hdr { background: #2e7d32; color: #fff; padding: 10px 16px; display: flex; justify-content: space-between; align-items: center; }\n' +
'.month-hdr .mname { font-size: 15px; font-weight: 700; }\n' +
'.month-hdr .mtotal { font-size: 15px; font-weight: 700; color: #b9f6ca; }\n' +
'.cust-section { border-top: 1px solid #e8f5e9; }\n' +
'.cust-hdr { background: #f1f8e9; padding: 8px 16px; display: flex; justify-content: space-between; align-items: center; cursor: pointer; user-select: none; }\n' +
'.cust-hdr:hover { background: #dcedc8; }\n' +
'.cname { font-weight: 700; font-size: 13px; color: #33691e; }\n' +
'.ctotal { font-weight: 700; font-size: 13px; color: #2e7d32; }\n' +
'.ctog { font-size: 11px; color: #66bb6a; margin-left: 8px; }\n' +
'.rec-wrap { display: none; overflow-x: auto; padding: 0 16px 10px; }\n' +
'.rec-wrap.open { display: block; }\n' +
'table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 8px; }\n' +
'thead tr { background: #388e3c; color: #fff; }\n' +
'thead th { padding: 7px 10px; text-align: left; font-weight: 600; white-space: nowrap; }\n' +
'tbody tr:nth-child(even) { background: #f9fbe7; }\n' +
'tbody tr:hover { background: #f1f8e9; }\n' +
'tbody td { padding: 7px 10px; border-bottom: 1px solid #e8f5e9; vertical-align: top; }\n' +
'td.amt { text-align: right; font-weight: 600; color: #2e7d32; white-space: nowrap; }\n' +
'.empty { padding: 40px; text-align: center; color: #9e9e9e; font-size: 14px; }\n' +
'.errmsg { padding: 16px 20px; color: #c62828; font-size: 13px; background: #ffebee; margin: 12px 20px; border-radius: 6px; }\n' +
'.loading { padding: 40px; text-align: center; color: #388e3c; font-size: 14px; }\n' +
'</style>\n' +
'</head>\n' +
'<body>\n' +
'<div class="header">&#128176; Expected Cashflow Report</div>\n' +
'<div class="controls">\n' +
'  <label>From:</label>\n' +
'  ' + buildMonthSel('fromMonth', fromMonth) + '\n' +
'  ' + buildYearSel('fromYear', fromYear) + '\n' +
'  <span style="color:#888">&rarr;</span>\n' +
'  <label>To:</label>\n' +
'  ' + buildMonthSel('toMonth', toMonth) + '\n' +
'  ' + buildYearSel('toYear', toYear) + '\n' +
'  <button class="btn-load" onclick="loadReport()">Load Report</button>\n' +
'</div>\n' +
'<div id="grandBar" class="grand-bar" style="display:none"></div>\n' +
'<div id="mainContent" class="content-area"><div class="loading">Select a range and click Load Report.</div></div>\n' +
'<script>\n' +
'function loadReport() {\n' +
'  var fm = parseInt(document.getElementById("fromMonth").value, 10);\n' +
'  var fy = parseInt(document.getElementById("fromYear").value,  10);\n' +
'  var tm = parseInt(document.getElementById("toMonth").value,   10);\n' +
'  var ty = parseInt(document.getElementById("toYear").value,    10);\n' +
'  document.getElementById("grandBar").style.display = "none";\n' +
'  document.getElementById("mainContent").innerHTML = "<div class=\'loading\'>Loading&#8230;</div>";\n' +
'  google.script.run\n' +
'    .withSuccessHandler(renderReport)\n' +
'    .withFailureHandler(function(err) {\n' +
'      var msg = (err && err.message) ? err.message : String(err);\n' +
'      document.getElementById("mainContent").innerHTML =\n' +
'        "<div class=\'errmsg\'>&#9888; Error: " + msg + "</div>";\n' +
'    })\n' +
'    .getCashflowData(fy, fm, ty, tm);\n' +
'}\n' +
'\n' +
'function renderReport(data) {\n' +
'  if (!data || !data.ok) {\n' +
'    document.getElementById("mainContent").innerHTML =\n' +
'      "<div class=\'errmsg\'>&#9888; " + (data ? data.error : "No response") + "</div>";\n' +
'    return;\n' +
'  }\n' +
'  var months = data.months;\n' +
'  if (!months || months.length === 0) {\n' +
'    document.getElementById("grandBar").style.display = "none";\n' +
'    document.getElementById("mainContent").innerHTML =\n' +
'      "<div class=\'empty\'>No expected payments found for this period.</div>";\n' +
'    return;\n' +
'  }\n' +
'  var bar = "<div class=\'grand-chip\'>Months: <span>" + months.length + "</span></div>";\n' +
'  bar += "<div class=\'grand-chip gtotal\'>Total expected: <span>&#8362; " + fmt(data.grandTotal) + "</span></div>";\n' +
'  document.getElementById("grandBar").innerHTML = bar;\n' +
'  document.getElementById("grandBar").style.display = "flex";\n' +
'\n' +
'  var html = "";\n' +
'  months.forEach(function(mo) {\n' +
'    html += "<div class=\'month-card\'>";\n' +
'    html += "<div class=\'month-hdr\'><span class=\'mname\'>" + esc(mo.label) + "</span><span class=\'mtotal\'>&#8362; " + fmt(mo.total) + "</span></div>";\n' +
'    mo.customers.forEach(function(cust, ci) {\n' +
'      var uid = mo.key.replace("-","") + "_" + ci;\n' +
'      html += "<div class=\'cust-section\'>";\n' +
'      html += "<div class=\'cust-hdr\' onclick=\'toggleRec(\\\"" + uid + "\\\")\' >";\n' +
'      html += "<span class=\'cname\'>" + esc(cust.name) + "</span>";\n' +
'      html += "<span><span class=\'ctotal\'>&#8362; " + fmt(cust.total) + "</span><span class=\'ctog\' id=\'tog_" + uid + "\'>&#9660; show</span></span>";\n' +
'      html += "</div>";\n' +
'      html += "<div class=\'rec-wrap\' id=\'rec_" + uid + "\'>";\n' +
'      html += "<table><thead><tr><th>PO #</th><th>Customer</th><th>Project</th><th>Description</th><th>Billing month</th><th>Terms</th><th>Amount (&#8362;)</th></tr></thead><tbody>";\n' +
'      cust.records.forEach(function(r) {\n' +
'        html += "<tr>";\n' +
'        html += "<td>" + esc(r.poNumber) + "</td>";\n' +
'        html += "<td>" + esc(r.customer) + "</td>";\n' +
'        html += "<td>" + esc(r.project) + "</td>";\n' +
'        html += "<td>" + esc(r.billingDesc) + "</td>";\n' +
'        html += "<td>" + esc(r.billingMonth) + "</td>";\n' +
'        html += "<td>" + (r.paymentTerms ? r.paymentTerms + "d" : "-") + "</td>";\n' +
'        html += "<td class=\'amt\'>&#8362; " + fmt(r.amount) + "</td>";\n' +
'        html += "</tr>";\n' +
'      });\n' +
'      html += "</tbody></table></div></div>";\n' +
'    });\n' +
'    html += "</div>";\n' +
'  });\n' +
'  document.getElementById("mainContent").innerHTML = html;\n' +
'}\n' +
'\n' +
'function toggleRec(uid) {\n' +
'  var w = document.getElementById("rec_" + uid);\n' +
'  var t = document.getElementById("tog_" + uid);\n' +
'  if (!w) return;\n' +
'  var open = w.classList.toggle("open");\n' +
'  if (t) t.innerHTML = open ? "&#9650; hide" : "&#9660; show";\n' +
'}\n' +
'\n' +
'function fmt(n) {\n' +
'  return Number(n).toLocaleString("he-IL", {minimumFractionDigits:0, maximumFractionDigits:2});\n' +
'}\n' +
'function esc(s) {\n' +
'  return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");\n' +
'}\n' +
'loadReport();\n' +
'</script>\n' +
'</body>\n' +
'</html>';
}
