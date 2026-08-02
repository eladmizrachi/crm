// ============================================================
// CRM_NewCustomer.gs — New Customer Creation
// Depends on: CRM_Setup.gs (for SHEETS, norm, runSearch)
//             CRM_NewPurchaseOrder.gs (for getPOActivities)
// ============================================================

// ── LOAD CODETABLE VALUES ──────────────────────────────────
// The CODETABLES sheet has NO header row — row 1 is already data.
// Columns: A=Industry, B=Activity, C=Manager, D=Recurring Type
function getCodeTableValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("CODETABLES");
  if (!sh) {
    SpreadsheetApp.getUi().alert('❌ Cannot find sheet named "CODETABLES"');
    return { industries: [], activities: [], managers: [] };
  }

  const data = sh.getDataRange().getValues();

  function colValues(colIdx) {
    return data
      .map(row => String(row[colIdx] || "").trim())
      .filter(v => v !== "");
  }

  return {
    industries: colValues(0), // Column A
    activities: colValues(1), // Column B
    managers:   colValues(2), // Column C
  };
}


// ── OPEN NEW CUSTOMER DIALOG ───────────────────────────────
function openNewCustomerDialog() {
  const codes      = getCodeTableValues();
  const poActivities = getPOActivities(); // from CRM_NewPurchaseOrder.gs

  const giResult   = getGreenInvoiceClients();
  const giClients  = giResult.clients || [];
  const giError    = giResult.ok ? '' : (giResult.error || 'Could not load billing customers');
  function buildGIOptions(clients) {
    return clients.map(c => `<option value="${escapeHtml(c.id)}" data-pt="${escapeHtml(c.paymentTerms || '')}">${escapeHtml(c.name)}</option>`).join('');
  }

  function buildOptions(arr) {
    return arr.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
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
        label .req-dynamic { color: #c62828; margin-left: 2px; display: none; }
        label .opt { color: #888; font-weight: 400; font-size: 11px; margin-left: 4px; }

        input[type=text], input[type=number], input[type=url], input[type=date],
        select, textarea {
          width: 100%; padding: 8px 10px; font-size: 13px;
          border: 1.5px solid #9fa8da; border-radius: 6px;
          background: #fff; color: #212121; outline: none;
          transition: border-color .15s;
        }
        textarea { resize: vertical; min-height: 52px; font-family: inherit; }
        input:focus, select:focus, textarea:focus { border-color: #3949ab; }
        input.error, select.error { border-color: #c62828 !important; }

        .skip-po-row {
          display: flex; align-items: center; gap: 8px;
          background: #fff8e1; border: 1px solid #ffe082;
          border-radius: 6px; padding: 8px 12px; margin-bottom: 10px;
          cursor: pointer;
        }
        .skip-po-row input[type=checkbox] {
          width: auto; margin: 0; cursor: pointer;
          accent-color: #f57f17;
        }
        .skip-po-row span { font-size: 12px; color: #5d4037; }

        .billing-hint {
          display: none; font-size: 11.5px; font-weight: 600;
          padding: 7px 10px; border-radius: 5px; margin-bottom: 8px;
        }
        .billing-hint.hourly    { background: #e3f2fd; color: #1565c0; border: 1px solid #90caf9; }
        .billing-hint.fixed     { background: #fff8e1; color: #f57f17; border: 1px solid #ffe082; }
        .billing-hint.recurring  { background: #f3e5f5; color: #6a1b9a; border: 1px solid #ce93d8; }
        .billing-hint.milestones { background: #e8f5e9; color: #2e7d32; border: 1px solid #81c784; }

        /* Currency selector */
        .currency-row { display: flex; gap: 8px; margin-bottom: 10px; }
        .currency-btn { flex: 1; padding: 7px 10px; font-size: 13px; font-weight: 700;
          border: 2px solid #9fa8da; border-radius: 6px;
          background: #fff; color: #555; cursor: pointer; text-align: center; transition: all .15s; }
        .currency-btn.active-nis { border-color: #1a237e; background: #1a237e; color: #fff; }
        .currency-btn.active-usd { border-color: #2e7d32; background: #2e7d32; color: #fff; }
        .rate-info { font-size: 11px; color: #5c6bc0; text-align: center;
          margin-bottom: 8px; min-height: 16px; }

        .total-box {
          background: #e8f5e9; border: 1.5px solid #81c784;
          border-radius: 6px; padding: 10px 14px; margin-top: 4px;
          display: flex; align-items: center; justify-content: space-between;
        }
        .total-box .total-label { font-weight: 700; color: #2e7d32; font-size: 13px; }
        .total-box .total-value { font-weight: 700; color: #2e7d32; font-size: 16px; }

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
        #emailStatus {
          margin-top: 6px; font-size: 11px;
          min-height: 14px; text-align: center;
        }
        #emailStatus.ok  { color: #2e7d32; }
        #emailStatus.err { color: #c62828; }
      </style>
    </head>
    <body>
      <h2>➕ New Customer</h2>
      <p class="sub">Fill in the details below to add a new customer record.</p>

      <!-- ── Organization ── -->
      <div class="section-title">Organization</div>
      <div class="field">
        <label>Organization Name <span class="req">*</span></label>
        <input type="text" id="orgName" placeholder="e.g. Acme Ltd." />
      </div>
      <div class="field">
        <label>Industry <span class="req">*</span></label>
        <select id="industry">
          <option value="">— Select —</option>
          ${buildOptions(codes.industries)}
        </select>
      </div>
      <div class="field">
        <label>Main Activity Type <span class="req">*</span></label>
        <select id="activity">
          <option value="">— Select —</option>
          ${buildOptions(codes.activities)}
        </select>
      </div>
      <div class="field">
        <label>Customer Manager <span class="req">*</span></label>
        <select id="manager">
          <option value="">— Select —</option>
          ${buildOptions(codes.managers)}
        </select>
      </div>

      <!-- ── Contact Details ── -->
      <div class="section-title">
        Contact Details
        <span style="font-weight:400;color:#888;font-size:10px;text-transform:none;letter-spacing:0">(one contact)</span>
      </div>
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
      <div class="field">
        <label>Email <span class="req">*</span></label>
        <input type="text" id="contactEmail" placeholder="e.g. john@company.com" />
      </div>
      <div class="row2">
        <div class="field">
          <label>Position <span class="opt">(optional)</span></label>
          <input type="text" id="contactPosition" placeholder="e.g. CEO" />
        </div>
        <div class="field">
          <label>Phone <span class="opt">(optional)</span></label>
          <input type="text" id="contactPhone" placeholder="e.g. 050-1234567" />
        </div>
      </div>
      <div class="field">
        <label>Date of Birth <span class="opt">(optional)</span></label>
        <input type="date" id="contactDob" />
      </div>
      <div class="field">
        <label>Address <span class="opt">(optional)</span></label>
        <input type="text" id="contactAddress" placeholder="e.g. 123 Main St, Tel Aviv" />
      </div>

      <!-- ── Purchase Order ── -->
      <div class="section-title">Purchase Order</div>

      <label class="skip-po-row" for="skipPO">
        <input type="checkbox" id="skipPO" onchange="togglePOSection()" />
        <span>Save customer without creating a Purchase Order</span>
      </label>

      <div id="poSection">
        <div class="field">
          <label>Project Name <span class="req">*</span></label>
          <input type="text" id="projectName" placeholder="e.g. CRM Implementation" />
        </div>
        <div class="row2">
          <div class="field">
            <label>Project <span class="req">*</span></label>
            <select id="project" onchange="onProjectChange()">
              <option value="">— Select Activity —</option>
              ${buildOptions(poActivities)}
            </select>
          </div>
          <div class="field">
            <label>PO Number <span class="req">*</span></label>
            <input type="number" id="poNumber" placeholder="e.g. 1001" min="1" />
          </div>
        </div>

        <div class="field">
          <label>Billing Type <span class="req">*</span></label>
          <select id="billingType" onchange="onBillingTypeChange()">
            <option value="">— Select Billing Type —</option>
            <option value="Hourly">Hourly</option>
            <option value="Fixed Project">Fixed Project</option>
            <option value="Recurring">Recurring</option>
          </select>
        </div>

        <div class="field">
          <label>Project Description <span class="opt">(optional)</span></label>
          <textarea id="projectDescription" placeholder="Brief description of the project scope…"></textarea>
        </div>

        <div class="row2">
          <div class="field">
            <label>Customer <span class="opt">(optional)</span></label>
            <input type="text" id="poCustomer" placeholder="e.g. Commbox" />
          </div>
          <div class="field">
            <label>Proposal Link <span class="opt">(optional)</span></label>
            <input type="url" id="proposalLink" placeholder="https://…" />
          </div>
        </div>

        <!-- Financial Details -->
        <div class="section-title">Financial Details</div>

        <!-- Currency selector -->
        <div class="currency-row">
          <button type="button" class="currency-btn active-nis" id="btnNIS" onclick="setCurrency('NIS')">&#8362; NIS</button>
          <button type="button" class="currency-btn" id="btnUSD" onclick="setCurrency('USD')">$ USD</button>
        </div>
        <div class="rate-info" id="rateInfo"></div>

        <div class="billing-hint" id="billingHint"></div>

        <div class="field">
          <label>
            Amount (<span id="currSymbolAmount">&#8362;</span>)
            <span class="req-dynamic" id="reqAmount">*</span>
            <span class="opt" id="optAmount">(optional)</span>
          </label>
          <input type="number" id="amount" placeholder="0" min="0" step="0.01" oninput="calcTotal()" />
        </div>

        <div class="field">
          <label>
            Recurring Amount (<span id="currSymbolRecurring">&#8362;</span>)
            <span class="req-dynamic" id="reqRecurring">*</span>
            <span class="opt" id="optRecurring">(optional)</span>
          </label>
          <input type="number" id="recurringAmount" placeholder="0" min="0" step="0.01" oninput="calcTotal()" />
        </div>

        <div class="field" id="fieldRecurringPeriod" style="display:none">
          <label>Recurring Period (months) <span class="req-dynamic" id="reqRecurringPeriod" style="display:inline">*</span></label>
          <select id="recurringPeriod" onchange="calcTotal()">
            <option value="">— Select Period —</option>
            <option value="1">1</option><option value="2">2</option><option value="3">3</option>
            <option value="4">4</option><option value="5">5</option><option value="6">6</option>
            <option value="7">7</option><option value="8">8</option><option value="9">9</option>
            <option value="10">10</option><option value="11">11</option><option value="12">12</option>
          </select>
        </div>

        <div class="row2">
          <div class="field">
            <label>
              Hours
              <span class="req-dynamic" id="reqHours">*</span>
              <span class="opt" id="optHours">(optional)</span>
            </label>
            <input type="number" id="hours" placeholder="0" min="0" step="0.5" oninput="calcTotal(); recalcAllMs();" />
          </div>
          <div class="field">
            <label>
              Price/Hour (<span id="currSymbolPricePerHour">&#8362;</span>)
              <span class="req-dynamic" id="reqPricePerHour">*</span>
              <span class="opt" id="optPricePerHour">(optional)</span>
            </label>
            <input type="number" id="pricePerHour" placeholder="0" min="0" step="0.01" oninput="calcTotal(); recalcAllMs();" />
          </div>
        </div>

        <div class="field">
          <label>Milestones <span class="opt">(optional)</span></label>
          <input type="text" id="milestones" placeholder="e.g. Phase 1, Phase 2, Phase 3" />
        </div>

        <div class="field">
          <label>
            Renewal Date
            <span class="req-dynamic" id="reqRenewalDate">*</span>
            <span class="opt" id="optRenewalDate">(optional)</span>
          </label>
          <input type="date" id="renewalDate" />
        </div>

        <!-- Start Date — shown for Recurring billing -->
        <div class="field" id="fieldStartDate" style="display:none">
          <label>Start Date <span class="req">*</span></label>
          <input type="date" id="startDate" />
        </div>

        <div class="field">
          <label>
            Commbox ARR (&#8362;)
            <span class="req-dynamic" id="reqCommboxARR">*</span>
            <span class="opt" id="optCommboxARR">(optional)</span>
          </label>
          <input type="number" id="commboxARR" placeholder="0" min="0" step="0.01" />
        </div>

        <div class="total-box" style="flex-direction:column; align-items:flex-start; gap:4px;">
          <div style="display:flex; justify-content:space-between; width:100%;">
            <span class="total-label">&#128176; Total Amount</span>
            <span class="total-value" id="totalDisplay">&#8362;0.00</span>
          </div>
          <div id="totalConversion" style="font-size:11px; color:#5c6bc0; display:none;"></div>
        </div>

        <!-- Billing Records -->
        <div class="section-title">Billing Records <span class="req">*</span></div>
        ${giError ? '<div style="font-size:11px;color:#c62828;margin-bottom:6px;">&#9888; ' + escapeHtml(giError) + '</div>' : ''}
        <div class="field">
          <label>Paying Customer (GreenInvoice) <span class="req">*</span></label>
          <select id="billPayingCustomer" onchange="onBillCustomerChange()">
            <option value="">&#8212; Select &#8212;</option>
            ${buildGIOptions(giClients)}
          </select>
        </div>
        <div class="field" id="fieldPaymentTerms" style="display:none">
          <label>Payment Terms <span class="opt">(days)</span></label>
          <input type="number" id="billPaymentTerms" placeholder="e.g. 30" min="0" />
        </div>
        <div class="row2" id="fieldBillingExtras" style="display:none">
          <div class="field">
            <label>Hours report required <span class="req">*</span></label>
            <select id="billHoursReport">
              <option value="">&#8212; Select &#8212;</option>
              <option value="Yes">Yes</option>
              <option value="No">No</option>
            </select>
          </div>
          <div class="field">
            <label>Initiate invoice type <span class="req">*</span></label>
            <select id="billInvoiceType">
              <option value="">&#8212; Select &#8212;</option>
              <option value="Tax invoice">Tax invoice</option>
              <option value="Proforma invoice">Proforma invoice</option>
            </select>
          </div>
        </div>
        <div class="field" id="fieldBillingPeriod" style="display:none">
          <label>Billing Period <span class="req">*</span></label>
          <select id="billPeriod">
            <option value="Monthly">Monthly</option>
            <option value="Quarterly">Quarterly (×3 per record)</option>
            <option value="Yearly">Yearly (×12 per record)</option>
          </select>
        </div>
        <div class="field" id="fieldMilestonesCheckbox" style="display:none">
          <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
            <input type="checkbox" id="isMilestones" onchange="onMilestonesToggle()" style="width:auto;margin:0;accent-color:#1a237e" />
            Milestones billing
          </label>
        </div>
        <div class="field" id="fieldBillMonth" style="display:none">
          <label>Billing Month <span class="req">*</span></label>
          <input type="month" id="billMonth" />
        </div>
        <div id="fieldMilestoneSection" style="display:none">
          <div class="field">
            <label>Number of Milestones <span class="req">*</span></label>
            <input type="number" id="milestoneCount" min="1" max="20" placeholder="e.g. 3" oninput="onMilestoneCountChange()" />
          </div>
          <div id="milestoneRows"></div>
        </div>
      </div><!-- /poSection -->

      <div class="btn-row">
        <button id="btnSave" onclick="doSave()">&#128190; Save</button>
        <button id="btnCancel" onclick="google.script.host.close()">Cancel</button>
      </div>
      <div id="status"></div>
      <div id="emailStatus"></div>

      <script>
        var RENEWAL_REQUIRED_PROJECTS = ['recurring implementation', 'outsourcing', 'support'];

        // ── Currency state ────────────────────────────────────
        var currentCurrency = 'NIS';
        var currentRate = 1;

        function setCurrency(cur) {
          currentCurrency = cur;
          document.getElementById('btnNIS').className = 'currency-btn' + (cur === 'NIS' ? ' active-nis' : '');
          document.getElementById('btnUSD').className = 'currency-btn' + (cur === 'USD' ? ' active-usd' : '');
          updateCurrencySymbols();
          if (cur === 'USD') {
            document.getElementById('rateInfo').innerText = 'Fetching live USD/NIS rate\u2026';
            google.script.run
              .withSuccessHandler(function(rate) {
                currentRate = rate || 1;
                document.getElementById('rateInfo').innerText = rate
                  ? '1 USD = ' + rate.toFixed(4) + ' NIS'
                  : 'Rate unavailable \u2014 using 1:1';
                calcTotal();
              })
              .withFailureHandler(function() {
                currentRate = 1;
                document.getElementById('rateInfo').innerText = 'Rate unavailable \u2014 using 1:1';
                calcTotal();
              })
              .getUsdToNisRate();
          } else {
            currentRate = 1;
            document.getElementById('rateInfo').innerText = '';
            calcTotal();
          }
        }

        function updateCurrencySymbols() {
          var sym = currentCurrency === 'USD' ? '$' : '\u20aa';
          ['currSymbolAmount', 'currSymbolRecurring', 'currSymbolPricePerHour'].forEach(function(id) {
            var el = document.getElementById(id);
            if (el) el.innerText = sym;
          });
        }

        function togglePOSection() {
          document.getElementById('poSection').style.display =
            document.getElementById('skipPO').checked ? 'none' : 'block';
        }

        function isRenewalRequired() {
          var project     = document.getElementById('project').value.trim().toLowerCase();
          var billingType = document.getElementById('billingType').value;
          return RENEWAL_REQUIRED_PROJECTS.indexOf(project) !== -1 || billingType === 'Recurring';
        }

        function updateRenewalRequired() {
          var req = isRenewalRequired();
          document.getElementById('reqRenewalDate').style.display = req ? 'inline' : 'none';
          document.getElementById('optRenewalDate').style.display = req ? 'none'   : 'inline';
          if (!req) document.getElementById('renewalDate').classList.remove('error');
        }

        function isCommboxARRRequired() {
          return document.getElementById('project').value.trim().toLowerCase() === 'support';
        }

        function updateCommboxARRRequired() {
          var req = isCommboxARRRequired();
          document.getElementById('reqCommboxARR').style.display = req ? 'inline' : 'none';
          document.getElementById('optCommboxARR').style.display = req ? 'none'   : 'inline';
          if (!req) document.getElementById('commboxARR').classList.remove('error');
        }

        function onProjectChange() {
          updateRenewalRequired();
          updateCommboxARRRequired();
        }

        function onBillingTypeChange() {
          var bt   = document.getElementById('billingType').value;
          var hint = document.getElementById('billingHint');

          ['Amount','Recurring','Hours','PricePerHour'].forEach(function(f) {
            document.getElementById('req' + f).style.display = 'none';
            document.getElementById('opt' + f).style.display = 'inline';
            var id = f === 'Amount' ? 'amount' : f === 'Recurring' ? 'recurringAmount'
                   : f === 'Hours'  ? 'hours'  : 'pricePerHour';
            document.getElementById(id).classList.remove('error');
          });

          hint.className = 'billing-hint';
          hint.style.display = 'none';
          hint.innerText = '';
          document.getElementById('fieldRecurringPeriod').style.display = 'none';
          document.getElementById('recurringPeriod').selectedIndex = 0;
          document.getElementById('recurringPeriod').classList.remove('error');
          if (bt !== 'Recurring') document.getElementById('billPeriod').selectedIndex = 0;

          if (bt === 'Hourly') {
            document.getElementById('reqHours').style.display        = 'inline';
            document.getElementById('optHours').style.display        = 'none';
            document.getElementById('reqPricePerHour').style.display = 'inline';
            document.getElementById('optPricePerHour').style.display = 'none';
            hint.className = 'billing-hint hourly'; hint.style.display = 'block';
            hint.innerText = 'Hourly billing: Hours and Price per Hour are required.';
          } else if (bt === 'Fixed Project') {
            document.getElementById('reqAmount').style.display = 'inline';
            document.getElementById('optAmount').style.display = 'none';
            hint.className = 'billing-hint fixed'; hint.style.display = 'block';
            hint.innerText = 'Fixed Project billing: Amount is required.';
          } else if (bt === 'Recurring') {
            document.getElementById('reqRecurring').style.display = 'inline';
            document.getElementById('optRecurring').style.display = 'none';
            document.getElementById('fieldRecurringPeriod').style.display = 'block';
            hint.className = 'billing-hint recurring'; hint.style.display = 'block';
            hint.innerText = 'Recurring billing: Recurring Amount, Period, Start Date and Renewal Date are required.';
          }

          var sdRow = document.getElementById('fieldStartDate');
          if (sdRow) sdRow.style.display = (bt === 'Recurring') ? 'block' : 'none';

          updateRenewalRequired();
          updateBillingFields();
          calcTotal();
        }

        function calcTotal() {
          var bt     = document.getElementById('billingType').value;
          var amt    = parseFloat(document.getElementById('amount').value)          || 0;
          var rec    = parseFloat(document.getElementById('recurringAmount').value) || 0;
          var period = parseInt(document.getElementById('recurringPeriod').value)   || 0;
          var hrs    = parseFloat(document.getElementById('hours').value)           || 0;
          var pph    = parseFloat(document.getElementById('pricePerHour').value)    || 0;
          var totalInCurrency;
          if      (bt === 'Fixed Project') totalInCurrency = amt;
          else if (bt === 'Recurring')     totalInCurrency = rec * period;
          else if (bt === 'Hourly')        totalInCurrency = hrs * pph;
          else                             totalInCurrency = 0;

          var totalNIS = totalInCurrency * currentRate;
          document.getElementById('totalDisplay').innerText =
            '\u20aa' + totalNIS.toLocaleString('he-IL', {minimumFractionDigits:2, maximumFractionDigits:2});

          var convDiv = document.getElementById('totalConversion');
          if (currentCurrency === 'USD' && totalInCurrency > 0 && currentRate > 1) {
            var convText;
            if (bt === 'Recurring' && period > 0) {
              convText = '(from $' + rec.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2})
                + '/month \u00d7 ' + period + ' months = $'
                + totalInCurrency.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2})
                + ', at rate ' + currentRate.toFixed(4) + ')';
            } else {
              convText = '(converted from $'
                + totalInCurrency.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2})
                + ' at rate ' + currentRate.toFixed(4) + ')';
            }
            convDiv.innerText = convText;
            convDiv.style.display = 'block';
          } else {
            convDiv.style.display = 'none';
          }
          return totalNIS;
        }

        function setStatus(msg, type) {
          var el = document.getElementById('status');
          el.innerText = msg; el.className = type || '';
        }

        function clearErrors() {
          ['orgName','industry','activity','manager','firstName','lastName','contactEmail',
           'projectName','project','poNumber','billingType','amount','recurringAmount','recurringPeriod',
           'hours','pricePerHour','startDate','renewalDate','commboxARR',
           'billPayingCustomer','billPaymentTerms','billHoursReport','billInvoiceType','billMonth','milestoneCount']
            .forEach(function(id) {
              var el = document.getElementById(id);
              if (el) el.classList.remove('error');
            });
        }

        // ── Billing handlers ─────────────────────────────────
        function onBillCustomerChange() {
          var sel = document.getElementById('billPayingCustomer');
          var pt  = (sel && sel.selectedIndex > 0)
            ? (sel.options[sel.selectedIndex].getAttribute('data-pt') || '')
            : '';
          var ptEl = document.getElementById('billPaymentTerms');
          if (ptEl) ptEl.value = pt;
          updateBillingFields();
        }

        function updateBillingFields() {
          var bt = document.getElementById('billingType').value;
          var hasCustomer = !!document.getElementById('billPayingCustomer').value;
          var isHourlyOrFixed = (bt === 'Fixed Project' || bt === 'Hourly');
          var isMsCb = document.getElementById('isMilestones');
          var isMilestones = hasCustomer && isHourlyOrFixed && isMsCb && isMsCb.checked;

          document.getElementById('fieldPaymentTerms').style.display =
            hasCustomer ? 'block' : 'none';
          document.getElementById('fieldBillingExtras').style.display =
            hasCustomer ? 'flex' : 'none';
          document.getElementById('fieldBillingPeriod').style.display =
            (hasCustomer && bt === 'Recurring') ? 'block' : 'none';
          document.getElementById('fieldMilestonesCheckbox').style.display =
            (hasCustomer && isHourlyOrFixed) ? 'block' : 'none';
          document.getElementById('fieldBillMonth').style.display =
            (hasCustomer && isHourlyOrFixed && !isMilestones) ? 'block' : 'none';
          document.getElementById('fieldMilestoneSection').style.display =
            isMilestones ? 'block' : 'none';

          if (!isMilestones) {
            var mc = document.getElementById('milestoneCount');
            if (mc) mc.value = '';
            document.getElementById('milestoneRows').innerHTML = '';
          }
          if (!hasCustomer || !isHourlyOrFixed) {
            if (isMsCb) isMsCb.checked = false;
          }
        }

        function onMilestonesToggle() { updateBillingFields(); }

        function onMilestoneCountChange() {
          var n = parseInt(document.getElementById('milestoneCount').value) || 0;
          var html = '';
          for (var i = 1; i <= n; i++) html += buildMilestoneRow(i);
          document.getElementById('milestoneRows').innerHTML = html;
        }

        function buildMilestoneRow(i) {
          var bt = document.getElementById('billingType').value;
          var isHourly = (bt === 'Hourly');
          var hoursField = isHourly
            ? '<div class="field"><label>Hours (auto)</label>' +
                '<input type="text" id="msHours_' + i + '" readonly ' +
                'style="background:#f0f4ff;color:#1a237e;font-weight:600" /></div>'
            : '<div class="field"></div>';
          return '<div style="border:1px solid #c5cae9;border-radius:6px;padding:10px;margin-bottom:8px;background:#f8f9ff">' +
            '<div style="font-size:11px;font-weight:700;color:#5c6bc0;margin-bottom:6px">Milestone ' + i + '</div>' +
            '<div class="row2">' +
              '<div class="field"><label>Name <span class="req">*</span></label><input type="text" id="msName_' + i + '" /></div>' +
              '<div class="field"><label>% of total <span class="req">*</span></label>' +
                '<input type="number" id="msPct_' + i + '" min="0.01" max="100" step="0.01" placeholder="e.g. 25" ' +
                'oninput="calcMs(' + i + ')" /></div>' +
            '</div>' +
            '<div class="row2">' +
              '<div class="field"><label>Amount &#8362; (auto)</label>' +
                '<input type="text" id="msAmount_' + i + '" readonly ' +
                'style="background:#f0f4ff;color:#1a237e;font-weight:600" /></div>' +
              hoursField +
            '</div>' +
            '<div class="field"><label>Description <span class="opt">(optional)</span></label><input type="text" id="msDesc_' + i + '" /></div>' +
            '<div class="field"><label>Billing Month <span class="req">*</span></label><input type="month" id="msMonth_' + i + '" /></div>' +
          '</div>';
        }

        function calcMs(i) {
          var pct   = parseFloat((document.getElementById('msPct_' + i) || {value:'0'}).value) || 0;
          var total = calcTotal();
          var amt   = Math.round(total * pct / 100 * 100) / 100;
          var amtEl = document.getElementById('msAmount_' + i);
          if (amtEl) amtEl.value = amt > 0 ? amt.toFixed(2) : '';
          var hrsEl = document.getElementById('msHours_' + i);
          if (hrsEl) {
            var hrs   = parseFloat((document.getElementById('hours') || {value:'0'}).value) || 0;
            var msHrs = Math.round(hrs * pct / 100 * 100) / 100;
            hrsEl.value = msHrs > 0 ? msHrs.toFixed(2) : '';
          }
        }

        function recalcAllMs() {
          var n = parseInt((document.getElementById('milestoneCount') || {value:'0'}).value) || 0;
          for (var k = 1; k <= n; k++) {
            if (document.getElementById('msPct_' + k)) calcMs(k);
          }
        }

        function doSave() {
          clearErrors();
          var skipPO = document.getElementById('skipPO').checked;

          // Customer fields
          var orgName         = document.getElementById('orgName').value.trim();
          var industry        = document.getElementById('industry').value;
          var activity        = document.getElementById('activity').value;
          var manager         = document.getElementById('manager').value;
          var firstName       = document.getElementById('firstName').value.trim();
          var lastName        = document.getElementById('lastName').value.trim();
          var contactEmail    = document.getElementById('contactEmail').value.trim();
          var contactPosition = document.getElementById('contactPosition').value.trim();
          var contactPhone    = document.getElementById('contactPhone').value.trim();
          var contactDob      = document.getElementById('contactDob').value;
          var contactAddress  = document.getElementById('contactAddress').value.trim();

          var valid = true;
          if (!orgName)      { document.getElementById('orgName').classList.add('error');      valid = false; }
          if (!industry)     { document.getElementById('industry').classList.add('error');     valid = false; }
          if (!activity)     { document.getElementById('activity').classList.add('error');     valid = false; }
          if (!manager)      { document.getElementById('manager').classList.add('error');      valid = false; }
          if (!firstName)    { document.getElementById('firstName').classList.add('error');    valid = false; }
          if (!lastName)     { document.getElementById('lastName').classList.add('error');     valid = false; }
          if (!contactEmail) { document.getElementById('contactEmail').classList.add('error'); valid = false; }

          // PO fields
          var projectName = '', project = '', poNumber = 0, billingType = '', projectDesc = '', poCustomer = '',
              proposalLink = '', amount = 0, recurringAmount = 0, recurringPeriod = 0,
              hours = 0, pricePerHour = 0, milestones = '', renewalDate = '', commboxARR = 0,
              total = 0, currency = 'NIS', exchangeRate = 1;
          // Billing fields
          var startDate = '', billPayingCustomerId = '', billPayingCustomer = '',
              billMonth = '', billPaymentTerms = '', billHoursReport = '', billInvoiceType = '',
              billPeriod = 'Monthly', isMilestones = false, billMilestones = [];

          if (!skipPO) {
            projectName     = document.getElementById('projectName').value.trim();
            project         = document.getElementById('project').value;
            poNumber        = parseInt(document.getElementById('poNumber').value, 10) || 0;
            billingType     = document.getElementById('billingType').value;
            projectDesc     = document.getElementById('projectDescription').value.trim();
            poCustomer      = document.getElementById('poCustomer').value.trim();
            proposalLink    = document.getElementById('proposalLink').value.trim();
            amount          = parseFloat(document.getElementById('amount').value)          || 0;
            recurringAmount = parseFloat(document.getElementById('recurringAmount').value) || 0;
            recurringPeriod = parseInt(document.getElementById('recurringPeriod').value)   || 0;
            hours           = parseFloat(document.getElementById('hours').value)           || 0;
            pricePerHour    = parseFloat(document.getElementById('pricePerHour').value)    || 0;
            milestones      = document.getElementById('milestones').value.trim();
            renewalDate     = document.getElementById('renewalDate').value;
            commboxARR      = parseFloat(document.getElementById('commboxARR').value)      || 0;
            total           = calcTotal();
            currency        = currentCurrency;
            exchangeRate    = currentCurrency === 'USD' ? currentRate : 1;

            if (!projectName) { document.getElementById('projectName').classList.add('error'); valid = false; }
            if (!project)     { document.getElementById('project').classList.add('error');     valid = false; }
            if (!poNumber)    { document.getElementById('poNumber').classList.add('error');    valid = false; }
            if (!billingType) { document.getElementById('billingType').classList.add('error'); valid = false; }

            if (isRenewalRequired() && !renewalDate)   { document.getElementById('renewalDate').classList.add('error');  valid = false; }
            if (isCommboxARRRequired() && !commboxARR) { document.getElementById('commboxARR').classList.add('error');   valid = false; }

            if (billingType === 'Hourly') {
              if (!hours)        { document.getElementById('hours').classList.add('error');        valid = false; }
              if (!pricePerHour) { document.getElementById('pricePerHour').classList.add('error'); valid = false; }
            } else if (billingType === 'Fixed Project') {
              if (!amount) { document.getElementById('amount').classList.add('error'); valid = false; }
            } else if (billingType === 'Recurring') {
              if (!recurringAmount) { document.getElementById('recurringAmount').classList.add('error'); valid = false; }
              if (!recurringPeriod) { document.getElementById('recurringPeriod').classList.add('error'); valid = false; }
            }

            // Billing fields (required)
            var billSel = document.getElementById('billPayingCustomer');
            billPayingCustomerId = billSel ? billSel.value : '';
            billPayingCustomer   = (billSel && billSel.selectedIndex > 0) ? billSel.options[billSel.selectedIndex].text : '';
            startDate        = (document.getElementById('startDate')        || {value:''}).value;
            billMonth        = (document.getElementById('billMonth')        || {value:''}).value;
            billPaymentTerms = (document.getElementById('billPaymentTerms') || {value:''}).value.trim();
            billHoursReport  = (document.getElementById('billHoursReport')  || {value:''}).value;
            billInvoiceType  = (document.getElementById('billInvoiceType')  || {value:''}).value;
            billPeriod       = (document.getElementById('billPeriod')       || {value:'Monthly'}).value || 'Monthly';
            var isMsCb = document.getElementById('isMilestones');
            isMilestones   = !!(isMsCb && isMsCb.checked);
            billMilestones = [];

            if (!billPayingCustomerId) {
              document.getElementById('billPayingCustomer').classList.add('error'); valid = false;
            }
            if (!billHoursReport) {
              document.getElementById('billHoursReport').classList.add('error'); valid = false;
            }
            if (!billInvoiceType) {
              document.getElementById('billInvoiceType').classList.add('error'); valid = false;
            }
            if (billPayingCustomerId) {
              if (billingType === 'Recurring' && !startDate) {
                document.getElementById('startDate').classList.add('error'); valid = false;
              }
              if ((billingType === 'Fixed Project' || billingType === 'Hourly') && !isMilestones && !billMonth) {
                document.getElementById('billMonth').classList.add('error'); valid = false;
              }
              if (isMilestones) {
                var msCount = parseInt((document.getElementById('milestoneCount') || {value:'0'}).value) || 0;
                if (!msCount) { document.getElementById('milestoneCount').classList.add('error'); valid = false; }
                for (var mi = 1; mi <= msCount; mi++) {
                  var msName   = (document.getElementById('msName_'  + mi) || {value:''}).value.trim();
                  var msPct    = parseFloat((document.getElementById('msPct_' + mi) || {value:'0'}).value) || 0;
                  var msAmt    = parseFloat((document.getElementById('msAmount_' + mi) || {value:'0'}).value) || 0;
                  var msDesc   = (document.getElementById('msDesc_'  + mi) || {value:''}).value.trim();
                  var msMo     = (document.getElementById('msMonth_' + mi) || {value:''}).value;
                  var msHoursEl = document.getElementById('msHours_' + mi);
                  var msHours  = msHoursEl ? (parseFloat(msHoursEl.value) || 0) : 0;
                  if (!msName) { var el = document.getElementById('msName_' + mi); if (el) el.classList.add('error'); valid = false; }
                  if (!msPct)  { var el = document.getElementById('msPct_'  + mi); if (el) el.classList.add('error'); valid = false; }
                  if (!msMo)   { var el = document.getElementById('msMonth_'+ mi); if (el) el.classList.add('error'); valid = false; }
                  billMilestones.push({
                    name:         msName,
                    amount:       msAmt,
                    description:  msDesc,
                    month:        msMo,
                    percentage:   msPct,
                    pricePerHour: billingType === 'Hourly' ? pricePerHour : 0,
                    hours:        msHours,
                  });
                }
              }
            }
          }

          if (!valid) { setStatus('\u26a0\ufe0f Please fill in all required fields.', 'err'); return; }

          setStatus('\u23f3 Saving\u2026');
          document.getElementById('emailStatus').innerText = '';
          document.getElementById('emailStatus').className = '';
          document.getElementById('btnSave').disabled = true;

          google.script.run
            .withSuccessHandler(function(result) {
              document.getElementById('btnSave').disabled = false;
              if (result.ok) {
                var es = document.getElementById('emailStatus');
                if (result.emailSent) {
                  es.innerText = '\u2709\ufe0f Email sent to: ' + result.emailRecipients;
                  es.className = 'ok';
                } else {
                  es.innerText = '\u26a0\ufe0f Email not sent: ' + (result.emailError || 'unknown error');
                  es.className = 'err';
                }
                setStatus('\u2705 ' + result.msg, 'ok');
                // Reset all fields
                ['orgName','contactEmail','contactPosition','contactPhone','contactAddress',
                 'firstName','lastName','projectName','projectDescription','poCustomer','proposalLink',
                 'milestones'].forEach(function(id) {
                  var el = document.getElementById(id);
                  if (el) el.value = '';
                });
                ['amount','recurringAmount','hours','pricePerHour','commboxARR','poNumber'].forEach(function(id) {
                  var el = document.getElementById(id);
                  if (el) el.value = '';
                });
                ['industry','activity','manager','project','billingType'].forEach(function(id) {
                  var el = document.getElementById(id);
                  if (el) el.selectedIndex = 0;
                });
                document.getElementById('contactDob').value = '';
                document.getElementById('renewalDate').value = '';
                document.getElementById('recurringPeriod').selectedIndex = 0;
                document.getElementById('fieldRecurringPeriod').style.display = 'none';
                document.getElementById('billingHint').style.display = 'none';
                document.getElementById('totalDisplay').innerText = '\u20aa0.00';
                document.getElementById('skipPO').checked = false;
                document.getElementById('poSection').style.display = 'block';
                // Reset billing section
                var bpc = document.getElementById('billPayingCustomer');
                if (bpc) bpc.selectedIndex = 0;
                var ptEl = document.getElementById('billPaymentTerms');
                if (ptEl) ptEl.value = '';
                var hrEl = document.getElementById('billHoursReport');
                if (hrEl) hrEl.selectedIndex = 0;
                var itEl = document.getElementById('billInvoiceType');
                if (itEl) itEl.selectedIndex = 0;
                var bpEl = document.getElementById('billPeriod');
                if (bpEl) bpEl.selectedIndex = 0;
                var sdEl = document.getElementById('startDate');
                if (sdEl) sdEl.value = '';
                var bmEl = document.getElementById('billMonth');
                if (bmEl) bmEl.value = '';
                var msCbEl = document.getElementById('isMilestones');
                if (msCbEl) msCbEl.checked = false;
                var mcEl = document.getElementById('milestoneCount');
                if (mcEl) mcEl.value = '';
                var mrEl = document.getElementById('milestoneRows');
                if (mrEl) mrEl.innerHTML = '';
                updateBillingFields();
                // Reset currency
                currentCurrency = 'NIS'; currentRate = 1;
                document.getElementById('btnNIS').className = 'currency-btn active-nis';
                document.getElementById('btnUSD').className = 'currency-btn';
                document.getElementById('rateInfo').innerText = '';
                updateCurrencySymbols();
              } else {
                setStatus('\u274c ' + result.msg, 'err');
              }
            })
            .withFailureHandler(function(e) {
              document.getElementById('btnSave').disabled = false;
              setStatus('\u274c Error: ' + e.message, 'err');
            })
            .saveNewCustomer({
              orgName: orgName, industry: industry, activity: activity, manager: manager,
              firstName: firstName, lastName: lastName, contactEmail: contactEmail,
              contactPosition: contactPosition, contactPhone: contactPhone, contactDob: contactDob, contactAddress: contactAddress,
              skipPO: skipPO,
              projectName: projectName, project: project, poNumber: poNumber, billingType: billingType,
              projectDescription: projectDesc, poCustomer: poCustomer, proposalLink: proposalLink,
              amount: amount, recurringAmount: recurringAmount, recurringPeriod: recurringPeriod,
              hours: hours, pricePerHour: pricePerHour, milestones: milestones,
              renewalDate: renewalDate, commboxARR: commboxARR || '', totalAmount: total,
              currency: currency, exchangeRate: exchangeRate,
              startDate: startDate, isMilestones: isMilestones,
              billPayingCustomer: billPayingCustomer, billPayingCustomerId: billPayingCustomerId,
              billMonth: billMonth, billMilestones: billMilestones,
              billPaymentTerms: billPaymentTerms,
              billHoursReport: billHoursReport, billInvoiceType: billInvoiceType,
              billPeriod: billPeriod,
            });
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(500).setHeight(1000).setTitle("New Customer");

  SpreadsheetApp.getUi().showModalDialog(html, "New Customer");
}


// ── SAVE NEW CUSTOMER ──────────────────────────────────────
function saveNewCustomer(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const custSheet = ss.getSheetByName(SRC.CUSTOMERS);
    if (!custSheet) return { ok: false, msg: 'Sheet "' + SRC.CUSTOMERS + '" not found.' };

    // Check for duplicate org name
    const allRows   = custSheet.getDataRange().getValues();
    const headers   = allRows[0];
    const orgColIdx = headers.findIndex(h => String(h).trim().replace(/\s+/g, ' ') === 'Organization' || String(h).trim() === 'Organization ');
    const duplicate = allRows.slice(1).find(r => norm(String(r[orgColIdx] || "")) === norm(data.orgName));
    if (duplicate) return { ok: false, msg: '"' + data.orgName + '" already exists in Customers.' };

    // Append to Customers
    custSheet.appendRow(buildRow(headers, {
      "Organization ":       data.orgName,
      "Industry ":           data.industry,
      "Main activity type ": data.activity,
      "Customer Manager ":   data.manager,
    }));

    // Append to Contacts
    if (data.firstName && data.lastName) {
      const contactSheet = ss.getSheetByName(SRC.CONTACTS);
      if (contactSheet) {
        const contactHeaders = contactSheet.getRange(1, 1, 1, contactSheet.getLastColumn()).getValues()[0];
        let dobFormatted = "";
        if (data.contactDob) {
          const parts = data.contactDob.split("-"); // yyyy-mm-dd → dd/mm/yyyy
          if (parts.length === 3) dobFormatted = parts[2] + "/" + parts[1] + "/" + parts[0];
        }
        contactSheet.appendRow(buildRow(contactHeaders, {
          "First Name":    data.firstName,
          "Last Name ":    data.lastName,
          "Date of Birth": dobFormatted,
          "Postion ":      data.contactPosition  || "",
          "Organization ": data.orgName,
          "Phone Number":  data.contactPhone     || "",
          "Email":         data.contactEmail     || "",
          "Address":       data.contactAddress   || "",
        }));
      }
    }

    // Append to Purchase Orders (if not skipped)
    var poSaved = false, billingCount = 0, billingError = '';
    if (!data.skipPO && data.poNumber) {
      const poSheet = ss.getSheetByName("Purchase Orders");
      if (poSheet) {
        const today = new Date();
        const dd    = String(today.getDate()).padStart(2, "0");
        const mm    = String(today.getMonth() + 1).padStart(2, "0");
        const yyyy  = today.getFullYear();
        const dateStr = dd + "/" + mm + "/" + yyyy;

        let startFormatted = "";
        if (data.startDate) {
          const parts = data.startDate.split("-");
          if (parts.length === 3) startFormatted = parts[2] + "/" + parts[1] + "/" + parts[0];
        }
        let renewalFormatted = "";
        if (data.renewalDate) {
          const parts = data.renewalDate.split("-");
          if (parts.length === 3) renewalFormatted = parts[2] + "/" + parts[1] + "/" + parts[0];
        }

        // Ensure all column headers exist (cols 14–22)
        if (!poSheet.getRange(1, 14).getValue()) poSheet.getRange(1, 14).setValue("Billing Type");
        if (!poSheet.getRange(1, 15).getValue()) poSheet.getRange(1, 15).setValue("Start Date");
        if (!poSheet.getRange(1, 16).getValue()) poSheet.getRange(1, 16).setValue("Renwal Date");
        if (!poSheet.getRange(1, 17).getValue()) poSheet.getRange(1, 17).setValue("Commbox ARR");
        if (!poSheet.getRange(1, 18).getValue()) poSheet.getRange(1, 18).setValue("Recurring Period");
        if (!poSheet.getRange(1, 19).getValue()) poSheet.getRange(1, 19).setValue("Currency");
        if (!poSheet.getRange(1, 20).getValue()) poSheet.getRange(1, 20).setValue("Exchange Rate");
        if (!poSheet.getRange(1, 21).getValue()) poSheet.getRange(1, 21).setValue("Project Name");
        if (!poSheet.getRange(1, 22).getValue()) poSheet.getRange(1, 22).setValue("Commbox Hourly Rate");

        poSheet.appendRow([
          dateStr,                       // col 1  — PO Date
          data.orgName,                  // col 2  — Organization
          data.project,                  // col 3  — Project
          data.projectDescription,       // col 4  — Project description
          data.poNumber,                 // col 5  — PO number
          data.amount          || "",    // col 6  — Amount
          data.recurringAmount || "",    // col 7  — Recurring Amount
          data.milestones,               // col 8  — Milestones
          data.hours           || "",    // col 9  — Hours
          data.pricePerHour    || "",    // col 10 — Price per hour
          data.totalAmount     || "",    // col 11 — Total Amount
          data.poCustomer      || "",    // col 12 — Customer
          data.proposalLink    || "",    // col 13 — Proposal link
          data.billingType     || "",    // col 14 — Billing Type
          startFormatted,                // col 15 — Start Date
          renewalFormatted,              // col 16 — Renwal Date
          data.commboxARR      || "",    // col 17 — Commbox ARR
          data.recurringPeriod || "",    // col 18 — Recurring Period
          data.currency        || "NIS", // col 19 — Currency
          data.exchangeRate    || 1,     // col 20 — Exchange Rate
          data.projectName     || "",    // col 21 — Project Name
          data.commboxHourlyRate || "",  // col 22 — Commbox Hourly Rate
        ]);

        poSaved = true;

        // Save billing records
        if (data.billingType && data.billPayingCustomerId) {
          const br = saveBillingRecords({
            organization:        data.orgName,
            poNumber:            data.poNumber         || '',
            billingType:         data.billingType,
            isMilestones:        data.isMilestones     || false,
            billPayingCustomer:  data.billPayingCustomer,
            billPayingCustomerId: data.billPayingCustomerId,
            startDate:           data.startDate        || '',
            billMonth:           data.billMonth        || '',
            renewalDate:         data.renewalDate      || '',
            recurringAmount:     data.recurringAmount  || 0,
            totalAmount:         data.totalAmount      || 0,
            amount:              data.amount           || 0,
            billMilestones:      data.billMilestones   || [],
            billPaymentTerms:    data.billPaymentTerms || '',
            project:             data.project          || '',
            billHoursReport:     data.billHoursReport  || '',
            billInvoiceType:     data.billInvoiceType  || '',
            billPeriod:          data.billPeriod       || 'Monthly',
            pricePerHour:        data.pricePerHour     || 0,
            hours:               data.hours            || 0,
          });
          billingCount = br.count || 0;
          billingError = br.ok ? '' : (br.error || 'Unknown billing error');
        }
      }
    }

    // Auto-search the new customer so results appear immediately
    runSearch(data.orgName);

    // Send notification email
    var emailSent = false, emailError = '', emailRecipients = '';
    try {
      const managerEmail = managerNameToEmail(data.manager);
      emailRecipients    = buildToList(managerEmail);
      const subject = 'New Customer: ' + data.orgName;
      var body =
        'A new customer has been added to the CRM.\n\n' +
        '-- Organization --\n' +
        'Name:             ' + data.orgName  + '\n' +
        'Industry:         ' + data.industry + '\n' +
        'Activity Type:    ' + data.activity + '\n' +
        'Customer Manager: ' + data.manager  + '\n\n' +
        '-- Contact Details --\n' +
        'Name:     ' + data.firstName + ' ' + data.lastName + '\n' +
        'Email:    ' + (data.contactEmail    || '-') + '\n' +
        'Position: ' + (data.contactPosition || '-') + '\n' +
        'Phone:    ' + (data.contactPhone    || '-') + '\n';
      if (poSaved) {
        body +=
          '\n-- Purchase Order --\n' +
          'PO Number:        ' + data.poNumber          + '\n' +
          'Project:          ' + data.project           + '\n' +
          'Billing Type:     ' + (data.billingType      || '-') + '\n' +
          'Amount:           ' + (data.amount           || '-') + '\n' +
          'Recurring Amount: ' + (data.recurringAmount  || '-') + '\n' +
          'Recurring Period: ' + (data.recurringPeriod  || '-') + '\n' +
          'Hours:            ' + (data.hours            || '-') + '\n' +
          'Price per Hour:   ' + (data.pricePerHour     || '-') + '\n' +
          'Total Amount:     ' + (data.totalAmount      || '-') + '\n';
      }
      MailApp.sendEmail(emailRecipients, subject, body);
      emailSent = true;
    } catch(mailErr) {
      emailError = mailErr.message;
    }

    const billingMsg = billingCount > 0 ? ' ' + billingCount + ' billing record(s) added.' : '';
    const msg = poSaved
      ? '"' + data.orgName + '" and PO #' + data.poNumber + ' saved successfully!' + billingMsg
      : '"' + data.orgName + '" saved successfully!';

    return { ok: true, msg: msg, emailSent: emailSent, emailError: emailError, emailRecipients: emailRecipients, billingError: billingError };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}


// ── HELPER: build a row array aligned to sheet headers ─────
function buildRow(headers, fieldMap) {
  return headers.map(h => {
    const key = String(h).trim();
    if (fieldMap.hasOwnProperty(key)) return fieldMap[key];
    const found = Object.keys(fieldMap).find(k => k.trim() === key);
    return found !== undefined ? fieldMap[found] : "";
  });
}


// ── HELPER: escape HTML special characters ─────────────────
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
