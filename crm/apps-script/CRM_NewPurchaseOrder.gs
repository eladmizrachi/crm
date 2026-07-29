// ============================================================
// CRM_NewPurchaseOrder.gs — New Purchase Order Creation
// Depends on: CRM_Setup.gs (for SHEETS, norm, runSearch)
// Add this file to the same Apps Script project.
// ============================================================


// ── LOAD EXISTING ORGANIZATIONS FROM CUSTOMERS SHEET ───────
function getPOOrganizations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Customers");
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const orgIdx  = headers.findIndex(h => String(h).trim().replace(/\s+/g,"") === "Organization");

  return data.slice(1)
    .map(row => String(row[orgIdx] || "").trim())
    .filter(v => v !== "")
    .sort();
}


// ── LOAD ACTIVITY VALUES FROM CODETABLES (Onboarding column) ─
function getPOActivities() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("CODETABLES");
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const actIdx = headers.findIndex(h => String(h).trim().toLowerCase().startsWith("onboarding"));
  if (actIdx === -1) return [];

  return data.slice(1)
    .map(row => String(row[actIdx] || "").trim())
    .filter(v => v !== "")
    .sort();
}


// ── OPEN NEW PURCHASE ORDER DIALOG ────────────────────────
// Called from menu (no arg) or after Save & New PO (org pre-selected).
function openNewPurchaseOrderDialogForOrg(orgName) {
  openNewPurchaseOrderDialog(orgName || '');
}

function openNewPurchaseOrderDialog(selectedOrg) {
  selectedOrg = selectedOrg || '';
  const orgs       = getPOOrganizations();
  const activities = getPOActivities();
  const giResult   = getGreenInvoiceClients();
  const giClients  = giResult.clients || [];
  const giError    = giResult.ok ? '' : (giResult.error || 'Could not load billing customers');

  function buildOptions(arr, selected) {
    selected = selected || '';
    return arr.map(v => `<option value="${escapeHtmlPO(v)}"${v === selected ? ' selected' : ''}>${escapeHtmlPO(v)}</option>`).join('');
  }

  function buildGIOptions(clients) {
    return clients.map(c => `<option value="${escapeHtmlPO(c.id)}" data-pt="${escapeHtmlPO(c.paymentTerms || '')}">${escapeHtmlPO(c.name)}</option>`).join('');
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

        input[type=text], input[type=number], input[type=url], select, textarea {
          width: 100%; padding: 8px 10px; font-size: 13px;
          border: 1.5px solid #9fa8da; border-radius: 6px;
          background: #fff; color: #212121; outline: none;
          transition: border-color .15s;
        }
        textarea { resize: vertical; min-height: 56px; font-family: inherit; }
        input:focus, select:focus, textarea:focus { border-color: #3949ab; }
        input.error, select.error { border-color: #c62828 !important; }

        /* Currency selector */
        .currency-row {
          display: flex; gap: 8px; margin-bottom: 10px;
        }
        .currency-btn {
          flex: 1; padding: 7px 10px; font-size: 13px; font-weight: 700;
          border: 2px solid #9fa8da; border-radius: 6px;
          background: #fff; color: #555; cursor: pointer; text-align: center;
          transition: all .15s;
        }
        .currency-btn.active-nis { border-color: #1a237e; background: #1a237e; color: #fff; }
        .currency-btn.active-usd { border-color: #2e7d32; background: #2e7d32; color: #fff; }
        .rate-info {
          font-size: 11px; color: #5c6bc0; text-align: center;
          margin-bottom: 8px; min-height: 16px;
        }

        /* Billing type hint banner */
        .billing-hint {
          display: none;
          font-size: 11.5px; font-weight: 600;
          padding: 7px 10px; border-radius: 5px;
          margin-bottom: 8px;
        }
        .billing-hint.hourly    { background: #e3f2fd; color: #1565c0; border: 1px solid #90caf9; }
        .billing-hint.fixed     { background: #fff8e1; color: #f57f17; border: 1px solid #ffe082; }
        .billing-hint.recurring  { background: #f3e5f5; color: #6a1b9a; border: 1px solid #ce93d8; }
        .billing-hint.milestones { background: #e8f5e9; color: #2e7d32; border: 1px solid #81c784; }

        .total-box {
          background: #e8f5e9; border: 1.5px solid #81c784;
          border-radius: 6px; padding: 10px 14px;
          margin-top: 4px; display: flex;
          align-items: center; justify-content: space-between;
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
      <h2>🛒 New Purchase Order</h2>
      <p class="sub">Create a new purchase order and link it to an existing customer.</p>

      <!-- Organization -->
      <div class="section-title">Organization</div>
      <div class="field">
        <label>Organization <span class="req">*</span></label>
        <select id="organization">
          <option value="">— Select Customer —</option>
          ${buildOptions(orgs, selectedOrg)}
        </select>
      </div>

      <!-- Order Details -->
      <div class="section-title">Order Details</div>
      <div class="field">
        <label>Project Name <span class="req">*</span></label>
        <input type="text" id="projectName" placeholder="e.g. CRM Implementation" />
      </div>
      <div class="row2">
        <div class="field">
          <label>Project <span class="req">*</span></label>
          <select id="project" onchange="onProjectChange()">
            <option value="">— Select Activity —</option>
            ${buildOptions(activities, '')}
          </select>
        </div>
        <div class="field">
          <label>PO Number <span class="req">*</span></label>
          <input type="number" id="poNumber" placeholder="e.g. 1001" min="1" />
        </div>
      </div>

      <!-- Billing Type -->
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
          <input type="text" id="customer" placeholder="e.g. Commbox" />
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

      <!-- Billing hint -->
      <div id="billingHint" class="billing-hint"></div>

      <!-- Amount — mandatory for Fixed Project -->
      <div class="field" id="fieldAmount">
        <label>
          Amount (<span id="currSymbolAmount">&#8362;</span>)
          <span class="req-dynamic" id="reqAmount">*</span>
          <span class="opt" id="optAmount">(optional)</span>
        </label>
        <input type="number" id="amount" placeholder="0" min="0" step="0.01" oninput="calcTotal()" />
      </div>

      <!-- Recurring Amount — mandatory for Recurring -->
      <div class="field" id="fieldRecurring">
        <label>
          Recurring Amount (<span id="currSymbolRecurring">&#8362;</span>)
          <span class="req-dynamic" id="reqRecurring">*</span>
          <span class="opt" id="optRecurring">(optional)</span>
        </label>
        <input type="number" id="recurringAmount" placeholder="0" min="0" step="0.01" oninput="calcTotal()" />
      </div>

      <!-- Recurring Period — shown when Billing Type = Recurring -->
      <div class="field" id="fieldRecurringPeriod" style="display:none">
        <label>Recurring Period (months) <span class="req-dynamic" id="reqRecurringPeriod" style="display:inline">*</span></label>
        <select id="recurringPeriod" onchange="calcTotal()">
          <option value="">— Select Period —</option>
          <option value="1">1</option>
          <option value="2">2</option>
          <option value="3">3</option>
          <option value="4">4</option>
          <option value="5">5</option>
          <option value="6">6</option>
          <option value="7">7</option>
          <option value="8">8</option>
          <option value="9">9</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
        </select>
      </div>

      <!-- Hours + Price per Hour — mandatory for Hourly -->
      <div class="row2">
        <div class="field" id="fieldHours">
          <label>
            Hours
            <span class="req-dynamic" id="reqHours">*</span>
            <span class="opt" id="optHours">(optional)</span>
          </label>
          <input type="number" id="hours" placeholder="0" min="0" step="0.5" oninput="calcTotal()" />
        </div>
        <div class="field" id="fieldPricePerHour">
          <label>
            Price per Hour (<span id="currSymbolPricePerHour">&#8362;</span>)
            <span class="req-dynamic" id="reqPricePerHour">*</span>
            <span class="opt" id="optPricePerHour">(optional)</span>
          </label>
          <input type="number" id="pricePerHour" placeholder="0" min="0" step="0.01" oninput="calcTotal()" />
        </div>
      </div>

      <div class="field">
        <label>Milestones <span class="opt">(optional)</span></label>
        <input type="text" id="milestones" placeholder="e.g. Phase 1, Phase 2, Phase 3" />
      </div>

      <!-- Start Date — shown for Recurring billing; stored in PO sheet -->
      <div class="field" id="fieldStartDate" style="display:none">
        <label>Start Date <span class="req" id="reqStartDate">*</span></label>
        <input type="date" id="startDate" />
      </div>

      <!-- Renewal Date — required for Recurring billing or certain project types -->
      <div class="field" id="fieldRenewalDate">
        <label>
          Renewal Date
          <span class="req-dynamic" id="reqRenewalDate">*</span>
          <span class="opt" id="optRenewalDate">(optional)</span>
        </label>
        <input type="date" id="renewalDate" />
      </div>

      <!-- Commbox ARR — required when project is Support -->
      <div class="field">
        <label>
          Commbox ARR (&#8362;)
          <span class="req-dynamic" id="reqCommboxARR">*</span>
          <span class="opt" id="optCommboxARR">(optional)</span>
        </label>
        <input type="number" id="commboxARR" placeholder="0" min="0" step="0.01" />
      </div>

      <!-- Commbox Hourly Rate — required when project is Support -->
      <div class="field">
        <label>
          Commbox Hourly Rate (<span id="currSymbolHourlyRate">&#8362;</span>)
          <span class="req-dynamic" id="reqCommboxHourlyRate">*</span>
          <span class="opt" id="optCommboxHourlyRate">(optional)</span>
        </label>
        <input type="number" id="commboxHourlyRate" placeholder="0" min="0" step="0.01" />
      </div>

      <!-- Total -->
      <div class="total-box" style="flex-direction:column; align-items:flex-start; gap:4px;">
        <div style="display:flex; justify-content:space-between; width:100%;">
          <span class="total-label">&#128176; Total Amount</span>
          <span class="total-value" id="totalDisplay">&#8362;0.00</span>
        </div>
        <div id="totalConversion" style="font-size:11px; color:#5c6bc0; display:none;"></div>
      </div>

      <!-- Billing Records -->
      <div class="section-title">Billing Records <span class="req">*</span></div>
      ${giError ? '<div style="font-size:11px;color:#c62828;margin-bottom:6px;">&#9888; ' + escapeHtmlPO(giError) + '</div>' : ''}
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

      <!-- Buttons -->
      <div class="btn-row">
        <button id="btnSave" onclick="doSave()">&#128190; Save Purchase Order</button>
        <button id="btnCancel" onclick="google.script.host.close()">Cancel</button>
      </div>
      <div id="status"></div>
      <div id="emailStatus"></div>

      <script>
        // Projects that require a Renewal Date
        var RENEWAL_REQUIRED_PROJECTS = ['recurring implementation', 'outsourcing', 'support'];

        // ── Currency state ────────────────────────────────────
        var currentCurrency = 'NIS';
        var currentRate = 1; // USD→NIS rate; 1 when NIS selected

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
          ['currSymbolAmount', 'currSymbolRecurring', 'currSymbolPricePerHour', 'currSymbolHourlyRate'].forEach(function(id) {
            var el = document.getElementById(id);
            if (el) el.innerText = sym;
          });
        }

        // Returns true if Renewal Date is currently required
        function isRenewalRequired() {
          var project     = document.getElementById('project').value.trim().toLowerCase();
          var billingType = document.getElementById('billingType').value;
          return RENEWAL_REQUIRED_PROJECTS.indexOf(project) !== -1 || billingType === 'Recurring';
        }

        // Updates the req/opt marker on the Renewal Date label
        function updateRenewalRequired() {
          var required = isRenewalRequired();
          document.getElementById('reqRenewalDate').style.display = required ? 'inline' : 'none';
          document.getElementById('optRenewalDate').style.display = required ? 'none'   : 'inline';
          if (!required) document.getElementById('renewalDate').classList.remove('error');
        }

        // Returns true if Commbox ARR is required (project = Support)
        function isCommboxARRRequired() {
          return document.getElementById('project').value.trim().toLowerCase() === 'support';
        }

        function updateCommboxARRRequired() {
          var required = isCommboxARRRequired();
          document.getElementById('reqCommboxARR').style.display = required ? 'inline' : 'none';
          document.getElementById('optCommboxARR').style.display = required ? 'none'   : 'inline';
          if (!required) document.getElementById('commboxARR').classList.remove('error');
        }

        function updateCommboxHourlyRateRequired() {
          var required = document.getElementById('project').value.trim().toLowerCase() === 'support';
          document.getElementById('reqCommboxHourlyRate').style.display = required ? 'inline' : 'none';
          document.getElementById('optCommboxHourlyRate').style.display = required ? 'none'   : 'inline';
          if (!required) document.getElementById('commboxHourlyRate').classList.remove('error');
        }

        // ── Project change handler ────────────────────────────
        function onProjectChange() {
          updateRenewalRequired();
          updateCommboxARRRequired();
          updateCommboxHourlyRateRequired();
        }

        // ── Billing type change handler ──────────────────────
        function onBillingTypeChange() {
          var bt   = document.getElementById('billingType').value;
          var hint = document.getElementById('billingHint');

          // Reset all dynamic required markers and opt labels
          var fields = ['Amount','Recurring','Hours','PricePerHour'];
          fields.forEach(function(f) {
            document.getElementById('req' + f).style.display = 'none';
            document.getElementById('opt' + f).style.display = 'inline';
            // clear any previous error state
            var inputId = f === 'Amount' ? 'amount'
                        : f === 'Recurring' ? 'recurringAmount'
                        : f === 'Hours' ? 'hours'
                        : 'pricePerHour';
            document.getElementById(inputId).classList.remove('error');
          });

          hint.className = 'billing-hint';
          hint.style.display = 'none';
          hint.innerText = '';

          // Reset recurring period field
          document.getElementById('fieldRecurringPeriod').style.display = 'none';
          document.getElementById('recurringPeriod').selectedIndex = 0;
          document.getElementById('recurringPeriod').classList.remove('error');
          // Reset billing period to default when not Recurring
          if (bt !== 'Recurring') document.getElementById('billPeriod').selectedIndex = 0;

          if (bt === 'Hourly') {
            // Hours + Price per Hour become required
            document.getElementById('reqHours').style.display        = 'inline';
            document.getElementById('optHours').style.display        = 'none';
            document.getElementById('reqPricePerHour').style.display = 'inline';
            document.getElementById('optPricePerHour').style.display = 'none';
            hint.className     = 'billing-hint hourly';
            hint.style.display = 'block';
            hint.innerText     = 'ℹ️  Hourly billing: Hours and Price per Hour are required.';

          } else if (bt === 'Fixed Project') {
            // Amount becomes required
            document.getElementById('reqAmount').style.display = 'inline';
            document.getElementById('optAmount').style.display = 'none';
            hint.className     = 'billing-hint fixed';
            hint.style.display = 'block';
            hint.innerText     = 'ℹ️  Fixed Project billing: Amount is required.';

          } else if (bt === 'Recurring') {
            // Recurring Amount + Recurring Period become required
            document.getElementById('reqRecurring').style.display = 'inline';
            document.getElementById('optRecurring').style.display = 'none';
            document.getElementById('fieldRecurringPeriod').style.display = 'block';
            hint.className     = 'billing-hint recurring';
            hint.style.display = 'block';
            hint.innerText     = 'ℹ️  Recurring billing: Recurring Amount, Recurring Period, Start Date and Renewal Date are required.';

          }

          // Show Start Date field for Recurring only
          var sdRow = document.getElementById('fieldStartDate');
          if (sdRow) sdRow.style.display = (bt === 'Recurring') ? 'block' : 'none';

          updateRenewalRequired();
          updateBillingFields();
          calcTotal();
        }

        // ── Total calculator ─────────────────────────────────
        function calcTotal() {
          var bt              = document.getElementById('billingType').value;
          var amount          = parseFloat(document.getElementById('amount').value)          || 0;
          var recurring       = parseFloat(document.getElementById('recurringAmount').value) || 0;
          var recurringPeriod = parseInt(document.getElementById('recurringPeriod').value)   || 0;
          var hours           = parseFloat(document.getElementById('hours').value)           || 0;
          var pricePerHour    = parseFloat(document.getElementById('pricePerHour').value)    || 0;

          var totalInCurrency;
          if      (bt === 'Fixed Project') totalInCurrency = amount;
          else if (bt === 'Recurring')     totalInCurrency = recurring * recurringPeriod;
          else if (bt === 'Hourly')        totalInCurrency = hours * pricePerHour;
          else                             totalInCurrency = 0;

          var totalNIS = totalInCurrency * currentRate;

          document.getElementById('totalDisplay').innerText =
            '\u20aa' + totalNIS.toLocaleString('he-IL', {minimumFractionDigits: 2, maximumFractionDigits: 2});

          var convDiv = document.getElementById('totalConversion');
          if (currentCurrency === 'USD' && totalInCurrency > 0 && currentRate > 1) {
            var convText;
            if (bt === 'Recurring' && recurringPeriod > 0) {
              convText = '(from $' + recurring.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2})
                + '/month \u00d7 ' + recurringPeriod + ' months = $'
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
          el.innerText = msg;
          el.className = type || '';
        }

        function clearErrors() {
          ['organization','projectName','project','billingType','poNumber','amount','recurringAmount','recurringPeriod','hours','pricePerHour','startDate','renewalDate','commboxARR','commboxHourlyRate','billPayingCustomer','billPaymentTerms','billHoursReport','billInvoiceType','billMonth','milestoneCount']
            .forEach(function(id) { var el = document.getElementById(id); if (el) el.classList.remove('error'); });
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
          var bt          = document.getElementById('billingType').value;
          var hasCustomer = !!document.getElementById('billPayingCustomer').value;
          var isHourlyOrFixed = (bt === 'Fixed Project' || bt === 'Hourly');
          var isMsCb      = document.getElementById('isMilestones');
          var isMilestones = hasCustomer && isHourlyOrFixed && isMsCb && isMsCb.checked;

          // Payment terms + billing extras: show whenever a paying customer is selected
          document.getElementById('fieldPaymentTerms').style.display =
            hasCustomer ? 'block' : 'none';
          document.getElementById('fieldBillingExtras').style.display =
            hasCustomer ? 'flex' : 'none';

          // Billing period: show only for Recurring with a paying customer
          document.getElementById('fieldBillingPeriod').style.display =
            (hasCustomer && bt === 'Recurring') ? 'block' : 'none';

          // Milestones checkbox: show only for Hourly / Fixed Project with a paying customer
          document.getElementById('fieldMilestonesCheckbox').style.display =
            (hasCustomer && isHourlyOrFixed) ? 'block' : 'none';

          // Billing month: only when NOT milestones
          document.getElementById('fieldBillMonth').style.display =
            (hasCustomer && isHourlyOrFixed && !isMilestones) ? 'block' : 'none';

          // Milestone rows: only when milestones checkbox is checked
          document.getElementById('fieldMilestoneSection').style.display =
            isMilestones ? 'block' : 'none';

          // Reset milestone fields when hidden
          if (!isMilestones) {
            var mc = document.getElementById('milestoneCount');
            if (mc) mc.value = '';
            document.getElementById('milestoneRows').innerHTML = '';
          }
          // Uncheck + reset when billing type or customer changes away
          if (!hasCustomer || !isHourlyOrFixed) {
            if (isMsCb) isMsCb.checked = false;
          }
        }

        function onMilestonesToggle() {
          updateBillingFields();
        }

        function onMilestoneCountChange() {
          var n = parseInt(document.getElementById('milestoneCount').value) || 0;
          var html = '';
          for (var i = 1; i <= n; i++) {
            html += buildMilestoneRow(i);
          }
          document.getElementById('milestoneRows').innerHTML = html;
        }

        function buildMilestoneRow(i) {
          return '<div style="border:1px solid #c5cae9;border-radius:6px;padding:10px;margin-bottom:8px;background:#f8f9ff">' +
            '<div style="font-size:11px;font-weight:700;color:#5c6bc0;margin-bottom:6px">Milestone ' + i + '</div>' +
            '<div class="row2">' +
              '<div class="field"><label>Name <span class="req">*</span></label><input type="text" id="msName_' + i + '" /></div>' +
              '<div class="field"><label>Amount (&#8362;) <span class="req">*</span></label><input type="number" id="msAmount_' + i + '" min="0" step="0.01" /></div>' +
            '</div>' +
            '<div class="field"><label>Description <span class="opt">(optional)</span></label><input type="text" id="msDesc_' + i + '" /></div>' +
            '<div class="field"><label>Billing Month <span class="req">*</span></label><input type="month" id="msMonth_' + i + '" /></div>' +
          '</div>';
        }

        // ── Save handler ─────────────────────────────────────
        function doSave() {
          clearErrors();
          var organization    = document.getElementById('organization').value;
          var projectName     = document.getElementById('projectName').value.trim();
          var project         = document.getElementById('project').value;
          var billingType     = document.getElementById('billingType').value;
          var poNumber        = document.getElementById('poNumber').value.trim();
          var projectDesc     = document.getElementById('projectDescription').value.trim();
          var customer        = document.getElementById('customer').value.trim();
          var proposalLink    = document.getElementById('proposalLink').value.trim();
          var amount          = parseFloat(document.getElementById('amount').value)         || 0;
          var recurringAmount = parseFloat(document.getElementById('recurringAmount').value) || 0;
          var hours           = parseFloat(document.getElementById('hours').value)          || 0;
          var pricePerHour    = parseFloat(document.getElementById('pricePerHour').value)   || 0;
          var milestones      = document.getElementById('milestones').value.trim();
          var renewalDate     = document.getElementById('renewalDate').value;
          var commboxARR          = parseFloat(document.getElementById('commboxARR').value)          || 0;
          var commboxHourlyRate   = parseFloat(document.getElementById('commboxHourlyRate').value)   || 0;
          var recurringPeriod = parseInt(document.getElementById('recurringPeriod').value) || 0;
          var total           = calcTotal();
          var currency        = currentCurrency;
          var exchangeRate    = currentCurrency === 'USD' ? currentRate : 1;

          var valid = true;

          // Always required
          if (!organization) { document.getElementById('organization').classList.add('error');  valid = false; }
          if (!projectName)  { document.getElementById('projectName').classList.add('error');   valid = false; }
          if (!project)      { document.getElementById('project').classList.add('error');       valid = false; }
          if (!billingType)  { document.getElementById('billingType').classList.add('error');   valid = false; }
          if (!poNumber)     { document.getElementById('poNumber').classList.add('error');      valid = false; }

          // Renewal Date required for specific project types or Recurring billing
          if (isRenewalRequired() && !renewalDate) { document.getElementById('renewalDate').classList.add('error'); valid = false; }

          // Commbox ARR required when project = Support
          if (isCommboxARRRequired() && !commboxARR) { document.getElementById('commboxARR').classList.add('error'); valid = false; }

          // Commbox Hourly Rate required when project = Support
          var isSupport = document.getElementById('project').value.trim().toLowerCase() === 'support';
          if (isSupport && !commboxHourlyRate) { document.getElementById('commboxHourlyRate').classList.add('error'); valid = false; }

          // Conditionally required based on billing type
          if (billingType === 'Hourly') {
            if (!hours)        { document.getElementById('hours').classList.add('error');        valid = false; }
            if (!pricePerHour) { document.getElementById('pricePerHour').classList.add('error'); valid = false; }
          } else if (billingType === 'Fixed Project') {
            if (!amount) { document.getElementById('amount').classList.add('error'); valid = false; }
          } else if (billingType === 'Recurring') {
            if (!recurringAmount) { document.getElementById('recurringAmount').classList.add('error'); valid = false; }
            if (!recurringPeriod) { document.getElementById('recurringPeriod').classList.add('error'); valid = false; }
          }

          // Billing fields
          var billSel              = document.getElementById('billPayingCustomer');
          var billPayingCustomerId = billSel ? billSel.value : '';
          var billPayingCustomer   = (billSel && billSel.selectedIndex > 0) ? billSel.options[billSel.selectedIndex].text : '';
          var startDate        = (document.getElementById('startDate')        || {value:''}).value;
          var billMonth        = (document.getElementById('billMonth')        || {value:''}).value;
          var billPaymentTerms = (document.getElementById('billPaymentTerms') || {value:''}).value.trim();
          var billHoursReport  = (document.getElementById('billHoursReport')  || {value:''}).value;
          var billInvoiceType  = (document.getElementById('billInvoiceType')  || {value:''}).value;
          var billPeriod       = (document.getElementById('billPeriod')       || {value:'Monthly'}).value || 'Monthly';
          var isMsCb         = document.getElementById('isMilestones');
          var isMilestones   = !!(isMsCb && isMsCb.checked);
          var billMilestones = [];

          // Billing customer is always required
          if (!billPayingCustomerId) {
            document.getElementById('billPayingCustomer').classList.add('error');
            valid = false;
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
                var msName   = (document.getElementById('msName_'   + mi) || {value:''}).value.trim();
                var msAmt    = parseFloat((document.getElementById('msAmount_' + mi) || {value:'0'}).value) || 0;
                var msDesc   = (document.getElementById('msDesc_'   + mi) || {value:''}).value.trim();
                var msMonth  = (document.getElementById('msMonth_'  + mi) || {value:''}).value;
                var msNameEl = document.getElementById('msName_'   + mi);
                var msAmtEl  = document.getElementById('msAmount_' + mi);
                var msMoEl   = document.getElementById('msMonth_'  + mi);
                if (!msName)  { if (msNameEl) msNameEl.classList.add('error'); valid = false; }
                if (!msAmt)   { if (msAmtEl)  msAmtEl.classList.add('error');  valid = false; }
                if (!msMonth) { if (msMoEl)   msMoEl.classList.add('error');   valid = false; }
                billMilestones.push({ name: msName, amount: msAmt, description: msDesc, month: msMonth });
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
              if (result.ok) {
                var statusMsg = '\u2705 ' + result.msg;
                if (result.billingError) statusMsg += ' (\u26a0\ufe0f Billing: ' + result.billingError + ')';
                setStatus(statusMsg, 'ok');
                var es = document.getElementById('emailStatus');
                if (result.emailSent) {
                  es.innerText = '\u2709\ufe0f Email sent to: ' + result.emailRecipients;
                  es.className = 'ok';
                } else {
                  es.innerText = '\u26a0\ufe0f Email not sent: ' + (result.emailError || 'unknown error');
                  es.className = 'err';
                }
                // Reset form
                ['organization','project','billingType'].forEach(function(id) {
                  document.getElementById(id).selectedIndex = 0;
                });
                ['projectName','poNumber','projectDescription','customer','proposalLink','milestones','renewalDate','commboxARR','commboxHourlyRate'].forEach(function(id) {
                  document.getElementById(id).value = '';
                });
                document.getElementById('recurringPeriod').selectedIndex = 0;
                document.getElementById('fieldRecurringPeriod').style.display = 'none';
                ['amount','recurringAmount','hours','pricePerHour'].forEach(function(id) {
                  document.getElementById(id).value = '';
                });
                document.getElementById('totalDisplay').innerText = '\u20aa0.00';
                // Reset billing hint and renewal date req/opt markers
                document.getElementById('billingHint').style.display = 'none';
                onBillingTypeChange(); // resets req/opt labels
                onProjectChange();     // resets renewal date + commbox ARR req/opt
                // Reset billing section
                document.getElementById('billPayingCustomer').selectedIndex = 0;
                document.getElementById('billPaymentTerms').value = '';
                document.getElementById('billHoursReport').selectedIndex = 0;
                document.getElementById('billInvoiceType').selectedIndex = 0;
                document.getElementById('billPeriod').selectedIndex = 0;
                document.getElementById('startDate').value = '';
                document.getElementById('billMonth').value = '';
                document.getElementById('isMilestones').checked = false;
                document.getElementById('milestoneCount').value = '';
                document.getElementById('milestoneRows').innerHTML = '';
                updateBillingFields();
              } else {
                setStatus('\u274c ' + result.msg, 'err');
              }
              document.getElementById('btnSave').disabled = false;
            })
            .withFailureHandler(function(e) {
              setStatus('\u274c Error: ' + e.message, 'err');
              document.getElementById('btnSave').disabled = false;
            })
            .saveNewPurchaseOrder({
              organization:       organization,
              projectName:        projectName,
              project:            project,
              billingType:        billingType,
              poNumber:           parseInt(poNumber, 10),
              projectDescription: projectDesc,
              customer:           customer,
              proposalLink:       proposalLink,
              amount:             amount,
              recurringAmount:    recurringAmount,
              milestones:         milestones,
              hours:              hours,
              pricePerHour:       pricePerHour,
              totalAmount:        total,
              renewalDate:        renewalDate,
              commboxARR:           commboxARR || "",
              commboxHourlyRate:    commboxHourlyRate || "",
              recurringPeriod:      recurringPeriod || "",
              currency:             currency,
              exchangeRate:         exchangeRate,
              startDate:            startDate,
              isMilestones:         isMilestones,
              billPayingCustomer:   billPayingCustomer,
              billPayingCustomerId: billPayingCustomerId,
              billMonth:            billMonth,
              billMilestones:       billMilestones,
              billPaymentTerms:     billPaymentTerms,
              billHoursReport:      billHoursReport,
              billInvoiceType:      billInvoiceType,
              billPeriod:           billPeriod,
            });
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(500).setHeight(1050).setTitle("New Purchase Order");

  SpreadsheetApp.getUi().showModalDialog(html, "New Purchase Order");
}


// ── SAVE NEW PURCHASE ORDER ────────────────────────────────
function saveNewPurchaseOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const poSheet = ss.getSheetByName("Purchase Orders");
    if (!poSheet) return { ok: false, msg: 'Sheet "Purchase Orders" not found.' };

    // Read headers — use fixed column count of 20 to cover all columns
    const lastCol = Math.max(poSheet.getLastColumn(), 20);
    const headers = poSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Format today's date as dd/mm/yyyy
    const today   = new Date();
    const dd      = String(today.getDate()).padStart(2, "0");
    const mm      = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy    = today.getFullYear();
    const dateStr = dd + "/" + mm + "/" + yyyy;

    // Format dates from yyyy-mm-dd to dd/mm/yyyy
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

    const newRow = [
      dateStr,                       // col 1  — PO Date
      data.organization,             // col 2  — Organization
      data.project,                  // col 3  — Project
      data.projectDescription,       // col 4  — Project description
      data.poNumber,                 // col 5  — PO number
      data.amount          || "",    // col 6  — Amount
      data.recurringAmount || "",    // col 7  — Recurring Amount
      data.milestones,               // col 8  — Milestones
      data.hours           || "",    // col 9  — Hours
      data.pricePerHour    || "",    // col 10 — Price per hour
      data.totalAmount     || "",    // col 11 — Total Amount
      data.customer,                 // col 12 — Customer
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
    ];

    poSheet.appendRow(newRow);

    // Auto-search the organization so the new PO appears immediately
    runSearch(data.organization);

    // Save billing records if paying customer was set
    let billingCount = 0;
    let billingError = '';
    if (data.billingType && data.billPayingCustomerId) {
      const br = saveBillingRecords({
        organization:        data.organization,
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
      });
      billingCount = br.count || 0;
      billingError = br.ok ? '' : (br.error || 'Unknown billing error');
    }

    // Send notification email
    var emailSent = false;
    var emailError = '';
    var emailRecipients = '';
    try {
      const managerName   = getManagerNameForOrg(data.organization);
      const managerEmail  = managerNameToEmail(managerName);
      emailRecipients     = buildToList(managerEmail);
      const subject = 'New Purchase Order #' + data.poNumber + ' - ' + data.organization;
      const body =
        'A new Purchase Order has been added to the CRM.\n\n' +
        '-- Order Details --\n' +
        'PO Number:           ' + data.poNumber             + '\n' +
        'PO Date:             ' + dateStr                   + '\n' +
        'Organization:        ' + data.organization         + '\n' +
        'Project Name:        ' + (data.projectName          || '-') + '\n' +
        'Project:             ' + data.project              + '\n' +
        'Billing Type:        ' + (data.billingType         || '-') + '\n' +
        'Customer:            ' + (data.customer            || '-') + '\n' +
        'Project Description: ' + (data.projectDescription  || '-') + '\n' +
        'Proposal Link:       ' + (data.proposalLink        || '-') + '\n\n' +
        '-- Financial Details --\n' +
        'Amount:              ' + (data.amount              || '-') + '\n' +
        'Recurring Amount:    ' + (data.recurringAmount     || '-') + '\n' +
        'Recurring Period:    ' + (data.recurringPeriod     || '-') + '\n' +
        'Hours:               ' + (data.hours               || '-') + '\n' +
        'Price per Hour:      ' + (data.pricePerHour        || '-') + '\n' +
        'Total Amount (NIS):  ' + (data.totalAmount         || '-') + '\n' +
        'Currency:            ' + (data.currency            || 'NIS') + '\n' +
        (data.currency === 'USD' ? 'Exchange Rate:       ' + (data.exchangeRate || '-') + ' (USD→NIS)\n' : '') +
        'Milestones:          ' + (data.milestones          || '-') + '\n' +
        'Renewal Date:        ' + (renewalFormatted         || '-') + '\n' +
        'Commbox ARR:         ' + (data.commboxARR          || '-') + '\n' +
        'Commbox Hourly Rate: ' + (data.commboxHourlyRate   || '-') + '\n';
      MailApp.sendEmail(emailRecipients, subject, body);
      emailSent = true;
    } catch(mailErr) {
      emailError = mailErr.message;
    }

    const billingMsg = billingCount > 0 ? ` ${billingCount} billing record(s) added.` : '';
    return {
      ok:              true,
      msg:             'PO #' + data.poNumber + ' for ' + data.organization + ' saved successfully!' + billingMsg,
      emailSent:       emailSent,
      emailError:      emailError,
      emailRecipients: emailRecipients,
      billingError:    billingError,
    };

  } catch (e) {
    return { ok: false, msg: e.message };
  }
}


// ── HELPER: build a row array aligned to sheet headers ─────
function buildPORow(headers, fieldMap) {
  return headers.map(h => {
    const key = String(h).trim();
    if (fieldMap.hasOwnProperty(key)) return fieldMap[key];
    // Fallback: match ignoring trailing spaces
    const found = Object.keys(fieldMap).find(k => k.trim() === key);
    return found !== undefined ? fieldMap[found] : "";
  });
}


// ── GET USD→NIS RATE FROM CODETABLES NAMED CELL ────────────
// Reads the named range "nisdollar" from the spreadsheet.
// To update the rate, just change that cell value in CODETABLES.
function getUsdToNisRate() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const range = ss.getRangeByName('nisdollar');
    if (!range) return null;
    const val = parseFloat(range.getValue());
    return isNaN(val) || val <= 0 ? null : val;
  } catch(e) {
    return null;
  }
}


// ── HELPER: escape HTML special characters ─────────────────
function escapeHtmlPO(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}