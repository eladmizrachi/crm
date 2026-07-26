// ============================================================
// CRM_Config.gs — Centralized CRM Settings (stored in sheet)
// ============================================================

const CONFIG_SHEET_NAME = '_CRM_Settings';

const CONFIG_DEFAULTS_ = {
  send_email_enabled:     'true',
  test_email_recipient:   'elad@teamiff.com',
  notification_always_to: 'elad@teamiff.com,anita@teamiff.com,gefen@teamiff.com',
  sheet_editors:          'elad@teamiff.com,david@teamiff.com',
};


// ── Read a single setting ─────────────────────────────────

function getCRMSetting_(key) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (sh) {
      const data = sh.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === key) return String(data[i][1]).trim();
      }
    }
  } catch(e) {}
  return CONFIG_DEFAULTS_[key] !== undefined ? CONFIG_DEFAULTS_[key] : '';
}


// ── Read all settings (used by dialog) ────────────────────

function getAllCRMSettings() {
  const result = Object.assign({}, CONFIG_DEFAULTS_);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (sh) {
      const data = sh.getDataRange().getValues();
      data.forEach(function(row) {
        const key = String(row[0]).trim();
        if (key && Object.prototype.hasOwnProperty.call(result, key)) {
          result[key] = String(row[1]).trim();
        }
      });
    }
  } catch(e) {}
  return result;
}


// ── Save all settings ─────────────────────────────────────

function saveAllCRMSettings(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(CONFIG_SHEET_NAME);
    }
    sh.clearContents();
    const rows = Object.keys(data).map(function(k) { return [k, data[k]]; });
    if (rows.length > 0) sh.getRange(1, 1, rows.length, 2).setValues(rows);
    try { sh.hideSheet(); } catch(e) {}
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}


// ── Convenience helpers used by other files ───────────────

function getAdminEmails_() {
  return getCRMSetting_('sheet_editors')
    .split(',').map(function(e) { return e.trim(); }).filter(Boolean);
}

function isCurrentUserAdmin_() {
  const cur = Session.getActiveUser().getEmail().toLowerCase();
  return getAdminEmails_().map(function(e) { return e.toLowerCase(); }).indexOf(cur) !== -1;
}


// ── Open Settings Dialog ──────────────────────────────────

function openCRMSettingsDialog() {
  const cfg = getAllCRMSettings();
  const html = HtmlService.createHtmlOutput(buildSettingsHtml_(cfg))
    .setTitle('CRM Settings')
    .setWidth(560)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'CRM Settings');
}


// ── HTML ──────────────────────────────────────────────────

function buildSettingsHtml_(cfg) {
  const emailEnabled = cfg.send_email_enabled !== 'false';
  const testEmail    = cfg.test_email_recipient || '';
  const distList     = cfg.notification_always_to.split(',').map(function(e){ return e.trim(); }).join('\n');
  const adminList    = cfg.sheet_editors.split(',').map(function(e){ return e.trim(); }).join('\n');

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #f5f7fa; color: #212121; font-size: 13px; }

  .header {
    background: #1a237e; color: #fff;
    padding: 15px 20px; font-size: 16px; font-weight: bold;
    display: flex; align-items: center; gap: 10px;
  }

  .body { padding: 18px 20px 20px; overflow-y: auto; max-height: 540px; }

  .section {
    background: #fff; border: 1px solid #e0e0e0;
    border-radius: 8px; margin-bottom: 16px; overflow: hidden;
  }
  .section-head {
    background: #e8eaf6; padding: 10px 14px;
    font-size: 12px; font-weight: 700; color: #3949ab;
    text-transform: uppercase; letter-spacing: .05em;
    border-bottom: 1px solid #c5cae9;
  }
  .section-body { padding: 14px; display: flex; flex-direction: column; gap: 12px; }

  label { display: block; font-weight: 600; margin-bottom: 4px; color: #1a237e; font-size: 12px; }
  .hint { font-size: 11px; color: #757575; margin-top: 3px; }

  input[type=email], input[type=text], textarea {
    width: 100%; padding: 8px 10px; font-size: 13px;
    border: 1.5px solid #9fa8da; border-radius: 6px;
    background: #fff; color: #212121; outline: none;
    font-family: inherit;
    transition: border-color .15s;
  }
  input:focus, textarea:focus { border-color: #3949ab; }
  textarea { resize: vertical; min-height: 80px; line-height: 1.5; }

  .toggle-row {
    display: flex; align-items: center; gap: 10px; cursor: pointer;
  }
  .toggle-row input[type=checkbox] {
    width: 18px; height: 18px; accent-color: #1a237e; cursor: pointer; flex-shrink: 0;
  }
  .toggle-row span { font-size: 13px; color: #212121; }

  #testEmailWrap { transition: opacity .2s; }

  .btn-row {
    display: flex; gap: 10px; padding: 0 20px 18px;
  }
  button {
    flex: 1; padding: 10px; font-size: 13px; font-weight: 700;
    border: none; border-radius: 6px; cursor: pointer;
  }
  #btnSave { background: #1a237e; color: #fff; }
  #btnSave:hover { background: #3949ab; }
  #btnCancel { background: #e8eaf6; color: #1a237e; }
  #btnCancel:hover { background: #c5cae9; }
  #status { padding: 0 20px 12px; font-size: 12px; min-height: 18px; text-align: center; }
  #status.ok  { color: #2e7d32; }
  #status.err { color: #c62828; }
</style>
</head>
<body>

<div class="header">⚙️ CRM Settings</div>

<div class="body">

  <!-- Email Settings -->
  <div class="section">
    <div class="section-head">📧 Email Settings</div>
    <div class="section-body">
      <div>
        <label class="toggle-row">
          <input type="checkbox" id="sendEmailEnabled" ${emailEnabled ? 'checked' : ''} onchange="onEmailToggle()" />
          <span>Enable email sending to all recipients</span>
        </label>
        <div class="hint">When unchecked, all emails are sent only to the test address below.</div>
      </div>
      <div id="testEmailWrap" style="opacity:${emailEnabled ? 0.5 : 1}">
        <label>Test email address (used when emails are disabled)</label>
        <input type="email" id="testEmailRecipient" value="${escapeHtmlCfg_(testEmail)}" placeholder="test@example.com" />
      </div>
      <div>
        <label>Distribution list — always notified on every action</label>
        <textarea id="notificationAlwaysTo" placeholder="one email per line">${escapeHtmlCfg_(distList)}</textarea>
        <div class="hint">One email address per line.</div>
      </div>
    </div>
  </div>

  <!-- Admin Users -->
  <div class="section">
    <div class="section-head">👥 Admin Users</div>
    <div class="section-body">
      <div>
        <label>Admin email addresses</label>
        <textarea id="sheetEditors" placeholder="one email per line">${escapeHtmlCfg_(adminList)}</textarea>
        <div class="hint">One email per line. Admins can edit data sheets directly and access the CRM Admin menu.</div>
      </div>
    </div>
  </div>

</div><!-- /body -->

<div id="status"></div>
<div class="btn-row">
  <button id="btnSave" onclick="doSave()">💾 Save Settings</button>
  <button id="btnCancel" onclick="google.script.host.close()">Cancel</button>
</div>

<script>
  function onEmailToggle() {
    var enabled = document.getElementById('sendEmailEnabled').checked;
    document.getElementById('testEmailWrap').style.opacity = enabled ? '0.5' : '1';
  }

  function doSave() {
    var data = {
      send_email_enabled:     document.getElementById('sendEmailEnabled').checked ? 'true' : 'false',
      test_email_recipient:   document.getElementById('testEmailRecipient').value.trim(),
      notification_always_to: linesToComma('notificationAlwaysTo'),
      sheet_editors:          linesToComma('sheetEditors'),
    };

    if (!data.test_email_recipient) {
      setStatus('⚠️ Test email address is required.', 'err'); return;
    }
    if (!data.notification_always_to) {
      setStatus('⚠️ Distribution list cannot be empty.', 'err'); return;
    }
    if (!data.sheet_editors) {
      setStatus('⚠️ Admin users list cannot be empty.', 'err'); return;
    }

    setStatus('⏳ Saving…');
    document.getElementById('btnSave').disabled = true;
    google.script.run
      .withSuccessHandler(function(res) {
        document.getElementById('btnSave').disabled = false;
        if (res.ok) {
          setStatus('✅ Settings saved successfully!', 'ok');
        } else {
          setStatus('❌ Error: ' + res.error, 'err');
        }
      })
      .withFailureHandler(function(e) {
        document.getElementById('btnSave').disabled = false;
        setStatus('❌ Error: ' + e.message, 'err');
      })
      .saveAllCRMSettings(data);
  }

  function linesToComma(id) {
    return document.getElementById(id).value
      .split('\\n').map(function(s){ return s.trim(); }).filter(Boolean).join(',');
  }

  function setStatus(msg, cls) {
    var el = document.getElementById('status');
    el.textContent = msg;
    el.className = cls || '';
  }
</script>
</body>
</html>`;
}

function escapeHtmlCfg_(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
