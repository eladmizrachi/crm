// ============================================================
// CRM_TestEmail.gs — Test email sending
// Run testSendEmail() from the Apps Script editor to verify
// that MailApp works and the correct address receives the mail.
// ============================================================

function testSendEmail() {
  const recipient = "accounting@teamiff.com";
  const subject   = "✅ CRM Email Test";
  const body      =
    "This is a test email from the CRM Google Apps Script.\n\n" +
    "If you received this, email notifications are working correctly.\n\n" +
    "Sent from: " + SpreadsheetApp.getActiveSpreadsheet().getName();

  try {
    MailApp.sendEmail(recipient, subject, body);
    SpreadsheetApp.getUi().alert("✅ Test email sent to " + recipient);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Failed to send email:\n\n" + e.message);
  }
}
