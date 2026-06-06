
function runAnalysisFromSheet() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName("UI");
  if (!sheet) {
    ui.alert('הטאב "UI" לא נמצא.');
    logAction("Run(Server) Error", 'UI sheet missing');
    return;
  }

  // קריאת פרמטרים מהטאב UI
  const selection = String(sheet.getRange("F8:G8").getValue() || "").trim(); // סוג ניתוח
  const sheetName = String(sheet.getRange("F5:G5").getValue() || "").trim(); // שם הגיליון שיווצר

  // לוג פתיחה
  logAction("Run(Server) Start", `selection=${selection || "(empty)"} | sheet=${sheetName || "(empty)"}`);

  // ולידציה בסיסית
  if (!selection) { ui.alert("אנא בחר סוג ניתוח בתא F8."); return; }
  if (!sheetName) { ui.alert("אנא מלא שם גיליון בתא F5."); return; }

  // בדיקת שם גיליון (בשרת)
  if (isSheetNameTaken(sheetName)) {
    ui.alert("שם הגיליון כבר קיים. אנא בחר שם אחר.");
    logAction("Run(Server) Aborted", "Sheet name already exists: " + sheetName);
    return;
  }

  // ניתוב לפי סוג ניתוח
  if (selection === "ניתוח לפי טווח תאריכים") {
    const fromDate = sheet.getRange("F11").getValue(); // תאריך התחלה
    const toDate   = sheet.getRange("F12").getValue(); // תאריך סיום

    if (!fromDate || !toDate || isNaN(new Date(fromDate)) || isNaN(new Date(toDate))) {
      ui.alert("אנא מלא תאריכים חוקיים בתאים F11 ו־F12.");
      logAction("Run(Server) Error", "Invalid or missing dates");
      return;
    }
    logAction("Run(Server) ByDate", `From=${fromDate} To=${toDate} -> ${sheetName}`);
    analyzeByDateRange(new Date(fromDate), new Date(toDate), sheetName);
    // analyzeByDateRange כבר מציגה alert עם תקציר

  } else if (selection === "ניתוח לפי מספרי שיחה") {
    const ids = sheet.getRange("C5:C150").getValues()
                  .flat()
                  .map(v => String(v).trim())
                  .filter(v => v);
    if (ids.length === 0) {
      ui.alert("אנא הזן מספרי שיחה בעמודה C (תאים C5:C150).");
      logAction("Run(Server) Error", "No IDs provided");
      return;
    }
    logAction("Run(Server) ByIDs", `Count=${ids.length} -> ${sheetName}`);
    analyzeByConversationIds(ids, sheetName);
    // analyzeByConversationIds מציגה alert עם תקציר

  } else if (selection === "ניתוח לפי קטגוריה") {
    /**
     * עבור הרצה מהשרת, קרא את ערכי הקטגוריה משני תאים (ניתן לשנות לפי הצורך):
     * F7 = שם כותרת (header) מתוך "Automatic analysis"
     * G7 = הערך המבוקש באותה עמודה
     */
    const headerName = String(sheet.getRange("F15").getValue() || "").trim();
    const value      = String(sheet.getRange("F16").getValue() || "").trim();
    if (!headerName || !value) {
      ui.alert('לניתוח לפי קטגוריה יש למלא: F15=כותרת, F16=ערך לחיפוש.');
      logAction("Run(Server) Error", "Missing category header/value (F15/F16)");
      return;
    }
    logAction("Run(Server) ByCategory", `${headerName} = ${value} -> ${sheetName}`);
    analyzeByCategory(headerName, value, sheetName);
    // analyzeByCategory תקרא פנימית ל-analyzeByConversationIds

  } else {
    ui.alert("ערך לא מוכר בתא F8. השתמש באחד: 'ניתוח לפי טווח תאריכים' / 'ניתוח לפי מספרי שיחה' / 'ניתוח לפי קטגוריה'.");
    logAction("Run(Server) Error", "Unknown selection: " + selection);
    return;
  }

  logAction("Run(Server) Done", `Started analysis for '${sheetName}'`);
}

function fetchTranscriptsToNewSheet() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const uiSheet = ss.getSheetByName("UI");

  if (!uiSheet) {
    ui.alert('The "UI" tab was not found.');
    return;
  }

  // Read conversation IDs from C5:C150
  const ids = uiSheet.getRange("C5:C150").getValues()
    .flat()
    .map(v => String(v).trim())
    .filter(v => v);

  if (ids.length === 0) {
    ui.alert("Please enter conversation IDs in column C (cells C5:C150).");
    return;
  }

  // Get streamId from script properties
  const streamId = PROPS.getProperty("STREAM_ID");
  if (!streamId) {
    ui.alert('Script property "STREAM_ID" is not set. Please add it in Project Settings > Script Properties.');
    return;
  }

  // Create a new sheet for transcripts
  const timestamp = Utilities.formatDate(new Date(), "Asia/Jerusalem", "yyyy-MM-dd HH:mm");
  const newSheetName = "Transcripts " + timestamp;
  let transcriptSheet = ss.getSheetByName(newSheetName);
  if (!transcriptSheet) {
    transcriptSheet = ss.insertSheet(newSheetName);
  }

  // Headers
  transcriptSheet.getRange(1, 1, 1, 3).setValues([["transcript", "conversationId", "status"]]);
  transcriptSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");
  transcriptSheet.setFrozenRows(1);
  transcriptSheet.setColumnWidth(1, 600);
  transcriptSheet.getRange("A:A").setWrap(true);

  let successCount = 0;
  let failCount = 0;

  ids.forEach((convId, index) => {
    const row = index + 2;
    const transcript = fetchTranscriptFromCommBox(streamId, convId);
    if (transcript) {
      transcriptSheet.getRange(row, 1, 1, 3).setValues([[transcript, convId, "OK"]]);
      successCount++;
    } else {
      transcriptSheet.getRange(row, 1, 1, 3).setValues([["", convId, "transcript not found"]]);
      failCount++;
    }
  });

  SpreadsheetApp.flush();
  ss.setActiveSheet(transcriptSheet);
  ui.alert(`Done!\n✅ Success: ${successCount}\n❌ Failed: ${failCount}\nSheet: "${newSheetName}"`);
}
