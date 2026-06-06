const PROPS = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = PROPS.getProperty("SPREADSHEET_ID");

function doPost(e) {
  try {
    logAction("Webhook Start", "Received incoming webhook");
    const body = JSON.parse(e.postData.contents);

    const activity = body.Event?.Object?.Activities?.[0];
    if (!activity) {
      logAction("Webhook Error", "No activity found in payload", body);
      return ContentService.createTextOutput("Error: No activity found in payload.");
    }

    const statusId = activity.Data.statusId;
    if (statusId !== 2) {
      logAction("Webhook Ignored", `Status ID ${statusId} ≠ 2 – skipped`);
      return ContentService.createTextOutput("Status is not 2, ignored.");
    }

    const conversationId = body.Event.Object.Id;
    const streamId = body.Event.Object.StreamId;
    const brand = body.Event.Brand;
    const timestamp = Utilities.formatDate(new Date(), "Asia/Jerusalem", "yyyy-MM-dd HH:mm:ss");

    const transcript = fetchTranscriptFromCommBox(streamId, conversationId);
    if (!transcript) {
      logAction("Transcript Missing", "Transcript not found", conversationId, streamId);
      return ContentService.createTextOutput("Transcript not found");
    }

    const result = analyzeTranscript(transcript);
    if (!result || !result.insights) {
      logAction("Analysis Failed", "GPT analysis returned empty or invalid result", conversationId, streamId);
      return ContentService.createTextOutput("Analysis failed");
    }

    const insights = result.insights;
    const usage = result.usage || {};

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Automatic analysis");
    if (!sheet) {
      sheet = ss.insertSheet("Automatic analysis");
    }

    if (sheet.getLastRow() === 0) {
      const headers = [
        "conversationId",
        "streamId",
        "brand",
        "timestamp",
        ...Object.keys(insights),
        "prompt_tokens",
        "completion_tokens"
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || "").toString().trim());

    const valueMap = Object.assign(
      {
        conversationId: `'${conversationId}`,
        streamId,
        brand,
        timestamp,
        prompt_tokens: usage.prompt_tokens ?? "",
        completion_tokens: usage.completion_tokens ?? ""
      },
      insights
    );

    const rowData = headers.map(h => valueMap[h] ?? "");

    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    SpreadsheetApp.flush();

    logAction("Analysis Success", "Insights saved to sheet", conversationId, streamId);
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    logAction("Webhook Error", `Webhook handling failed: ${error}`);
    return ContentService.createTextOutput("Invalid Webhook format.").setMimeType(ContentService.MimeType.TEXT);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧠 Conversation Analysis')
    .addItem('Open Analysis Interface', 'showUI')
    .addSeparator()
    .addItem('Run (server) from sheet', 'runAnalysisFromSheet')
    .addSeparator()
    .addItem('Get Transcripts from IDs (column C)', 'fetchTranscriptsToNewSheet')
    .addToUi();
}

function showUI() {
  const html = HtmlService.createHtmlOutputFromFile('UI')
                         .setTitle('Conversation Analysis Interface')
                         .setWidth(520)
                         .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}