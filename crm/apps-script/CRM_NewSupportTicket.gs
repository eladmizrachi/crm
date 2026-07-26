function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Service Tickets");

    const object = data.Event.Object;
    const activity = object.Activities && object.Activities[0];

    const conversationId = object.Id || "";
    const rawCreateDate = activity ? activity.CreateDate || "" : "";
    const createDate = addThreeHours(rawCreateDate);

    const statusId = activity && activity.Data ? Number(activity.Data.statusId) : "";

    const statusMap = {
      1: "פתוח",
      2: "טופל",
      4: "טופל",
      8: "בטיפול"
    };

    const statusText = statusMap[statusId] || statusId;

    const lastRow = sheet.getLastRow();
    let existingRow = null;

    if (conversationId && lastRow > 0) {
      const conversationIds = sheet.getRange(1, 1, lastRow, 1).getValues();

      for (let i = 0; i < conversationIds.length; i++) {
        if (String(conversationIds[i][0]).trim() === String(conversationId).trim()) {
          existingRow = i + 1;
          break;
        }
      }
    }

    if (!existingRow) {
      sheet.appendRow([
        conversationId,
        "",
        createDate,
        "",
        "",
        statusText
      ]);

      existingRow = sheet.getLastRow();
    } else {
      sheet.getRange(existingRow, 6).setValue(statusText);

      if (statusText === "טופל") {
        sheet.getRange(existingRow, 4).setValue(createDate);
      }
    }

    setStatusDropdownAndColor(sheet, existingRow, statusText);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function addThreeHours(dateString) {
  if (!dateString) return "";

  const parts = dateString.split(" ");
  const dateParts = parts[0].split("-");
  const timeParts = parts[1].split(":");

  const day = Number(dateParts[0]);
  const month = Number(dateParts[1]) - 1;
  const year = Number(dateParts[2]);

  const hour = Number(timeParts[0]);
  const minute = Number(timeParts[1]);
  const second = Number(timeParts[2]);

  const date = new Date(year, month, day, hour, minute, second);
  date.setHours(date.getHours() + 3);

  return Utilities.formatDate(date, "Asia/Jerusalem", "dd-MM-yyyy HH:mm:ss");
}

function setStatusDropdownAndColor(sheet, row, statusText) {
  const cell = sheet.getRange(row, 6);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["פתוח", "טופל", "בטיפול"], true)
    .setAllowInvalid(false)
    .build();

  cell.setDataValidation(rule);

  if (statusText === "פתוח") {
    cell.setBackground("#FFF2CC"); // צהוב
  } else if (statusText === "טופל") {
    cell.setBackground("#D9EAD3"); // ירוק
  } else if (statusText === "בטיפול") {
    cell.setBackground("#D9D2E9"); // סגול
  } else {
    cell.setBackground(null);
  }
}
// test auto deploy