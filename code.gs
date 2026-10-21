const SPREADSHEET_ID = "11XpLhQPE5af98HqsbF_A7RUKKpQBhtDFtSs3OD-kZP4";
const SHEET_NAME = "시트1";

function doGet() {
  return ContentService
    .createTextOutput("Response service is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("요청 데이터가 비어 있습니다.");
    }

    const body = JSON.parse(e.postData.contents);
    const response = String(body.response || "").trim();

    if (!response || response.length > 5000) {
      throw new Error("응답이 비어 있거나 5000자를 초과했습니다.");
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error("시트를 찾을 수 없습니다: " + SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["제출 시각", "응답 내용"]);
      sheet.setFrozenRows(1);
    }

    const safeResponse = /^[=+\-@\t\r]/.test(response)
      ? "'" + response
      : response;

    sheet.appendRow([new Date(), safeResponse]);
    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("응답 저장 실패: " + (error.stack || error.message || error));
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(error.message || error) }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
