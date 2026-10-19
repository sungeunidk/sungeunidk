const SPREADSHEET_ID =
  "11XpLhQPE5af98HqsbF_A7RUKKpQBhtDFtSs3OD-kZP4";

const SHEET_NAME = "응답";

function doGet() {
  return ContentService
    .createTextOutput("Response service is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    const body = JSON.parse(
      e.postData.contents
    );

    const response = String(
      body.response || ""
    ).trim();

    if (!response || response.length > 5000) {
      throw new Error("Invalid response");
    }

    lock.waitLock(10000);

    const spreadsheet =
      SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet =
      spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["제출 시각", "응답 내용"]);
      sheet.setFrozenRows(1);
    }

    // 수식으로 해석될 수 있는 입력을 방지
    const safeResponse = /^[=+\-@\t\r]/.test(response)
      ? "'" + response
      : response;

    sheet.appendRow([
      new Date(),
      safeResponse
    ]);

    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(
        JSON.stringify({ ok: true })
      )
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error(error);

    return ContentService
      .createTextOutput(
        JSON.stringify({ ok: false })
      )
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
