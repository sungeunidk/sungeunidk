
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
const SHEET_NAME = "응답";

function doPost(e) {
  try {
    const response = String(
      e.parameter.response || ""
    ).trim();

    if (!response || response.length > 5000) {
      return HtmlService.createHtmlOutput(
        "잘못된 응답입니다."
      );
    }

    const spreadsheet =
      SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      sheet.appendRow(["제출 시각", "응답 내용"]);
      sheet.setFrozenRows(1);
    }

    // 응답은 일반 텍스트로 저장
    const safeResponse = response.replace(
      /^[=+\-@\t\r]/,
      "'$&"
    );

    sheet.appendRow([
      new Date(),
      safeResponse
    ]);

    return HtmlService.createHtmlOutput(
      "응답이 저장되었습니다."
    );
  } catch (error) {
    console.error(error);

    return HtmlService.createHtmlOutput(
      "저장 중 오류가 발생했습니다."
    );
  }
}
