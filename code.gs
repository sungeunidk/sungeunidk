
const SPREADSHEET_ID = "11XpLhQPE5af98HqsbF_A7RUKKpQBhtDFtSs3OD-kZP4";
const SHEET_NAME = "시트1";

const HEADERS = [
  "제출 시각",
  "응답 내용",
  "응답 ID",
  "이전 응답 ID"
];

/**
 * 웹 앱 진입점
 * 기본 주소: 전시용 연결망 화면
 * ?action=graph : 연결망 JSON 데이터
 */
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : "";

  if (action === "graph") {
    return ContentService
      .createTextOutput(JSON.stringify(getGraphData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("응답 연결망")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * 사용자 응답 저장
 * 기존 POST 요청 형식: { "response": "응답 내용" }
 */
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

    const sheet = getResponseSheet_();
    migrateRows_(sheet);

    const lastRow = sheet.getLastRow();

    // 가장 최근 응답의 ID를 가져온다.
    const previousId = lastRow >= 2
      ? String(sheet.getRange(lastRow, 3).getDisplayValue())
      : "";

    const responseId = Utilities.getUuid();

    // 스프레드시트 수식 실행 방지
    const safeResponse = /^[=+\-@\t\r]/.test(response)
      ? "'" + response
      : response;

    sheet.appendRow([
      new Date(),
      safeResponse,
      responseId,
      previousId
    ]);

    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        id: responseId,
        previousId: previousId
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error(
      "응답 저장 실패: " +
      (error.stack || error.message || error)
    );

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: String(error.message || error)
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * 스프레드시트 가져오기 및 헤더 확인
 */
function getResponseSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("시트를 찾을 수 없습니다: " + SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    // 기존 A, B열은 유지하고 C, D열 헤더만 보완한다.
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 기존 응답에 ID와 이전 응답 ID를 부여한다.
 * 이미 ID가 있는 행은 기존 값을 유지한다.
 */
function migrateRows_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 4);
  const rows = range.getValues();

  let previousId = "";

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    // 내용이 없는 행은 연결하지 않는다.
    if (!String(row[1] || "").trim()) continue;

    let id = String(row[2] || "").trim();

    if (!id) {
      id = Utilities.getUuid();
      row[2] = id;
    }

    if (!String(row[3] || "").trim() && previousId) {
      row[3] = previousId;
    }

    previousId = id;
  }

  range.setValues(rows);
}

/**
 * 전시 화면에 전달할 연결망 데이터
 */
function getGraphData() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const sheet = getResponseSheet_();
    migrateRows_(sheet);

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        ok: true,
        nodes: [],
        edges: [],
        total: 0
      };
    }

    const rows = sheet
      .getRange(2, 1, lastRow - 1, 4)
      .getValues();

    const nodes = [];
    const edges = [];

    for (const row of rows) {
      const timestamp = row[0];
      const rawResponse = String(row[1] || "");
      const id = String(row[2] || "").trim();
      const previousId = String(row[3] || "").trim();

      if (!rawResponse.trim() || !id) continue;

      // 시트 수식 방지를 위한 앞쪽 작은따옴표는
      // 화면에서 표시하지 않는다.
      const response = rawResponse.replace(/^'/, "");

      nodes.push({
        id: id,
        previousId: previousId,
        text: response,
        timestamp: timestamp instanceof Date
          ? timestamp.toISOString()
          : String(timestamp || "")
      });

      if (previousId) {
        edges.push({
          source: previousId,
          target: id
        });
      }
    }

    return {
      ok: true,
      nodes: nodes,
      edges: edges,
      total: nodes.length
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
