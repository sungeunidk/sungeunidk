
const SPREADSHEET_ID = "11XpLhQPE5af98HqsbF_A7RUKKpQBhtDFtSs3OD-kZP4";
const SHEET_NAME = "시트1";

const HEADERS = [
  "제출 시각",
  "응답 내용",
  "응답 ID",
  "이전 응답 ID"
];

/**
 * 웹 앱 화면
 * 기본 접속: 응답 작성 화면
 * ?page=done&id=응답ID : 제출 완료 및 연결망 화면
 */
function doGet(e) {
  const page = e && e.parameter ? e.parameter.page : "";

  if (page === "done") {
    return HtmlService
      .createHtmlOutputFromFile("done")
      .setTitle("응답과 응답 사이")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (e && e.parameter && e.parameter.action === "graph") {
    return ContentService
      .createTextOutput(JSON.stringify(getGraphData()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("응답 남기기")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * HTML 응답 작성 화면에서 호출
 */
function submitResponse(response) {
  return saveResponse_(response);
}

/**
 * 기존 외부 POST 요청도 지원
 * 요청 형식: {"response":"응답 내용"}
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("요청 데이터가 비어 있습니다.");
    }

    const body = JSON.parse(e.postData.contents);
    const result = saveResponse_(body.response);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("응답 저장 실패: " + (error.stack || error.message || error));

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: String(error.message || error)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 응답 저장 및 이전 응답과 연결
 */
function saveResponse_(input) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    if (input === null || input === undefined) {
      throw new Error("응답을 입력해 주세요.");
    }

    const response = String(input).trim();

    if (!response || response.length > 5000) {
      throw new Error("응답이 비어 있거나 5000자를 초과했습니다.");
    }

    const sheet = getResponseSheet_();

    // 기존 응답도 연결 ID가 있도록 보완
    migrateRows_(sheet);

    const lastRow = sheet.getLastRow();

    const previousId = lastRow >= 2
      ? String(sheet.getRange(lastRow, 3).getDisplayValue()).trim()
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

    return {
      ok: true,
      id: responseId,
      previousId: previousId
    };

  } catch (error) {
    console.error("응답 저장 실패: " + (error.stack || error.message || error));
    throw error;

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * 스프레드시트 및 헤더 확인
 */
function getResponseSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("시트를 찾을 수 없습니다: " + SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    // 기존 A, B열 데이터는 유지하고 C, D열을 추가
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 기존 응답에 ID와 이전 응답 ID를 부여
 */
function migrateRows_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 4);
  const rows = range.getValues();

  let previousId = "";

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    if (!String(row[1] || "").trim()) continue;

    let id = String(row[2] || "").trim();

    if (!id) {
      id = Utilities.getUuid();
      row[2] = id;
    }

    // 기존 연결 정보가 없는 경우에만 직전 응답과 연결
    if (!String(row[3] || "").trim() && previousId) {
      row[3] = previousId;
    }

    previousId = id;
  }

  range.setValues(rows);
}

/**
 * 전시 화면용 연결망 데이터 반환
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

      // 수식 방지용 접두 작은따옴표 제거
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
