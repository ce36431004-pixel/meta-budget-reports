// 이 코드를 Google Apps Script(script.google.com)에 붙여넣고 웹앱으로 배포하세요.
// 별도로 시트를 미리 만들 필요 없음 — 처음 실행될 때 스프레드시트를 자동으로 생성합니다.

const SHEET_NAME = 'Decisions';

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let ss;
  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create('메타 예산 제안 - 담당자 판단 기록');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['reportId', 'itemId', 'itemName', 'decision', 'decidedAt']);
  }
  return sheet;
}

function doGet(e) {
  const reportId = e.parameter.reportId;
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < rows.length; i++) {
    const [rId, itemId, itemName, decision, decidedAt] = rows[i];
    if (rId === reportId) {
      result[itemId] = { itemName, decision, decidedAt };
    }
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const { reportId, itemId, itemName, decision } = body;
  if (!reportId || !itemId || !decision) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'reportId, itemId, decision은 필수입니다' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const decidedAt = new Date().toISOString();
  let updated = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === reportId && rows[i][1] === itemId) {
      sheet.getRange(i + 1, 3, 1, 3).setValues([[itemName || '', decision, decidedAt]]);
      updated = true;
      break;
    }
  }
  if (!updated) {
    sheet.appendRow([reportId, itemId, itemName || '', decision, decidedAt]);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, saved: { itemId, itemName, decision, decidedAt } }))
    .setMimeType(ContentService.MimeType.JSON);
}
