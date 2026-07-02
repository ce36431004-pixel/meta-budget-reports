// 이 코드를 Google Apps Script(script.google.com)에 붙여넣고 웹앱으로 배포하세요.
// 별도로 시트를 미리 만들 필요 없음 — 처음 실행될 때 스프레드시트를 자동으로 생성합니다.
// 기준 입력 사이트(criteria)와 리포트 담당자 판단(decisions) 저장을 모두 이 스크립트 하나로 처리합니다.

const DECISIONS_SHEET = 'Decisions';
const CRITERIA_SHEET = 'Criteria';

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) return SpreadsheetApp.openById(ssId);
  const ss = SpreadsheetApp.create('메타 예산 제안 - 기준/판단 기록');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function getSheet_(name, headerRow) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headerRow);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const type = e.parameter.type;

  if (type === 'criteria') {
    const sheet = getSheet_(CRITERIA_SHEET, ['updatedAt', 'payloadJson']);
    const rows = sheet.getDataRange().getValues();
    const records = rows.slice(1).map(r => JSON.parse(r[1])).reverse(); // 최신 먼저
    if (e.parameter.action === 'history') return json_(records);
    return json_(records.length ? records[0] : null); // latest
  }

  if (type === 'decisions') {
    const reportId = e.parameter.reportId;
    const sheet = getSheet_(DECISIONS_SHEET, ['reportId', 'itemId', 'itemName', 'decision', 'decidedAt']);
    const rows = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < rows.length; i++) {
      const [rId, itemId, itemName, decision, decidedAt] = rows[i];
      if (rId === reportId) result[itemId] = { itemName, decision, decidedAt };
    }
    return json_(result);
  }

  return json_({ error: 'type 파라미터가 필요합니다 (criteria 또는 decisions)' });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.type === 'criteria') {
    const sheet = getSheet_(CRITERIA_SHEET, ['updatedAt', 'payloadJson']);
    const record = { ...body, type: undefined, updatedAt: new Date().toISOString() };
    delete record.type;
    sheet.appendRow([record.updatedAt, JSON.stringify(record)]);
    return json_({ ok: true, saved: record });
  }

  if (body.type === 'decisions') {
    const { reportId, itemId, itemName, decision } = body;
    if (!reportId || !itemId || !decision) return json_({ error: 'reportId, itemId, decision은 필수입니다' });
    const sheet = getSheet_(DECISIONS_SHEET, ['reportId', 'itemId', 'itemName', 'decision', 'decidedAt']);
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
    if (!updated) sheet.appendRow([reportId, itemId, itemName || '', decision, decidedAt]);
    return json_({ ok: true, saved: { itemId, itemName, decision, decidedAt } });
  }

  return json_({ error: 'type 필드가 필요합니다 (criteria 또는 decisions)' });
}
