// 이 코드를 Google Apps Script(script.google.com)에 붙여넣고 웹앱으로 배포하세요.
// 별도로 시트를 미리 만들 필요 없음 — 처음 실행될 때 스프레드시트를 자동으로 생성합니다.
// 기준 입력 사이트(criteria), 리포트 담당자 판단(decisions), 실행 여부(executedAt)를 모두 이 스크립트 하나로 처리합니다.

const DECISIONS_SHEET = 'Decisions';
const CRITERIA_SHEET = 'Criteria';
const DECISIONS_HEADER = ['reportId', 'itemId', 'itemName', 'decision', 'decidedAt', 'entityType', 'currentBudgetKRW', 'proposedBudgetKRW', 'confirmedBudgetKRW', 'executedAt'];

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
  } else {
    // 기존 시트에 새 컬럼이 추가된 경우 헤더만 안전하게 확장 (기존 데이터는 건드리지 않음)
    const existingHeader = sheet.getRange(1, 1, 1, headerRow.length).getValues()[0];
    if (headerRow.some((h, i) => existingHeader[i] !== h)) {
      sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    }
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
    const sheet = getSheet_(DECISIONS_SHEET, DECISIONS_HEADER);
    const rows = sheet.getDataRange().getValues();

    // action=all: 리포트 구분 없이 지금까지의 모든 결정 이력을 최신순으로 반환 ("실행 이력" 탭용)
    if (e.parameter.action === 'all') {
      const records = [];
      for (let i = 1; i < rows.length; i++) {
        const [reportId, itemId, itemName, decision, decidedAt, entityType, currentBudgetKRW, proposedBudgetKRW, confirmedBudgetKRW, executedAt] = rows[i];
        if (!reportId) continue;
        records.push({ reportId, itemId, itemName, decision, decidedAt, entityType, currentBudgetKRW, proposedBudgetKRW, confirmedBudgetKRW, executedAt: executedAt || null });
      }
      records.sort((a, b) => (b.decidedAt || '').localeCompare(a.decidedAt || ''));
      return json_(records);
    }

    const reportId = e.parameter.reportId;
    const result = {};
    for (let i = 1; i < rows.length; i++) {
      const [rId, itemId, itemName, decision, decidedAt, entityType, currentBudgetKRW, proposedBudgetKRW, confirmedBudgetKRW, executedAt] = rows[i];
      if (rId === reportId) {
        result[itemId] = { itemName, decision, decidedAt, entityType, currentBudgetKRW, proposedBudgetKRW, confirmedBudgetKRW, executedAt: executedAt || null };
      }
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
    const sheet = getSheet_(DECISIONS_SHEET, DECISIONS_HEADER);

    // 특정 리포트의 저장된 결정을 전부 삭제 (담당자가 확정 전 상태로 롤백하고 싶을 때)
    if (body.action === 'clearReport') {
      if (!body.reportId) return json_({ error: 'reportId는 필수입니다' });
      const rows = sheet.getDataRange().getValues();
      let deleted = 0;
      for (let i = rows.length - 1; i >= 1; i--) {
        if (rows[i][0] === body.reportId) {
          sheet.deleteRow(i + 1);
          deleted++;
        }
      }
      return json_({ ok: true, cleared: body.reportId, deleted });
    }

    const { reportId, itemId } = body;
    if (!reportId || !itemId) return json_({ error: 'reportId, itemId는 필수입니다' });
    const rows = sheet.getDataRange().getValues();

    // 실행 완료 표시만 하는 호출 (executedAt만 옴, decision 등은 안 건드림)
    if (body.markExecuted) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === reportId && rows[i][1] === itemId) {
          sheet.getRange(i + 1, 10).setValue(new Date().toISOString());
          return json_({ ok: true, executed: true });
        }
      }
      return json_({ error: '해당 항목을 찾을 수 없습니다' });
    }

    // 일반 저장 (확정 예산 입력/변경 시 자동 호출됨)
    const { itemName, decision, entityType, currentBudgetKRW, proposedBudgetKRW, confirmedBudgetKRW } = body;
    if (!decision) return json_({ error: 'decision은 필수입니다' });
    const decidedAt = new Date().toISOString();
    const rowValues = [itemName || '', decision, decidedAt, entityType || '', currentBudgetKRW ?? '', proposedBudgetKRW ?? '', confirmedBudgetKRW ?? '', ''];
    let updated = false;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === reportId && rows[i][1] === itemId) {
        sheet.getRange(i + 1, 3, 1, 8).setValues([rowValues]);
        updated = true;
        break;
      }
    }
    if (!updated) sheet.appendRow([reportId, itemId, ...rowValues]);
    return json_({ ok: true, saved: { itemId, itemName, decision, decidedAt, entityType, currentBudgetKRW, proposedBudgetKRW, confirmedBudgetKRW } });
  }

  return json_({ error: 'type 필드가 필요합니다 (criteria 또는 decisions)' });
}
