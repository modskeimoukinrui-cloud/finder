const SHEET_ID = '11QnXAozq3e-BYjOu-oAsk6hTqGcvx99ahK68qUciZkI';

const SHEET_NAMES = {
  MASTER: '①施設マスター',
  DETAIL: '②施設詳細・ケア体制',
  TIMELINE: '③口コミ・タイムライン',
};

// 各シートのヘッダー行番号（1行目=作業メモ、2行目=列名の場合は2）
const HEADER_ROWS = {
  MASTER: 2,
  DETAIL: 2,
  TIMELINE: 2,
};

function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function sheetToObjects(sheet, headerRow) {
  headerRow = headerRow || 1;
  const values = sheet.getDataRange().getValues();
  const headers = values[headerRow - 1];
  const rows = values.slice(headerRow);
  return rows
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

// 列見出しの表記ゆれ（"施設固有ID\n【変更禁止】"のような改行・注記付き）に対応するため、
// 「ID」を含む列名を実際のヘッダーから動的に探す（'施設ID'固定文字列には依存しない）
function findIdKey(obj) {
  return Object.keys(obj).find(k => k.replace(/\s/g, '').includes('ID'));
}

function getFacilities() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const masterSheet = ss.getSheetByName(SHEET_NAMES.MASTER);
  const detailSheet = ss.getSheetByName(SHEET_NAMES.DETAIL);

  const masters = sheetToObjects(masterSheet, HEADER_ROWS.MASTER);
  const details = sheetToObjects(detailSheet, HEADER_ROWS.DETAIL);

  const masterIdKey = masters.length ? findIdKey(masters[0]) : null;
  const detailIdKey = details.length ? findIdKey(details[0]) : null;

  const detailMap = {};
  if (detailIdKey) {
    details.forEach(d => { detailMap[d[detailIdKey]] = d; });
  }

  const merged = masters.map(m => ({
    ...m,
    ...((masterIdKey && detailMap[m[masterIdKey]]) || {}),
  }));

  return merged;
}

function getTimeline(facilityId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TIMELINE);
  const all = sheetToObjects(sheet, HEADER_ROWS.TIMELINE);

  if (facilityId) {
    const idKey = all.length ? findIdKey(all[0]) : null;
    if (!idKey) return [];
    return all.filter(row => row[idKey] === facilityId);
  }
  return all;
}

function postTimeline(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TIMELINE);

  if (sheet.getLastRow() === 0) {
    throw new Error('タイムラインシートにヘッダーがありません');
  }

  const headers = sheet.getRange(HEADER_ROWS.TIMELINE, 1, 1, sheet.getLastColumn()).getValues()[0];
  const now = new Date();
  const row = headers.map(h => {
    if (h === '投稿日時') return now;
    return payload[h] !== undefined ? payload[h] : '';
  });

  sheet.appendRow(row);
  return { success: true, timestamp: now.toISOString() };
}

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || '';

  try {
    let data;

    if (action === 'facilities') {
      data = getFacilities();
    } else if (action === 'timeline') {
      data = getTimeline(params.id || null);
    } else {
      data = { error: 'Unknown action. Use action=facilities or action=timeline' };
    }

    return jsonResponse(data);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || (e.postData ? JSON.parse(e.postData.contents).action : '');

    let payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (_) {}
    }
    Object.assign(payload, params);

    let data;

    if (action === 'postTimeline') {
      data = postTimeline(payload);
    } else {
      data = { error: 'Unknown action. Use action=postTimeline' };
    }

    return jsonResponse(data);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}
