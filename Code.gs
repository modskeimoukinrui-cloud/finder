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

// タイムライン投稿に添付する写真の保存先（Google Driveフォルダ「finder_photos」、事前に手動作成済み）
const PHOTO_FOLDER_ID = '1kB8ixkgfcHtMVvvgljqJwc4yMpDUhZNK';

function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// 日付セルはそのままJSONに入れるとUTCのISO文字列になり、スプレッドシートの
// タイムゾーン設定によっては日付が1日ずれて見える。シート上の表示と一致させるため、
// スプレッドシート自身のタイムゾーンで文字列に整形してから返す。
function formatCellValue(v, timeZone) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    const hms = Utilities.formatDate(v, timeZone, 'HH:mm:ss');
    return hms === '00:00:00'
      ? Utilities.formatDate(v, timeZone, 'yyyy-MM-dd')
      : Utilities.formatDate(v, timeZone, 'yyyy-MM-dd HH:mm:ss');
  }
  return v;
}

// 今日の日付（Asia/Tokyo）を yyyy-MM-dd で返す
function todayInTokyo() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function sheetToObjects(sheet, headerRow) {
  headerRow = headerRow || 1;
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const values = sheet.getDataRange().getValues();
  const headers = values[headerRow - 1];
  const rows = values.slice(headerRow);
  return rows
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = formatCellValue(row[i], timeZone); });
      return obj;
    });
}

// 列見出しの表記ゆれ（"施設固有ID\n【変更禁止】"のような改行・注記付き）に対応するため、
// 「ID」を含む列名を実際のヘッダーから動的に探す（'施設ID'固定文字列には依存しない）
function findIdKey(obj) {
  return Object.keys(obj).find(k => k.replace(/\s/g, '').includes('ID'));
}

// ③口コミ・タイムラインシートには「投稿ID」と「施設固有ID」の2つのID列があり、
// findIdKey()の「IDを含む」だけの判定だと先に出現する「投稿ID」に誤マッチしてしまう。
// そのため施設IDの照合には「施設固有ID」への完全一致（前後の空白・改行除去のうえ）を優先する。
function findTimelineFacilityIdKey(obj) {
  if (Object.prototype.hasOwnProperty.call(obj, '施設固有ID')) return '施設固有ID';
  return Object.keys(obj).find(k => k.replace(/\s/g, '') === '施設固有ID');
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

  // デバッグ：実際の最終行番号と、ヘッダー行（HEADER_ROWS.TIMELINE行目）の中身を確認する
  Logger.log('getTimeline: sheet.getLastRow() = ' + sheet.getLastRow());
  const headerRowValues = sheet.getRange(HEADER_ROWS.TIMELINE, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('getTimeline: ヘッダー行(' + HEADER_ROWS.TIMELINE + '行目) = ' + JSON.stringify(headerRowValues));

  const all = sheetToObjects(sheet, HEADER_ROWS.TIMELINE);
  Logger.log('getTimeline: sheetToObjects()で読めたデータ行数 = ' + all.length);

  if (facilityId) {
    const idKey = all.length ? findTimelineFacilityIdKey(all[0]) : null;
    Logger.log('getTimeline: 施設ID照合に使うキー = ' + idKey);
    if (!idKey) return [];
    const filtered = all.filter(row => row[idKey] === facilityId);
    Logger.log('getTimeline: facilityId=' + facilityId + ' に一致した件数 = ' + filtered.length);
    return filtered;
  }
  return all;
}

// 列名を突き合わせるため、空白・改行を取り除いて正規化する
function squashKey(s) {
  return String(s).split(/[\s　]+/).join('');
}

// ②施設詳細・ケア体制 の1行を、アプリの「施設情報を更新」フォームの内容で更新する。
// 施設固有IDで対象行を特定し、送られてきた項目だけを書き換える。
function updateFacilityDetail(payload) {
  const facilityId = String(payload.facilityId || '').trim();
  if (!facilityId) throw new Error('facilityId が指定されていません');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.DETAIL);
  if (!sheet) throw new Error('②施設詳細・ケア体制シートが見つかりません');

  const headerRow = HEADER_ROWS.DETAIL;
  const values = sheet.getDataRange().getValues();
  const headers = values[headerRow - 1];

  const colOf = {};
  headers.forEach((h, i) => { colOf[squashKey(h)] = i + 1; });
  const findCol = (name) => colOf[squashKey(name)] || 0;

  const idCol = headers.findIndex(h => squashKey(h).indexOf('ID') !== -1) + 1;
  if (!idCol) throw new Error('ID列が見つかりません');

  let rowNum = 0;
  for (let r = headerRow; r < values.length; r++) {
    if (String(values[r][idCol - 1]).trim() === facilityId) { rowNum = r + 1; break; }
  }
  if (!rowNum) throw new Error('該当する施設が見つかりません: ' + facilityId);

  const current = values[rowNum - 1];
  const updated = [];

  // 指定された列に値を書き込む（列が無ければ黙って飛ばす）。
  // 実際に値が変わったときだけ true を返す。
  function put(colName, value) {
    const col = findCol(colName);
    if (!col) return false;
    const before = current[col - 1];
    const beforeStr = (before === null || before === undefined) ? '' : String(before).trim();
    const afterStr = (value === null || value === undefined) ? '' : String(value).trim();
    if (beforeStr === afterStr) return false;
    sheet.getRange(rowNum, col).setValue(value === null || value === undefined ? '' : value);
    updated.push(colName);
    return true;
  }

  const numOrBlank = (v) => (v === '' || v === null || v === undefined) ? '' : Number(v);

  // --- 空床状況 ---
  // 空床状況が「変化したときだけ」空床確認日を今日（日本時間）で更新する
  if (payload.vacancy !== undefined) {
    const beforeVacancy = String(current[(findCol('空床状況') || 1) - 1] || '').trim();
    const afterVacancy = String(payload.vacancy || '').trim();
    put('空床状況', afterVacancy);
    if (afterVacancy && afterVacancy !== beforeVacancy) {
      put('空床確認日', todayInTokyo());
    }
  }
  if (payload.vacancyNote !== undefined) put('空床メモ', payload.vacancyNote || '');

  // --- 費用 ---
  // 入居時費用・月額下限・月額上限・費用特記のいずれかが変わったら費用更新日を記録する
  let costChanged = false;
  if (payload.nyukyoFee !== undefined) costChanged = put('入居時費用', payload.nyukyoFee || '') || costChanged;
  if (payload.costMin !== undefined) costChanged = put('月額下限（円）', numOrBlank(payload.costMin)) || costChanged;
  if (payload.costMax !== undefined) costChanged = put('月額上限（円）', numOrBlank(payload.costMax)) || costChanged;
  if (payload.costNote !== undefined) costChanged = put('費用に関する信頼性・特記事項', payload.costNote || '') || costChanged;
  if (costChanged) put('費用更新日', todayInTokyo());

  // --- ケア体制14項目 ---
  // 14項目のいずれかが変わったらケア体制更新日を記録する
  let careChanged = false;
  if (payload.care && typeof payload.care === 'object') {
    Object.keys(payload.care).forEach(label => {
      if (put(label, payload.care[label] || '')) careChanged = true;
    });
  }
  // 「その他」を選んだ項目がある場合に項目名と内容を書く共通の自由記述欄。
  // ケア体制の一部という扱いなので、変更検知もケア体制更新日に含める。
  if (payload.careOtherNote !== undefined) {
    if (put('ケア体制その他メモ', payload.careOtherNote || '')) careChanged = true;
  }
  if (careChanged) put('ケア体制更新日', todayInTokyo());

  // --- 更新の記録（全体の最終更新。新着順ソート用に項目別の日付とは別に持つ） ---
  // ここまでに何か1つでも実際に値が変わっていた（updatedが空でない）ときだけ記録する。
  // 無条件で書いてしまうと、内容を何も変えずに保存しただけで「新着」扱いになってしまうため。
  if (updated.length > 0) {
    put('最終更新日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'));
    if (payload.updatedBy) put('最終更新者', payload.updatedBy);
  }

  return {
    success: true, facilityId: facilityId, row: rowNum,
    costChanged: costChanged, careChanged: careChanged, updated: updated
  };
}

// ③口コミ・タイムラインの列位置を、ヘッダー名から引けるようにして返す
function timelineColumns(sheet) {
  const headers = sheet.getRange(HEADER_ROWS.TIMELINE, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[squashKey(h)] = i + 1; });
  return {
    headers: headers,
    postId: col['投稿ID'] || 0,
    facilityId: col['施設固有ID'] || 0,
    text: col['投稿テキスト'] || 0,
    author: col['投稿者名'] || 0,
    photoUrl: col['写真URL'] || 0,
  };
}

// 写真URL（https://drive.google.com/uc?id={fileId} 形式）からDriveファイルIDを取り出す
function extractDriveFileId(url) {
  const m = String(url || '').match(/[?&]id=([^&]+)/);
  return m ? m[1] : '';
}

// カンマ区切りの写真URL文字列を受け取り、対応するDriveファイルをすべてゴミ箱へ移動する。
// 「リンクを知っている全員が閲覧可」の共有設定はゴミ箱に移動しただけでは解除されず、
// リンク経由でアクセスできてしまうため、ゴミ箱へ移動する前に共有を非公開へ戻す。
// （完全な物理削除ではなくゴミ箱移動にとどめているのは、誤操作時にオーナーが
// 　Drive側から手動で復元できる余地を残すため）
// 1件の失敗（既に削除済み・IDが不正等）で全体を止めないよう、ファイルごとにtry/catchする。
function trashDrivePhotosByUrls(urlsCsv) {
  String(urlsCsv || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (url) {
    const fileId = extractDriveFileId(url);
    if (!fileId) return;
    try {
      const file = DriveApp.getFileById(fileId);
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      file.setTrashed(true);
    } catch (e) {
      Logger.log('写真の削除に失敗（続行します）: ' + url + ' / ' + e);
    }
  });
}

// 投稿IDを「既存IDの最大値＋1」で採番する。
// 行数から計算すると、投稿を削除したあとに既存IDと衝突するため。
function nextPostId(sheet, postIdCol) {
  const lastRow = sheet.getLastRow();
  let maxSeq = 0;
  if (lastRow > HEADER_ROWS.TIMELINE && postIdCol) {
    const ids = sheet.getRange(HEADER_ROWS.TIMELINE + 1, postIdCol, lastRow - HEADER_ROWS.TIMELINE, 1).getValues();
    ids.forEach(function (r) {
      const m = String(r[0] === null || r[0] === undefined ? '' : r[0]).match(/(\d+)/);
      if (m) maxSeq = Math.max(maxSeq, parseInt(m[1], 10));
    });
  }
  return 'P' + String(maxSeq + 1).padStart(5, '0');
}

// タイムライン投稿に添付された写真（リサイズ・圧縮済みのJPEG base64データURL配列）をDriveへ保存し、
// 表示用URLをカンマ区切りの文字列にして返す。ファイル名は {施設ID}_{タイムスタンプ}_{連番}.jpg。
//
// 注意：doPostの実行時間には上限があるため、一度に大量（目安10枚以上）の写真を送ると
// タイムアウトする可能性がある。base64化された画像は元データの約1.37倍のサイズになるため、
// POSTペイロードが大きくなりすぎないようフロント側でリサイズ・圧縮してから送る前提。
function savePhotosToDrive(facilityId, photosBase64Array) {
  const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
  const urls = [];

  photosBase64Array.forEach(function (dataUrl, i) {
    if (!dataUrl) return;
    const base64 = String(dataUrl).replace(/^data:image\/\w+;base64,/, '');
    const fileName = facilityId + '_' + timestamp + '_' + (i + 1) + '.jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls.push('https://drive.google.com/uc?id=' + file.getId());
  });

  return urls.join(',');
}

function postTimeline(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TIMELINE);

  if (sheet.getLastRow() === 0) {
    throw new Error('タイムラインシートにヘッダーがありません');
  }

  const cols = timelineColumns(sheet);
  const postId = nextPostId(sheet, cols.postId);

  // 写真が添付されている場合はDriveへ保存し、URLをカンマ区切りで写真URL列へ記録する
  let photoUrls = '';
  if (payload.photos && payload.photos.length > 0) {
    photoUrls = savePhotosToDrive(payload.facilityId || postId, payload.photos);
  }

  // ③口コミ・タイムラインシートの実際の列順：
  // 投稿ID / 施設固有ID / 施設名（参照用） / 投稿テキスト / 写真URL / 確認年月日 / 投稿者名 / システム通知フラグ
  const row = [
    postId,
    payload.facilityId || '',
    payload.facilityName || '',
    payload.text || '',
    photoUrls,
    payload.date || '',
    payload.author || '',
    '',
  ];

  sheet.appendRow(row);
  return { success: true, postId: postId, photoCount: photoUrls ? photoUrls.split(',').length : 0 };
}

// ③口コミ・タイムラインから投稿を1件削除する。
// 投稿IDで行を特定し、削除前に行の内容を実行ログへ残す（誤削除時に復元する手がかりにする）。
//
// 権限について：このWeb APIは匿名アクセスできるため、ここでの投稿者名の照合は
// 「アプリ側の誤操作を防ぐ」ためのものであり、技術的なアクセス制御ではない。
function deleteTimeline(payload) {
  const postId = String(payload.postId || '').trim();
  if (!postId) throw new Error('postId が指定されていません');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TIMELINE);
  if (!sheet) throw new Error('③口コミ・タイムラインシートが見つかりません');

  const cols = timelineColumns(sheet);
  if (!cols.postId) throw new Error('投稿ID列が見つかりません');

  const values = sheet.getDataRange().getValues();
  let rowNum = 0;
  for (let r = HEADER_ROWS.TIMELINE; r < values.length; r++) {
    if (String(values[r][cols.postId - 1]).trim() === postId) { rowNum = r + 1; break; }
  }
  if (!rowNum) throw new Error('該当する投稿が見つかりません: ' + postId);

  const target = values[rowNum - 1];
  const author = cols.author ? String(target[cols.author - 1] || '').trim() : '';

  // 投稿者本人か、管理者としての操作でなければ削除しない（誤操作防止）
  const requestedBy = String(payload.requestedBy || '').trim();
  const asAdmin = payload.asAdmin === true || payload.asAdmin === 'true';
  if (!asAdmin && author && requestedBy !== author) {
    throw new Error('投稿者本人ではないため削除できません（投稿者: ' + author + ' / 操作者: ' + (requestedBy || '不明') + '）');
  }

  // --- 削除前に内容をログへ残す ---
  // 復元時にそのまま貼り戻せるよう、日付はシートの表示と同じ形に整えておく
  const timeZone = ss.getSpreadsheetTimeZone();
  const snapshot = {};
  cols.headers.forEach(function (h, i) {
    const v = formatCellValue(target[i], timeZone);
    if (v !== '' && v !== null && v !== undefined) snapshot[String(h).replace(/\n/g, ' ')] = String(v);
  });
  Logger.log('=== タイムライン削除 ===');
  Logger.log('投稿ID: ' + postId + ' / シート行: ' + rowNum);
  Logger.log('操作者: ' + (requestedBy || '(不明)') + ' / 管理者操作: ' + asAdmin);
  Logger.log('削除する行の内容: ' + JSON.stringify(snapshot));

  if (payload.dryRun === true || payload.dryRun === 'true') {
    Logger.log('※ dryRun のため削除していません');
    return { success: true, dryRun: true, postId: postId, row: rowNum, deleted: snapshot };
  }

  // 添付写真があればDrive上の実ファイルもゴミ箱へ移動する（失敗しても行削除は続行する）
  if (cols.photoUrl) {
    trashDrivePhotosByUrls(target[cols.photoUrl - 1]);
  }

  sheet.deleteRow(rowNum);
  Logger.log('行' + rowNum + ' を削除しました。残り件数: ' + Math.max(0, sheet.getLastRow() - HEADER_ROWS.TIMELINE));

  return { success: true, postId: postId, row: rowNum, deleted: snapshot };
}

// ③口コミ・タイムラインの1投稿から、写真を1枚だけ取り除く（投稿テキスト自体は残す）。
// 対応するDrive上の実ファイルもゴミ箱へ移動する。
//
// 権限の考え方はdeleteTimelineと同じ：投稿者本人か管理者操作のみ許可する（誤操作防止目的）。
function deletePhoto(payload) {
  const postId = String(payload.postId || '').trim();
  const photoUrl = String(payload.photoUrl || '').trim();
  if (!postId) throw new Error('postId が指定されていません');
  if (!photoUrl) throw new Error('photoUrl が指定されていません');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.TIMELINE);
  if (!sheet) throw new Error('③口コミ・タイムラインシートが見つかりません');

  const cols = timelineColumns(sheet);
  if (!cols.postId || !cols.photoUrl) throw new Error('必要な列（投稿ID/写真URL）が見つかりません');

  const values = sheet.getDataRange().getValues();
  let rowNum = 0;
  for (let r = HEADER_ROWS.TIMELINE; r < values.length; r++) {
    if (String(values[r][cols.postId - 1]).trim() === postId) { rowNum = r + 1; break; }
  }
  if (!rowNum) throw new Error('該当する投稿が見つかりません: ' + postId);

  const target = values[rowNum - 1];
  const author = cols.author ? String(target[cols.author - 1] || '').trim() : '';

  const requestedBy = String(payload.requestedBy || '').trim();
  const asAdmin = payload.asAdmin === true || payload.asAdmin === 'true';
  if (!asAdmin && author && requestedBy !== author) {
    throw new Error('投稿者本人ではないため削除できません（投稿者: ' + author + ' / 操作者: ' + (requestedBy || '不明') + '）');
  }

  const currentUrls = String(target[cols.photoUrl - 1] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const remaining = currentUrls.filter(function (u) { return u !== photoUrl; });
  if (remaining.length === currentUrls.length) {
    throw new Error('指定された写真が見つかりません: ' + photoUrl);
  }

  trashDrivePhotosByUrls(photoUrl);
  sheet.getRange(rowNum, cols.photoUrl).setValue(remaining.join(','));

  Logger.log('写真を1枚削除しました。投稿ID: ' + postId + ' / 残り枚数: ' + remaining.length);

  return { success: true, postId: postId, remainingCount: remaining.length };
}

// ①②シート双方をスキャンし、施設固有IDの数値部分の最大値+1をF0001形式で返す。
// タイムライン投稿ID（nextPostId）と同じ「既存最大値+1」方式。
function nextFacilityId(ss) {
  const sheets = [ss.getSheetByName(SHEET_NAMES.MASTER), ss.getSheetByName(SHEET_NAMES.DETAIL)];
  let maxNum = 0;
  sheets.forEach(function (sheet) {
    if (!sheet) return;
    const headerRow = HEADER_ROWS.MASTER; // ①②とも2行目
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return;
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idCol = headers.findIndex(function (h) { return squashKey(h).indexOf('ID') !== -1; }) + 1;
    if (!idCol) return;
    const ids = sheet.getRange(headerRow + 1, idCol, lastRow - headerRow, 1).getValues();
    ids.forEach(function (r) {
      const m = String(r[0] === null || r[0] === undefined ? '' : r[0]).match(/(\d+)/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
  });
  return 'F' + String(maxNum + 1).padStart(4, '0');
}

// 列名（表記ゆれはsquashKeyで吸収）に従って値をセットした行を作り、appendRowする。
// 該当する列が無いキーは黙って無視する。
function appendRowByColumnNames(sheet, headerRow, valuesByColName) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  headers.forEach(function (h, i) {
    const key = Object.keys(valuesByColName).find(function (k) { return squashKey(k) === squashKey(h); });
    if (key !== undefined) row[i] = valuesByColName[key];
  });
  sheet.appendRow(row);
}

// 新規施設を①施設マスター・②施設詳細ケア体制の両方に登録する。
// ID採番とマスター行・詳細行の最小限の追加はここで行い、費用・ケア体制など
// フォームで入力された詳細項目の書き込みは既存のupdateFacilityDetail()にそのまま委譲する
// （変更検知・各種更新日の記録ロジックを再利用するため）。
function addFacility(payload) {
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('施設名が指定されていません');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const masterSheet = ss.getSheetByName(SHEET_NAMES.MASTER);
  const detailSheet = ss.getSheetByName(SHEET_NAMES.DETAIL);
  if (!masterSheet) throw new Error('①施設マスターシートが見つかりません');
  if (!detailSheet) throw new Error('②施設詳細・ケア体制シートが見つかりません');

  const facilityId = nextFacilityId(ss);
  const idNumMatch = facilityId.match(/(\d+)/);
  const idNum = idNumMatch ? String(parseInt(idNumMatch[1], 10)) : '';
  const registeredBy = String(payload.registeredBy || '').trim();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  appendRowByColumnNames(masterSheet, HEADER_ROWS.MASTER, {
    '施設固有ID【変更禁止】': facilityId,
    '通し番号': idNum,
    'データソース': 'アプリ登録',
    '施設名': name,
    '施設類型': payload.type || '',
    '所在地': payload.address || '',
    '行政区': payload.ward || '',
    '電話番号': payload.tel || '',
    '更新事由': '新規登録' + (registeredBy ? '（' + registeredBy + '）' : ''),
  });

  // 最終更新日時・最終更新者はここで無条件にセットする（詳細項目を何も入力せず
  // 登録した場合でも新着順ソートに乗るように。個別項目の更新日はこの後の
  // updateFacilityDetail()が、実際に値が入っていれば記録する）
  appendRowByColumnNames(detailSheet, HEADER_ROWS.DETAIL, {
    '施設固有ID【変更禁止】': facilityId,
    '施設名（参照用）': name,
    '運用ステータス（運用中／廃止）': '運用中',
    '最終更新日時': now,
    '最終更新者': registeredBy,
  });

  const detailPayload = Object.assign({}, payload, { facilityId: facilityId, updatedBy: registeredBy });
  updateFacilityDetail(detailPayload);

  return { success: true, facilityId: facilityId };
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
    } else if (action === 'updateFacilityDetail') {
      data = updateFacilityDetail(payload);
    } else if (action === 'deleteTimeline') {
      data = deleteTimeline(payload);
    } else if (action === 'deletePhoto') {
      data = deletePhoto(payload);
    } else if (action === 'addFacility') {
      data = addFacility(payload);
    } else {
      data = { error: 'Unknown action. Use action=postTimeline, updateFacilityDetail, deleteTimeline, deletePhoto or addFacility' };
    }

    return jsonResponse(data);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}
