// 勤怠アプリ クラウド同期用 Google Apps Script
// このコードをスプレッドシートの「拡張機能 → Apps Script」に貼り付けて、
// 「デプロイ → 新しいデプロイ → ウェブアプリ」で公開してください。
// 実行ユーザー: 自分 / アクセスできるユーザー: 全員

const SHEET_NAME = 'logs';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'createdAt', 'staff', 'type', 'json']);
  }
  return sheet;
}

// 記録の一覧を返す（アプリが読み込むとき）
function doGet() {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < rows.length; i++) {
    try { logs.push(JSON.parse(rows[i][4])); } catch (e) {}
  }
  return ContentService
    .createTextOutput(JSON.stringify({ logs: logs }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 記録の追加・承認・削除（アプリが送信するとき）
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (data.action === 'add' && data.log && data.log.id) {
      sheet.appendRow([
        data.log.id,
        data.log.createdAt || '',
        data.log.staff || '',
        data.log.type || '',
        JSON.stringify(data.log)
      ]);
    }

    if (data.action === 'approve' && data.id) {
      const row = findRow_(sheet, data.id);
      if (row > 0) {
        const log = JSON.parse(sheet.getRange(row, 5).getValue());
        log.approved = true;
        sheet.getRange(row, 5).setValue(JSON.stringify(log));
      }
    }

    if (data.action === 'delete' && data.id) {
      const row = findRow_(sheet, data.id);
      if (row > 0) sheet.deleteRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function findRow_(sheet, id) {
  const ids = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 1;
  }
  return -1;
}
