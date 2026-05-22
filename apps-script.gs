// =====================================================
// Google Apps Script — ネタ帳 API v2
// =====================================================

const SPREADSHEET_ID = '14XxhaLyv6VU5F8fiZTRg1_vfIah7gimZMNXZ3EC7hak';
const SHEET_NAME     = 'ネタ帳';

// ── POST: ネタ追加 / ステータス変更 ──────────────────
function doPost(e) {
  try {
    const action = (e.parameter.action || 'add').trim();
    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet  = getOrCreateSheet(ss);

    // ステータス変更
    if (action === 'updateStatus') {
      const ts        = (e.parameter.timestamp || '').trim();
      const newStatus = (e.parameter.status    || '').trim();
      if (!ts || !newStatus) return jsonRes({ status: 'error', message: 'missing params' });

      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === ts) {
          sheet.getRange(i + 1, 6).setValue(newStatus);
          return jsonRes({ status: 'ok' });
        }
      }
      return jsonRes({ status: 'error', message: 'row not found' });
    }

    // ネタ追加（デフォルト）
    const keyword  = (e.parameter.keyword  || '').trim();
    const genre    = (e.parameter.genre    || 'その他').trim();
    const priority = (e.parameter.priority || '中').trim();
    const memo     = (e.parameter.memo     || '').trim();

    if (!keyword) return jsonRes({ status: 'error', message: 'keyword is required' });

    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    sheet.appendRow([now, keyword, genre, priority, memo, 'ネタ']);
    return jsonRes({ status: 'ok' });

  } catch (err) {
    return jsonRes({ status: 'error', message: err.toString() });
  }
}

// ── GET: 一覧取得（JSONP対応） ────────────────────────
function doGet(e) {
  const action   = (e.parameter.action   || '').trim();
  const callback = (e.parameter.callback || '').trim();

  if (action === 'list') {
    try {
      const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);

      if (!sheet || sheet.getLastRow() <= 1) {
        return jsonpRes('[]', callback);
      }

      const data    = sheet.getDataRange().getValues();
      const headers = data[0].map(String);
      const rows    = data.slice(1)
        .reverse()                          // 新しい順
        .slice(0, 100)                      // 最大100件
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = String(row[i] ?? ''); });
          return obj;
        });

      return jsonpRes(JSON.stringify(rows), callback);

    } catch (err) {
      return jsonpRes('[]', callback);
    }
  }

  // ヘルスチェック
  return ContentService.createTextOutput('ネタ帳 API v2 ✅');
}

// ── ヘルパー ─────────────────────────────────────────
function getOrCreateSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const header = sheet.getRange(1, 1, 1, 6);
    header.setValues([['日時', 'キーワード', 'ジャンル', '優先度', 'メモ', 'ステータス']]);
    header.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 260);
    sheet.setColumnWidth(3, 90);  sheet.setColumnWidth(4, 70);
    sheet.setColumnWidth(5, 300); sheet.setColumnWidth(6, 90);
  }
  return sheet;
}

function jsonRes(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpRes(json, callback) {
  const body = callback ? `${callback}(${json})` : json;
  const mime = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(body).setMimeType(mime);
}
