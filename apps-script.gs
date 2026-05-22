// =====================================================
// Google Apps Script — ネタ帳 API
// スプレッドシートの「拡張機能」→「Apps Script」に
// このコードを貼り付けて使ってください。
// =====================================================

const SPREADSHEET_ID = '14XxhaLyv6VU5F8fiZTRg1_vfIah7gimZMNXZ3EC7hak';
const SHEET_NAME     = 'ネタ帳';

// POST: ネタを追加する
function doPost(e) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName(SHEET_NAME);

    // シートが存在しない場合は自動作成
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['日時', 'キーワード', 'ジャンル', '優先度', 'メモ', 'ステータス']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setColumnWidth(1, 140);
      sheet.setColumnWidth(2, 260);
      sheet.setColumnWidth(3, 90);
      sheet.setColumnWidth(4, 70);
      sheet.setColumnWidth(5, 300);
      sheet.setColumnWidth(6, 90);
    }

    const keyword  = (e.parameter.keyword  || '').trim();
    const genre    = (e.parameter.genre    || 'その他').trim();
    const priority = (e.parameter.priority || '中').trim();
    const memo     = (e.parameter.memo     || '').trim();

    if (!keyword) {
      return jsonResponse({ status: 'error', message: 'keyword is required' });
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    sheet.appendRow([now, keyword, genre, priority, memo, 'ネタ']);

    return jsonResponse({ status: 'ok' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// GET: 動作確認用
function doGet() {
  return ContentService.createTextOutput('ネタ帳 API is running ✅');
}

// ヘルパー
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
