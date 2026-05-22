// =====================================================
// Google Apps Script — ネタ帳 API v3
// =====================================================

const SPREADSHEET_ID  = '14XxhaLyv6VU5F8fiZTRg1_vfIah7gimZMNXZ3EC7hak';
const SHEET_NETA      = 'ネタ帳';
const SHEET_CONTENT   = 'コンテンツ管理';
const DRIVE_FOLDER    = 'ネタ帳コンテンツ';

// ── POST ────────────────────────────────────────────
function doPost(e) {
  try {
    const action = (e.parameter.action || 'add').trim();
    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ① ネタを追加
    if (action === 'add') {
      const sheet   = getOrCreateNetaSheet(ss);
      const keyword = (e.parameter.keyword  || '').trim();
      const genre   = (e.parameter.genre    || 'その他').trim();
      const prio    = (e.parameter.priority || '中').trim();
      const memo    = (e.parameter.memo     || '').trim();
      if (!keyword) return jsonRes({ status: 'error', message: 'keyword required' });
      sheet.appendRow([now(), keyword, genre, prio, memo, 'ネタ']);
      return jsonRes({ status: 'ok' });
    }

    // ② ネタのステータス変更
    if (action === 'updateStatus') {
      const sheet  = getOrCreateNetaSheet(ss);
      const ts     = (e.parameter.timestamp || '').trim();
      const status = (e.parameter.status    || '').trim();
      return updateRowField(sheet, 0, ts, 5, status);
    }

    // ③ ネタ→コンテンツ化（Google Doc自動生成）
    if (action === 'createContent') {
      const keyword = (e.parameter.keyword || '').trim();
      const genre   = (e.parameter.genre   || '').trim();
      const memo    = (e.parameter.memo    || '').trim();
      if (!keyword) return jsonRes({ status: 'error', message: 'keyword required' });

      // Google Docを作成
      const docUrl = createGoogleDoc(keyword, genre, memo);

      // コンテンツ管理シートに追加
      const sheet = getOrCreateContentSheet(ss);
      sheet.appendRow([
        now(),       // A: 作成日
        keyword,     // B: タイトル
        genre,       // C: ジャンル
        docUrl,      // D: ドキュメントURL
        '下書き',    // E: noteステータス
        '',          // F: note URL
        '',          // G: note公開日
        '未収録',    // H: Voicyステータス
        '',          // I: Voicy URL
        '',          // J: Voicy公開日
        '有料',      // K: 有料/無料
        memo         // L: メモ
      ]);
      return jsonRes({ status: 'ok', docUrl });
    }

    // ④ コンテンツ情報を更新
    if (action === 'updateContent') {
      const sheet    = getOrCreateContentSheet(ss);
      const ts       = (e.parameter.timestamp || '').trim();
      const field    = (e.parameter.field     || '').trim();
      const value    = (e.parameter.value     || '').trim();

      const fieldMap = {
        noteStatus:  4,  // E（0-indexed）
        noteUrl:     5,  // F
        noteDate:    6,  // G
        voicyStatus: 7,  // H
        voicyUrl:    8,  // I
        voicyDate:   9,  // J
        paid:        10, // K
        memo:        11  // L
      };
      const col = fieldMap[field];
      if (col === undefined) return jsonRes({ status: 'error', message: 'unknown field' });
      return updateRowField(sheet, 0, ts, col, value);
    }

    return jsonRes({ status: 'error', message: 'unknown action' });

  } catch (err) {
    return jsonRes({ status: 'error', message: err.toString() });
  }
}

// ── GET（JSONP対応） ──────────────────────────────────
function doGet(e) {
  const action   = (e.parameter.action   || '').trim();
  const callback = (e.parameter.callback || '').trim();

  // ネタ帳一覧
  if (action === 'list') {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NETA);
    const rows  = sheetToJson(sheet, 100);
    return jsonpRes(JSON.stringify(rows), callback);
  }

  // コンテンツ管理一覧
  if (action === 'contents') {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_CONTENT);
    const rows  = sheetToJson(sheet, 200);
    return jsonpRes(JSON.stringify(rows), callback);
  }

  return ContentService.createTextOutput('ネタ帳 API v3 ✅');
}

// ── Google Doc作成 ────────────────────────────────────
function createGoogleDoc(keyword, genre, memo) {
  // フォルダ取得（なければ作成）
  let folder;
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);

  // ドキュメント作成
  const doc  = DocumentApp.create(keyword);
  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(folder);

  const body = doc.getBody();
  body.clear();

  // ── テンプレート ──
  const h1 = DocumentApp.ParagraphHeading.HEADING1;
  const h2 = DocumentApp.ParagraphHeading.HEADING2;
  const h3 = DocumentApp.ParagraphHeading.HEADING3;

  body.appendParagraph(keyword).setHeading(h1);
  body.appendParagraph(`ジャンル：${genre}　｜　作成日：${now()}`).setItalic(true);
  if (memo) body.appendParagraph(`メモ：${memo}`).setItalic(true);
  body.appendHorizontalRule();

  // ── note記事 ──
  body.appendParagraph('📝 note 記事').setHeading(h1);

  body.appendParagraph('① 冒頭（読者が共感できる問題提起）').setHeading(h2);
  body.appendParagraph('※ 読者が「あるある！」と思える失敗や状況から書き始める。');

  body.appendParagraph('② 体験談（具体的な数字・費用・期間を入れる）').setHeading(h2);
  body.appendParagraph('※ 費用：円、期間：日/週/月、回数など、具体的な数字を必ず入れる。');

  body.appendParagraph('③ 解決策・学んだこと').setHeading(h2);
  body.appendParagraph('※ 失敗から得た知見を書く。同じ失敗をしないための情報。');

  body.appendParagraph('④ 読者へのアドバイス').setHeading(h2);
  body.appendParagraph('※ 「もし私がもう一度やるなら…」という視点でまとめる。');

  body.appendHorizontalRule();

  // ── Voicy台本 ──
  body.appendParagraph('🎙️ Voicy 台本').setHeading(h1);

  body.appendParagraph('【入り口トーク（〜2分）】').setHeading(h2);
  body.appendParagraph('※ 記事より軽いトーンで話す。リスナーが「もっと知りたい」と思う入り口だけ話す。\n例：「今日はXXXについて話すんですが、実はめちゃくちゃ詰まったことがあって…」');

  body.appendParagraph('【本題（〜3分）】').setHeading(h2);
  body.appendParagraph('※ 2〜3つのポイントを話す。詳細はnoteに誘導する。\n・ポイント1：\n・ポイント2：\n・ポイント3：');

  body.appendParagraph('【締め（〜30秒）】').setHeading(h2);
  body.appendParagraph('「詳しくはnoteの記事にまとめています。概要欄のリンクからぜひ読んでみてください。今日も聴いてくれてありがとうございました！」');

  doc.saveAndClose();

  return `https://docs.google.com/document/d/${doc.getId()}/edit`;
}

// ── ヘルパー ──────────────────────────────────────────
function now() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function jsonRes(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpRes(json, callback) {
  const body = callback ? `${callback}(${json})` : json;
  const mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(body).setMimeType(mime);
}

function sheetToJson(sheet, limit) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  return data.slice(1).reverse().slice(0, limit).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = String(row[i] ?? ''); });
    return obj;
  });
}

function updateRowField(sheet, keyCol, keyVal, targetCol, newVal) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyCol]).trim() === keyVal) {
      sheet.getRange(i + 1, targetCol + 1).setValue(newVal);
      return jsonRes({ status: 'ok' });
    }
  }
  return jsonRes({ status: 'error', message: 'row not found' });
}

// ネタ帳シートのセットアップ
function getOrCreateNetaSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NETA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NETA);
    setupSheet(sheet, ['日時','キーワード','ジャンル','優先度','メモ','ステータス'],
      [140, 260, 80, 70, 280, 80]);
  }
  return sheet;
}

// コンテンツ管理シートのセットアップ
function getOrCreateContentSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_CONTENT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONTENT);
    setupSheet(sheet,
      ['作成日','タイトル','ジャンル','ドキュメントURL','noteステータス','note URL','note公開日','Voicyステータス','Voicy URL','Voicy公開日','有料/無料','メモ'],
      [130, 240, 80, 260, 90, 220, 90, 90, 220, 90, 70, 200]);
  }
  return sheet;
}

function setupSheet(sheet, headers, widths) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}
