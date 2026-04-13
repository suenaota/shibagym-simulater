/**
 * ジム比較ツール — GAS バックエンド
 *
 * 【設定】
 *   SPREADSHEET_ID : 記録先スプレッドシートの ID（URLの /d/XXXXX/ 部分）
 *   SHEET_NAME     : 書き込むシート名
 */
var SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
var SHEET_NAME     = '比較ログ';

/* ヘッダー定義 */
var HEADERS = [
  '記録日時(JST)',
  'トリガー',
  'タブ',
  '比較ジム名',
  'しばジム総額',
  '比較ジム総額',
  '差額',
  'しばジム総額(換算)',
  '比較ジム総額(換算)',
  '換算回数',
  '送信時刻(UTC)'
];

/* ─────────────────────────────────────
   doPost — フロントエンドから呼ばれるエンドポイント
───────────────────────────────────── */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // ヘッダー行がなければ追加
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
           .setFontWeight('bold')
           .setBackground('#f0f0f0');
      sheet.setFrozenRows(1);
    }

    var jst = Utilities.formatDate(
      new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'
    );

    var trigger    = data.trigger    || 'auto';
    var tab        = data.tab        || 'personal';
    var normCount  = data.normCount  || 0;
    var gyms       = data.gyms       || [];
    var totals     = data.totals     || [];
    var normTotals = data.normTotals || [];

    var ownTotal     = totals[0]     || 0;
    var ownNormTotal = normTotals[0] || 0;

    // しばジム以外のジムごとに1行記録
    var recorded = false;
    for (var i = 1; i < gyms.length; i++) {
      var gymName      = gyms[i]       || '未入力';
      var gymTotal     = totals[i]     || 0;
      var gymNormTotal = normTotals[i] || 0;

      sheet.appendRow([
        jst,
        trigger,
        tab,
        gymName,
        ownTotal,
        gymTotal,
        gymTotal - ownTotal,
        ownNormTotal,
        gymNormTotal,
        normCount,
        data.timestamp || ''
      ]);
      recorded = true;
    }

    // 比較ジムが0社の場合でも記録（しばジムのみ閲覧）
    if (!recorded) {
      sheet.appendRow([
        jst, trigger, tab, '（未比較）',
        ownTotal, 0, 0,
        ownNormTotal, 0, normCount,
        data.timestamp || ''
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error('doPost error:', err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ─────────────────────────────────────
   doGet — GETアクセス時の簡易応答
───────────────────────────────────── */
function doGet(e) {
  return ContentService
    .createTextOutput('このエンドポイントは POST 専用です。')
    .setMimeType(ContentService.MimeType.TEXT);
}
