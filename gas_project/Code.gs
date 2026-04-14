/* ================================================================
   共通設定
================================================================ */
const SPREADSHEET_ID    = '1yIjoC8_LGSe97jNICheADhjs5LRnkoXrv6pj6D0PVQQ';
const LINE_NOTIFY_TOKEN = 'YOUR_LINE_NOTIFY_TOKEN'; // ← LINE Notifyトークンをここに設定

/* ================================================================
   doPost — エントリーポイント
   type フィールドでルーティング
================================================================ */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'comparison') return handleComparison(data);
    return handleDietSim(data);
  } catch (error) {
    return jsonResponse({ status: 'error', message: String(error) });
  }
}

/* ================================================================
   ジム比較ログ
   シート名: ジム比較
================================================================ */
function handleComparison(data) {
  const SHEET_NAME = 'ジム比較';
  const HEADERS = [
    '記録日時(JST)', 'トリガー', 'タブ', '比較ジム名',
    'しばジム総額', '比較ジム総額', '差額',
    'しばジム総額(換算)', '比較ジム総額(換算)', '換算回数', '送信時刻(UTC)'
  ];

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setFontWeight('bold').setBackground('#fff2e0').setFontColor('#c95318');
    sheet.setFrozenRows(1);
  }

  const jst        = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const trigger    = data.trigger    || 'auto';
  const tab        = data.tab        || 'personal';
  const normCount  = data.normCount  || 0;
  const gyms       = data.gyms       || [];
  const totals     = data.totals     || [];
  const normTotals = data.normTotals || [];

  const ownTotal     = totals[0]     || 0;
  const ownNormTotal = normTotals[0] || 0;

  let recorded = false;
  for (let i = 1; i < gyms.length; i++) {
    const gymName      = gyms[i]       || '未入力';
    const gymTotal     = totals[i]     || 0;
    const gymNormTotal = normTotals[i] || 0;
    sheet.appendRow([
      jst, trigger, tab, gymName,
      ownTotal, gymTotal, gymTotal - ownTotal,
      ownNormTotal, gymNormTotal, normCount,
      data.timestamp || ''
    ]);
    recorded = true;
  }

  if (!recorded) {
    sheet.appendRow([
      jst, trigger, tab, '（未比較）',
      ownTotal, 0, 0,
      ownNormTotal, 0, normCount,
      data.timestamp || ''
    ]);
  }

  return jsonResponse({ status: 'ok' });
}


/* ================================================================
   減量シミュレーター（既存処理）
================================================================ */
function handleDietSim(data) {
  const SHEET_NAME = '減量シミュ履歴';
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`シート '${SHEET_NAME}' が見つかりません。`);

  let eventName = '不明';
  if (data.eventType === 'form_calculated') eventName = '照会完了';
  else if (data.eventType === 'line_clicked') eventName = 'LINEタップ！';

  const rowData = [
    data.timestamp||new Date().toLocaleString('ja-JP'), eventName,
    data.sessionId||'', data.gender||'', data.age||'', data.height||'',
    data.weight||'', data.bmi||'', data.guaranteeKg||'', data.expectedKg||'',
    data.matchCount||'', data.matchAvg||'', data.utm_source||'',
    data.utm_medium||'', data.utm_campaign||'', data.utm_term||'',
    data.utm_content||'', data.referrer||'', data.deviceOs||'',
    data.browser||'', data.timeSpent||'', data.screenSize||'',
    data.userLanguage||'', data.network||'', data.rawUserAgent||''
  ];
  sheet.appendRow(rowData);
  if (data.eventType === 'line_clicked') {
    sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn()).setBackground('#e0f7fa');
  }
  return jsonResponse({ status: 'success' });
}

/* ================================================================
   共通レスポンス
================================================================ */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
