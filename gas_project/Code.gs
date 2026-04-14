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
    '記録日時(JST)', 'セッションID', 'トリガー', 'タブ',
    '比較ジム名', '比較ジム数',
    /* しばジム */
    'しばジム総額', 'しばジム換算額', 'しばジム1回単価', 'しばジム回数', 'しばジムオプション',
    /* 比較ジム */
    '比較ジム総額', '比較ジム換算額', '比較ジム1回単価', '比較ジム回数', '比較ジムオプション',
    '差額',
    /* 比較ジム 各行入力値 */
    '入力データ(JSON)',
    /* 換算 */
    '換算回数',
    /* デバイス情報 */
    'デバイス', '画面サイズ', '流入元', '滞在時間(秒)',
    '送信時刻(UTC)'
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
  const sessionId  = data.sessionId  || '';
  const timeSpent  = data.timeSpent  || 0;
  const device     = data.device     || '';
  const screenSize = data.screenSize || '';
  const referrer   = data.referrer   || '';
  const details    = data.gymDetails || [];

  /* しばジム情報（details[0]） */
  const own = details[0] || {};
  const ownTotal     = own.total     || 0;
  const ownNormTotal = own.normTotal || 0;
  const ownUnit      = own.unitPrice != null ? own.unitPrice : '';
  const ownCount     = own.count     != null ? own.count : '';
  const ownOpts      = (own.options || []).join(', ');

  /* 比較ジムごとに1行 */
  let recorded = false;
  for (let i = 1; i < details.length; i++) {
    const g = details[i] || {};
    const gymName      = g.name       || '未入力';
    const gymTotal     = g.total      || 0;
    const gymNormTotal = g.normTotal  || 0;
    const gymUnit      = g.unitPrice  != null ? g.unitPrice : '';
    const gymCount     = g.count      != null ? g.count : '';
    const gymOpts      = (g.options || []).join(', ');
    const rowsJson     = JSON.stringify(g.rows || {});

    sheet.appendRow([
      jst, sessionId, trigger, tab,
      gymName, data.gymCount || 0,
      ownTotal, ownNormTotal, ownUnit, ownCount, ownOpts,
      gymTotal, gymNormTotal, gymUnit, gymCount, gymOpts,
      gymTotal - ownTotal,
      rowsJson,
      normCount,
      device, screenSize, referrer, timeSpent,
      data.timestamp || ''
    ]);
    recorded = true;
  }

  if (!recorded) {
    sheet.appendRow([
      jst, sessionId, trigger, tab,
      '（未比較）', 0,
      ownTotal, ownNormTotal, ownUnit, ownCount, ownOpts,
      0, 0, '', '', '',
      0,
      '',
      normCount,
      device, screenSize, referrer, timeSpent,
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
