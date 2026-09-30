// ===== 設定 =====
var CONFIG = {
  senderName: 'チケモ運営事務局',
  contactEmail: 'chikemo.info@gmail.com',
  resendFrom: 'チケモ運営事務局 <chikemo.info@chikemo.net>',
  shippingMethod: '日本郵便 レターパックライト',
};

var COLUMN_FALLBACKS = {
  '入金': 19, // S列
  '発送通知済み': 30, // AD列
  '発送通知日時': 31, // AE列
  '発送通知エラー': 32, // AF列
  'キャンセル通知済み': 33, // AG列
  'キャンセル通知日時': 34, // AH列
  'キャンセル通知エラー': 35, // AI列
  '処理監視': 36, // AJ列
};

// ===== メイン処理：入金列が変更されたら自動でメール送信 =====
// NOTE: 関数名を onEdit にすると simple trigger として自動発火し、
// AuthMode.LIMITED で GmailApp が権限エラーになるため handleEdit にしている。
function handleEdit(e) {
  if (!e || !e.range || e.range.getRow() <= 1) return;

  // 同時編集で取りこぼさないよう、処理を直列化する
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    processEdit_(e);
  } finally {
    lock.releaseLock();
  }
}

function processEdit_(e) {
  try {
    if (!e || !e.range || e.range.getRow() <= 1) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'シート1') return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = e.range.getRow();
    var paymentCol = findColumn_(headers, '入金');

    // 編集範囲に入金列が含まれない場合は無視する。
    // 複数列にまたがる貼り付けでも入金列を取りこぼさないよう、範囲の先頭列ではなく包含で判定する。
    if (paymentCol === 0 || paymentCol < e.range.getColumn() || paymentCol > e.range.getLastColumn()) return;

    // 編集イベントの値ではなく、セルの現在値を読む（複数列貼り付けで先頭セルが入金列とは限らないため）
    var value = normalizePayment_(sheet.getRange(row, paymentCol).getValue());
    console.log('handleEdit row=' + row + ' rows=' + e.range.getNumRows() + ' value=' + value);

    // コピペ等で複数行同時編集された場合、先頭行のみ処理し残り行に警告を書く
    if (e.range.getNumRows() > 1) {
      var rest = sheet.getRange(row + 1, paymentCol, e.range.getNumRows() - 1, 1).getValues();
      for (var i = 0; i < rest.length; i++) {
        var restValue = normalizePayment_(rest[i][0]);
        if (restValue !== 'OK' && restValue !== 'NG') continue;
        var prefix = restValue === 'OK' ? '発送通知' : 'キャンセル通知';
        setCell_(sheet, row + 1 + i, headers, prefix + 'エラー',
          '複数行まとめて ' + restValue + ' が入力されました。この行は未処理です。個別に ' + restValue + ' を入れ直してください');
      }
    }

    if (value === 'OK') sendShippingNotification_(sheet, row, headers);
    else if (value === 'NG') sendCancellationNotification_(sheet, row, headers);
    else if (value) console.warn('入金列の値が OK/NG ではないため処理しません: value=' + value + ' row=' + row);
  } catch (err) {
    console.error('handleEdit failed:', err, 'row=', e && e.range && e.range.getRow());
    throw err;
  }
}

// スプレッドシートの = 比較は大文字小文字を区別しないため、AJ列は "ok" でも警告を出す。
// 処理側も同じ基準（全角半角・前後空白・大文字小文字を無視）で判定して取りこぼしを防ぐ。
function normalizePayment_(value) {
  return String(value).normalize('NFKC').trim().toUpperCase();
}

// ===== 発送通知メール =====
function sendShippingNotification_(sheet, row, headers) {
  if (getCell_(sheet, row, headers, '発送通知済み') === '送信済み') return;

  var email = getCell_(sheet, row, headers, 'メールアドレス');
  var tracking = getCell_(sheet, row, headers, '追跡番号');
  if (!email) return setError_(sheet, row, headers, '発送通知', 'メールアドレスが空');
  if (!tracking) return setError_(sheet, row, headers, '発送通知', '追跡番号が空');

  var name = getRecipientName_(sheet, row, headers);
  if (!name) return setError_(sheet, row, headers, '発送通知', '宛名が空（システム表示名 / お名前（スペースなし） / お名前 / 商品お届け先名）');
  var item = getCell_(sheet, row, headers, 'ご購入商品');
  var qty = getCell_(sheet, row, headers, '購入枚数');
  var toName = getCell_(sheet, row, headers, '商品お届け先名');
  var toAddr = getCell_(sheet, row, headers, '商品お届け先住所');

  var body =
    name + '様\n\n' +
    'チケモをご利用いただき、誠にありがとうございます。\n' +
    'ご入金が確認できましたので、お知らせいたします。\n\n' +
    '【ご注文商品】\n' +
    '・商品名：' + item + '\n' +
    '・購入枚数：' + qty + '\n\n' +
    '【発送について】\n' +
    '商品の発送は3日以内に行います。\n' +
    '追跡番号のご連絡は原則当日中にいたします。\n' +
    '・送付先名：' + toName + '\n' +
    '・送付先住所：' + toAddr + '\n' +
    '・発送方法：' + CONFIG.shippingMethod + '\n' +
    '・到着予定：発送から1〜3日\n' +
    '・追跡番号：' + tracking + '\n' +
    '※ポスト投函でのお届けとなります（受取サイン不要）\n' +
    '※追跡番号の反映はポスト投函から半日程度時間を要します\n\n' +
    'ご不明な点がございましたら、お名前を添えて下記までお問い合わせください。\n' +
    CONFIG.contactEmail + '\n\n' +
    'この度はご利用いただき誠にありがとうございました。\n' +
    '今後とも、チケモをよろしくお願いいたします。\n\n' +
    CONFIG.senderName;

  sendEmail_(sheet, row, headers, '発送通知', email, '【追跡番号のお知らせ】ご入金ありがとうございます', body);
}

// ===== キャンセル通知メール =====
function sendCancellationNotification_(sheet, row, headers) {
  if (getCell_(sheet, row, headers, 'キャンセル通知済み') === '送信済み') return;

  var email = getCell_(sheet, row, headers, 'メールアドレス');
  if (!email) return setError_(sheet, row, headers, 'キャンセル通知', 'メールアドレスが空');

  var name = getRecipientName_(sheet, row, headers);
  if (!name) return setError_(sheet, row, headers, 'キャンセル通知', '宛名が空（システム表示名 / お名前（スペースなし） / お名前 / 商品お届け先名）');
  var item = getCell_(sheet, row, headers, 'ご購入商品');
  var qty = getCell_(sheet, row, headers, '購入枚数');
  var toName = getCell_(sheet, row, headers, '商品お届け先名');
  var toAddr = getCell_(sheet, row, headers, '商品お届け先住所');

  var body =
    name + '様\n\n' +
    'チケモをご利用いただき、誠にありがとうございます。\n' +
    '以下のご注文につきまして、大変恐れ入りますが、キャンセルとさせていただきます。\n\n' +
    '【ご注文内容】\n' +
    '・商品名：' + item + '\n' +
    '・購入枚数：' + qty + '\n' +
    '・送付先名：' + toName + '\n' +
    '・送付先住所：' + toAddr + '\n\n' +
    '商品をご希望の際は、改めてLINEよりご注文ください。\n' +
    'チケモLINE公式アカウント：https://lin.ee/nbdod08F\n\n' +
    'この度はご利用いただき誠にありがとうございました。\n' +
    '今後とも、チケモをよろしくお願いいたします。\n\n' +
    CONFIG.senderName;

  sendEmail_(sheet, row, headers, 'キャンセル通知', email, '【チケモ】ご注文キャンセルのお知らせ', body);
}

// ===== メール送信 & ステータス記録 =====
function sendEmail_(sheet, row, headers, type, to, subject, body) {
  try {
    GmailApp.sendEmail(to, subject, body, {
      name: CONFIG.senderName,
      replyTo: CONFIG.contactEmail,
    });
    setCell_(sheet, row, headers, type + '済み', '送信済み');
    setCell_(sheet, row, headers, type + '日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
    setCell_(sheet, row, headers, type + 'エラー', '');
  } catch (gmailErr) {
    if (isQuotaError_(gmailErr)) {
      try {
        sendViaResend_(to, subject, body);
        setCell_(sheet, row, headers, type + '済み', '送信済み');
        setCell_(sheet, row, headers, type + '日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
        setCell_(sheet, row, headers, type + 'エラー', '');
      } catch (resendErr) {
        setCell_(sheet, row, headers, type + '済み', 'エラー');
        setCell_(sheet, row, headers, type + '日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
        setCell_(sheet, row, headers, type + 'エラー', 'Resend fallback失敗: ' + String(resendErr));
      }
    } else {
      setCell_(sheet, row, headers, type + '済み', 'エラー');
      setCell_(sheet, row, headers, type + '日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
      setCell_(sheet, row, headers, type + 'エラー', String(gmailErr));
    }
  }
}

// ===== Gmail クォータエラー判定 =====
function isQuotaError_(err) {
  var msg = String(err).toLowerCase();
  return msg.indexOf('limit') !== -1
    || msg.indexOf('quota') !== -1
    || msg.indexOf('too many') !== -1;
}

// ===== Resend API によるメール送信 =====
function sendViaResend_(to, subject, body) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
  if (!apiKey) throw new Error('RESEND_API_KEY が未設定');

  var res = UrlFetchApp.fetch('https://api.resend.com/emails', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey },
    payload: JSON.stringify({
      from: CONFIG.resendFrom,
      to: [to],
      subject: subject,
      text: body,
      reply_to: CONFIG.contactEmail,
    }),
    muteHttpExceptions: true,
  });

  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Resend API error ' + code + ': ' + res.getContentText());
  }
}

function setError_(sheet, row, headers, type, message) {
  setCell_(sheet, row, headers, type + '済み', 'エラー');
  setCell_(sheet, row, headers, type + '日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  setCell_(sheet, row, headers, type + 'エラー', message);
}

// ===== セル読み書き =====
function getRecipientName_(sheet, row, headers) {
  return getCell_(sheet, row, headers, 'システム表示名')
    || getCell_(sheet, row, headers, 'お名前（スペースなし）')
    || getCell_(sheet, row, headers, 'お名前')
    || getCell_(sheet, row, headers, '商品お届け先名');
}

function getCell_(sheet, row, headers, name) {
  var col = findColumn_(headers, name);
  return col > 0 ? String(sheet.getRange(row, col).getValue()).trim() : '';
}

function setCell_(sheet, row, headers, name, value) {
  var col = findColumn_(headers, name);
  if (col > 0) {
    sheet.getRange(row, col).setValue(value);
  } else {
    console.warn('列が見つからないため書き込みをスキップ:', name, 'row=', row);
  }
}

function findColumn_(headers, name) {
  var normalizedName = normalizeHeader_(name);

  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader_(headers[i]) === normalizedName) return i + 1;
  }

  var aliases = getHeaderAliases_(name);
  for (var a = 0; a < aliases.length; a++) {
    var normalizedAlias = normalizeHeader_(aliases[a]);
    for (var j = 0; j < headers.length; j++) {
      if (normalizeHeader_(headers[j]) === normalizedAlias) return j + 1;
    }
  }

  return COLUMN_FALLBACKS[name] || 0;
}

function normalizeHeader_(value) {
  return String(value)
    .replace(/[ 　\t\r\n]/g, '')
    .trim();
}

function getHeaderAliases_(name) {
  var aliases = {
    '入金': ['入金確認', '入金ステータス'],
    '発送通知済み': ['発送通知済', '発送通知送信済み', '発送通知ステータス'],
    '発送通知日時': ['発送通知日', '発送通知送信日時'],
    '発送通知エラー': ['発送通知エラー内容'],
    'キャンセル通知済み': ['キャンセル通知済', 'キャンセル通知送信済み', 'キャンセル通知ステータス'],
    'キャンセル通知日時': ['キャンセル通知日', 'キャンセル通知送信日時'],
    'キャンセル通知エラー': ['キャンセル通知エラー内容'],
  };
  return aliases[name] || [];
}

// ===== トリガー管理 =====
// 各 remove は対象の関数のトリガーだけを削除する。
// 全トリガーを削除すると、別のトリガー（sweepMissedEdits）まで消えてしまうため。
function setupTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  Logger.log('トリガー設定完了');
}

function setupAutomation() {
  setupTrigger();
  setupSweeper();
  setupPaymentDropdown();
  setupMonitoringFormula();
  Logger.log('自動処理セットアップ完了');
}

function removeTrigger() {
  removeTriggersByHandler_('handleEdit');
}

function removeTriggersByHandler_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(t);
  });
}

// ===== 取りこぼし検知（5分おき）=====
// 編集トリガーは、同時編集・連続編集・一時的な実行失敗などで発火しないことがある。
// 前回の巡回時点から入金列が新たに OK/NG になった行のうち、通知が未処理のものだけを送信する。
// 過去から未送信のまま残っている行は対象外（AJ列の警告を見て reprocessUnsent で人が判断する）。
// 初回は現在の状態を記録するだけで、送信しない。
var SWEEP_SNAPSHOT_PREFIX = 'PAYMENT_SNAPSHOT_';
var SWEEP_CHUNK_SIZE = 8000;
var SWEEP_MAX_SENDS = 10; // 1回の巡回の上限。行の挿入・削除で状態がずれた場合の大量送信を防ぐ

function setupSweeper() {
  removeTriggersByHandler_('sweepMissedEdits');
  ScriptApp.newTrigger('sweepMissedEdits')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('取りこぼし検知トリガー設定完了');
}

function removeSweeper() {
  removeTriggersByHandler_('sweepMissedEdits');
  clearSweepSnapshot_();
}

function sweepMissedEdits() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    console.warn('sweepMissedEdits: ロックを取得できないため次回に回します');
    return;
  }
  try {
    sweepMissedEdits_();
  } finally {
    lock.releaseLock();
  }
}

function sweepMissedEdits_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  if (!sheet || sheet.getLastRow() < 2) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var paymentCol = findColumn_(headers, '入金');
  if (paymentCol === 0) return;

  var values = sheet.getRange(2, paymentCol, sheet.getLastRow() - 1, 1).getValues();
  var current = '';
  for (var i = 0; i < values.length; i++) {
    var v = normalizePayment_(values[i][0]);
    current += v === 'OK' ? 'O' : (v === 'NG' ? 'N' : '-');
  }

  var previous = loadSweepSnapshot_();
  if (previous === null) {
    saveSweepSnapshot_(current);
    console.log('sweepMissedEdits: 初回のため現在の状態を記録しました rows=' + current.length);
    return;
  }

  var changed = [];
  for (var j = 0; j < current.length; j++) {
    var before = j < previous.length ? previous.charAt(j) : '-';
    if (current.charAt(j) !== '-' && current.charAt(j) !== before) changed.push(j);
  }

  if (changed.length > SWEEP_MAX_SENDS) {
    console.warn('sweepMissedEdits: 変化が ' + changed.length + ' 行あり上限を超えたため送信せず状態だけ更新します（行の挿入・削除の可能性）');
    saveSweepSnapshot_(current);
    return;
  }

  changed.forEach(function(index) {
    var row = 2 + index;
    try {
      if (current.charAt(index) === 'O') {
        // 状態が空で、エラーも書かれていない行だけ（エラー行や複数行貼り付けの警告行は再送しない）
        if (getCell_(sheet, row, headers, '発送通知済み') === '' && getCell_(sheet, row, headers, '発送通知エラー') === '') {
          console.log('sweepMissedEdits: 発送通知を補完 row=' + row);
          sendShippingNotification_(sheet, row, headers);
        }
      } else if (getCell_(sheet, row, headers, 'キャンセル通知済み') === '' && getCell_(sheet, row, headers, 'キャンセル通知エラー') === '') {
        console.log('sweepMissedEdits: キャンセル通知を補完 row=' + row);
        sendCancellationNotification_(sheet, row, headers);
      }
    } catch (err) {
      console.error('sweepMissedEdits failed: row=' + row, err);
    }
  });

  saveSweepSnapshot_(current);
}

function loadSweepSnapshot_() {
  var props = PropertiesService.getScriptProperties();
  var count = props.getProperty(SWEEP_SNAPSHOT_PREFIX + 'COUNT');
  if (count === null) return null;
  var text = '';
  for (var i = 0; i < Number(count); i++) {
    text += props.getProperty(SWEEP_SNAPSHOT_PREFIX + i) || '';
  }
  return text;
}

function saveSweepSnapshot_(text) {
  clearSweepSnapshot_();
  var props = PropertiesService.getScriptProperties();
  var data = {};
  var count = Math.ceil(text.length / SWEEP_CHUNK_SIZE);
  for (var i = 0; i < count; i++) {
    data[SWEEP_SNAPSHOT_PREFIX + i] = text.substr(i * SWEEP_CHUNK_SIZE, SWEEP_CHUNK_SIZE);
  }
  data[SWEEP_SNAPSHOT_PREFIX + 'COUNT'] = String(count);
  props.setProperties(data);
}

function clearSweepSnapshot_() {
  var props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).forEach(function(key) {
    if (key.indexOf(SWEEP_SNAPSHOT_PREFIX) === 0) props.deleteProperty(key);
  });
}

function setupArrayFormulas() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');

  // 既存の値をクリア（ヘッダーは残す）
  var lastRow = sheet.getMaxRows();
  var colsToClear = [18, 22, 23, 24, 25, 26, 27]; // R, V, W, X, Y, Z, AA
  colsToClear.forEach(function(col) {
    if (lastRow > 1) sheet.getRange(2, col, lastRow - 1).clearContent();
  });

  // ARRAYFORMULA を2行目に設定
  sheet.getRange('R2').setFormula('=ARRAYFORMULA(IF(H2:H="","",IF(AC2:AC<>"",AC2:AC,(H2:H*999)+2900)))');
  sheet.getRange('V2').setFormula('=ARRAYFORMULA(IF(M2:M="","",M2:M))');
  sheet.getRange('W2').setFormula('=ARRAYFORMULA(IF(N2:N="","",N2:N))');
  sheet.getRange('X2').setFormula('=ARRAYFORMULA(IF(G2:G="","",G2:G))');
  sheet.getRange('Y2').setFormula('=ARRAYFORMULA(IF(H2:H="","",H2:H))');
  sheet.getRange('Z2').setFormula('=ARRAYFORMULA(IF(I2:I="","",I2:I))');
  sheet.getRange('AA2').setFormula('=ARRAYFORMULA(IF(J2:J="","",J2:J))');

  Logger.log('ARRAYFORMULA 設定完了');
}

function setupPaymentDropdown() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = findColumn_(headers, '入金');
  if (col === 0) { Logger.log('入金列が見つかりません'); return; }

  var range = sheet.getRange(2, col, sheet.getMaxRows() - 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OK', 'NG'], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  Logger.log('入金列（' + col + '列目）にプルダウン設定完了');
}

function setupMonitoringFormula() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var lastRow = sheet.getMaxRows();

  sheet.getRange('AJ1').setValue('処理監視');
  if (lastRow > 1) sheet.getRange(2, 36, lastRow - 1).clearContent();

  sheet.getRange('AJ2').setFormula(
    '=ARRAYFORMULA(IF(S2:S="","",' +
      'IF((S2:S="OK")*(AD2:AD="エラー"),"要確認：発送通知エラー",' +
        'IF((S2:S="OK")*(AD2:AD<>"送信済み"),"要確認：発送通知が未送信です。GASのreprocessUnsent関数を実行してください",' +
          'IF((S2:S="NG")*(AG2:AG="エラー"),"要確認：キャンセル通知エラー",' +
            'IF((S2:S="NG")*(AG2:AG<>"送信済み"),"要確認：キャンセル通知が未送信です。GASのreprocessUnsent関数を実行してください",""))))))'
  );

  Logger.log('処理監視列（AJ列）に ARRAYFORMULA 設定完了');
}

// 入金=OK かつ 発送通知済みが「送信済み」以外、
// および 入金=NG かつ キャンセル通知済みが「送信済み」以外の行を一括再送する緊急リカバリ関数。
// トリガー失敗/無言スキップが疑われる時にエディタから手動実行する。
// 実行前に対象行（AJ列の警告行）を確認すること。
// send*Notification_ 側で既送信行は自動スキップされるので二重送信にならない。
function reprocessUnsent() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('対象行なし'); return; }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var paymentCol = findColumn_(headers, '入金');
  var shipStatusCol = findColumn_(headers, '発送通知済み');
  var cancelStatusCol = findColumn_(headers, 'キャンセル通知済み');
  if (paymentCol === 0 || shipStatusCol === 0 || cancelStatusCol === 0) {
    Logger.log('入金 / 発送通知済み / キャンセル通知済み 列が見つかりません');
    return;
  }

  var payments = sheet.getRange(2, paymentCol, lastRow - 1, 1).getValues();
  var shipStatuses = sheet.getRange(2, shipStatusCol, lastRow - 1, 1).getValues();
  var cancelStatuses = sheet.getRange(2, cancelStatusCol, lastRow - 1, 1).getValues();
  var shipped = 0;
  var cancelled = 0;

  for (var i = 0; i < payments.length; i++) {
    var payment = String(payments[i][0]).trim();
    if (payment === 'OK' && String(shipStatuses[i][0]).trim() !== '送信済み') {
      sendShippingNotification_(sheet, 2 + i, headers);
      shipped++;
    } else if (payment === 'NG' && String(cancelStatuses[i][0]).trim() !== '送信済み') {
      sendCancellationNotification_(sheet, 2 + i, headers);
      cancelled++;
    }
  }

  Logger.log('reprocessUnsent 完了: 発送通知 ' + shipped + ' 件 / キャンセル通知 ' + cancelled + ' 件処理');
}

function testResend() {
  sendViaResend_(
    CONFIG.contactEmail,
    '【テスト】Resend API 送信テスト',
    'このメールは Resend API のテスト送信です。\n受信できていれば正常に動作しています。'
  );
  Logger.log('テスト送信完了');
}

function checkStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var msg = 'トリガー数: ' + triggers.length + '\n';
  triggers.forEach(function(t) {
    msg += '- ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')\n';
  });
  Logger.log(msg);
}
