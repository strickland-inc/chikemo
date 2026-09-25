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
  try {
    if (!e || !e.range || e.range.getRow() <= 1) return;

    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'シート1') return;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = e.range.getRow();
    var paymentCol = findColumn_(headers, '入金');

    // 入金列以外の編集は無視
    if (paymentCol === 0 || e.range.getColumn() !== paymentCol) return;
    var value = String(e.range.getValue()).trim();

    // コピペ等で複数行同時編集された場合、先頭行のみ処理し残り行に警告を書く
    if (e.range.getNumRows() > 1 && (value === 'OK' || value === 'NG')) {
      var prefix = value === 'OK' ? '発送通知' : 'キャンセル通知';
      for (var r = row + 1; r <= e.range.getLastRow(); r++) {
        setCell_(sheet, r, headers, prefix + 'エラー',
          '複数行まとめて ' + value + ' が入力されました。この行は未処理です。個別に OK を入れ直してください');
      }
    }

    if (value === 'OK') {
      var shippingLockResult = withChikemoMainSendLock_(function() {
        sendShippingNotification_(sheet, row, headers);
      });
      if (!shippingLockResult.locked) {
        setLockBusyError_(sheet, row, headers, '発送通知');
      }
    }
    if (value === 'NG') {
      var cancelLockResult = withChikemoMainSendLock_(function() {
        sendCancellationNotification_(sheet, row, headers);
      });
      if (!cancelLockResult.locked) {
        setLockBusyError_(sheet, row, headers, 'キャンセル通知');
      }
    }
  } catch (err) {
    console.error('handleEdit failed:', err, 'row=', e && e.range && e.range.getRow());
    throw err;
  }
}

// 送信済みの確認・送信・記録を1つのスクリプトロックで囲む共通ヘルパー。
// fn内でさらにこのロックを取り直さないこと（二重取得を避ける。reprocessUnsentはループ全体を1回だけ囲む）。
function withChikemoMainSendLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { locked: false };
  try {
    fn();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
  return { locked: true };
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
  // 送信と記録を同じ try に入れない。送信後の書き込みが「too many」等で失敗すると
  // クォータ切れと誤判定し、Gmail で届いたメールを Resend でもう一度送ってしまうため。
  try {
    GmailApp.sendEmail(to, subject, body, {
      name: CONFIG.senderName,
      replyTo: CONFIG.contactEmail,
    });
  } catch (gmailErr) {
    if (!isQuotaError_(gmailErr)) {
      return setError_(sheet, row, headers, type, String(gmailErr));
    }
    try {
      sendViaResend_(to, subject, body, computeChikemoMainIdempotencyKey_(type, to, subject, body));
    } catch (resendErr) {
      return setError_(sheet, row, headers, type, 'Resend fallback失敗: ' + String(resendErr));
    }
  }

  setCell_(sheet, row, headers, type + '済み', '送信済み');
  SpreadsheetApp.flush();
  setCell_(sheet, row, headers, type + '日時', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
  setCell_(sheet, row, headers, type + 'エラー', '');
}

// ===== Gmail クォータエラー判定 =====
function isQuotaError_(err) {
  var msg = String(err).toLowerCase();
  return msg.indexOf('limit') !== -1
    || msg.indexOf('quota') !== -1
    || msg.indexOf('too many') !== -1;
}

// ===== Resend API によるメール送信 =====
function sendViaResend_(to, subject, body, idempotencyKey) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
  if (!apiKey) throw new Error('RESEND_API_KEY が未設定');

  var headers = { 'Authorization': 'Bearer ' + apiKey };
  if (idempotencyKey) headers['Idempotency-Key'] = idempotencyKey;

  var res = UrlFetchApp.fetch('https://api.resend.com/emails', {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
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
  if (code === 409) {
    throw new Error('同じ内容のメールを別の処理が送信中です。二重送信を防ぐため送っていません。入金欄は入れ直さず、しばらく後に送信結果を確認してください: ' + res.getContentText());
  }
  if (code < 200 || code >= 300) {
    throw new Error('Resend API error ' + code + ': ' + res.getContentText());
  }
}

// type（発送通知／キャンセル通知）ごとに接頭辞を変え、宛先・件名・本文からIdempotency-Keyを作る。
function computeChikemoMainIdempotencyKey_(type, to, subject, body) {
  var prefix = type === '発送通知' ? 'chikemo/shipping/' : 'chikemo/cancel/';
  var text = [String(to).toLowerCase().trim(), subject, body].join('\n');
  var digestBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < digestBytes.length; i++) {
    hex += ('0' + ((digestBytes[i] + 256) % 256).toString(16)).slice(-2);
  }
  return (prefix + hex).slice(0, 256);
}

// ロックが取れなかった行にはエラー欄だけを書く。「〜済み」欄を「エラー」にすると、
// 実行中の処理が書いた「送信済み」を消し、次の編集で再送してしまうため触らない。
function setLockBusyError_(sheet, row, headers, type) {
  setCell_(sheet, row, headers, type + 'エラー', '別の送信処理が実行中のため未処理です。少し待ってから入金欄を入れ直してください');
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
  setupPaymentDropdown();
  setupMonitoringFormula();
  Logger.log('自動処理セットアップ完了');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
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
        'IF((S2:S="OK")*(AD2:AD<>"送信済み"),"要確認：GASのsetupTrigger関数を実行して権限を承認してください（発送通知未完了）",' +
          'IF((S2:S="NG")*(AG2:AG="エラー"),"要確認：キャンセル通知エラー",' +
            'IF((S2:S="NG")*(AG2:AG<>"送信済み"),"要確認：GASのsetupTrigger関数を実行して権限を承認してください（キャンセル通知未完了）",""))))))'
  );

  Logger.log('処理監視列（AJ列）に ARRAYFORMULA 設定完了');
}

// 入金=OK かつ 発送通知済みが「送信済み」以外の行を一括再送する緊急リカバリ関数。
// トリガー失敗/無言スキップが疑われる時にエディタから手動実行する。
// sendShippingNotification_ 側で既送信行は自動スキップされるので二重送信にならない。
function reprocessUnsent() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('対象行なし'); return; }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var paymentCol = findColumn_(headers, '入金');
  var statusCol = findColumn_(headers, '発送通知済み');
  if (paymentCol === 0 || statusCol === 0) {
    Logger.log('入金 or 発送通知済み 列が見つかりません');
    return;
  }

  var payments = sheet.getRange(2, paymentCol, lastRow - 1, 1).getValues();
  var statuses = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
  var processed = 0;

  // ループ全体を1回のロックで囲む。内側の sendShippingNotification_ は
  // ロックを取り直さない（withChikemoMainSendLock_ の二重取得を避けるため）。
  var lockResult = withChikemoMainSendLock_(function() {
    for (var i = 0; i < payments.length; i++) {
      var payment = String(payments[i][0]).trim();
      var status = String(statuses[i][0]).trim();
      if (payment === 'OK' && status !== '送信済み') {
        sendShippingNotification_(sheet, 2 + i, headers);
        processed++;
      }
    }
  });

  if (!lockResult.locked) {
    Logger.log('別の送信処理が実行中のため reprocessUnsent を中断しました');
    return;
  }

  Logger.log('reprocessUnsent 完了: ' + processed + ' 件処理');
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
