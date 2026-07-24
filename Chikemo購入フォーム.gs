var CHIKEMO_PURCHASE_FORM = {
  spreadsheetId: '1Pf2GPmzRdf32QlhX-OutkTD_fdkWcDfjVhOluNZy_0Q',
  sheetName: 'シート1',
  senderEmail: 'chikemo.info@chikemo.net',
  senderName: 'チケモ運営事務局',
  subject: '【チケモ】追跡番号のお知らせ',
  firstDataRow: 3,
  columns: {
    quantity: 8,         // H
    name: 9,             // I
    email: 11,           // K
    deliveryName: 13,    // M
    deliveryAddress: 14, // N
    payment: 21,         // U
    tracking: 23,        // W
    sendResult: 30,      // AD
    sendMessage: 31,     // AE
    paymentDate: 32,     // AF
  },
};

// インストール型編集トリガー専用。simple trigger化を避けるため onEdit と命名しない。
function handleChikemoPurchaseFormEdit(e) {
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== CHIKEMO_PURCHASE_FORM.sheetName) return;
  if (sheet.getParent().getId() !== CHIKEMO_PURCHASE_FORM.spreadsheetId) return;
  if (range.getRow() < CHIKEMO_PURCHASE_FORM.firstDataRow) return;
  if (range.getColumn() !== CHIKEMO_PURCHASE_FORM.columns.payment) return;

  var value = String(range.getValue()).trim();
  var paymentValuesToSend = ['OK', 'OK【トット】', 'OK【モット】'];
  if (paymentValuesToSend.indexOf(value) === -1) return; // 「完了」やNGでは送らない

  // 複数行貼り付けは先頭行だけ処理し、残りは誤送信防止のため警告する。
  if (range.getNumRows() > 1) {
    for (var row = range.getRow() + 1; row <= range.getLastRow(); row++) {
      setChikemoPurchaseFormError_(
        sheet,
        row,
        '複数行まとめてOKが入力されたため未処理。1行ずつOKを入れ直してください',
      );
    }
  }

  sendChikemoPurchaseFormPaymentEmail_(sheet, range.getRow());
}

function sendChikemoPurchaseFormPaymentEmail_(sheet, row) {
  var columns = CHIKEMO_PURCHASE_FORM.columns;
  if (getChikemoPurchaseFormCell_(sheet, row, columns.sendResult) === '送信済み') return;

  var data = {
    quantity: getChikemoPurchaseFormCell_(sheet, row, columns.quantity),
    name: getChikemoPurchaseFormCell_(sheet, row, columns.name),
    email: getChikemoPurchaseFormCell_(sheet, row, columns.email),
    deliveryName: getChikemoPurchaseFormCell_(sheet, row, columns.deliveryName),
    deliveryAddress: getChikemoPurchaseFormCell_(sheet, row, columns.deliveryAddress),
    tracking: getChikemoPurchaseFormCell_(sheet, row, columns.tracking),
  };

  var required = [
    ['メールアドレス', data.email],
    ['お名前', data.name],
    ['購入枚数', data.quantity],
    ['送付先名', data.deliveryName],
    ['送付先住所', data.deliveryAddress],
    ['追跡番号', data.tracking],
  ];
  for (var i = 0; i < required.length; i++) {
    if (!required[i][1]) {
      setChikemoPurchaseFormError_(sheet, row, required[i][0] + 'が空');
      return;
    }
  }

  var body = buildChikemoPurchaseFormBody_(data);

  try {
    assertChikemoPurchaseFormSender_();
    GmailApp.sendEmail(data.email, CHIKEMO_PURCHASE_FORM.subject, body, {
      name: CHIKEMO_PURCHASE_FORM.senderName,
      replyTo: CHIKEMO_PURCHASE_FORM.senderEmail,
    });
  } catch (gmailError) {
    if (isChikemoPurchaseFormQuotaError_(gmailError)) {
      try {
        sendChikemoPurchaseFormViaResend_(data.email, CHIKEMO_PURCHASE_FORM.subject, body);
      } catch (resendError) {
        setChikemoPurchaseFormError_(sheet, row, 'Resend fallback失敗: ' + String(resendError));
        return;
      }
    } else {
      setChikemoPurchaseFormError_(sheet, row, String(gmailError));
      return;
    }
  }

  sheet.getRange(row, columns.sendResult).setValue('送信済み');
  sheet.getRange(row, columns.sendMessage).setValue('');
  sheet.getRange(row, columns.paymentDate).setValue(
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
  );
}

function buildChikemoPurchaseFormBody_(data) {
  return '【追跡番号】' + data.tracking + '\n\n' +
    data.name + '様\n\n' +
    'いつもチケモをご利用いただき、誠にありがとうございます。\n' +
    'ご入金が確認できましたので、商品の発送についてお知らせいたします。\n\n' +
    '【ご注文商品】\n' +
    '・商品名：全国百貨店共通商品券（1,000円分）\n' +
    '・購入枚数：' + data.quantity + '\n\n' +
    '【発送について】\n' +
    '商品の発送はご入金から2日以内に行います。\n' +
    '・送付先名：' + data.deliveryName + '\n' +
    '・送付先住所：' + data.deliveryAddress + '\n' +
    '・発送方法：日本郵便 レターパックライト\n' +
    '・到着予定：発送から1〜3日\n' +
    '※ポスト投函でのお届けとなります（受取サイン不要）\n' +
    '※追跡番号の反映はポスト投函から半日程度時間を要します\n\n' +
    'ご不明な点がございましたら、「必ずお名前を添えて」下記メールアドレスまでお問い合わせください。\n' +
    CHIKEMO_PURCHASE_FORM.senderEmail + '\n\n' +
    'この度はご利用いただき誠にありがとうございました。\n' +
    '今後とも、チケモをよろしくお願いいたします。\n\n' +
    CHIKEMO_PURCHASE_FORM.senderName;
}

function getChikemoPurchaseFormCell_(sheet, row, column) {
  return String(sheet.getRange(row, column).getValue()).trim();
}

function setChikemoPurchaseFormError_(sheet, row, message) {
  sheet.getRange(row, CHIKEMO_PURCHASE_FORM.columns.sendResult).setValue('エラー');
  sheet.getRange(row, CHIKEMO_PURCHASE_FORM.columns.sendMessage).setValue(message);
}

function assertChikemoPurchaseFormSender_() {
  var email = String(Session.getEffectiveUser().getEmail()).toLowerCase();
  if (email !== CHIKEMO_PURCHASE_FORM.senderEmail) {
    throw new Error(
      '実行アカウンが ' + CHIKEMO_PURCHASE_FORM.senderEmail + 'ではありません: ' + email,
    );
  }
}

function isChikemoPurchaseFormQuotaError_(error) {
  var message = String(error).toLowerCase();
  return message.indexOf('limit') !== -1 ||
    message.indexOf('quota') !== -1 ||
    message.indexOf('too many') !== -1;
}

function sendChikemoPurchaseFormViaResend_(to, subject, body) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
  if (!apiKey) throw new Error('RESEND_API_KEY が未設定');

  var response = UrlFetchApp.fetch('https://api.resend.com/emails', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify({
      from: CHIKEMO_PURCHASE_FORM.senderName + ' <' + CHIKEMO_PURCHASE_FORM.senderEmail + '>',
      to: [to],
      subject: subject,
      text: body,
      reply_to: CHIKEMO_PURCHASE_FORM.senderEmail,
    }),
    muteHttpExceptions: true,
  });

  var status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Resend API error ' + status + ': ' + response.getContentText());
  }
}

// Apps Scriptエディタから chikemo.info@chikemo.net で1回だけ実行する。
function setupChikemoPurchaseFormAutomation() {
  assertChikemoPurchaseFormSender_();

  var spreadsheet = SpreadsheetApp.openById(CHIKEMO_PURCHASE_FORM.spreadsheetId);
  var sheet = spreadsheet.getSheetByName(CHIKEMO_PURCHASE_FORM.sheetName);
  if (!sheet) throw new Error('対象シートが見つかりません: ' + CHIKEMO_PURCHASE_FORM.sheetName);

  verifyChikemoPurchaseFormHeaders_(sheet);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'handleChikemoPurchaseFormEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleChikemoPurchaseFormEdit')
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  console.log('Chikemo_購入フォームの自動送信トリガーを設定しました');
}

function verifyChikemoPurchaseFormHeaders_(sheet) {
  var expected = [
    ['H1', '購入枚数'],
    ['I1', 'お名前（スペースなし）'],
    ['K1', 'メールアドレス'],
    ['M1', '商品お届け先名'],
    ['N1', '商品お届け先住所'],
    ['U2', '入金'],
    ['W2', '追跡番号'],
    ['AD2', '送信結果'],
    ['AE2', '送信メッセージ'],
    ['AF2', '入金日'],
  ];

  expected.forEach(function(item) {
    var actual = String(sheet.getRange(item[0]).getDisplayValue()).trim();
    if (actual !== item[1]) {
      throw new Error(item[0] + 'のヘッダーが予想と異なります。期待: ' + item[1] + ' / 実際: ' + actual);
    }
  });
}

// Resend設定確認用。管理アドレスにテストメールを1通送る。
function testChikemoResendConfiguration() {
  sendChikemoPurchaseFormViaResend_(
    'i7811832616@gmail.com',
    '【テスト】Chikemo Resend設定確認',
    'このメールはChikemo購入フォームのResend設定確認です。受信できていれば正常です。'
  );
  console.log('Resendテスト送信成功');
}

// 8466行目のテストデータだけをResend経由で送信する。
function sendTestRow8466ViaResend() {
  var row = 8466;
  var spreadsheet = SpreadsheetApp.openById(CHIKEMO_PURCHASE_FORM.spreadsheetId);
  var sheet = spreadsheet.getSheetByName(CHIKEMO_PURCHASE_FORM.sheetName);
  if (!sheet) throw new Error('対象シートが見つかりません');

  var columns = CHIKEMO_PURCHASE_FORM.columns;
  if (getChikemoPurchaseFormCell_(sheet, row, columns.sendResult) === '送信済み') {
    throw new Error('8466行目は送信済みです');
  }

  var data = {
    quantity: getChikemoPurchaseFormCell_(sheet, row, columns.quantity),
    name: getChikemoPurchaseFormCell_(sheet, row, columns.name),
    email: getChikemoPurchaseFormCell_(sheet, row, columns.email),
    deliveryName: getChikemoPurchaseFormCell_(sheet, row, columns.deliveryName),
    deliveryAddress: getChikemoPurchaseFormCell_(sheet, row, columns.deliveryAddress),
    tracking: getChikemoPurchaseFormCell_(sheet, row, columns.tracking),
  };

  var required = [
    ['メールアドレス', data.email],
    ['お名前', data.name],
    ['購入枚数', data.quantity],
    ['送付先名', data.deliveryName],
    ['送付先住所', data.deliveryAddress],
    ['追跡番号', data.tracking],
  ];
  for (var i = 0; i < required.length; i++) {
    if (!required[i][1]) throw new Error(required[i][0] + 'が空');
  }

  try {
    sendChikemoPurchaseFormViaResend_(
      data.email,
      CHIKEMO_PURCHASE_FORM.subject,
      buildChikemoPurchaseFormBody_(data)
    );
    sheet.getRange(row, columns.sendResult).setValue('送信済み');
    sheet.getRange(row, columns.sendMessage).setValue('');
    sheet.getRange(row, columns.paymentDate).setValue(
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    );
    console.log('8466行目のResend送信成功');
  } catch (error) {
    setChikemoPurchaseFormError_(sheet, row, 'Resendテスト失敗: ' + String(error));
    throw error;
  }
}
