import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const sourcePath = process.env.CHIKEMO_PURCHASE_FORM_SOURCE
  ? path.resolve(process.env.CHIKEMO_PURCHASE_FORM_SOURCE)
  : path.resolve('Chikemo購入フォーム.gs');

function createRange(sheet, row, column, value, numRows = 1) {
  return {
    getColumn: () => column,
    getLastRow: () => row + numRows - 1,
    getNumRows: () => numRows,
    getRow: () => row,
    getSheet: () => sheet,
    getValue: () => value,
  };
}

function loadScript({
  cells = {},
  effectiveUser = 'chikemo.info@chikemo.net',
  gmailError = null,
  resendStatus = 200,
} = {}) {
  const writes = [];
  const sent = [];
  const resendRequests = [];
  const createdTriggers = [];
  const deletedTriggers = [];

  const headers = {
    H1: '購入枚数',
    I1: 'お名前（スペースなし）',
    K1: 'メールアドレス',
    M1: '商品お届け先名',
    N1: '商品お届け先住所',
    U2: '入金',
    W2: '追跡番号',
    AD2: '送信結果',
    AE2: '送信メッセージ',
    AF2: '入金日',
  };

  const sheet = {
    getName: () => 'シート1',
    getParent: () => ({ getId: () => '1Pf2GPmzRdf32QlhX-OutkTD_fdkWcDfjVhOluNZy_0Q' }),
    getRange(row, column) {
      if (typeof row === 'string') {
        return { getDisplayValue: () => headers[row] ?? '' };
      }
      return {
        getValue: () => cells[`${row}:${column}`] ?? '',
        setValue(value) {
          cells[`${row}:${column}`] = value;
          writes.push({ row, column, value });
        },
      };
    },
  };

  const spreadsheet = {
    getId: () => '1Pf2GPmzRdf32QlhX-OutkTD_fdkWcDfjVhOluNZy_0Q',
    getSheetByName: (name) => (name === 'シート1' ? sheet : null),
  };

  const context = {
    CONFIG: {
      contactEmail: 'chikemo.info@gmail.com',
      resendFrom: 'チケモ運営事務局 <chikemo.info@chikemo.net>',
      senderName: 'チケモ運営事務局',
    },
    console,
    GmailApp: {
      sendEmail(to, subject, body, options) {
        sent.push({ to, subject, body, options });
        if (gmailError) throw gmailError;
      },
    },
    isQuotaError_: (error) => /limit|quota|too many/i.test(String(error)),
    ScriptApp: {
      deleteTrigger(trigger) {
        deletedTriggers.push(trigger);
      },
      getProjectTriggers: () => [],
      newTrigger(handler) {
        const trigger = { handler };
        return {
          forSpreadsheet(target) {
            trigger.spreadsheet = target;
            return this;
          },
          onEdit() {
            return this;
          },
          create() {
            createdTriggers.push(trigger);
            return trigger;
          },
        };
      },
    },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => 'test-resend-key' }),
    },
    sendViaResend_(to, subject, body) {
      resendRequests.push({
        url: 'https://api.resend.com/emails',
        options: {
          payload: JSON.stringify({
            from: 'チケモ運営事務局 <chikemo.info@chikemo.net>',
            to: [to],
            subject,
            text: body,
          }),
        },
      });
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => effectiveUser }) },
    SpreadsheetApp: { openById: () => spreadsheet },
    UrlFetchApp: {
      fetch(url, options) {
        resendRequests.push({ url, options });
        return {
          getContentText: () => 'response body',
          getResponseCode: () => resendStatus,
        };
      },
    },
    Utilities: { formatDate: () => '2026/07/23 12:34:56' },
  };

  vm.createContext(context);
  vm.runInContext(fs.readFileSync(sourcePath, 'utf8'), context, { filename: sourcePath });
  return { context, createdTriggers, deletedTriggers, resendRequests, sent, sheet, spreadsheet, writes };
}

test('U列をOKにすると指定本文で追跡番号メールを1通送る', () => {
  const cells = {
    '3:8': '2',
    '3:9': '山田太郎',
    '3:11': 'customer@example.com',
    '3:13': '山田太郎',
    '3:14': '東京都千代田区1-1',
    '3:21': 'OK',
    '3:23': 'ABCD123456JP',
  };
  const { context, sent, sheet, writes } = loadScript({ cells });
  const event = { range: createRange(sheet, 3, 21, 'OK') };

  context.handleChikemoPurchaseFormEdit(event);

  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, 'customer@example.com');
  assert.equal(sent[0].subject, '【チケモ】追跡番号のお知らせ');
  assert.equal(sent[0].body, [
    '【追跡番号】ABCD123456JP',
    '',
    '山田太郎様',
    '',
    'いつもチケモをご利用いただき、誠にありがとうございます。',
    'ご入金が確認できましたので、商品の発送についてお知らせいたします。',
    '',
    '【ご注文商品】',
    '・商品名：全国百貨店共通商品券（1,000円分）',
    '・購入枚数：2',
    '',
    '【発送について】',
    '商品の発送はご入金から2日以内に行います。',
    '・送付先名：山田太郎',
    '・送付先住所：東京都千代田区1-1',
    '・発送方法：日本郵便 レターパックライト',
    '・到着予定：発送から1〜3日',
    '※ポスト投函でのお届けとなります（受取サイン不要）',
    '※追跡番号の反映はポスト投函から半日程度時間を要します',
    '',
    'ご不明な点がございましたら、「必ずお名前を添えて」下記メールアドレスまでお問い合わせください。',
    'chikemo.info@chikemo.net',
    '',
    'この度はご利用いただき誠にありがとうございました。',
    '今後とも、チケモをよろしくお願いいたします。',
    '',
    'チケモ運営事務局',
  ].join('\n'));
  assert.equal(sent[0].options.name, 'チケモ運営事務局');
  assert.equal(sent[0].options.replyTo, 'chikemo.info@chikemo.net');
  assert.deepEqual(
    writes.map(({ row, column, value }) => ({ row, column, value })),
    [
      { row: 3, column: 30, value: '送信済み' },
      { row: 3, column: 31, value: '' },
      { row: 3, column: 32, value: '2026/07/23 12:34:56' },
    ],
  );
});

test('追跡番号が空なら送信せずAD/AEにエラーを記録する', () => {
  const cells = {
    '3:8': '2',
    '3:9': '山田太郎',
    '3:11': 'customer@example.com',
    '3:13': '山田太郎',
    '3:14': '東京都千代田区1-1',
    '3:21': 'OK',
  };
  const { context, sent, sheet, writes } = loadScript({ cells });

  context.handleChikemoPurchaseFormEdit({ range: createRange(sheet, 3, 21, 'OK') });

  assert.equal(sent.length, 0);
  assert.deepEqual(
    writes.map(({ row, column, value }) => ({ row, column, value })),
    [
      { row: 3, column: 30, value: 'エラー' },
      { row: 3, column: 31, value: '追跡番号が空' },
    ],
  );
});

test('送信済み行は二重送信しない', () => {
  const cells = { '3:21': 'OK', '3:30': '送信済み' };
  const { context, sent, sheet, writes } = loadScript({ cells });

  context.handleChikemoPurchaseFormEdit({ range: createRange(sheet, 3, 21, 'OK') });

  assert.equal(sent.length, 0);
  assert.equal(writes.length, 0);
});

test('Gmailクォータ超過時はResendで同じ内容を送る', () => {
  const cells = {
    '3:8': '2',
    '3:9': '山田太郎',
    '3:11': 'customer@example.com',
    '3:13': '山田太郎',
    '3:14': '東京都千代田区1-1',
    '3:21': 'OK',
    '3:23': 'ABCD123456JP',
  };
  const { context, resendRequests, sheet, writes } = loadScript({
    cells,
    gmailError: new Error('Service invoked too many times: email quota'),
  });

  context.handleChikemoPurchaseFormEdit({ range: createRange(sheet, 3, 21, 'OK') });

  assert.equal(resendRequests.length, 1);
  assert.equal(resendRequests[0].url, 'https://api.resend.com/emails');
  const payload = JSON.parse(resendRequests[0].options.payload);
  assert.equal(payload.from, 'チケモ運営事務局 <chikemo.info@chikemo.net>');
  assert.deepEqual(payload.to, ['customer@example.com']);
  assert.equal(payload.subject, '【チケモ】追跡番号のお知らせ');
  assert.equal(writes.at(-3).value, '送信済み');
});

test('別シート・別列・OK以外は送信しない', () => {
  const { context, sent, sheet, writes } = loadScript();

  context.handleChikemoPurchaseFormEdit({ range: createRange(sheet, 3, 23, 'OK') });
  context.handleChikemoPurchaseFormEdit({ range: createRange(sheet, 3, 21, 'NG') });
  const otherSheet = { ...sheet, getName: () => '追跡番号' };
  context.handleChikemoPurchaseFormEdit({ range: createRange(otherSheet, 3, 21, 'OK') });

  assert.equal(sent.length, 0);
  assert.equal(writes.length, 0);
});

test('U列のOK【トット】とOK【モット】も送信対象にする', () => {
  for (const paymentValue of ['OK【トット】', 'OK【モット】']) {
    const cells = {
      '3:8': '1',
      '3:9': 'テスト申込者',
      '3:11': 'test@example.com',
      '3:13': 'テスト申込者',
      '3:14': 'テスト住所',
      '3:21': paymentValue,
      '3:23': 'TEST-TRACKING',
    };
    const { context, sent, sheet } = loadScript({ cells });

    context.handleChikemoPurchaseFormEdit({
      range: createRange(sheet, 3, 21, paymentValue),
    });

    assert.equal(sent.length, 1, paymentValue);
  }
});

test('チケモアカウント以外ではトリガーを作成しない', () => {
  const { context, createdTriggers } = loadScript({ effectiveUser: 'other@example.com' });

  assert.throws(
    () => context.setupChikemoPurchaseFormAutomation(),
    /chikemo\.info@chikemo\.net/,
  );
  assert.equal(createdTriggers.length, 0);
});

test('専用セットアップは対象スプレッドシートの編集トリガーを1つ作る', () => {
  const { context, createdTriggers, spreadsheet } = loadScript();

  context.setupChikemoPurchaseFormAutomation();

  assert.equal(createdTriggers.length, 1);
  assert.equal(createdTriggers[0].handler, 'handleChikemoPurchaseFormEdit');
  assert.equal(createdTriggers[0].spreadsheet, spreadsheet);
});

test('購入フォーム専用clasp設定が本番プロジェクトを参照する', () => {
  const config = JSON.parse(
    fs.readFileSync(path.resolve('.clasp.chikemo-purchase.json'), 'utf8'),
  );
  assert.equal(
    config.scriptId,
    '1D8gGYjvfinrWlm3wGXVNLIewoEVFX3Hs06Z0PRRrzmgunBQAw5WKzIKU',
  );
});

test('Resend送信はChikemoの安定運用どおりメールAPIだけを使用する', () => {
  const source = fs.readFileSync(sourcePath, 'utf8');
  assert.match(source, /https:\/\/api\.resend\.com\/emails/);
  assert.doesNotMatch(source, /https:\/\/api\.resend\.com\/domains/);
  assert.doesNotMatch(source, /re_[A-Za-z0-9]{10,}/);
});
