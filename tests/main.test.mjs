import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const mainSourcePath = process.env.CHIKEMO_MAIN_SOURCE
  ? path.resolve(process.env.CHIKEMO_MAIN_SOURCE)
  : path.resolve('メイン.gs');

const purchaseFormSourcePath = process.env.CHIKEMO_PURCHASE_FORM_SOURCE
  ? path.resolve(process.env.CHIKEMO_PURCHASE_FORM_SOURCE)
  : path.resolve('Chikemo購入フォーム.gs');

const HEADERS = [
  'メールアドレス',
  '追跡番号',
  'お名前（スペースなし）',
  'ご購入商品',
  '購入枚数',
  '商品お届け先名',
  '商品お届け先住所',
  '入金',
  '発送通知済み',
  '発送通知日時',
  '発送通知エラー',
  'キャンセル通知済み',
  'キャンセル通知日時',
  'キャンセル通知エラー',
];
const COL = {};
HEADERS.forEach((name, i) => { COL[name] = i + 1; });

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

function createSheet(cells, writes, lastRow) {
  return {
    getName: () => 'シート1',
    getLastColumn: () => HEADERS.length,
    getLastRow: () => lastRow,
    getRange(row, column, numRows, numCols) {
      if (numRows !== undefined) {
        if (row === 1) return { getValues: () => [HEADERS] };
        const cols = numCols || 1;
        const values = [];
        for (let r = 0; r < numRows; r++) {
          const rowValues = [];
          for (let c = 0; c < cols; c++) {
            rowValues.push(cells[`${row + r}:${column + c}`] ?? '');
          }
          values.push(rowValues);
        }
        return { getValues: () => values };
      }
      return {
        getValue: () => cells[`${row}:${column}`] ?? '',
        setValue(value) {
          cells[`${row}:${column}`] = value;
          writes.push({ row, column, value });
          if (value === '送信済み') writes.push({ marker: 'write:sendResult' });
        },
      };
    },
  };
}

function loadScript({
  cells = {},
  lastRow = 10,
  gmailError = null,
  resendStatus = 200,
  lockAcquirable = true,
} = {}) {
  const writes = [];
  const sent = [];
  const resendRequests = [];
  const callOrder = [];

  const sheet = createSheet(cells, writes, lastRow);

  const context = {
    console,
    Logger: { log() {} },
    CONFIG: {
      senderName: 'チケモ運営事務局',
      contactEmail: 'chikemo.info@gmail.com',
      resendFrom: 'チケモ運営事務局 <chikemo.info@chikemo.net>',
      shippingMethod: '日本郵便 レターパックライト',
    },
    GmailApp: {
      sendEmail(to, subject, body, options) {
        sent.push({ to, subject, body, options });
        if (gmailError) throw gmailError;
      },
    },
    LockService: {
      getScriptLock: () => ({
        tryLock() {
          callOrder.push('tryLock');
          return lockAcquirable;
        },
        releaseLock() {
          callOrder.push('releaseLock');
        },
      }),
    },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => 'test-resend-key' }),
    },
    ScriptApp: {
      getProjectTriggers: () => [],
      deleteTrigger() {},
      newTrigger() {
        return {
          forSpreadsheet() { return this; },
          onEdit() { return this; },
          create() {},
        };
      },
    },
    SpreadsheetApp: {
      getActive: () => ({ getSheetByName: () => sheet }),
      flush() {
        callOrder.push('flush');
      },
    },
    UrlFetchApp: {
      fetch(url, options) {
        resendRequests.push({ url, options });
        return {
          getContentText: () => 'response body',
          getResponseCode: () => resendStatus,
        };
      },
    },
    Utilities: {
      formatDate: () => '2026/07/23 12:34:56',
      computeDigest: (algorithm, text) => Array.from(crypto.createHash('sha256').update(text, 'utf8').digest()),
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
    },
  };

  vm.createContext(context);
  vm.runInContext(fs.readFileSync(mainSourcePath, 'utf8'), context, { filename: mainSourcePath });
  return { callOrder, context, resendRequests, sent, sheet, writes };
}

function baseShippingCells(overrides = {}) {
  return {
    [`3:${COL['メールアドレス']}`]: 'customer@example.com',
    [`3:${COL['追跡番号']}`]: 'ABCD123456JP',
    [`3:${COL['お名前（スペースなし）']}`]: '山田太郎',
    [`3:${COL['ご購入商品']}`]: 'チケット',
    [`3:${COL['購入枚数']}`]: '2',
    [`3:${COL['商品お届け先名']}`]: '山田太郎',
    [`3:${COL['商品お届け先住所']}`]: '東京都千代田区1-1',
    [`3:${COL['入金']}`]: 'OK',
    ...overrides,
  };
}

test('入金をOKにすると発送通知メールを1通送る', () => {
  const cells = baseShippingCells();
  const { context, sent, sheet } = loadScript({ cells });

  context.handleEdit({ range: createRange(sheet, 3, COL['入金'], 'OK') });

  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, 'customer@example.com');
});

test('送信はロック内で確認・送信・記録が行われ、releaseLockの前にflushが呼ばれる', () => {
  const cells = baseShippingCells();
  const { callOrder, context, sheet, writes } = loadScript({ cells });

  context.handleEdit({ range: createRange(sheet, 3, COL['入金'], 'OK') });

  assert.ok(writes.some((w) => w.marker === 'write:sendResult'));
  const tryLockIndex = callOrder.indexOf('tryLock');
  const lastFlushIndex = callOrder.lastIndexOf('flush');
  const releaseLockIndex = callOrder.indexOf('releaseLock');
  assert.ok(tryLockIndex === 0);
  assert.ok(lastFlushIndex < releaseLockIndex);
  assert.ok(callOrder.includes('flush'));
});

test('ロックが取れないとGmailもResendも呼ばれず、発送通知済み欄は触らずエラー欄だけに書く', () => {
  const cells = baseShippingCells();
  const { context, resendRequests, sent, sheet, writes } = loadScript({ cells, lockAcquirable: false });

  context.handleEdit({ range: createRange(sheet, 3, COL['入金'], 'OK') });

  assert.equal(sent.length, 0);
  assert.equal(resendRequests.length, 0);
  // 実行中の処理が書いた「送信済み」を消すと次の編集で再送するため、済み欄と日時欄には書かない。
  assert.equal(writes.some((w) => w.column === COL['発送通知済み'] || w.column === COL['発送通知日時']), false);
  const errorWrite = writes.find((w) => w.column === COL['発送通知エラー']);
  assert.equal(
    errorWrite.value,
    '別の送信処理が実行中のため未処理です。少し待ってから入金欄を入れ直してください',
  );
});

test('1回目の実行で送信済みになった行を2回目の実行が送らない', () => {
  const cells = baseShippingCells();
  const { context, sent, sheet } = loadScript({ cells });
  const event = { range: createRange(sheet, 3, COL['入金'], 'OK') };

  context.handleEdit(event);
  context.handleEdit(event);

  assert.equal(sent.length, 1);
});

test('Resend送信のIdempotency-Keyは発送通知とキャンセル通知で接頭辞が異なる', () => {
  const gmailError = new Error('Service invoked too many times: email quota');

  const shippingCells = baseShippingCells();
  const shipping = loadScript({ cells: shippingCells, gmailError });
  shipping.context.handleEdit({ range: createRange(shipping.sheet, 3, COL['入金'], 'OK') });

  const cancelCells = baseShippingCells({ [`3:${COL['入金']}`]: 'NG' });
  const cancel = loadScript({ cells: cancelCells, gmailError });
  cancel.context.handleEdit({ range: createRange(cancel.sheet, 3, COL['入金'], 'NG') });

  const shippingKey = shipping.resendRequests[0].options.headers['Idempotency-Key'];
  const cancelKey = cancel.resendRequests[0].options.headers['Idempotency-Key'];

  assert.match(shippingKey, /^chikemo\/shipping\/[0-9a-f]+$/);
  assert.match(cancelKey, /^chikemo\/cancel\/[0-9a-f]+$/);
});

test('Resend送信は同じ内容なら同じIdempotency-Keyになり、本文が1文字違えば変わる', () => {
  const gmailError = new Error('Service invoked too many times: email quota');

  const cellsA = baseShippingCells();
  const runA = loadScript({ cells: cellsA, gmailError });
  runA.context.handleEdit({ range: createRange(runA.sheet, 3, COL['入金'], 'OK') });

  const cellsB = baseShippingCells();
  const runB = loadScript({ cells: cellsB, gmailError });
  runB.context.handleEdit({ range: createRange(runB.sheet, 3, COL['入金'], 'OK') });

  const cellsC = baseShippingCells({ [`3:${COL['追跡番号']}`]: 'ABCD123456JQ' });
  const runC = loadScript({ cells: cellsC, gmailError });
  runC.context.handleEdit({ range: createRange(runC.sheet, 3, COL['入金'], 'OK') });

  const keyA = runA.resendRequests[0].options.headers['Idempotency-Key'];
  const keyB = runB.resendRequests[0].options.headers['Idempotency-Key'];
  const keyC = runC.resendRequests[0].options.headers['Idempotency-Key'];

  assert.equal(keyA, keyB);
  assert.notEqual(keyA, keyC);
});

test('Resendが409を返したら二重送信防止のエラーメッセージを記録する', () => {
  const cells = baseShippingCells();
  const { context, sheet, writes } = loadScript({
    cells,
    gmailError: new Error('Service invoked too many times: email quota'),
    resendStatus: 409,
  });

  context.handleEdit({ range: createRange(sheet, 3, COL['入金'], 'OK') });

  const errorWrite = writes.find((w) => w.column === COL['発送通知エラー']);
  assert.match(errorWrite.value, /同じ内容のメールを別の処理が送信中です。二重送信を防ぐため送っていません/);
});

test('reprocessUnsentはループ全体を1回のロックで囲み、未送信行だけ処理する', () => {
  const cells = {
    ...baseShippingCells({ [`3:${COL['入金']}`]: 'OK' }),
    [`4:${COL['入金']}`]: 'OK',
    [`4:${COL['メールアドレス']}`]: 'other@example.com',
    [`4:${COL['追跡番号']}`]: 'ZZZZ999999JP',
    [`4:${COL['お名前（スペースなし）']}`]: '佐藤花子',
    [`4:${COL['ご購入商品']}`]: 'チケット',
    [`4:${COL['購入枚数']}`]: '1',
    [`4:${COL['商品お届け先名']}`]: '佐藤花子',
    [`4:${COL['商品お届け先住所']}`]: '大阪府大阪市1-1',
    [`4:${COL['発送通知済み']}`]: '送信済み',
  };
  const { callOrder, context, sent } = loadScript({ cells, lastRow: 4 });

  context.reprocessUnsent();

  assert.equal(sent.length, 1);
  assert.equal(sent[0].to, 'customer@example.com');
  assert.equal(callOrder.filter((c) => c === 'tryLock').length, 1);
});

test('reprocessUnsentはロックが取れなければ何も送らない', () => {
  const cells = baseShippingCells();
  const { context, sent } = loadScript({ cells, lastRow: 3, lockAcquirable: false });

  context.reprocessUnsent();

  assert.equal(sent.length, 0);
});

test('メイン.gsとChikemo購入フォーム.gsは1つのコンテキストで評価しても関数名が衝突しない', () => {
  const mainSource = fs.readFileSync(mainSourcePath, 'utf8');
  const purchaseFormSource = fs.readFileSync(purchaseFormSourcePath, 'utf8');

  const extractTopLevelNames = (source) => {
    const names = [];
    const funcRe = /^function\s+([A-Za-z0-9_]+)\s*\(/gm;
    const varRe = /^var\s+([A-Za-z0-9_]+)\s*=/gm;
    let m;
    while ((m = funcRe.exec(source))) names.push(m[1]);
    while ((m = varRe.exec(source))) names.push(m[1]);
    return names;
  };

  const mainNames = extractTopLevelNames(mainSource);
  const purchaseFormNames = extractTopLevelNames(purchaseFormSource);
  const overlap = mainNames.filter((name) => purchaseFormNames.includes(name));
  assert.deepEqual(overlap, []);

  const combinedContext = {
    console,
    Logger: { log() {} },
    GmailApp: { sendEmail() {} },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'test-resend-key' }) },
    ScriptApp: {
      getProjectTriggers: () => [],
      deleteTrigger() {},
      newTrigger() {
        return { forSpreadsheet() { return this; }, onEdit() { return this; }, create() {} };
      },
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'chikemo.info@chikemo.net' }) },
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: () => null }),
      getActive: () => ({ getSheetByName: () => null }),
      flush() {},
    },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getContentText: () => '' }) },
    Utilities: {
      formatDate: () => '',
      computeDigest: (algorithm, text) => Array.from(crypto.createHash('sha256').update(text, 'utf8').digest()),
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
    },
  };

  vm.createContext(combinedContext);
  vm.runInContext(purchaseFormSource + '\n' + mainSource, combinedContext, { filename: 'combined' });

  assert.equal(typeof combinedContext.handleChikemoPurchaseFormEdit, 'function');
  assert.equal(typeof combinedContext.handleEdit, 'function');
  assert.equal(typeof combinedContext.withChikemoPurchaseFormSendLock_, 'function');
  assert.equal(typeof combinedContext.withChikemoMainSendLock_, 'function');
});
