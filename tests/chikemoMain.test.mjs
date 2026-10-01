import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import vm from 'node:vm';

const source = fs.readFileSync(path.resolve('メイン.gs'), 'utf8');

const COL = { email: 11, name: 5, tracking: 20, payment: 19, shipStatus: 30, shipError: 32, cancelStatus: 33, cancelError: 35 };

function load({ lastRow = 10 } = {}) {
  const cells = new Map();
  const sent = [];
  const props = new Map();
  const triggers = [{ fn: 'handleEdit' }, { fn: 'other' }];
  const deleted = [];
  const headerRow = [];
  headerRow[COL.email - 1] = 'メールアドレス';
  headerRow[COL.name - 1] = 'システム表示名';
  headerRow[COL.tracking - 1] = '追跡番号';

  const sheet = {
    getName: () => 'シート1',
    getLastColumn: () => 40,
    getLastRow: () => lastRow,
    getRange(row, col, numRows = 1, numCols = 1) {
      return {
        getValues() {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const line = [];
            for (let c = 0; c < numCols; c++) {
              line.push(row === 1 ? (headerRow[col - 1 + c] ?? '') : (cells.get(`${row + r}:${col + c}`) ?? ''));
            }
            out.push(line);
          }
          return out;
        },
        getValue() { return this.getValues()[0][0]; },
        setValue(v) { cells.set(`${row}:${col}`, v); },
      };
    },
  };

  const context = {
    console: { log() {}, warn() {}, error() {} },
    Logger: { log() {} },
    SpreadsheetApp: { getActive: () => ({ getSheetByName: (n) => (n === 'シート1' ? sheet : null) }) },
    LockService: { getDocumentLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
    GmailApp: { sendEmail: (to, subject) => sent.push({ to, subject }) },
    Utilities: { formatDate: () => '2026/09/30 00:00:00' },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (props.has(k) ? props.get(k) : null),
        getProperties: () => Object.fromEntries(props),
        setProperties: (o) => Object.entries(o).forEach(([k, v]) => props.set(k, v)),
        deleteProperty: (k) => props.delete(k),
      }),
    },
    ScriptApp: {
      getProjectTriggers: () => triggers.map((t) => ({ ...t, getHandlerFunction: () => t.fn })),
      deleteTrigger: (t) => deleted.push(t.fn),
    },
  };
  vm.createContext(context);
  vm.runInContext(source, context);

  const setRow = (row, { payment = '', tracking = 'T-1' } = {}) => {
    cells.set(`${row}:${COL.email}`, `u${row}@example.com`);
    cells.set(`${row}:${COL.name}`, `name${row}`);
    cells.set(`${row}:${COL.tracking}`, tracking);
    cells.set(`${row}:${COL.payment}`, payment);
  };
  const range = (row, col, numRows = 1, numCols = 1) => ({
    getRow: () => row,
    getColumn: () => col,
    getLastColumn: () => col + numCols - 1,
    getNumRows: () => numRows,
    getSheet: () => sheet,
  });
  return { context, cells, sent, deleted, setRow, range };
}

test('複数列にまたがる貼り付けでも入金列を処理する', () => {
  const t = load();
  t.setRow(3, { payment: 'OK' });
  t.context.handleEdit({ range: t.range(3, 18, 1, 4) });
  assert.equal(t.sent.length, 1);
  assert.equal(t.cells.get(`3:${COL.shipStatus}`), '送信済み');
});

test('大文字小文字・全角の OK/NG も処理する', () => {
  const t = load();
  t.setRow(3, { payment: 'ok' });
  t.setRow(4, { payment: 'ＮＧ' });
  t.context.handleEdit({ range: t.range(3, COL.payment) });
  t.context.handleEdit({ range: t.range(4, COL.payment) });
  assert.equal(t.cells.get(`3:${COL.shipStatus}`), '送信済み');
  assert.equal(t.cells.get(`4:${COL.cancelStatus}`), '送信済み');
});

test('複数行貼り付けは先頭行だけ処理し、残りに警告を書く', () => {
  const t = load();
  [3, 4, 5].forEach((r) => t.setRow(r, { payment: 'OK' }));
  t.context.handleEdit({ range: t.range(3, COL.payment, 3) });
  assert.equal(t.sent.length, 1);
  assert.match(t.cells.get(`4:${COL.shipError}`), /未処理/);
  assert.match(t.cells.get(`5:${COL.shipError}`), /未処理/);
  assert.equal(t.cells.get(`4:${COL.shipStatus}`), undefined);
});

test('removeTrigger は handleEdit のトリガーだけを削除する', () => {
  const t = load();
  t.context.removeTrigger();
  assert.deepEqual(t.deleted, ['handleEdit']);
});

test('手動送信済みの行は再送しない', () => {
  const t = load();
  t.setRow(3, { payment: 'OK' });
  t.setRow(4, { payment: 'NG' });
  t.cells.set(`3:${COL.shipStatus}`, '手動送信済み');
  t.cells.set(`4:${COL.cancelStatus}`, '手動送信済み');
  t.context.handleEdit({ range: t.range(3, COL.payment) });
  t.context.handleEdit({ range: t.range(4, COL.payment) });
  assert.equal(t.sent.length, 0);
});

test('reprocessUnsent は手動送信済みを除き、OK/NGの未送信行だけ送る', () => {
  const t = load();
  t.setRow(2, { payment: 'OK' });
  t.setRow(3, { payment: 'NG' });
  t.setRow(4, { payment: 'NG' });
  t.cells.set(`3:${COL.cancelStatus}`, '手動送信済み');
  t.context.reprocessUnsent();
  assert.deepEqual(t.sent.map((m) => m.to), ['u2@example.com', 'u4@example.com']);
});
