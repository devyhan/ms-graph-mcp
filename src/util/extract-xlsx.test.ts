import assert from 'node:assert/strict';
import test from 'node:test';

import { MAX_XLSX_COLS, MAX_XLSX_ROWS } from './extract.js';
import { cellText, extractXlsx } from './extract-xlsx.js';

test('cells render: saved formula results win, rich text joins', () => {
  assert.equal(cellText('plain'), 'plain');
  assert.equal(cellText(42), '42');
  assert.equal(cellText(true), 'TRUE');
  assert.equal(cellText(null), '');
  assert.equal(cellText({ formula: 'A1+B1', result: 7 }), '7 [=A1+B1]');
  assert.equal(cellText({ formula: 'NOW()', result: undefined }), '[=NOW()]');
  assert.equal(cellText({ richText: [{ text: 'a' }, { text: 'b' }] }), 'ab');
  assert.equal(cellText({ text: 'link', hyperlink: 'https://example.com' }), 'link');
});

test('one unit per sheet with dimensions in the title', async () => {
  const result = await extractXlsx(new Uint8Array([1]), 'book.xlsx', {
    readSheets: async () => [
      { name: 'Q1', rows: [[, 'name', 'rev'], [, 'EMEA', 9]] },
      { name: 'Notes', rows: [[, 'hello']] },
    ],
  });
  assert.equal(result.units.length, 2);
  assert.equal(result.units[0]?.id, 'Q1');
  assert.equal(result.units[0]?.title, 'Q1 (2 rows × 2 cols)');
  assert.equal(result.units[0]?.text, 'name\trev\nEMEA\t9');
  assert.equal(result.units[1]?.text, 'hello');
  assert.deepEqual(result.warnings, []);
});

test('the row cap stops the read and says so', async () => {
  const rows: unknown[][] = Array.from({ length: MAX_XLSX_ROWS + 10 }, (_, i) => [, `r${i}`]);
  const result = await extractXlsx(new Uint8Array([1]), 'big.xlsx', {
    readSheets: async () => [{ name: 'Data', rows }],
  });
  assert.match(result.units[0]?.text ?? '', /r0/);
  assert.ok(!(result.units[0]?.text ?? '').includes(`r${MAX_XLSX_ROWS}`));
  assert.ok(result.warnings.some((warning) => warning.includes('100,000 rows')));
});

test('wide rows drop past the column cap with a warning', async () => {
  const rows: unknown[][] = [[, ...Array.from({ length: MAX_XLSX_COLS + 5 }, (_, i) => `c${i}`)]];
  const result = await extractXlsx(new Uint8Array([1]), 'wide.xlsx', {
    readSheets: async () => [{ name: 'W', rows }],
  });
  assert.ok(!(result.units[0]?.text ?? '').includes(`c${MAX_XLSX_COLS}`));
  assert.ok(result.warnings.some((warning) => warning.includes('Columns past')));
});

test('an encrypted workbook maps to the protection refusal', async () => {
  await assert.rejects(
    () =>
      extractXlsx(new Uint8Array([1]), 'locked.xlsx', {
        readSheets: async () => {
          throw new Error('File is encrypted');
        },
      }),
    /password-protected or rights-managed/,
  );
});
