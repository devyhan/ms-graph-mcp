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
  assert.equal(
    result.units[0]?.text,
    '| Field | Value |\n| --- | --- |\n| name | EMEA |\n| rev | 9 |',
  );
  assert.equal(result.units[1]?.text, '| hello |\n| --- |');
  assert.deepEqual(result.warnings, []);
});

test('multi-row sheets render as md-tables with padded ragged rows', async () => {
  const result = await extractXlsx(new Uint8Array([1]), 'book.xlsx', {
    readSheets: async () => [
      { name: 'Q1', rows: [[, 'name', 'rev'], [, 'EMEA', 9], [, 'APAC']] },
    ],
  });
  assert.equal(
    result.units[0]?.text,
    '| name | rev |\n| --- | --- |\n| EMEA | 9 |\n| APAC |  |',
  );
});

test('hidden sheets skip with a warning; hidden rows count and skip', async () => {
  const result = await extractXlsx(new Uint8Array([1]), 'book.xlsx', {
    readSheets: async () => [
      { name: 'Secret', rows: [], hidden: true },
      { name: 'Open', rows: [[, 'a'], [, 'b']], skippedHiddenRows: 2 },
    ],
  });
  assert.equal(result.units.length, 1);
  assert.equal(result.units[0]?.id, 'Open');
  assert.ok(result.warnings.some((warning) => warning.includes('"Secret" is hidden')));
  assert.ok(result.warnings.some((warning) => warning.includes('2 hidden rows')));
});

test('an all-hidden workbook is a friendly error, not an empty unit', async () => {
  await assert.rejects(
    () =>
      extractXlsx(new Uint8Array([1]), 'hidden.xlsx', {
        readSheets: async () => [{ name: 'Secret', rows: [], hidden: true }],
      }),
    /only hidden worksheets/,
  );
});

test('merged ranges fill down from the top-left value', async () => {
  const result = await extractXlsx(new Uint8Array([1]), 'merged.xlsx', {
    readSheets: async () => [
      {
        name: 'M',
        rows: [[, 'Region', 'Q1'], [, 'EMEA'], [, undefined, 5]],
        merges: ['A2:A3'],
      },
    ],
  });
  assert.equal(
    result.units[0]?.text,
    '| Region | Q1 |\n| --- | --- |\n| EMEA |  |\n| EMEA | 5 |',
  );
});

test('cells honour number formats: percent, currency, grouped', () => {
  assert.equal(cellText(0.125, '0%'), '13%');
  assert.equal(cellText(0.125, '0.00%'), '12.50%');
  assert.equal(cellText(1234.5, '"$"#,##0.00'), '$1,234.50');
  assert.equal(cellText(1234567, '#,##0'), '1,234,567');
  assert.equal(cellText(1.5, '0'), '2');
  assert.equal(cellText(7, 'General'), '7');
  assert.equal(cellText({ formula: 'A1*2', result: 0.5 }, '0%'), '50% [=A1*2]');
});

test('dates render as seen, with the 1904 system honoured', () => {
  // Serial 45292 is 2024-01-01 in the 1900 system.
  assert.equal(cellText(45292, 'yyyy-mm-dd'), '2024-01-01');
  assert.equal(cellText(45292, 'mm/dd/yyyy'), '01/01/2024');
  assert.equal(cellText(45292, 'd mmm yyyy'), '1 Jan 2024');
  assert.equal(cellText(new Date('2024-03-05T00:00:00.000Z'), 'yyyy-mm-dd'), '2024-03-05');
  // The same serial under 1904 is four years and a day later.
  assert.equal(cellText(45292, 'yyyy-mm-dd', true), '2028-01-02');
  // Bare dates keep their legacy ISO shape.
  assert.equal(cellText(new Date('2024-03-05T00:00:00.000Z')), '2024-03-05T00:00:00.000Z');
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

test('a real workbook reads end to end: strings, formats, hidden, merges', async () => {
  const mod = (await import('exceljs')) as unknown as {
    default?: { Workbook: new () => RealWorkbook };
    Workbook: new () => RealWorkbook;
  };
  interface RealWorksheet {
    addRow: (values: unknown[]) => { hidden: boolean };
    getCell: (address: string) => { numFmt: string };
    getRow: (num: number) => { hidden: boolean };
    mergeCells: (ref: string) => void;
    state: string;
  }
  interface RealWorkbook {
    addWorksheet: (name: string) => RealWorksheet;
    xlsx: { writeBuffer: () => Promise<Uint8Array> };
  }
  const Workbook = mod.default?.Workbook ?? mod.Workbook;
  const wb = new Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow(['name', 'share']);
  ws.addRow(['EMEA', 0.125]);
  ws.addRow(['HID', 0]);
  ws.addRow(['APAC', 0.5]);
  ws.getCell('B2').numFmt = '0%';
  ws.getRow(3).hidden = true;
  ws.getCell('B4').numFmt = '0%';
  ws.mergeCells('B4:C4');
  const secret = wb.addWorksheet('Secret');
  secret.addRow(['classified']);
  secret.state = 'hidden';

  const result = await extractXlsx(new Uint8Array(await wb.xlsx.writeBuffer()), 'real.xlsx');
  assert.equal(result.units.length, 1);
  assert.equal(result.units[0]?.id, 'Data');
  assert.equal(
    result.units[0]?.text,
    '| name | share |  |\n| --- | --- | --- |\n| EMEA | 13% |  |\n| APAC | 50% | 50% |',
  );
  assert.ok(result.warnings.some((warning) => warning.includes('"Secret" is hidden')));
  assert.ok(result.warnings.some((warning) => warning.includes('1 hidden row')));
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
