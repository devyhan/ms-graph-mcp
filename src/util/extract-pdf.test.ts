import assert from 'node:assert/strict';
import test from 'node:test';

import { MAX_PDF_PAGES } from './extract.js';
import { extractPdf, rowsOf, type PdfEngine } from './extract-pdf.js';

test('rows group by y and order by x', () => {
  const rows = rowsOf([
    { str: 'world', x: 60, y: 700 },
    { str: 'Hello', x: 10, y: 700 },
    { str: 'next line', x: 10, y: 680 },
  ]);
  assert.deepEqual(rows, ['Hello world', 'next line']);
});

test('hasEOL breaks rows even at the same y', () => {
  const rows = rowsOf([
    { str: 'a', x: 10, y: 700, hasEOL: true },
    { str: 'b', x: 10, y: 700 },
  ]);
  assert.deepEqual(rows, ['a', 'b']);
});

function stubEngine(pages: string[][]): PdfEngine {
  return {
    pageCount: async () => pages.length,
    pageItems: async (_bytes, pageNumber) =>
      (pages[pageNumber - 1] ?? []).map((str, i) => ({ str, x: 10, y: 700 - i * 20, hasEOL: true })),
  };
}

test('one unit per page with page titles', async () => {
  const result = await extractPdf(
    new Uint8Array([1]),
    'report.pdf',
    stubEngine([
      ['Annual report', 'Revenue grew.'],
      ['Risks'],
    ]),
  );
  assert.equal(result.units.length, 2);
  assert.deepEqual(result.units[0], { id: 'page-1', title: 'Page 1', text: 'Annual report\nRevenue grew.' });
  assert.deepEqual(result.warnings, []);
});

test('long documents stop at the page cap and say so', async () => {
  const pages = Array.from({ length: MAX_PDF_PAGES + 5 }, (_, i) => [`page ${i + 1}`]);
  const result = await extractPdf(new Uint8Array([1]), 'long.pdf', stubEngine(pages));
  assert.equal(result.units.length, MAX_PDF_PAGES);
  assert.ok(result.warnings.some((warning) => warning.includes(`first ${MAX_PDF_PAGES} of`)));
});

test('a textless PDF is the scanned-paper refusal, not empty units', async () => {
  await assert.rejects(
    () => extractPdf(new Uint8Array([1]), 'scan.pdf', stubEngine([[], []])),
    /scanned PDF/,
  );
});

test('a password exception maps to the protection refusal', async () => {
  const engine: PdfEngine = {
    pageCount: async () => {
      const error = new Error('Needs password') as Error & { name: string };
      error.name = 'PasswordException';
      throw error;
    },
    pageItems: async () => [],
  };
  await assert.rejects(() => extractPdf(new Uint8Array([1]), 'locked.pdf', engine), /password-protected/);
});
