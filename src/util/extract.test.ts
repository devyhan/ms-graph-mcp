import assert from 'node:assert/strict';
import test from 'node:test';

import { detectKind, toMarkdownTable } from './extract.js';

const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];

/** OLE directory entries store stream names as UTF-16LE. */
function oleWithStreams(...names: string[]): Uint8Array {
  const parts: number[] = [...OLE_MAGIC];
  for (const name of names) {
    for (const ch of Buffer.from(name, 'utf16le')) parts.push(ch);
  }
  return new Uint8Array(parts);
}

void test('an encrypted OOXML package is refused as protected, not legacy', () => {
  const bytes = oleWithStreams('EncryptedPackage', 'EncryptionInfo');
  assert.throws(
    () => detectKind('deck.pptx', undefined, bytes),
    /password-protected or rights-managed/,
  );
});

void test('a legacy OLE document without encryption streams keeps the legacy error', () => {
  const bytes = oleWithStreams('PowerPoint Document');
  assert.throws(() => detectKind('deck.ppt', undefined, bytes), /legacy OLE document/);
});

void test('zip and pdf detection is unchanged', () => {
  assert.equal(detectKind('deck.pptx', undefined, new Uint8Array([...ZIP_MAGIC, 0])), 'pptx');
  assert.equal(
    detectKind('doc.pdf', undefined, new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d])),
    'pdf',
  );
});

void test('tables render as Markdown with padded short rows', () => {
  assert.equal(
    toMarkdownTable(['name', 'rev'], [['EMEA', '9'], ['APAC']]),
    '| name | rev |\n| --- | --- |\n| EMEA | 9 |\n| APAC |  |',
  );
});

void test('a single-record block transposes to Field/Value form', () => {
  assert.equal(
    toMarkdownTable(['name', 'rev'], [['EMEA', '9']]),
    '| Field | Value |\n| --- | --- |\n| name | EMEA |\n| rev | 9 |',
  );
});

void test('pipes and newlines in cells cannot break the table', () => {
  assert.equal(
    toMarkdownTable(['a|b', 'c'], [['x\ny', 'z']]),
    '| Field | Value |\n| --- | --- |\n| a\\|b | x y |\n| c | z |',
  );
});

void test('an empty table renders nothing', () => {
  assert.equal(toMarkdownTable([], []), '');
});
