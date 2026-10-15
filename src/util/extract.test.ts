import assert from 'node:assert/strict';
import test from 'node:test';

import { detectKind } from './extract.js';

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
