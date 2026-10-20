import assert from 'node:assert/strict';
import test from 'node:test';

import {
  AIP_COVER_PHRASES,
  ProtectedDocumentError,
  detectKind,
  isAipCoverText,
  isAipWrapperPdf,
  protectionRefusal,
  toMarkdownTable,
} from './extract.js';

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

/** Minimal fake PDF bytes: magic plus whatever markers the case needs. */
function pdfWith(...chunks: string[]): Uint8Array {
  return new Uint8Array(Buffer.from(['%PDF-1.7', ...chunks].join('\n'), 'latin1'));
}

void test('a Purview wrapper PDF is detected from its byte markers', () => {
  const bytes = pdfWith('1 0 obj', '<< /Filter /Aadrm /V 4 >>', 'trailer << /Encrypt 1 0 R >>');
  assert.equal(isAipWrapperPdf(bytes), true);
});

void test('wrapper marker matching is case-insensitive', () => {
  assert.equal(isAipWrapperPdf(pdfWith('/ENCRYPT 1 0 R', 'contains irm license')), true);
  assert.equal(isAipWrapperPdf(pdfWith('/encrypt 7 0 R', 'RMS-protected')), true);
  assert.equal(isAipWrapperPdf(pdfWith('/Encrypt 3 0 R', 'usage RIGHTS attached')), true);
});

void test('a password-style encrypted PDF without IRM markers is not a wrapper', () => {
  assert.equal(isAipWrapperPdf(pdfWith('trailer << /Encrypt 5 0 R /ID [<>] >>')), false);
});

void test('an unencrypted PDF is not a wrapper', () => {
  assert.equal(isAipWrapperPdf(pdfWith('plain content here')), false);
});

void test('non-PDF bytes with markers are not wrappers', () => {
  const bytes = new Uint8Array(Buffer.from('/Encrypt with rms markers but no magic', 'latin1'));
  assert.equal(isAipWrapperPdf(bytes), false);
});

void test('every cover phrase matches through whitespace and case mangling', () => {
  assert.ok(AIP_COVER_PHRASES.length >= 2, 'the detector needs EN and KO phrases to be worth anything');
  for (const phrase of AIP_COVER_PHRASES) {
    const mangled = `Cover notice:\n\t ${phrase.replace(/ /g, '  \n\t')}  \nFooter`.toUpperCase();
    assert.equal(isAipCoverText(mangled), true, `phrase missed: ${phrase}`);
  }
});

void test('ordinary text — including near-miss fragments — is not a cover', () => {
  assert.equal(isAipCoverText('Quarterly report\nRevenue grew 9%.'), false);
  assert.equal(isAipCoverText(''), false);
  // A lone fragment of a longer phrase must not trip the detector.
  assert.equal(isAipCoverText('this document is protected'), false);
});

void test('the default refusal keeps its message and gains structure', () => {
  const error = protectionRefusal('locked.pdf');
  assert.ok(error instanceof ProtectedDocumentError);
  assert.match(error.message, /password-protected or rights-managed/);
  assert.equal(error.code, 'PROTECTED_DOCUMENT');
  assert.equal(error.protection, 'unknown');
  assert.equal(error.fileName, 'locked.pdf');
});

void test('the AIP refusal names the decryptor option and the native-app fallback', () => {
  const error = protectionRefusal('secret.pdf', 'aip-rms');
  assert.ok(error instanceof ProtectedDocumentError);
  assert.equal(error.protection, 'aip-rms');
  assert.match(error.message, /Purview|AIP\/RMS/);
  assert.match(error.message, /decrypt/i);
  assert.match(error.message, /native app/);
});
