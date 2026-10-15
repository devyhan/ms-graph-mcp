import assert from 'node:assert/strict';
import test from 'node:test';

import { extractDocx, splitSections } from './extract-docx.js';

test('headings become titled sections, lists stay line-shaped', () => {
  const sections = splitSections(
    '<h1>Quarterly review</h1><p>Revenue is up.</p><ul><li>EMEA leads</li><li>APAC follows</li></ul>' +
      '<h2>Risks</h2><p>Supply chain.</p>',
  );
  assert.equal(sections.length, 2);
  assert.equal(sections[0]?.title, 'Quarterly review');
  assert.equal(sections[0]?.text, 'Revenue is up.\nEMEA leads\nAPAC follows');
  assert.equal(sections[1]?.title, 'Risks');
  assert.equal(sections[1]?.text, 'Supply chain.');
});

test('a heading-less document chunks instead of becoming one unit', () => {
  const html = `<p>${'x'.repeat(5900)}</p><p>${'y'.repeat(5900)}</p>`;
  const sections = splitSections(html);
  assert.ok(sections.length >= 2, 'long prose must page across units');
  assert.ok(sections.every((section) => section.text.length <= 6000 + 5900));
});

test('extractDocx projects sections to units and passes warnings through', async () => {
  const result = await extractDocx(new Uint8Array([1, 2]), 'memo.docx', {
    convertHtml: async () => ({
      html: '<h1>Memo</h1><p>Hello.</p>',
      warnings: ['Unrecognised element kept as-is'],
    }),
  });
  assert.deepEqual(result.units, [{ id: 'section-1', title: 'Memo', text: 'Hello.' }]);
  assert.deepEqual(result.warnings, ['Unrecognised element kept as-is']);
});

test('an image-only document is a friendly error, not an empty unit', async () => {
  await assert.rejects(
    () =>
      extractDocx(new Uint8Array([1]), 'scan.docx', {
        convertHtml: async () => ({ html: '<p></p>', warnings: [] }),
      }),
    /only images/,
  );
});

test('a password failure maps to the protection refusal', async () => {
  await assert.rejects(
    () =>
      extractDocx(new Uint8Array([1]), 'locked.docx', {
        convertHtml: async () => {
          throw new Error('Could not decrypt the document (password protected)');
        },
      }),
    /password-protected or rights-managed/,
  );
});

test('a real minimal docx converts through mammoth', async () => {
  // A stored (uncompressed) zip hand-built so the test needs no fixture file:
  // [Content_Types].xml + word/document.xml is all mammoth requires.
  const { default: JSZip } = (await import('jszip')) as unknown as {
    default: typeof import('jszip');
  };
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  );
  zip.file(
    'word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Real heading</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Hello from a real docx.</w:t></w:r></w:p></w:body></w:document>',
  );
  const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
  const result = await extractDocx(bytes, 'real.docx');
  assert.equal(result.units.length, 1);
  assert.equal(result.units[0]?.title, 'Real heading');
  assert.match(result.units[0]?.text ?? '', /Hello from a real docx/);
});
