import assert from 'node:assert/strict';
import test from 'node:test';

import { collectRevisions, extractDocx, splitSections } from './extract-docx.js';

function zipOf(parts: Record<string, string>): {
  openZip: (bytes: Uint8Array) => Promise<{
    listParts: () => Promise<string[]>;
    readText: (path: string) => Promise<string | null>;
  }>;
} {
  return {
    openZip: async () => ({
      listParts: async () => Object.keys(parts),
      readText: async (path: string) => parts[path] ?? null,
    }),
  };
}

const REVISION_DOC =
  '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:body><w:p><w:r><w:t>Keep </w:t></w:r>' +
  '<w:ins w:author="Ada"><w:r><w:t>this addition</w:t></w:r></w:ins>' +
  '<w:del w:author="Ada"><w:r><w:t>that removal</w:t></w:r></w:del>' +
  '<w:moveFrom w:author="Ada"><w:r><w:t>moved away</w:t></w:r></w:moveFrom>' +
  '<w:moveTo w:author="Ada"><w:r><w:t>moved here</w:t></w:r></w:moveTo>' +
  '</w:p></w:body></w:document>';

async function realParser(): Promise<(xml: string) => Record<string, unknown>> {
  const { XMLParser } = (await import('fast-xml-parser')) as unknown as {
    XMLParser: new (options: Record<string, unknown>) => { parse: (xml: string) => Record<string, unknown> };
  };
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', trimValues: false });
  return (xml: string) => parser.parse(xml);
}

test('headings become titled sections, lists stay line-shaped', () => {
  const sections = splitSections(
    '<h1>Quarterly review</h1><p>Revenue is up.</p><ul><li>EMEA leads</li><li>APAC follows</li></ul>' +
      '<h2>Risks</h2><p>Supply chain.</p>',
  );
  assert.equal(sections.length, 2);
  assert.equal(sections[0]?.title, 'Quarterly review');
  assert.equal(sections[0]?.text, 'Revenue is up.\nEMEA leads\nAPAC follows');
  assert.equal(sections[1]?.title, 'Quarterly review > Risks');
  assert.equal(sections[1]?.text, 'Supply chain.');
});

test('nested headings carry the full path, siblings reset it', () => {
  const sections = splitSections(
    '<h1>Part</h1><h2>Chapter</h2><p>Deep.</p><h2>Next</h2><p>Shallow.</p><h1>Alone</h1><p>Top.</p>',
  );
  assert.deepEqual(
    sections.map((section) => section.title),
    ['Part', 'Part > Chapter', 'Part > Next', 'Alone'],
  );
  assert.equal(sections[0]?.text, '');
});

test('word tables become md-tables, single-record blocks KV', () => {
  const sections = splitSections(
    '<h1>Data</h1><table><tr><th>name</th><th>rev</th></tr>' +
      '<tr><td>EMEA</td><td>9</td></tr><tr><td>APAC</td><td>7</td></tr></table>',
  );
  assert.equal(sections[0]?.title, 'Data');
  assert.equal(
    sections[0]?.text,
    '| name | rev |\n| --- | --- |\n| EMEA | 9 |\n| APAC | 7 |',
  );
  const kv = splitSections(
    '<h1>One</h1><table><tr><th>name</th><th>rev</th></tr><tr><td>EMEA</td><td>9</td></tr></table>',
  );
  assert.equal(
    kv[0]?.text,
    '| Field | Value |\n| --- | --- |\n| name | EMEA |\n| rev | 9 |',
  );
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

test('revisions collect in document order: ins/moveTo add, del/moveFrom cut', async () => {
  const parse = await realParser();
  const revisions = collectRevisions(parse(REVISION_DOC) as Parameters<typeof collectRevisions>[0]);
  assert.deepEqual(revisions, [
    { kind: 'ins', text: 'this addition' },
    { kind: 'del', text: 'that removal' },
    { kind: 'del', text: 'moved away' },
    { kind: 'ins', text: 'moved here' },
  ]);
});

test('insertions tag inline, deletions collect into a marked unit', async () => {
  const result = await extractDocx(
    new Uint8Array([1, 2]),
    'tracked.docx',
    { convertHtml: async () => ({ html: '<p>Keep this addition.</p>', warnings: [] }) },
    zipOf({ 'word/document.xml': REVISION_DOC }),
  );
  assert.equal(result.units[0]?.text, 'Keep [+this addition+].');
  const tracked = result.units.find((unit) => unit.id === 'tracked-changes');
  assert.equal(tracked?.title, 'Tracked changes');
  assert.equal(tracked?.text, '[-that removal-]\n[-moved away-]\n[+moved here+]');
});

test('headers and footers append as marked units in numeric order', async () => {
  const part = (text: string): string =>
    '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`;
  const result = await extractDocx(
    new Uint8Array([1, 2]),
    'letter.docx',
    { convertHtml: async () => ({ html: '<h1>Body</h1><p>Letter text.</p>', warnings: [] }) },
    zipOf({
      'word/document.xml':
        '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:p><w:r><w:t>Plain.</w:t></w:r></w:p></w:body></w:document>',
      'word/footer2.xml': part('Footer two'),
      'word/header1.xml': part('Header one'),
      'word/footer1.xml': part('Footer one'),
    }),
  );
  assert.deepEqual(
    result.units.map((unit) => [unit.id, unit.title]),
    [
      ['section-1', 'Body'],
      ['header-1', 'Header 1'],
      ['footer-1', 'Footer 1'],
      ['footer-2', 'Footer 2'],
    ],
  );
  assert.equal(result.units[1]?.text, 'Header one');
  assert.equal(result.units[3]?.text, 'Footer two');
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

test('comments append at the end through the comment-reference style map', async () => {
  // Installed mammoth 1.13.0 ignores comment references by default; the
  // engine passes `styleMap: ['comment-reference => sup']`, which appends
  // referenced comments in a trailing `<dl>` — pinned here.
  const { default: JSZip } = (await import('jszip')) as unknown as {
    default: typeof import('jszip');
  };
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      '</Types>',
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
      '</Relationships>',
  );
  zip.file(
    'word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:r><w:t>Reviewed text.</w:t></w:r>' +
      '<w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/>' +
      '<w:r><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>',
  );
  zip.file(
    'word/comments.xml',
    '<?xml version="1.0"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:comment w:id="0" w:author="Ada" w:initials="A"><w:p><w:r><w:t>Check this figure.</w:t></w:r></w:p></w:comment>' +
      '</w:comments>',
  );
  const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
  const result = await extractDocx(bytes, 'commented.docx');
  assert.equal(result.units.length, 1);
  const text = result.units[0]?.text ?? '';
  assert.match(text, /Reviewed text/);
  assert.match(text, /Check this figure/);
  assert.ok(text.indexOf('Reviewed text') < text.indexOf('Check this figure'));
});

test('headers, footers and revisions survive a real end-to-end pass', async () => {
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
      '<w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Keep </w:t></w:r>' +
      '<w:ins w:author="Ada"><w:r><w:t>this addition</w:t></w:r></w:ins>' +
      '<w:del w:author="Ada"><w:r><w:t>that removal</w:t></w:r></w:del></w:p></w:body></w:document>',
  );
  zip.file(
    'word/header1.xml',
    '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r><w:t>Draft header</w:t></w:r></w:p></w:hdr>',
  );
  zip.file(
    'word/footer1.xml',
    '<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r><w:t>Page footer</w:t></w:r></w:p></w:ftr>',
  );
  const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
  const result = await extractDocx(bytes, 'full.docx');
  assert.deepEqual(
    result.units.map((unit) => unit.id),
    ['section-1', 'tracked-changes', 'header-1', 'footer-1'],
  );
  assert.match(result.units[0]?.text ?? '', /\[\+this addition\+\]/);
  assert.match(result.units[1]?.text ?? '', /\[-that removal-\]/);
  assert.equal(result.units[2]?.text, 'Draft header');
  assert.equal(result.units[3]?.text, 'Page footer');
});
