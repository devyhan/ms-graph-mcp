import assert from 'node:assert/strict';
import test from 'node:test';

import { extractPptx } from './extract-pptx.js';

const SLIDE_2 =
  '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<p:cSld><p:spTree>' +
  '<p:sp><p:txBody><a:p><a:r><a:t>Second slide</a:t></a:r></a:p>' +
  '<a:p><a:r><a:t>with a table</a:t></a:r></a:p></p:txBody></p:sp>' +
  '<p:graphicFrame><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>cell one</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></p:graphicFrame>' +
  '</p:spTree></p:cSld></p:sld>';

const SLIDE_1 =
  '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First slide</a:t></a:r></a:p></p:txBody></p:sp>' +
  '<p:grpSp><p:sp><p:txBody><a:p><a:r><a:t>grouped shape</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>' +
  '</p:spTree></p:cSld></p:sld>';

function zipOf(parts: Record<string, string>): { openZip: (bytes: Uint8Array) => Promise<{ readText: (path: string) => Promise<string | null> }> } {
  return {
    openZip: async () => ({
      readText: async (path: string) => parts[path] ?? null,
    }),
  };
}

// Note the sldIdLst order: rId3 (slides/slide2.xml) FIRST, so zip order and
// presentation order disagree and the test proves which one wins.
const PARTS: Record<string, string> = {
  'ppt/presentation.xml':
    '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<p:sldIdLst><p:sldId id="257" r:id="rId3"/><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>',
  'ppt/_rels/presentation.xml.rels':
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId2" Target="slides/slide1.xml"/><Relationship Id="rId3" Target="slides/slide2.xml"/>' +
    '</Relationships>',
  'ppt/slides/slide1.xml': SLIDE_1,
  'ppt/slides/slide2.xml': SLIDE_2,
};

test('slides come out in presentation order with tables and grouped shapes', async () => {
  const result = await extractPptx(new Uint8Array([1]), 'deck.pptx', zipOf(PARTS));
  assert.equal(result.units.length, 2);
  assert.equal(result.units[0]?.id, '257');
  assert.equal(result.units[0]?.title, 'Second slide');
  assert.equal(result.units[0]?.text, 'Second slide\nwith a table\ncell one');
  assert.equal(result.units[1]?.text, 'First slide\ngrouped shape');
  assert.deepEqual(result.warnings, []);
});

test('speaker notes append with a marker', async () => {
  const parts: Record<string, string> = {
    ...PARTS,
    'ppt/slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Target="../notesSlides/notesSlide1.xml"/></Relationships>',
    'ppt/notesSlides/notesSlide1.xml':
      '<?xml version="1.0"?><p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>do not read this aloud</a:t></a:r></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld></p:notes>',
  };
  const result = await extractPptx(new Uint8Array([1]), 'deck.pptx', zipOf(parts));
  const first = result.units.find((unit) => unit.id === '256');
  assert.match(first?.text ?? '', /Speaker notes: do not read this aloud/);
});

test('a chart reference warns and names the rel target', async () => {
  const parts: Record<string, string> = {
    ...PARTS,
    'ppt/slides/slide1.xml': SLIDE_1.replace(
      '</p:spTree>',
      '<p:graphicFrame><a:graphic xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></a:graphic></p:graphicFrame></p:spTree>',
    ),
    'ppt/slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Target="../charts/chart1.xml"/></Relationships>',
  };
  const result = await extractPptx(new Uint8Array([1]), 'deck.pptx', zipOf(parts));
  assert.equal(result.warnings.length, 1);
  assert.match(result.warnings[0] ?? '', /chart1\.xml/);
  assert.match(result.warnings[0] ?? '', /not extracted/);
});

test('a package without a presentation is a friendly error', async () => {
  await assert.rejects(
    () => extractPptx(new Uint8Array([1]), 'empty.pptx', zipOf({})),
    /no ppt\/presentation\.xml/,
  );
});

test('a real minimal pptx parses end to end through jszip', async () => {
  const { default: JSZip } = (await import('jszip')) as unknown as {
    default: typeof import('jszip');
  };
  const zip = new JSZip();
  for (const [path, text] of Object.entries(PARTS)) zip.file(path, text);
  const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
  const result = await extractPptx(bytes, 'real.pptx');
  assert.equal(result.units.length, 2);
  assert.equal(result.units[0]?.text, 'Second slide\nwith a table\ncell one');
});
