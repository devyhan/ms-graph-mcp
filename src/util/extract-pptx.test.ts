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

function zipOf(
  parts: Record<string, string>,
  bytes: Record<string, Uint8Array> = {},
): {
  openZip: (input: Uint8Array) => Promise<{
    readText: (path: string) => Promise<string | null>;
    readBytes: (path: string) => Promise<Uint8Array | null>;
  }>;
} {
  return {
    openZip: async () => ({
      readText: async (path: string) => parts[path] ?? null,
      readBytes: async (path: string) => bytes[path] ?? null,
    }),
  };
}

function shape(text: string): string {
  return `<p:sp><p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

function presentationWith(slides: string[]): Record<string, string> {
  const ids = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('');
  const rels = slides.map((_, i) => `<Relationship Id="rId${i + 2}" Target="slides/slide${i + 1}.xml"/>`).join('');
  const parts: Record<string, string> = {
    'ppt/presentation.xml':
      '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `${rels}</Relationships>`,
  };
  slides.forEach((body, i) => {
    parts[`ppt/slides/slide${i + 1}.xml`] =
      '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<p:cSld><p:spTree>${body}</p:spTree></p:cSld></p:sld>`;
  });
  return parts;
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

test('cached chart values resolve in place and clear the warning', async () => {
  const frame =
    '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9"/>' +
    '</a:graphicData></a:graphic></p:graphicFrame>';
  const parts: Record<string, string> = {
    ...presentationWith([`${shape('Before')}${frame}${shape('After')}`]),
    'ppt/slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Target="../charts/chart1.xml"/></Relationships>',
    'ppt/charts/chart1.xml':
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:numCache>' +
      '<c:pt idx="0"><c:v>42</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt>' +
      '</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
  };
  const result = await extractPptx(new Uint8Array([1]), 'chart.pptx', zipOf(parts));
  // fast-xml-parser groups same-named shapes, so the frame's values trail the
  // shape paragraphs — the same grouping the old paragraph walk always had.
  assert.equal(result.units[0]?.text, 'Before\nAfter\nChart values: 42, 7');
  assert.deepEqual(result.warnings, []);
});

test('SmartArt text resolves and clears the warning', async () => {
  const frame =
    '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
    '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="rId9" r:lo="rId10"/>' +
    '</a:graphicData></a:graphic></p:graphicFrame>';
  const parts: Record<string, string> = {
    ...presentationWith([`${shape('Process')}${frame}`]),
    'ppt/slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Target="../diagrams/data1.xml"/>' +
      '<Relationship Id="rId10" Target="../diagrams/layout1.xml"/></Relationships>',
    'ppt/diagrams/data1.xml':
      '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
      '<dgm:pt><dgm:t>Step one</dgm:t></dgm:pt><dgm:pt><dgm:t>Step two</dgm:t></dgm:pt></dgm:dataModel>',
  };
  const result = await extractPptx(new Uint8Array([1]), 'smart.pptx', zipOf(parts));
  assert.equal(result.units[0]?.text, 'Process\nSmartArt: Step one | Step two');
  assert.deepEqual(result.warnings, []);
});

test('an uncached chart keeps the warning naming the rel target', async () => {
  const frame =
    '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9"/>' +
    '</a:graphicData></a:graphic></p:graphicFrame>';
  const parts: Record<string, string> = {
    ...presentationWith([`${shape('Title')}${frame}`]),
    'ppt/slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Target="../charts/chart9.xml"/></Relationships>',
    'ppt/charts/chart9.xml':
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>',
  };
  const result = await extractPptx(new Uint8Array([1]), 'empty-chart.pptx', zipOf(parts));
  assert.equal(result.warnings.length, 1);
  assert.match(result.warnings[0] ?? '', /chart9\.xml/);
  assert.match(result.warnings[0] ?? '', /not extracted/);
});

test('an embedded workbook behind a chart extracts as capped text', async () => {
  const mod = (await import('exceljs')) as unknown as {
    default?: { Workbook: new () => { addWorksheet: (name: string) => { addRow: (row: unknown[]) => void }; xlsx: { writeBuffer: () => Promise<Uint8Array> } } };
    Workbook: new () => { addWorksheet: (name: string) => { addRow: (row: unknown[]) => void }; xlsx: { writeBuffer: () => Promise<Uint8Array> } };
  };
  const Workbook = mod.default?.Workbook ?? mod.Workbook;
  const wb = new Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(['q', 'rev']);
  ws.addRow(['Q1', 9]);
  const book = new Uint8Array(await wb.xlsx.writeBuffer());
  const frame =
    '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9"/>' +
    '</a:graphicData></a:graphic></p:graphicFrame>';
  const parts: Record<string, string> = {
    ...presentationWith([`${shape('Title')}${frame}`]),
    'ppt/slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Target="../charts/chart1.xml"/></Relationships>',
    'ppt/charts/chart1.xml':
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>',
    'ppt/charts/_rels/chart1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Target="../embeddings/data.xlsx"/></Relationships>',
  };
  const result = await extractPptx(new Uint8Array([1]), 'embedded.pptx', zipOf(parts, { 'ppt/embeddings/data.xlsx': book }));
  assert.match(result.units[0]?.text ?? '', /Embedded workbook:/);
  assert.match(result.units[0]?.text ?? '', /Q1/);
});

test('footer boilerplate stays on the first slide only', async () => {
  const parts = presentationWith([
    `${shape('Alpha')}${shape('Confidential')}`,
    `${shape('Beta')}${shape('Confidential')}`,
    `${shape('Gamma')}${shape('Confidential')}`,
  ]);
  const result = await extractPptx(new Uint8Array([1]), 'deck.pptx', zipOf(parts));
  assert.equal(result.units[0]?.text, 'Alpha\nConfidential');
  assert.equal(result.units[1]?.text, 'Beta');
  assert.equal(result.units[2]?.text, 'Gamma');
  // Titles are fixed before dedupe, so the outline is untouched.
  assert.deepEqual(
    result.units.map((unit) => unit.title),
    ['Alpha', 'Beta', 'Gamma'],
  );
});

test('note placeholders skip; the speaker body survives', async () => {
  const noteShape = (text: string, ph: string): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="2"/><p:cNvSpPr/><p:nvPr><p:ph${ph}/></p:nvPr></p:nvSpPr>` +
    `<p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  const parts: Record<string, string> = {
    ...presentationWith([shape('Title')]),
    'ppt/slides/_rels/slide1.xml.rels':
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Target="../notesSlides/notesSlide1.xml"/></Relationships>',
    'ppt/notesSlides/notesSlide1.xml':
      '<?xml version="1.0"?><p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<p:cSld><p:spTree>' +
      noteShape('say this', ' type="body"') +
      noteShape('2024-01-01', ' type="dt"') +
      noteShape('Header text', ' type="hdr"') +
      noteShape('alt text', ' type="sldImg"') +
      '</p:spTree></p:cSld></p:notes>',
  };
  const result = await extractPptx(new Uint8Array([1]), 'notes.pptx', zipOf(parts));
  assert.equal(result.units[0]?.text, 'Title\nSpeaker notes: say this');
});

test('a textless slide still gets a deterministic title', async () => {
  const parts = presentationWith(['<p:sp></p:sp>']);
  const result = await extractPptx(new Uint8Array([1]), 'blank.pptx', zipOf(parts));
  assert.equal(result.units[0]?.title, 'Slide 1');
  assert.equal(result.units[0]?.text, '');
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
