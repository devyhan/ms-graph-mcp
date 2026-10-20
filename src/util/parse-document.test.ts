import assert from 'node:assert/strict';
import test from 'node:test';

import type { GraphClient, GraphRequestOptions, GraphResponse } from '../contracts.js';
import { AIP_COVER_PHRASES, MAX_DOCUMENT_BYTES, ProtectedDocumentError } from './extract.js';
import {
  decodeParseCursor,
  parseDocument,
  type ParseDeps,
} from './parse-document.js';

const PDF_BYTES = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31]);

/** Minimal fake Purview wrapper: magic + /Encrypt + an IRM marker. Never real bytes. */
const WRAPPER_PDF_BYTES = new Uint8Array(
  Buffer.from('%PDF-1.7\n1 0 obj\n<< /Filter /Aadrm >>\ntrailer << /Encrypt 1 0 R >>', 'latin1'),
);

const COVER_TEXT = `Protective cover\n${AIP_COVER_PHRASES[0] ?? 'missing phrase'}\nEnd of cover`;

function stubGraph(routes: Record<string, unknown>): GraphClient & { paths: string[] } {
  const paths: string[] = [];
  return {
    paths,
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      paths.push(opts.path);
      // /shares/{token}/driveItem answers the same driveItem shape everywhere.
      const key = opts.path.startsWith('/shares/') ? '/shares/*' : opts.path;
      if (!(key in routes)) throw new Error(`stub has no route for ${opts.path}`);
      return { status: 200, data: routes[key] as T };
    },
    async batch() {
      throw new Error('unused');
    },
    async follow() {
      throw new Error('unused');
    },
  };
}

function deps(graph: GraphClient, bytes = PDF_BYTES): ParseDeps & { graph: GraphClient & { paths: string[] } } {
  return {
    graph: graph as GraphClient & { paths: string[] },
    fileScopes: ['Files.Read.All'],
    mailScopes: ['Mail.Read'],
    fetchBytes: async () => ({ bytes, nextOffset: bytes.length }),
    extract: async (kind) => ({
      units: [
        { id: 'page-1', title: 'Page 1', text: `first ${kind} page` },
        { id: 'page-2', title: 'Page 2', text: 'second page' },
        { id: 'page-3', title: 'Page 3', text: 'third page' },
      ],
      warnings: [],
    }),
  };
}

const DRIVE_META = {
  id: 'ITEM',
  name: 'report.pdf',
  size: 6,
  file: { mimeType: 'application/pdf' },
  parentReference: { driveId: 'DRIVE' },
};

test('a drive address parses to unit rows plus a manifest', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const result = await parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, deps(graph));

  assert.equal(result.fileName, 'report.pdf');
  assert.equal(result.documentKind, 'pdf');
  assert.equal(result.unitKind, 'page');
  assert.equal(result.units.length, 3);
  assert.equal(result.units[0]?.i, 0);
  assert.equal(result.units[0]?.text, 'first pdf page');
  assert.deepEqual(result.manifest.units[0], { i: 0, id: 'page-1', chars: 14, truncated: false });
  assert.equal(result.manifest.unitCount, 3);
  assert.deepEqual(result.manifest.returnedRange, { from: 0, to: 2 });
  assert.equal(result.manifest.reason, 'complete');
  assert.equal(result.manifest.cursor, undefined);
});

test('a sharing link resolves through /shares first', async () => {
  const graph = stubGraph({ '/shares/*': DRIVE_META, '/drives/DRIVE/items/ITEM': DRIVE_META });
  const result = await parseDocument(
    { source: { shareUrl: 'https://tenant.sharepoint.com/:w:/g/x' } },
    deps(graph),
  );
  assert.match(graph.paths[0] ?? '', /^\/shares\/u!/);
  assert.equal(result.driveId, 'DRIVE');
  assert.equal(result.itemId, 'ITEM');
});

test('a mail file attachment reads through /$value', async () => {
  const graph = stubGraph({
    '/me/messages/M/attachments/A': {
      id: 'A',
      name: 'deck.pdf',
      contentType: 'application/pdf',
      size: 6,
      '@odata.type': '#microsoft.graph.fileAttachment',
    },
  });
  const d = deps(graph);
  const seen: string[] = [];
  d.fetchBytes = async (opts) => {
    seen.push(opts.path);
    return { bytes: PDF_BYTES, nextOffset: PDF_BYTES.length };
  };
  const result = await parseDocument({ source: { messageId: 'M', attachmentId: 'A' } }, d);
  assert.deepEqual(seen, ['/me/messages/M/attachments/A/$value']);
  assert.equal(result.fileName, 'deck.pdf');
});

test('a mail attachment that overflows the fetch budget fails like a drive file', async () => {
  const graph = stubGraph({
    '/me/messages/M/attachments/A': {
      id: 'A',
      name: 'big.pdf',
      contentType: 'application/pdf',
      '@odata.type': '#microsoft.graph.fileAttachment',
    },
  });
  const d = deps(graph);
  d.fetchBytes = async () => ({
    bytes: new Uint8Array(MAX_DOCUMENT_BYTES),
    nextOffset: MAX_DOCUMENT_BYTES,
    totalLength: MAX_DOCUMENT_BYTES + 1,
  });
  await assert.rejects(
    () => parseDocument({ source: { messageId: 'M', attachmentId: 'A' } }, d),
    /exceeds the 16 MiB/,
  );
});

test('a mail reference attachment points at shareUrl, never parsed directly', async () => {
  const graph = stubGraph({
    '/me/messages/M/attachments/A': {
      id: 'A',
      name: 'linked.docx',
      '@odata.type': '#microsoft.graph.referenceAttachment',
    },
  });
  await assert.rejects(
    () => parseDocument({ source: { messageId: 'M', attachmentId: 'A' } }, deps(graph)),
    /shareUrl/,
  );
});

test('units page with a cursor that replays the source', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const first = await parseDocument(
    { source: { driveId: 'DRIVE', itemId: 'ITEM' }, unitLimit: 2 },
    deps(graph),
  );
  assert.equal(first.units.length, 2);
  assert.equal(first.manifest.reason, 'maxItems');
  assert.ok(first.manifest.cursor !== undefined);

  const cursor = decodeParseCursor(first.manifest.cursor);
  assert.deepEqual(cursor.source, { driveId: 'DRIVE', itemId: 'ITEM' });
  const second = await parseDocument(
    { source: cursor.source, unitOffset: cursor.unitOffset, unitLimit: cursor.unitLimit },
    deps(graph),
  );
  assert.deepEqual(
    second.units.map((unit) => unit.i),
    [2],
  );
  assert.equal(second.manifest.reason, 'complete');
});

test('per-unit truncation flags instead of silently cutting', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const result = await parseDocument(
    { source: { driveId: 'DRIVE', itemId: 'ITEM' }, maxCharsPerUnit: 200 },
    {
      ...deps(graph),
      extract: async () => ({ units: [{ id: 'p1', text: 'x'.repeat(500) }], warnings: [] }),
    },
  );
  assert.equal(result.units[0]?.truncated, true);
  assert.match(result.units[0]?.text ?? '', /\[truncated/);
});

test('oversize and folder sources fail with friendly errors', async () => {
  const big = stubGraph({
    '/drives/DRIVE/items/ITEM': { ...DRIVE_META, size: MAX_DOCUMENT_BYTES + 1 },
  });
  await assert.rejects(
    () => parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, deps(big)),
    /over 16 MiB/,
  );
  const folder = stubGraph({ '/drives/DRIVE/items/ITEM': { id: 'ITEM', name: 'F', folder: {} } });
  await assert.rejects(
    () => parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, deps(folder)),
    /folder/,
  );
});

test('a bogus cursor fails before any request', () => {
  assert.throws(() => decodeParseCursor('not-a-cursor!!'), /Invalid cursor/);
});

test('the manifest carries a whole-document titles outline beyond the slice', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const result = await parseDocument(
    { source: { driveId: 'DRIVE', itemId: 'ITEM' }, unitLimit: 1 },
    deps(graph),
  );
  // The slice holds one unit, but the outline surveys all three.
  assert.equal(result.units.length, 1);
  assert.deepEqual(result.manifest.titles, [
    { i: 0, id: 'page-1', title: 'Page 1' },
    { i: 1, id: 'page-2', title: 'Page 2' },
    { i: 2, id: 'page-3', title: 'Page 3' },
  ]);
  assert.equal(result.manifest.unitCount, 3);
});

test('untitled units appear in the outline without a title key', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const result = await parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, {
    ...deps(graph),
    extract: async () => ({ units: [{ id: 'p1', text: 'no heading here' }], warnings: [] }),
  });
  assert.deepEqual(result.manifest.titles, [{ i: 0, id: 'p1' }]);
});

test('an ordinary document reports unprotected in the result and the manifest', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const result = await parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, deps(graph));
  assert.equal(result.protected, false);
  assert.equal(result.protection, undefined);
  assert.equal(result.decrypted, false);
  assert.equal(result.manifest.protected, false);
  assert.equal(result.manifest.protection, undefined);
  assert.equal(result.manifest.decrypted, false);
});

test('a wrapper without a decryptor is refused with structure, never returned as cover units', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const error = await parseDocument(
    { source: { driveId: 'DRIVE', itemId: 'ITEM' } },
    deps(graph, WRAPPER_PDF_BYTES),
  ).then(
    () => null,
    (thrown: unknown) => thrown,
  );
  assert.ok(error instanceof ProtectedDocumentError);
  assert.equal(error.code, 'PROTECTED_DOCUMENT');
  assert.equal(error.protection, 'aip-rms');
  assert.equal(error.fileName, 'report.pdf');
  assert.match(error.message, /decrypt/i);
});

test('a wrapper with a decryptor parses the plaintext and reports it', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  const seen: Array<{ bytes: Uint8Array; fileName: string }> = [];
  const result = await parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, {
    ...deps(graph, WRAPPER_PDF_BYTES),
    decrypt: async (bytes, fileName) => {
      seen.push({ bytes, fileName });
      return PDF_BYTES;
    },
    extract: async () => ({ units: [{ id: 'page-1', title: 'Page 1', text: 'decrypted content' }], warnings: [] }),
  });
  assert.equal(seen.length, 1);
  assert.equal(seen[0]?.fileName, 'report.pdf');
  assert.deepEqual(seen[0]?.bytes, WRAPPER_PDF_BYTES);
  assert.equal(result.units[0]?.text, 'decrypted content');
  assert.equal(result.protected, true);
  assert.equal(result.protection, 'aip-rms');
  assert.equal(result.decrypted, true);
  assert.equal(result.manifest.protected, true);
  assert.equal(result.manifest.protection, 'aip-rms');
  assert.equal(result.manifest.decrypted, true);
});

test('a cover page without byte markers is refused without a decryptor', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  await assert.rejects(
    () =>
      parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, {
        ...deps(graph),
        extract: async () => ({ units: [{ id: 'page-1', text: COVER_TEXT }], warnings: [] }),
      }),
    /Purview|AIP\/RMS/,
  );
});

test('a cover page with a decryptor re-extracts the plaintext', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  let calls = 0;
  const result = await parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, {
    ...deps(graph),
    decrypt: async () => PDF_BYTES,
    extract: async () => {
      calls += 1;
      return calls === 1
        ? { units: [{ id: 'page-1', text: COVER_TEXT }], warnings: [] }
        : { units: [{ id: 'page-1', text: 'real content after decrypt' }], warnings: [] };
    },
  });
  assert.equal(calls, 2);
  assert.equal(result.units[0]?.text, 'real content after decrypt');
  assert.equal(result.protected, true);
  assert.equal(result.protection, 'aip-rms');
  assert.equal(result.decrypted, true);
});

test('decrypted bytes that still read as a cover are refused', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  await assert.rejects(
    () =>
      parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, {
        ...deps(graph, WRAPPER_PDF_BYTES),
        decrypt: async () => PDF_BYTES,
        extract: async () => ({ units: [{ id: 'page-1', text: COVER_TEXT }], warnings: [] }),
      }),
    /Purview|AIP\/RMS/,
  );
});

test('a decrypt output that is still a wrapper is refused', async () => {
  const graph = stubGraph({ '/drives/DRIVE/items/ITEM': DRIVE_META });
  await assert.rejects(
    () =>
      parseDocument({ source: { driveId: 'DRIVE', itemId: 'ITEM' } }, {
        ...deps(graph, WRAPPER_PDF_BYTES),
        decrypt: async () => WRAPPER_PDF_BYTES,
      }),
    /Purview|AIP\/RMS/,
  );
});
