import assert from 'node:assert/strict';
import test from 'node:test';

import type { GraphClient, GraphRequestOptions, GraphResponse } from '../contracts.js';
import { MAX_DOCUMENT_BYTES } from './extract.js';
import {
  decodeParseCursor,
  parseDocument,
  type ParseDeps,
} from './parse-document.js';

const PDF_BYTES = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31]);

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
