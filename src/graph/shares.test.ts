import assert from 'node:assert/strict';
import test from 'node:test';

import type { GraphClient, GraphRequestOptions, GraphResponse } from '../contracts.js';
import { resolveShare, shareTokenToUrl, shareUrlToToken } from './shares.js';

test('a sharing link encodes to a u! base64url token and back', () => {
  const url = 'https://tenant.sharepoint.com/:w:/g/personal/ada/Ea1b2C3d4E5';
  const token = shareUrlToToken(url);
  assert.ok(token.startsWith('u!'), `token should carry the u! prefix, got ${token}`);
  assert.ok(!/[+/=]/.test(token), `token must be base64url without padding, got ${token}`);
  assert.equal(shareTokenToUrl(token), url);
});

test('blank and non-URL links fail before any request', () => {
  for (const bad of ['', '   ', 'not a url', 'ftp://files.example.com/x.docx']) {
    assert.throws(() => shareUrlToToken(bad), /shareUrl/);
  }
});

test('resolveShare requests /shares/{token}/driveItem and lifts the ids', async () => {
  const seen: GraphRequestOptions[] = [];
  const graph: GraphClient = {
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      seen.push(opts);
      return {
        status: 200,
        data: {
          id: 'ITEM',
          name: 'q3.docx',
          size: 41,
          file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
          parentReference: { driveId: 'DRIVE' },
          webUrl: 'https://tenant.sharepoint.com/q3.docx',
        } as T,
      };
    },
    async batch() {
      throw new Error('unused');
    },
    async follow() {
      throw new Error('unused');
    },
  };

  const item = await resolveShare(graph, ['Files.Read.All'], 'https://tenant.sharepoint.com/:w:/g/x');
  assert.equal(item.driveId, 'DRIVE');
  assert.equal(item.itemId, 'ITEM');
  assert.equal(item.name, 'q3.docx');
  assert.equal(item.kind, 'file');
  assert.match(seen[0]?.path ?? '', /^\/shares\/u![A-Za-z0-9_-]+\/driveItem$/);
  assert.deepEqual(seen[0]?.scopes, ['Files.Read.All']);
});

test('a resolved item without ids is a deleted-link error, not an empty result', async () => {
  const graph: GraphClient = {
    async request<T = unknown>(): Promise<GraphResponse<T>> {
      return { status: 200, data: { id: null, parentReference: null } as T };
    },
    async batch() {
      throw new Error('unused');
    },
    async follow() {
      throw new Error('unused');
    },
  };
  await assert.rejects(() => resolveShare(graph, ['Files.Read.All'], 'https://example.com/x'), /deleted/);
});
