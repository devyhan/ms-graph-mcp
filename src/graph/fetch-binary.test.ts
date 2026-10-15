import assert from 'node:assert/strict';
import test from 'node:test';

import type { AuthProvider, ServerConfig } from '../contracts.js';
import { fetchContentRange } from './fetch-binary.js';

const config: Pick<ServerConfig, 'graphHost' | 'graphVersion'> = {
  graphHost: 'graph.microsoft.com',
  graphVersion: 'v1.0',
};

const auth: AuthProvider = {
  async getToken() {
    return 'TOKEN';
  },
  async getAccount() {
    return null;
  },
  async login() {
    throw new Error('unused');
  },
  async logout() {},
  async beginDeviceLogin() {
    throw new Error('unused');
  },
  async deviceLoginStatus() {
    return { state: 'none' as const };
  },
};

interface Seen {
  url: string;
  init: RequestInit;
}

function stubFetch(handler: (seen: Seen[]) => Response | Promise<Response>): {
  seen: Seen[];
  impl: typeof fetch;
} {
  const seen: Seen[] = [];
  const impl = (async (url: unknown, init?: RequestInit) => {
    seen.push({ url: String(url), init: init ?? {} });
    return handler(seen);
  }) as typeof fetch;
  return { seen, impl };
}

function rangeOf(init: RequestInit): string | null {
  return new Headers(init.headers).get('range');
}

function authOf(init: RequestInit): string | null {
  return new Headers(init.headers).get('authorization');
}

test('a 302 re-issues Range on the Location without the bearer', async () => {
  const payload = new Uint8Array([1, 2, 3, 4]).buffer;
  const { seen, impl } = stubFetch((calls) =>
    calls.length === 1
      ? new Response(null, { status: 302, headers: { location: 'https://storage.example.net/dl?sig=abc' } })
      : new Response(payload, { status: 206, headers: { 'content-range': 'bytes 10-13/100' } }),
  );

  const result = await fetchContentRange(auth, config, {
    path: '/drives/D/items/I/content',
    scopes: ['Files.Read.All'],
    offset: 10,
    byteBudget: 4,
    fetchImpl: impl,
  });

  assert.equal(seen.length, 2);
  assert.match(seen[0]?.url ?? '', /^https:\/\/graph\.microsoft\.com\/v1\.0\/drives\//);
  assert.equal(authOf(seen[0]?.init ?? {}), 'Bearer TOKEN');
  assert.equal(rangeOf(seen[0]?.init ?? {}), 'bytes=10-13');
  // The storage hop carries the same Range and must NOT carry the token.
  assert.equal(seen[1]?.url, 'https://storage.example.net/dl?sig=abc');
  assert.equal(rangeOf(seen[1]?.init ?? {}), 'bytes=10-13');
  assert.equal(authOf(seen[1]?.init ?? {}), null);

  assert.deepEqual([...result.bytes], [1, 2, 3, 4]);
  assert.equal(result.nextOffset, 14);
  assert.equal(result.totalLength, 100);
  assert.equal(result.truncated, true);
});

test('a direct 200 needs no second hop and reports its cursor', async () => {
  const { seen, impl } = stubFetch(() => new Response(new Uint8Array([9]).buffer, { status: 200 }));
  const result = await fetchContentRange(auth, config, {
    path: '/me/messages/M/attachments/A/$value',
    scopes: ['Mail.Read'],
    fetchImpl: impl,
  });
  assert.equal(seen.length, 1);
  assert.equal(result.nextOffset, 1);
  assert.equal(result.totalLength, undefined);
  assert.equal(result.truncated, false);
});

test('a redirect without a Location fails loudly', async () => {
  const { impl } = stubFetch(() => new Response(null, { status: 302 }));
  await assert.rejects(
    () =>
      fetchContentRange(auth, config, {
        path: '/drives/D/items/I/content',
        scopes: ['Files.Read.All'],
        fetchImpl: impl,
      }),
    /without a Location/,
  );
});
