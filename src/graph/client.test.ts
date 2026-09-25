import assert from 'node:assert/strict';
import { afterEach, test } from 'node:test';

import { GraphError } from '../contracts.js';
import type { AuthProvider, GraphBatchRequest, ServerConfig } from '../contracts.js';
import { createGraphClient, extractCollection } from './client.js';

const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

// `graphHost` is a bare hostname, matching CLOUDS in src/config.ts. A value
// carrying a scheme would make graphBaseUrl() produce `https://https://...`.
const config: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['mail'],
  readOnly: false,
  graphVersion: 'v1.0',
  allowBeta: false,
  discovery: false,
  orgMode: false,
  maxOutputChars: 25_000,
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'auto',
  authPort: 0,
  allowGenericWrite: false,
};

function stubAuth(): AuthProvider & { readonly tokens: string[] } {
  const tokens: string[] = [];
  return {
    tokens,
    async getToken(): Promise<string> {
      const token = `token-${tokens.length + 1}`;
      tokens.push(token);
      return token;
    },
    async getAccount() {
      return null;
    },
    async login() {
      throw new Error('login is not exercised by these tests');
    },
    async logout() {
      // no session to clear
    },
  };
}

interface RecordedCall {
  url: string;
  init: RequestInit;
}

function installFetch(
  handler: (url: string, init: RequestInit) => Response | Promise<Response>,
): RecordedCall[] {
  const calls: RecordedCall[] = [];
  globalThis.fetch = (async (input: unknown, init: RequestInit = {}) => {
    const url = typeof input === 'string' ? input : String(input);
    calls.push({ url, init });
    return await handler(url, init);
  }) as unknown as typeof fetch;
  return calls;
}

function json(
  body: unknown,
  init: { status?: number; headers?: Record<string, string> } = {},
): Response {
  return new Response(JSON.stringify(body), {
    status: init.status ?? 200,
    headers: { 'content-type': 'application/json', ...init.headers },
  });
}

function headerOf(init: RequestInit, name: string): string | null {
  return new Headers(init.headers ?? {}).get(name);
}

test('a 429 with Retry-After is retried and then succeeds', async () => {
  let attempt = 0;
  const calls = installFetch(() => {
    attempt += 1;
    if (attempt === 1) {
      return new Response('', { status: 429, headers: { 'retry-after': '0' } });
    }
    return json({ value: [{ id: 'AAA' }] });
  });

  const client = createGraphClient({ auth: stubAuth(), config });
  const response = await client.request({ path: '/me/messages', scopes: ['Mail.Read'] });

  assert.equal(calls.length, 2);
  assert.equal(response.status, 200);
  assert.deepEqual(extractCollection(response.data), [{ id: 'AAA' }]);
});

test('a 404 throws a GraphError carrying the status, code and a hint', async () => {
  installFetch(() =>
    json(
      {
        error: {
          code: 'itemNotFound',
          message: 'The specified object was not found in the store.',
          innerError: { 'request-id': 'req-404', date: '2026-09-03T00:00:00' },
        },
      },
      { status: 404 },
    ),
  );

  const client = createGraphClient({ auth: stubAuth(), config });

  await assert.rejects(
    () => client.request({ path: '/me/messages/does-not-exist', scopes: ['Mail.Read'] }),
    (error: unknown) => {
      assert.ok(error instanceof GraphError, 'expected a GraphError');
      assert.equal(error.status, 404);
      assert.equal(error.code, 'itemNotFound');
      assert.equal(error.requestId, 'req-404');
      assert.equal(error.path, '/me/messages/does-not-exist');
      assert.ok(typeof error.hint === 'string' && error.hint.length > 0, 'expected a hint');
      return true;
    },
  );
});

test('a 401 is retried exactly once with a fresh token, then throws', async () => {
  const calls = installFetch(() =>
    json(
      { error: { code: 'InvalidAuthenticationToken', message: 'Access token has expired.' } },
      { status: 401 },
    ),
  );

  const auth = stubAuth();
  const client = createGraphClient({ auth, config });

  await assert.rejects(
    () => client.request({ path: '/me', scopes: ['User.Read'] }),
    (error: unknown) => {
      assert.ok(error instanceof GraphError);
      assert.equal(error.status, 401);
      assert.match(error.hint ?? '', /sign in again/i);
      return true;
    },
  );

  assert.equal(calls.length, 2, 'a 401 is replayed once, not more');
  assert.equal(auth.tokens.length, 2, 'the replay asks the provider for a token again');
  assert.equal(headerOf(calls[1]?.init ?? {}, 'authorization'), 'Bearer token-2');
});

test('a POST is never replayed after a network error', async () => {
  const calls = installFetch(() => {
    throw new TypeError('fetch failed');
  });

  const client = createGraphClient({ auth: stubAuth(), config });

  await assert.rejects(
    () => client.request({ path: '/me/sendMail', method: 'POST', body: { message: {} }, scopes: ['Mail.Send'] }),
    (error: unknown) => error instanceof TypeError,
  );
  assert.equal(calls.length, 1, 'a mutating call must not be sent twice on a transport failure');
});

test('a GET is replayed after a network error', async () => {
  let attempt = 0;
  const calls = installFetch(() => {
    attempt += 1;
    if (attempt === 1) throw new TypeError('fetch failed');
    return json({ value: [] });
  });

  const client = createGraphClient({ auth: stubAuth(), config });
  const response = await client.request({ path: '/me/messages', scopes: ['Mail.Read'] });

  assert.equal(calls.length, 2);
  assert.equal(response.status, 200);
});

test('paging follows @odata.nextLink and merges the value arrays', async () => {
  const nextLink = 'https://graph.microsoft.com/v1.0/me/messages?$top=2&$skiptoken=page2';
  const deltaLink = 'https://graph.microsoft.com/v1.0/me/messages/delta?$deltatoken=xyz';
  const calls = installFetch((url) => {
    if (url.includes('skiptoken=page2')) {
      return json({ value: [{ id: 'c' }], '@odata.deltaLink': deltaLink });
    }
    return json({ value: [{ id: 'a' }, { id: 'b' }], '@odata.nextLink': nextLink });
  });

  const client = createGraphClient({ auth: stubAuth(), config });
  const response = await client.request({
    path: '/me/messages',
    query: { $top: 2 },
    scopes: ['Mail.Read'],
    maxPages: 3,
  });

  assert.equal(calls.length, 2);
  assert.deepEqual(extractCollection(response.data), [{ id: 'a' }, { id: 'b' }, { id: 'c' }]);
  assert.equal(response.nextLink, undefined, 'no cursor remains once the collection is exhausted');
  assert.equal(response.deltaLink, deltaLink);
});

test('paging stops at maxPages and hands back the remaining cursor', async () => {
  const nextLink = 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=page2';
  const calls = installFetch(() => json({ value: [{ id: 'a' }], '@odata.nextLink': nextLink }));

  const client = createGraphClient({ auth: stubAuth(), config });
  const response = await client.request({ path: '/me/messages', scopes: ['Mail.Read'], maxPages: 2 });

  assert.equal(calls.length, 2);
  assert.equal(extractCollection(response.data).length, 2);
  assert.equal(response.nextLink, nextLink);
});

test('a nextLink pointing at another host is rejected before it is fetched', async () => {
  const calls = installFetch(() =>
    json({
      value: [{ id: 'a' }],
      '@odata.nextLink': 'https://evil.example.com/v1.0/me/messages?$skiptoken=page2',
    }),
  );

  const client = createGraphClient({ auth: stubAuth(), config });

  await assert.rejects(
    () => client.request({ path: '/me/messages', scopes: ['Mail.Read'], maxPages: 5 }),
    (error: unknown) => error instanceof Error,
  );
  assert.equal(calls.length, 1, 'the foreign host never receives the bearer token');
});

test('follow() refuses a link on a lookalike host', async () => {
  const calls = installFetch(() => json({ value: [] }));
  const client = createGraphClient({ auth: stubAuth(), config });

  await assert.rejects(
    () => client.follow('https://graph.microsoft.com.evil.example/v1.0/me/messages', ['Mail.Read']),
    (error: unknown) => error instanceof Error,
  );
  assert.equal(calls.length, 0);
});

test('batch() splits 25 requests into chunks of 20 and 5 and returns 25 responses in order', async () => {
  const requests: GraphBatchRequest[] = Array.from({ length: 25 }, (_, index) => ({
    id: String(index + 1),
    method: 'GET' as const,
    url: `/me/messages/${index + 1}`,
  }));

  const chunkSizes: number[] = [];
  const calls = installFetch(async (_url, init) => {
    const payload = JSON.parse(String(init.body ?? '{}')) as { requests: Array<{ id: string }> };
    chunkSizes.push(payload.requests.length);
    // Answer out of order, the way Graph does, to prove the caller's ordering is restored.
    const responses = payload.requests
      .map((entry) => ({ id: entry.id, status: 200, body: { id: entry.id } }))
      .reverse();
    return json({ responses });
  });

  const client = createGraphClient({ auth: stubAuth(), config });
  const responses = await client.batch(requests, ['Mail.Read']);

  assert.equal(calls.length, 2);
  assert.deepEqual(chunkSizes, [20, 5]);
  assert.ok(
    calls.every((call) => call.url === 'https://graph.microsoft.com/v1.0/$batch'),
    'every chunk posts to /$batch',
  );
  assert.equal(responses.length, 25);
  assert.deepEqual(
    responses.map((response) => response.id),
    requests.map((request) => request.id),
  );
  assert.ok(responses.every((response) => response.status === 200));
});

test('batch() surfaces a per-entry failure instead of throwing, and retries a 429 entry once', async () => {
  let round = 0;
  const calls = installFetch(async (_url, init) => {
    round += 1;
    const payload = JSON.parse(String(init.body ?? '{}')) as { requests: Array<{ id: string }> };
    if (round === 1) {
      return json({
        responses: [
          { id: '1', status: 200, body: { id: '1' } },
          {
            id: '2',
            status: 429,
            headers: { 'Retry-After': '0' },
            body: { error: { code: 'activityLimitReached' } },
          },
          { id: '3', status: 403, body: { error: { code: 'accessDenied', message: 'no' } } },
        ],
      });
    }
    assert.deepEqual(
      payload.requests.map((entry) => entry.id),
      ['2'],
      'only the throttled entry is replayed',
    );
    return json({ responses: [{ id: '2', status: 200, body: { id: '2' } }] });
  });

  const client = createGraphClient({ auth: stubAuth(), config });
  const responses = await client.batch(
    [
      { id: '1', method: 'GET', url: '/me/messages/1' },
      { id: '2', method: 'GET', url: '/me/messages/2' },
      { id: '3', method: 'GET', url: '/me/messages/3' },
    ],
    ['Mail.Read'],
  );

  assert.equal(calls.length, 2, 'exactly one follow-up batch for the throttled entry');
  assert.deepEqual(
    responses.map((response) => [response.id, response.status]),
    [
      ['1', 200],
      ['2', 200],
      ['3', 403],
    ],
  );
});

test('a body-carrying batch entry is sent with a JSON Content-Type header', async () => {
  let sent: { requests: Array<Record<string, unknown>> } = { requests: [] };
  installFetch(async (_url, init) => {
    sent = JSON.parse(String(init.body ?? '{}')) as { requests: Array<Record<string, unknown>> };
    return json({ responses: [{ id: '1', status: 202 }] });
  });

  const client = createGraphClient({ auth: stubAuth(), config });
  await client.batch(
    [{ id: '1', method: 'POST', url: '/me/sendMail', body: { message: { subject: 'hi' } } }],
    ['Mail.Send'],
  );

  assert.deepEqual(sent.requests[0]?.['headers'], { 'Content-Type': 'application/json' });
});

test('a $count query gets the ConsistencyLevel header Graph requires', async () => {
  const calls = installFetch(() => json({ value: [], '@odata.count': 0 }));
  const client = createGraphClient({ auth: stubAuth(), config });

  await client.request({
    path: '/users',
    query: { $count: true, $top: 1 },
    scopes: ['User.Read.All'],
  });

  assert.equal(headerOf(calls[0]?.init ?? {}, 'consistencylevel'), 'eventual');
  assert.match(calls[0]?.url ?? '', /^https:\/\/graph\.microsoft\.com\/v1\.0\/users\?/);
});

test('a query string inlined in the path is parsed rather than encoded into a segment', async () => {
  const calls = installFetch(() => json({ value: [] }));
  const client = createGraphClient({ auth: stubAuth(), config });

  await client.request({ path: '/me/messages?$top=3', scopes: ['Mail.Read'] });

  const url = new URL(calls[0]?.url ?? '');
  assert.equal(url.pathname, '/v1.0/me/messages');
  assert.equal(url.searchParams.get('$top'), '3');
});

test('beta is refused unless the flag is set, and nothing is sent', async () => {
  const calls = installFetch(() => json({}));
  const client = createGraphClient({ auth: stubAuth(), config });

  await assert.rejects(
    () => client.request({ path: '/me', version: 'beta', scopes: ['User.Read'] }),
    /--beta/,
  );
  assert.equal(calls.length, 0);
});

test('a 204 yields a null body', async () => {
  installFetch(() => new Response(null, { status: 204 }));
  const client = createGraphClient({ auth: stubAuth(), config });

  const response = await client.request({
    path: '/me/messages/1',
    method: 'DELETE',
    scopes: ['Mail.ReadWrite'],
  });

  assert.equal(response.status, 204);
  assert.equal(response.data, null);
});

test('extractCollection returns [] for anything that is not a collection', () => {
  assert.deepEqual(extractCollection({ value: [1, 2] }), [1, 2]);
  assert.deepEqual(extractCollection({ id: 'x' }), []);
  assert.deepEqual(extractCollection(null), []);
  assert.deepEqual(extractCollection('nope'), []);
});

test('a string body is sent verbatim so file content is not JSON-wrapped', async () => {
  // Graph's /drive/…/:/content endpoints store the request body as the file's
  // bytes. JSON-encoding a string there would write a file that begins and ends
  // with a quote and whose newlines are literal backslash-n.
  const text = 'line one\nline two\t"quoted"';
  const calls = installFetch(() => json({ id: 'ITEM', name: 'a.txt' }));

  const client = createGraphClient({ auth: stubAuth(), config });
  await client.request({
    path: '/me/drive/root:/notes.txt:/content',
    method: 'PUT',
    body: text,
    scopes: ['Files.ReadWrite'],
  });

  assert.equal(calls.length, 1);
  assert.equal(calls[0]!.init.body, text);
  assert.equal(headerOf(calls[0]!.init, 'content-type'), 'text/plain');
});

test('an explicit Content-Type from the caller is not overwritten', async () => {
  const calls = installFetch(() => json({ id: 'ITEM' }));

  const client = createGraphClient({ auth: stubAuth(), config });
  await client.request({
    path: '/me/drive/root:/notes.md:/content',
    method: 'PUT',
    body: '# heading',
    headers: { 'Content-Type': 'text/markdown; charset=utf-8' },
    scopes: ['Files.ReadWrite'],
  });

  assert.equal(headerOf(calls[0]!.init, 'content-type'), 'text/markdown; charset=utf-8');
});

test('an object body is still JSON-encoded with a JSON Content-Type', async () => {
  const calls = installFetch(() => json({ id: 'MSG' }));

  const client = createGraphClient({ auth: stubAuth(), config });
  await client.request({
    path: '/me/messages',
    method: 'POST',
    body: { subject: 'hi', importance: 'normal' },
    scopes: ['Mail.ReadWrite'],
  });

  assert.equal(calls[0]!.init.body, '{"subject":"hi","importance":"normal"}');
  assert.equal(headerOf(calls[0]!.init, 'content-type'), 'application/json');
});

test('the caller can never override the Authorization header', async () => {
  const calls = installFetch(() => json({ ok: true }));

  const client = createGraphClient({ auth: stubAuth(), config });
  await client.request({
    path: '/me',
    headers: { Authorization: 'Bearer attacker-supplied' },
    scopes: ['User.Read'],
  });

  assert.notEqual(headerOf(calls[0]!.init, 'authorization'), 'Bearer attacker-supplied');
});
