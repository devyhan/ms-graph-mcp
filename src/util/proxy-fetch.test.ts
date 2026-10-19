import assert from 'node:assert/strict';
import net from 'node:net';
import { Duplex } from 'node:stream';
import test from 'node:test';

import {
  createProxiedFetch,
  decodeChunked,
  installProxyFetch,
  isBypassed,
  isProxyFetchInstalled,
  parseStatusAndHeaders,
  resolveProxyConfig,
  toBodyBuffer,
  uninstallProxyFetch,
} from './proxy-fetch.js';
import type { ProxyTransport } from './proxy-fetch.js';
import type { AuthProvider, ServerConfig } from '../contracts.js';
import { fetchContentRange } from '../graph/fetch-binary.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * A closed loopback port: dials fail fast with no listen needed. High ports
 * only — undici refuses to dial the WHATWG bad-port list before connecting.
 */
interface ClosedPort {
  port: number;
  /** ECONNREFUSED normally; a syscall sandbox reports EPERM instead. */
  code: string;
}

async function findClosedPort(skip: readonly number[] = []): Promise<ClosedPort> {
  for (const port of [54321, 54322, 54323, 54324, 54325, 59123, 59124, 59125]) {
    if (skip.includes(port)) continue;
    const outcome = await new Promise<'open' | string>((resolve) => {
      const sock = net.connect(port, '127.0.0.1');
      sock.once('connect', () => {
        sock.destroy();
        resolve('open');
      });
      sock.once('error', (error: NodeJS.ErrnoException) => resolve(error.code ?? 'UNKNOWN'));
    });
    if (outcome === 'open') continue;
    return { port, code: outcome };
  }
  throw new Error('test setup: no closed loopback port found');
}

/** Asserts a rejection shaped like a native transport failure. */
function isFetchFailedWith(error: unknown, code: string, port?: number): boolean {
  assert.ok(error instanceof TypeError, `expected TypeError, got ${String(error)}`);
  assert.equal((error as Error).message, 'fetch failed');
  const cause = (error as { cause?: { code?: unknown; port?: unknown } }).cause;
  assert.equal(cause?.code, code);
  if (port !== undefined) assert.equal(cause?.port, port);
  return true;
}

/** One observed HTTP request: the request line, lowercased headers, raw body. */
interface ObservedRequest {
  requestLine: string;
  headers: Map<string, string>;
  body: Buffer;
}

interface ObservedBox {
  request?: ObservedRequest;
}

/** Parses a complete request, or null when more bytes are still needed. */
function parseObservedRequest(received: Buffer): ObservedRequest | null {
  const idx = received.indexOf('\r\n\r\n');
  if (idx === -1) return null;
  const lines = received.subarray(0, idx).toString('latin1').split('\r\n');
  const requestLine = lines.shift() ?? '';
  const headers = new Map<string, string>();
  for (const line of lines) {
    const separator = line.indexOf(':');
    if (separator !== -1) {
      headers.set(line.slice(0, separator).trim().toLowerCase(), line.slice(separator + 1).trim());
    }
  }
  const need = Number(headers.get('content-length') ?? '0');
  const bodyStart = idx + 4;
  if (!Number.isFinite(need) || received.length < bodyStart + need) return null;
  return { requestLine, headers, body: received.subarray(bodyStart, bodyStart + need) };
}

function httpResponse(statusLine: string, headerLines: string[], body: Buffer | string): Buffer {
  const payload = typeof body === 'string' ? Buffer.from(body) : body;
  return Buffer.concat([
    Buffer.from(`HTTP/1.1 ${statusLine}\r\n${headerLines.join('\r\n')}\r\n\r\n`, 'latin1'),
    payload,
  ]);
}

type ScriptResult = { reply: Buffer; closeAfter: boolean } | null;

/** A socket that answers one request from a script instead of the network. */
class ScriptedSocket extends Duplex {
  received = Buffer.alloc(0);
  #respond: (received: Buffer) => ScriptResult;
  #replied = false;

  constructor(respond: (received: Buffer) => ScriptResult) {
    super();
    this.#respond = respond;
  }

  override _write(
    chunk: object,
    _encoding: BufferEncoding,
    callback: (error?: Error | null) => void,
  ): void {
    this.received = Buffer.concat([this.received, chunk as Buffer]);
    if (!this.#replied) {
      const out = this.#respond(this.received);
      if (out !== null) {
        this.#replied = true;
        // Deferred so the reader's listeners are attached before bytes flow.
        setImmediate(() => {
          this.push(out.reply);
          if (out.closeAfter) this.push(null);
        });
      }
    }
    callback();
  }

  override _read(): void {}
}

/** Records the observed request and answers with the canned response. */
function answering(seen: ObservedBox, reply: Buffer): (received: Buffer) => ScriptResult {
  return (received: Buffer) => {
    const request = parseObservedRequest(received);
    if (request === null) return null;
    seen.request = request;
    return { reply, closeAfter: true };
  };
}

interface TransportCalls {
  dials: Array<{ host: string; port: number }>;
  secures: string[];
}

/**
 * A transport whose dial/secure calls consume responders in order: an HTTPS
 * hop takes two (CONNECT, then the origin request), a plain-HTTP hop one.
 */
function scriptedTransport(
  responders: Array<(received: Buffer) => ScriptResult>,
  calls: TransportCalls,
): ProxyTransport {
  let next = 0;
  const take = (): ((received: Buffer) => ScriptResult) => responders[next++] ?? (() => null);
  return {
    dial: (host: string, port: number): Promise<net.Socket> => {
      calls.dials.push({ host, port });
      return Promise.resolve(new ScriptedSocket(take()) as unknown as net.Socket);
    },
    secure: (_sock: net.Socket, servername: string): Promise<net.Socket> => {
      calls.secures.push(servername);
      return Promise.resolve(new ScriptedSocket(take()) as unknown as net.Socket);
    },
  };
}

function proxiedVia(
  proxy: string,
  noProxy: readonly string[],
  transport: ProxyTransport,
  nativeFetch: typeof fetch = globalThis.fetch,
): typeof fetch {
  return createProxiedFetch({ proxy: new URL(proxy), noProxy, nativeFetch, transport });
}

// ---------------------------------------------------------------------------
// Env resolution and NO_PROXY matching
// ---------------------------------------------------------------------------

test('resolveProxyConfig prefers HTTPS_PROXY over the lowercase and ALL variants', () => {
  const config = resolveProxyConfig({
    HTTPS_PROXY: 'http://a:1',
    https_proxy: 'http://b:2',
    ALL_PROXY: 'http://c:3',
    all_proxy: 'http://d:4',
  });
  assert.equal(config?.source, 'HTTPS_PROXY');
  assert.equal(config?.url.href, 'http://a:1/');
});

test('resolveProxyConfig falls through the variants in order', () => {
  assert.equal(resolveProxyConfig({ all_proxy: 'http://d:4' })?.source, 'all_proxy');
  assert.equal(
    resolveProxyConfig({ https_proxy: 'http://b:2', ALL_PROXY: 'http://c:3' })?.source,
    'https_proxy',
  );
  assert.equal(resolveProxyConfig({})?.source, undefined);
  assert.equal(resolveProxyConfig({ HTTPS_PROXY: '' })?.source, undefined);
});

test('resolveProxyConfig ignores HTTP_PROXY: only the HTTPS-capable vars route', () => {
  assert.equal(resolveProxyConfig({ HTTP_PROXY: 'http://x:1' }), undefined);
});

test('resolveProxyConfig tolerates a bare host:port and skips garbage', () => {
  assert.equal(resolveProxyConfig({ HTTPS_PROXY: '127.0.0.1:8080' })?.url.href, 'http://127.0.0.1:8080/');
  assert.equal(resolveProxyConfig({ HTTPS_PROXY: 'http://exa mple' }), undefined);
});

test('resolveProxyConfig skips values that cannot serve a tunnel', () => {
  // An empty host and a non-HTTP scheme both parse as URLs, but neither can
  // answer a CONNECT: treat them like typos rather than dialing wrongly.
  assert.equal(resolveProxyConfig({ HTTPS_PROXY: 'http://?x' }), undefined);
  assert.equal(resolveProxyConfig({ ALL_PROXY: 'socks5://127.0.0.1:1080' }), undefined);
  assert.equal(resolveProxyConfig({ ALL_PROXY: 'ftp://127.0.0.1:21' }), undefined);
});

test('resolveProxyConfig falls past a malformed entry to a valid later one', () => {
  const config = resolveProxyConfig({
    HTTPS_PROXY: 'http://exa mple',
    ALL_PROXY: 'http://127.0.0.1:9',
  });
  assert.equal(config?.source, 'ALL_PROXY');
  assert.equal(config?.url.href, 'http://127.0.0.1:9/');
});

test('resolveProxyConfig parses NO_PROXY into trimmed lowercase rules', () => {
  const config = resolveProxyConfig({
    HTTPS_PROXY: 'http://127.0.0.1:8080',
    NO_PROXY: 'Example.COM, 10.0.0.1 ,, ',
  });
  assert.deepEqual(config?.noProxy, ['example.com', '10.0.0.1']);
  const lower = resolveProxyConfig({
    HTTPS_PROXY: 'http://127.0.0.1:8080',
    no_proxy: 'a.example',
  });
  assert.deepEqual(lower?.noProxy, ['a.example']);
});

test('isBypassed matches *, exact hosts and domain suffixes', () => {
  assert.equal(isBypassed('anything.example', ['*']), true);
  assert.equal(isBypassed('example.com', ['example.com']), true);
  assert.equal(isBypassed('API.EXAMPLE.COM', ['example.com']), true);
  assert.equal(isBypassed('api.example.com', ['example.com']), true);
  assert.equal(isBypassed('api.example.com', ['.example.com']), true);
  assert.equal(isBypassed('example.com', ['.example.com']), true);
  assert.equal(isBypassed('example.com.evil.net', ['example.com']), false);
  assert.equal(isBypassed('other.example', ['example.com']), false);
  assert.equal(isBypassed('example.com', []), false);
  assert.equal(isBypassed('example.com', ['']), false);
});

// ---------------------------------------------------------------------------
// Body and framing units
// ---------------------------------------------------------------------------

test('toBodyBuffer passes strings, buffers, params and views through', () => {
  assert.deepEqual(toBodyBuffer('hi'), Buffer.from('hi'));
  assert.deepEqual(toBodyBuffer(new URLSearchParams({ a: '1', b: '2' })), Buffer.from('a=1&b=2'));
  const buf = Buffer.from([1, 2]);
  assert.ok(toBodyBuffer(buf) === buf);
  assert.deepEqual(toBodyBuffer(new Uint8Array([7]).buffer), Buffer.from([7]));
  const view = new Uint8Array([0, 1, 2, 3]).subarray(1, 3);
  assert.deepEqual(toBodyBuffer(view), Buffer.from([1, 2]));
  assert.equal(toBodyBuffer(undefined), undefined);
  assert.equal(toBodyBuffer(null), undefined);
});

test('toBodyBuffer throws on unsupported bodies rather than missending', () => {
  assert.throws(() => toBodyBuffer(123), TypeError);
  assert.throws(() => toBodyBuffer({}), TypeError);
});

test('decodeChunked joins chunks and tolerates extensions', () => {
  assert.equal(
    decodeChunked(Buffer.from('5\r\nhello\r\n6\r\n world\r\n0\r\n\r\n')).toString(),
    'hello world',
  );
  assert.equal(decodeChunked(Buffer.from('5;ext=1\r\nhello\r\n0\r\n\r\n')).toString(), 'hello');
  assert.equal(decodeChunked(Buffer.from('A\r\n0123456789\r\n0\r\n\r\n')).toString(), '0123456789');
  assert.equal(decodeChunked(Buffer.from('0\r\n\r\n')).length, 0);
});

test('parseStatusAndHeaders reads the status and folds repeated headers', () => {
  const { status, headers } = parseStatusAndHeaders(
    'HTTP/1.1 404 Not Found\r\nContent-Type: text/plain\r\nX-Multi: a\r\nX-Multi: b',
  );
  assert.equal(status, 404);
  assert.equal(headers.get('content-type'), 'text/plain');
  assert.equal(headers.get('x-multi'), 'a, b');
});

test('parseStatusAndHeaders rejects a garbage status line', () => {
  assert.throws(() => parseStatusAndHeaders('GARBAGE'), /bad status line/);
});

// ---------------------------------------------------------------------------
// Scripted-transport round trips (no sockets at all)
// ---------------------------------------------------------------------------

const CONNECT_OK = Buffer.from('HTTP/1.1 200 Connection Established\r\n\r\n');

test('an https GET tunnels with CONNECT and forwards headers', async () => {
  const connect: ObservedBox = {};
  const origin: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering(connect, CONNECT_OK),
        answering(origin, httpResponse('200 OK', ['Content-Length: 2', 'X-Stub: yes'], 'ok')),
      ],
      calls,
    ),
  );

  const response = await fetchVia('https://api.example/items?q=1', {
    headers: { 'X-Test': 'forward-me', Authorization: 'Bearer abc' },
  });
  assert.equal(response.status, 200);
  assert.equal(await response.text(), 'ok');
  assert.equal(response.headers.get('x-stub'), 'yes');

  assert.deepEqual(calls.dials, [{ host: '127.0.0.1', port: 8080 }]);
  assert.deepEqual(calls.secures, ['api.example']);
  assert.equal(connect.request?.requestLine, 'CONNECT api.example:443 HTTP/1.1');
  assert.equal(connect.request?.headers.get('host'), 'api.example:443');
  assert.equal(origin.request?.requestLine, 'GET /items?q=1 HTTP/1.1');
  assert.equal(origin.request?.headers.get('host'), 'api.example');
  assert.equal(origin.request?.headers.get('x-test'), 'forward-me');
  assert.equal(origin.request?.headers.get('authorization'), 'Bearer abc');
  assert.equal(origin.request?.headers.get('connection'), 'close');
});

test('proxy credentials ride Proxy-Authorization on CONNECT', async () => {
  const connect: ObservedBox = {};
  const origin: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://user:p%40ss@127.0.0.1:8080',
    [],
    scriptedTransport(
      [answering(connect, CONNECT_OK), answering(origin, httpResponse('200 OK', [], ''))],
      calls,
    ),
  );

  await fetchVia('https://api.example/');
  const expected = `Basic ${Buffer.from('user:p@ss').toString('base64')}`;
  assert.equal(connect.request?.headers.get('proxy-authorization'), expected);
});

test('a credential-less proxy sends no Proxy-Authorization anywhere', async () => {
  const connect: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchTls = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [answering(connect, CONNECT_OK), answering({}, httpResponse('200 OK', [], ''))],
      calls,
    ),
  );
  await fetchTls('https://api.example/');
  assert.equal(connect.request?.headers.has('proxy-authorization'), false);

  const origin: ObservedBox = {};
  const plain: TransportCalls = { dials: [], secures: [] };
  const fetchPlain = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport([answering(origin, httpResponse('200 OK', [], ''))], plain),
  );
  await fetchPlain('http://origin.example/');
  assert.equal(origin.request?.headers.has('proxy-authorization'), false);
});

test('a plain-http POST uses absolute-form and keeps the body intact', async () => {
  const origin: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://user:pass@127.0.0.1:8080',
    [],
    scriptedTransport(
      [answering(origin, httpResponse('201 Created', ['Content-Length: 2'], 'ok'))],
      calls,
    ),
  );

  const response = await fetchVia('http://origin.example/submit?q=1', {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain', 'X-Test': '1' },
    body: 'hello-proxy',
  });
  assert.equal(response.status, 201);
  assert.equal(await response.text(), 'ok');
  assert.deepEqual(calls.dials, [{ host: '127.0.0.1', port: 8080 }]);
  assert.deepEqual(calls.secures, []);
  assert.equal(origin.request?.requestLine, 'POST http://origin.example/submit?q=1 HTTP/1.1');
  assert.equal(origin.request?.headers.get('host'), 'origin.example');
  assert.equal(origin.request?.headers.get('x-test'), '1');
  assert.equal(
    origin.request?.headers.get('proxy-authorization'),
    `Basic ${Buffer.from('user:pass').toString('base64')}`,
  );
  assert.equal(origin.request?.body.toString(), 'hello-proxy');
});

test('URLSearchParams and Buffer bodies serialize intact', async () => {
  for (const [body, expected] of [
    [new URLSearchParams({ a: '1', b: '2' }), 'a=1&b=2'],
    [Buffer.from([0, 255, 1]), null],
  ] as const) {
    const origin: ObservedBox = {};
    const calls: TransportCalls = { dials: [], secures: [] };
    const fetchVia = proxiedVia(
      'http://127.0.0.1:8080',
      [],
      scriptedTransport([answering(origin, httpResponse('200 OK', ['Content-Length: 0'], ''))], calls),
    );
    await fetchVia('http://origin.example/', { method: 'POST', body });
    if (expected === null) {
      assert.deepEqual(origin.request?.body, body as Buffer);
    } else {
      assert.equal(origin.request?.body.toString(), expected);
    }
  }
});

test('binary bodies survive intact through arrayBuffer', async () => {
  const payload = Buffer.from([0, 1, 2, 250, 251, 252, 253, 254, 255]);
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse(
            '200 OK',
            ['Content-Type: application/octet-stream', `Content-Length: ${payload.length}`],
            payload,
          ),
        ),
      ],
      calls,
    ),
  );

  const response = await fetchVia('https://api.example/file');
  assert.deepEqual(Buffer.from(await response.arrayBuffer()), payload);
});

test('a Request object input keeps its method, headers and body', async () => {
  const connect: ObservedBox = {};
  const origin: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [answering(connect, CONNECT_OK), answering(origin, httpResponse('200 OK', [], ''))],
      calls,
    ),
  );

  const request = new Request('https://api.example/items', {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: 'payload',
  });
  await fetchVia(request);
  assert.equal(origin.request?.requestLine, 'POST /items HTTP/1.1');
  assert.equal(origin.request?.headers.get('content-type'), 'text/plain');
  assert.equal(origin.request?.body.toString(), 'payload');
});

test('chunked and close-delimited bodies decode fully', async () => {
  const chunked: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchChunked = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering(chunked, CONNECT_OK),
        answering(
          {},
          Buffer.from(
            'HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n' +
              '5\r\nhello\r\n6\r\n world\r\n0\r\n\r\n',
          ),
        ),
      ],
      calls,
    ),
  );
  assert.equal(await (await fetchChunked('https://api.example/')).text(), 'hello world');

  const delimited: TransportCalls = { dials: [], secures: [] };
  const fetchDelimited = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering({}, Buffer.from('HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\n\r\ntrail')),
      ],
      delimited,
    ),
  );
  assert.equal(await (await fetchDelimited('https://api.example/')).text(), 'trail');
});

test('a 204 arrives with an empty body instead of throwing', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [answering({}, CONNECT_OK), answering({}, Buffer.from('HTTP/1.1 204 No Content\r\n\r\n'))],
      calls,
    ),
  );
  const response = await fetchVia('https://api.example/item', { method: 'DELETE' });
  assert.equal(response.status, 204);
  assert.equal(await response.text(), '');
});

test('interim 1xx responses are skipped for the final one', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          Buffer.from(
            'HTTP/1.1 103 Early Hints\r\nLink: </style.css>\r\n\r\n' +
              'HTTP/1.1 200 OK\r\nContent-Length: 2\r\n\r\nok',
          ),
        ),
      ],
      calls,
    ),
  );
  const response = await fetchVia('https://api.example/');
  assert.equal(response.status, 200);
  assert.equal(await response.text(), 'ok');
});

test('a same-origin redirect is followed with headers intact', async () => {
  const first: ObservedBox = {};
  const second: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(first, httpResponse('302 Found', ['Location: /new', 'Content-Length: 0'], '')),
        answering({}, CONNECT_OK),
        answering(second, httpResponse('200 OK', ['Content-Length: 2'], 'ok')),
      ],
      calls,
    ),
  );

  const response = await fetchVia('https://a.example/old', {
    headers: { Authorization: 'Bearer abc' },
  });
  assert.equal(response.status, 200);
  assert.equal(await response.text(), 'ok');
  assert.equal(first.request?.requestLine, 'GET /old HTTP/1.1');
  assert.equal(second.request?.requestLine, 'GET /new HTTP/1.1');
  assert.equal(second.request?.headers.get('authorization'), 'Bearer abc');
});

test('a 303 turns a POST into a GET without the body headers', async () => {
  const first: ObservedBox = {};
  const second: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          first,
          httpResponse('303 See Other', ['Location: https://a.example/done', 'Content-Length: 0'], ''),
        ),
        answering({}, CONNECT_OK),
        answering(second, httpResponse('200 OK', ['Content-Length: 2'], 'ok')),
      ],
      calls,
    ),
  );

  await fetchVia('https://a.example/submit', {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: 'data',
  });
  assert.equal(first.request?.requestLine, 'POST /submit HTTP/1.1');
  assert.equal(second.request?.requestLine, 'GET /done HTTP/1.1');
  assert.equal(second.request?.headers.has('content-length'), false);
  assert.equal(second.request?.headers.has('content-type'), false);
});

test('a 301 or 302 turns a POST into a GET without the body', async () => {
  for (const status of [301, 302]) {
    const first: ObservedBox = {};
    const second: ObservedBox = {};
    const calls: TransportCalls = { dials: [], secures: [] };
    const fetchVia = proxiedVia(
      'http://127.0.0.1:8080',
      [],
      scriptedTransport(
        [
          answering({}, CONNECT_OK),
          answering(
            first,
            httpResponse(
              `${status} Redirect`,
              ['Location: https://a.example/done', 'Content-Length: 0'],
              '',
            ),
          ),
          answering({}, CONNECT_OK),
          answering(second, httpResponse('200 OK', ['Content-Length: 2'], 'ok')),
        ],
        calls,
      ),
    );

    await fetchVia('https://a.example/submit', {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain', Authorization: 'Bearer test-token' },
      body: 'data',
    });
    assert.equal(first.request?.requestLine, 'POST /submit HTTP/1.1');
    assert.equal(second.request?.requestLine, 'GET /done HTTP/1.1');
    assert.equal(second.request?.body.length, 0);
    assert.equal(second.request?.headers.has('content-length'), false);
    assert.equal(second.request?.headers.has('content-type'), false);
    // Same origin, so the credential survives the rewrite.
    assert.equal(second.request?.headers.get('authorization'), 'Bearer test-token');
  }
});

test('a 307 or 308 replays the POST with method, body and headers', async () => {
  for (const status of [307, 308]) {
    const second: ObservedBox = {};
    const calls: TransportCalls = { dials: [], secures: [] };
    const fetchVia = proxiedVia(
      'http://127.0.0.1:8080',
      [],
      scriptedTransport(
        [
          answering({}, CONNECT_OK),
          answering(
            {},
            httpResponse(
              `${status} Redirect`,
              ['Location: https://a.example/done', 'Content-Length: 0'],
              '',
            ),
          ),
          answering({}, CONNECT_OK),
          answering(second, httpResponse('200 OK', ['Content-Length: 2'], 'ok')),
        ],
        calls,
      ),
    );

    await fetchVia('https://a.example/submit', {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain', Authorization: 'Bearer test-token' },
      body: 'data',
    });
    assert.equal(second.request?.requestLine, 'POST /done HTTP/1.1');
    assert.equal(second.request?.body.toString(), 'data');
    assert.equal(second.request?.headers.get('content-type'), 'text/plain');
    assert.equal(second.request?.headers.get('authorization'), 'Bearer test-token');
  }
});

test('a 303 leaves HEAD alone', async () => {
  const second: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse('303 See Other', ['Location: https://a.example/done', 'Content-Length: 0'], ''),
        ),
        answering({}, CONNECT_OK),
        answering(second, httpResponse('200 OK', ['Content-Length: 0'], '')),
      ],
      calls,
    ),
  );

  await fetchVia('https://a.example/submit', { method: 'HEAD' });
  assert.equal(second.request?.requestLine, 'HEAD /done HTTP/1.1');
});

test('a cross-origin redirect strips authorization', async () => {
  const second: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse('302 Found', ['Location: https://b.example/', 'Content-Length: 0'], ''),
        ),
        answering({}, CONNECT_OK),
        answering(second, httpResponse('200 OK', ['Content-Length: 2'], 'ok')),
      ],
      calls,
    ),
  );

  await fetchVia('https://a.example/', { headers: { Authorization: 'Bearer abc' } });
  assert.equal(second.request?.headers.has('authorization'), false);
});

test('a redirect to a different port strips authorization too', async () => {
  const second: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse('302 Found', ['Location: https://a.example:8443/next', 'Content-Length: 0'], ''),
        ),
        answering({}, CONNECT_OK),
        answering(second, httpResponse('200 OK', ['Content-Length: 2'], 'ok')),
      ],
      calls,
    ),
  );

  // Same host, but the origin — and therefore the trust boundary — changed.
  await fetchVia('https://a.example/', { headers: { Authorization: 'Bearer test-token' } });
  assert.deepEqual(calls.secures, ['a.example', 'a.example']);
  assert.equal(second.request?.requestLine, 'GET /next HTTP/1.1');
  assert.equal(second.request?.headers.has('authorization'), false);
});

test('redirect loops fail instead of spinning forever', async () => {
  const responders: Array<(received: Buffer) => ScriptResult> = [];
  for (let i = 0; i < 21; i += 1) {
    responders.push(answering({}, CONNECT_OK));
    responders.push(
      answering(
        {},
        httpResponse('302 Found', ['Location: https://loop.example/round', 'Content-Length: 0'], ''),
      ),
    );
  }
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia('http://127.0.0.1:8080', [], scriptedTransport(responders, calls));
  await assert.rejects(fetchVia('https://loop.example/round'), /too many redirects/);
  assert.equal(calls.dials.length, 21);
});

test('manual redirect mode returns the redirect itself', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse('302 Found', ['Location: /new', 'Content-Length: 0'], ''),
        ),
      ],
      calls,
    ),
  );
  const response = await fetchVia('https://a.example/old', { redirect: 'manual' });
  assert.equal(response.status, 302);
  assert.equal(response.headers.get('location'), '/new');
  assert.equal(await response.text(), '');
  assert.equal(calls.dials.length, 1);
});

test('fetch-binary range reads ride the proxy: manual 302 plus binary body', async () => {
  const graph: ObservedBox = {};
  const storage: ObservedBox = {};
  const calls: TransportCalls = { dials: [], secures: [] };
  const payload = Buffer.from([0, 255, 1, 2, 159]);
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          graph,
          httpResponse(
            '302 Found',
            ['Location: https://storage.example.net/dl?sig=abc', 'Content-Length: 0'],
            '',
          ),
        ),
        answering({}, CONNECT_OK),
        answering(
          storage,
          httpResponse(
            '206 Partial Content',
            ['Content-Range: bytes 10-14/100', `Content-Length: ${payload.length}`],
            payload,
          ),
        ),
      ],
      calls,
    ),
  );

  const graphConfig: Pick<ServerConfig, 'graphHost' | 'graphVersion'> = {
    graphHost: 'graph.microsoft.com',
    graphVersion: 'v1.0',
  };
  const stubAuth: AuthProvider = {
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

  const result = await fetchContentRange(stubAuth, graphConfig, {
    path: '/drives/D/items/I/content',
    scopes: ['Files.Read.All'],
    offset: 10,
    byteBudget: 5,
    fetchImpl: fetchVia,
  });

  assert.deepEqual(Buffer.from(result.bytes), payload);
  assert.equal(result.nextOffset, 15);
  assert.equal(result.totalLength, 100);
  assert.equal(result.truncated, true);
  assert.equal(graph.request?.requestLine, 'GET /v1.0/drives/D/items/I/content HTTP/1.1');
  assert.equal(graph.request?.headers.get('range'), 'bytes=10-14');
  // The storage hop re-issues Range but must NOT forward the bearer.
  assert.equal(storage.request?.headers.get('range'), 'bytes=10-14');
  assert.equal(storage.request?.headers.has('authorization'), false);
});

test('an unparseable Location returns the redirect as-is', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse('302 Found', ['Location: http://exa mple', 'Content-Length: 4'], 'lost'),
        ),
      ],
      calls,
    ),
  );
  const response = await fetchVia('https://a.example/old');
  assert.equal(response.status, 302);
  assert.equal(await response.text(), 'lost');
  assert.equal(calls.dials.length, 1);
});

test('error redirect mode rejects instead of following', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse('302 Found', ['Location: /new', 'Content-Length: 0'], ''),
        ),
      ],
      calls,
    ),
  );
  await assert.rejects(fetchVia('https://a.example/old', { redirect: 'error' }), /redirect/);
});

test('a redirect onto a NO_PROXY host hands back to the native fetch', async () => {
  // The handoff target must be the true native fetch, not this module's wrapper.
  uninstallProxyFetch();
  const native = globalThis.fetch;
  const seen: string[] = [];
  const nativeFetch: typeof fetch = ((input: unknown, init?: unknown) => {
    seen.push(String((input as { url?: unknown })?.url ?? input));
    return native(input as never, init as never);
  }) as typeof fetch;
  const calls: TransportCalls = { dials: [], secures: [] };
  const closed = await findClosedPort();
  const fetchVia = proxiedVia(
    `http://127.0.0.1:${closed.port}`,
    ['127.0.0.1'],
    scriptedTransport(
      [
        answering({}, CONNECT_OK),
        answering(
          {},
          httpResponse(
            '302 Found',
            [`Location: http://127.0.0.1:${closed.port}/direct`, 'Content-Length: 0'],
            '',
          ),
        ),
      ],
      calls,
    ),
    nativeFetch,
  );

  // The handoff itself is the assertion: native fetch dials direct and refuses.
  await assert.rejects(fetchVia('https://a.example/old'), (error: unknown) =>
    isFetchFailedWith(error, closed.code, closed.port),
  );
  assert.deepEqual(seen, [`http://127.0.0.1:${closed.port}/direct`]);
});

test('a CONNECT refusal fails as a TypeError carrying the status', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport(
      [answering({}, Buffer.from('HTTP/1.1 407 Proxy Authentication Required\r\n\r\n'))],
      calls,
    ),
  );
  await assert.rejects(
    fetchVia('https://api.example/'),
    (error: unknown) => {
      assert.ok(error instanceof TypeError);
      assert.match(String((error as { cause?: unknown }).cause), /407/);
      return true;
    },
  );
  assert.deepEqual(calls.secures, []);
});

test('an already-aborted signal rejects without dialing', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia(
    'http://127.0.0.1:8080',
    [],
    scriptedTransport([], calls),
  );
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(fetchVia('https://api.example/', { signal: controller.signal }), {
    name: 'AbortError',
  });
  assert.deepEqual(calls.dials, []);
});

test('aborting mid-request rejects without waiting for the proxy', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  // A responder that never answers: the read hangs until the abort lands.
  const fetchVia = proxiedVia('http://127.0.0.1:8080', [], scriptedTransport([() => null], calls));
  const controller = new AbortController();
  const pending = fetchVia('https://api.example/', { signal: controller.signal });
  setImmediate(() => controller.abort());
  await assert.rejects(pending, { name: 'AbortError' });
  assert.equal(calls.dials.length, 1);
});

test('a mid-request abort reports the signal reason', async () => {
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia('http://127.0.0.1:8080', [], scriptedTransport([() => null], calls));
  const controller = new AbortController();
  const pending = fetchVia('https://api.example/', { signal: controller.signal });
  const reason = new Error('stop now');
  setImmediate(() => controller.abort(reason));
  await assert.rejects(pending, (error: unknown) => {
    assert.ok(error === reason);
    return true;
  });
});

test('non-http(s) urls and unparseable urls reach the native fetch', async () => {
  const seen: unknown[] = [];
  const nativeFetch: typeof fetch = ((input: unknown, init?: unknown) => {
    seen.push(input);
    return Promise.resolve(new Response('native'));
  }) as typeof fetch;
  const calls: TransportCalls = { dials: [], secures: [] };
  const fetchVia = proxiedVia('http://127.0.0.1:8080', [], scriptedTransport([], calls), nativeFetch);

  const response = await fetchVia('data:text/plain,hi');
  assert.equal(await response.text(), 'native');
  assert.deepEqual(calls.dials, []);
  assert.equal(seen.length, 1);
});

// ---------------------------------------------------------------------------
// Global install, activation log, interception (no listen)
// ---------------------------------------------------------------------------

test('install leaves the global fetch untouched when no proxy is configured', () => {
  uninstallProxyFetch();
  try {
    assert.equal(isProxyFetchInstalled(), false);
    const before = globalThis.fetch;
    assert.equal(installProxyFetch({}), false);
    assert.equal(installProxyFetch({ HTTP_PROXY: 'http://127.0.0.1:1' }), false);
    assert.equal(installProxyFetch({ HTTPS_PROXY: 'http://exa mple' }), false);
    assert.ok(globalThis.fetch === before);
    assert.equal(isProxyFetchInstalled(), false);
  } finally {
    uninstallProxyFetch();
  }
});

test('install logs one redacted line on activation, and only once', () => {
  uninstallProxyFetch();
  const writes: string[] = [];
  const stderr = process.stderr as unknown as { write: (chunk: string) => boolean };
  const originalWrite = stderr.write;
  stderr.write = (chunk: string): boolean => {
    writes.push(String(chunk));
    return true;
  };
  try {
    assert.equal(installProxyFetch({ HTTPS_PROXY: 'http://user:s3cret@127.0.0.1:9' }), true);
    assert.equal(isProxyFetchInstalled(), true);
    // Idempotent: neither re-wraps nor re-logs.
    assert.equal(installProxyFetch({ HTTPS_PROXY: 'http://other:10' }), true);
  } finally {
    uninstallProxyFetch();
    stderr.write = originalWrite;
  }
  assert.equal(writes.length, 1);
  assert.ok(writes[0]?.includes('127.0.0.1:9'));
  assert.ok(!writes[0]?.includes('s3cret'));
  assert.ok(!writes[0]?.includes('user'));
});

test('install warns and stays native on a malformed proxy URL', () => {
  uninstallProxyFetch();
  const writes: string[] = [];
  const stderr = process.stderr as unknown as { write: (chunk: string) => boolean };
  const originalWrite = stderr.write;
  stderr.write = (chunk: string): boolean => {
    writes.push(String(chunk));
    return true;
  };
  try {
    const before = globalThis.fetch;
    // Neither a garbage URL nor an unspeakable scheme installs or throws.
    assert.equal(installProxyFetch({ HTTPS_PROXY: 'http://exa mple' }), false);
    assert.ok(globalThis.fetch === before);
    assert.equal(isProxyFetchInstalled(), false);
    assert.equal(installProxyFetch({ ALL_PROXY: 'socks5://127.0.0.1:1080' }), false);
    assert.ok(globalThis.fetch === before);
    assert.equal(isProxyFetchInstalled(), false);
  } finally {
    uninstallProxyFetch();
    stderr.write = originalWrite;
  }
  assert.equal(writes.length, 2);
  assert.ok(writes[0]?.includes('WARN'));
  assert.ok(writes[0]?.includes('$HTTPS_PROXY'));
  assert.ok(writes[1]?.includes('$ALL_PROXY'));
  for (const line of writes) {
    // The variable is named; its value — possibly credentialed — never is.
    assert.ok(!line.includes('exa mple'));
    assert.ok(!line.includes('socks5'));
  }
});

test('a proxied https request dials the proxy, not the target', async () => {
  uninstallProxyFetch();
  try {
    const proxy = await findClosedPort();
    assert.equal(installProxyFetch({ HTTPS_PROXY: `http://127.0.0.1:${proxy.port}` }), true);
    // Direct DNS for an .invalid host would be ENOTFOUND; the proxy dial
    // refuses first, proving the request was intercepted.
    await assert.rejects(
      fetch('https://proxied.example/', { signal: AbortSignal.timeout(10000) }),
      (error: unknown) => isFetchFailedWith(error, proxy.code, proxy.port),
    );
  } finally {
    uninstallProxyFetch();
  }
});

test('a NO_PROXY match bypasses the proxy for the real target', async () => {
  uninstallProxyFetch();
  try {
    const proxy = await findClosedPort();
    const target = await findClosedPort([proxy.port]);
    installProxyFetch({
      HTTPS_PROXY: `http://127.0.0.1:${proxy.port}`,
      NO_PROXY: '127.0.0.1',
    });
    // Bypassed: native fetch dials the target itself, refusing on ITS port.
    await assert.rejects(
      fetch(`http://127.0.0.1:${target.port}/x`, { signal: AbortSignal.timeout(10000) }),
      (error: unknown) => isFetchFailedWith(error, target.code, target.port),
    );
    // Not bypassed: the proxy dial refuses on the PROXY port instead.
    await assert.rejects(
      fetch('https://proxied.example/', { signal: AbortSignal.timeout(10000) }),
      (error: unknown) => isFetchFailedWith(error, proxy.code, proxy.port),
    );
  } finally {
    uninstallProxyFetch();
  }
});

// ---------------------------------------------------------------------------
// Stub TCP proxy over real loopback sockets
// ---------------------------------------------------------------------------

async function startStub(
  onConnection: (sock: net.Socket) => void,
): Promise<{ port: number; close: () => Promise<void> }> {
  const server = net.createServer(onConnection);
  server.on('error', () => {});
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolve());
  });
  const address = server.address();
  assert.ok(typeof address === 'object' && address !== null, 'expected an Address');
  return {
    port: address.port,
    close: () =>
      new Promise<void>((resolve, reject) => {
        server.close((error) => (error === undefined ? resolve() : reject(error)));
      }),
  };
}

function readFullRequest(sock: net.Socket): Promise<ObservedRequest> {
  return new Promise((resolve, reject) => {
    let acc = Buffer.alloc(0);
    const onData = (chunk: Buffer): void => {
      acc = Buffer.concat([acc, chunk]);
      const request = parseObservedRequest(acc);
      if (request === null) return;
      cleanup();
      resolve(request);
    };
    const onError = (error: Error): void => {
      cleanup();
      reject(error);
    };
    const onClose = (): void => {
      cleanup();
      reject(new Error('test stub: closed before the request completed'));
    };
    const cleanup = (): void => {
      sock.off('data', onData);
      sock.off('error', onError);
      sock.off('close', onClose);
    };
    sock.on('data', onData);
    sock.on('error', onError);
    sock.on('close', onClose);
  });
}

/** Fails the wait instead of hanging the suite when the stub never reports. */
async function withTimeout<T>(promise: Promise<T>, ms: number, message: string): Promise<T> {
  let timer: NodeJS.Timeout | undefined;
  try {
    return await Promise.race([
      promise,
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => reject(new Error(message)), ms);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

test('stub proxy: an http POST arrives in absolute-form with headers and body', async (t) => {
  uninstallProxyFetch();
  let observed: ObservedRequest | undefined;
  let stub: { port: number; close: () => Promise<void> };
  try {
    stub = await startStub((sock) => {
      void (async () => {
        try {
          observed = await readFullRequest(sock);
          sock.end(
            'HTTP/1.1 201 Created\r\nContent-Type: text/plain\r\nContent-Length: 2\r\n' +
              'X-Stub: yes\r\nConnection: close\r\n\r\nok',
          );
        } catch {
          sock.destroy();
        }
      })();
    });
  } catch (error) {
    if ((error as NodeJS.ErrnoException)?.code === 'EPERM') {
      t.skip('this runner cannot bind loopback at all (sandboxed)');
      return;
    }
    throw error;
  }
  try {
    installProxyFetch({ HTTPS_PROXY: `http://127.0.0.1:${stub.port}` });
    const response = await fetch('http://origin.example/submit?q=1', {
      method: 'POST',
      headers: { 'X-Test': 'forward-me' },
      body: 'hello-proxy',
      signal: AbortSignal.timeout(10000),
    });
    assert.equal(response.status, 201);
    assert.equal(await response.text(), 'ok');
    assert.equal(response.headers.get('x-stub'), 'yes');
    assert.ok(observed !== undefined);
    assert.equal(observed.requestLine, 'POST http://origin.example/submit?q=1 HTTP/1.1');
    assert.equal(observed.headers.get('host'), 'origin.example');
    assert.equal(observed.headers.get('x-test'), 'forward-me');
    assert.equal(observed.body.toString(), 'hello-proxy');
  } finally {
    uninstallProxyFetch();
    await stub.close();
  }
});

test('stub proxy: an https CONNECT names the target and precedes TLS', async (t) => {
  uninstallProxyFetch();
  interface Report {
    connect?: ObservedRequest;
    helloFirstByte?: number;
  }
  let report!: (value: Report) => void;
  const reported = new Promise<Report>((resolve) => {
    report = resolve;
  });
  let stub: { port: number; close: () => Promise<void> };
  try {
    stub = await startStub((sock) => {
      void (async () => {
        try {
          const connect = await readFullRequest(sock);
          sock.write('HTTP/1.1 200 Connection Established\r\n\r\n');
          // The next bytes must be a TLS ClientHello, not plain HTTP.
          const hello = await new Promise<Buffer>((resolve, reject) => {
            sock.once('data', (chunk: Buffer) => resolve(chunk));
            sock.once('error', reject);
            sock.once('close', () => reject(new Error('closed before ClientHello')));
          });
          report({ connect, helloFirstByte: hello[0] });
        } catch {
          report({});
        } finally {
          sock.destroy();
        }
      })();
    });
  } catch (error) {
    if ((error as NodeJS.ErrnoException)?.code === 'EPERM') {
      t.skip('this runner cannot bind loopback at all (sandboxed)');
      return;
    }
    throw error;
  }
  try {
    installProxyFetch({ HTTPS_PROXY: `http://user:pass@127.0.0.1:${stub.port}` });
    // The stub tears the tunnel down after the ClientHello, so the fetch
    // itself must fail — the CONNECT observations are the assertion.
    await assert.rejects(
      fetch('https://tls.example:8443/x', { signal: AbortSignal.timeout(10000) }),
      (error: unknown) => {
        assert.ok(error instanceof TypeError);
        return true;
      },
    );
    const got = await withTimeout(reported, 10000, 'test stub: no CONNECT observed');
    assert.equal(got.connect?.requestLine, 'CONNECT tls.example:8443 HTTP/1.1');
    assert.equal(
      got.connect?.headers.get('proxy-authorization'),
      `Basic ${Buffer.from('user:pass').toString('base64')}`,
    );
    assert.equal(got.helloFirstByte, 0x16);
  } finally {
    uninstallProxyFetch();
    await stub.close();
  }
});
