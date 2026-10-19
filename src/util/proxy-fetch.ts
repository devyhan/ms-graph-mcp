/**
 * Proxy-aware `fetch` for the whole process.
 *
 * Node's `fetch` ignores proxy environment variables, so behind a proxy every
 * Entra/Graph call fails with ENOTFOUND while curl and npm work. MSAL, the
 * graph client and the update check all use the global `fetch`, so wrapping it
 * once here routes all three. `index.ts` imports this module first so the
 * wrapper is in place before any of them evaluates.
 *
 * HTTPS targets are tunnelled with CONNECT (plus `Proxy-Authorization` when
 * the proxy URL carries credentials); plain-HTTP targets are sent in
 * absolute-form. Anything else — no proxy env, a NO_PROXY match, a
 * non-HTTP(S) URL — falls through to the native fetch untouched.
 */

import net from 'node:net';
import tls from 'node:tls';
import type { HeadersInit, RequestRedirect } from 'undici-types';

import { logger } from './logger.js';

/** Env vars consulted for the proxy URL, in precedence order. */
const PROXY_VARS = ['HTTPS_PROXY', 'https_proxy', 'ALL_PROXY', 'all_proxy'] as const;

/** Statuses that carry a `Location` worth following. */
const REDIRECT_STATUSES: ReadonlySet<number> = new Set([301, 302, 303, 307, 308]);

/** Native fetch follows 20 redirects before giving up; match that. */
const MAX_REDIRECTS = 20;

/** Statuses whose response must carry a null body. */
const NULL_BODY_STATUSES: ReadonlySet<number> = new Set([204, 304]);

export interface ProxyConfig {
  /** The proxy URL. Never logged verbatim: it may carry credentials. */
  url: URL;
  /** Which env var it came from, e.g. `HTTPS_PROXY`. */
  source: string;
  /** Lowercased NO_PROXY rules. */
  noProxy: string[];
}

/**
 * Resolves the proxy configuration from the environment, or undefined when
 * none is set. A malformed URL is skipped rather than fatal: crashing the
 * process at import time over a typo'd env var would be worse than going
 * direct, and a bare `host:port` gains the `http://` scheme curl would assume.
 * A value that parses but cannot serve a CONNECT tunnel — an empty host, or a
 * scheme this module does not speak such as `socks5:` — is skipped the same
 * way; `installProxyFetch` warns about the leftovers.
 */
export function resolveProxyConfig(env: NodeJS.ProcessEnv = process.env): ProxyConfig | undefined {
  let url: URL | undefined;
  let source = '';
  for (const name of PROXY_VARS) {
    const raw = env[name]?.trim();
    if (raw === undefined || raw === '') continue;
    const withScheme = /^[a-z][a-z0-9+.-]*:\/\//i.test(raw) ? raw : `http://${raw}`;
    let candidate: URL;
    try {
      candidate = new URL(withScheme);
    } catch {
      continue;
    }
    if (candidate.hostname === '' || (candidate.protocol !== 'http:' && candidate.protocol !== 'https:')) {
      continue;
    }
    url = candidate;
    source = name;
    break;
  }
  if (url === undefined) return undefined;
  const noProxy = (env['NO_PROXY'] ?? env['no_proxy'] ?? '')
    .split(',')
    .map((rule) => rule.trim().toLowerCase())
    .filter((rule) => rule !== '');
  return { url, source, noProxy };
}

/**
 * NO_PROXY matching: `*` bypasses everything, otherwise an exact host match or
 * a domain-suffix match (`example.com` covers `api.example.com`; a leading
 * dot on the rule is accepted and ignored).
 */
export function isBypassed(hostname: string, rules: readonly string[]): boolean {
  const host = hostname.toLowerCase();
  return rules.some((rule) => {
    const normalized = rule.toLowerCase();
    if (normalized === '*') return true;
    // A leading dot is accepted and ignored: `.example.com` covers the bare
    // domain as well as its subdomains, as in curl.
    const domain = normalized.replace(/^\./, '');
    if (domain === '') return false;
    return host === domain || host.endsWith(`.${domain}`);
  });
}

/**
 * Normalizes a fetch body to a Buffer. Covers what this codebase sends
 * (strings) plus the binary shapes fetch accepts; anything else throws rather
 * than silently sending the wrong bytes.
 */
export function toBodyBuffer(body: unknown): Buffer | undefined {
  if (body === undefined || body === null) return undefined;
  if (typeof body === 'string') return Buffer.from(body);
  if (body instanceof URLSearchParams) return Buffer.from(body.toString());
  if (Buffer.isBuffer(body)) return body;
  if (body instanceof ArrayBuffer) return Buffer.from(body);
  if (ArrayBuffer.isView(body)) return Buffer.from(body.buffer, body.byteOffset, body.byteLength);
  throw new TypeError(`proxy-fetch: unsupported body type ${Object.prototype.toString.call(body)}`);
}

/** Parses an HTTP status line plus headers into a status and a Headers. */
export function parseStatusAndHeaders(head: string): { status: number; headers: Headers } {
  const lines = head.split('\r\n');
  const statusLine = lines.shift() ?? '';
  const status = /^HTTP\/\d(?:\.\d)?\s+(\d{3})/.exec(statusLine)?.[1];
  if (status === undefined) throw new Error(`proxy-fetch: bad status line: ${statusLine}`);
  const headers = new Headers();
  for (const line of lines) {
    const separator = line.indexOf(':');
    if (separator === -1) continue;
    headers.append(line.slice(0, separator).trim(), line.slice(separator + 1).trim());
  }
  return { status: Number(status), headers };
}

/** Decodes one complete chunked body; extensions and trailers are ignored. */
export function decodeChunked(raw: Buffer): Buffer {
  const out: Buffer[] = [];
  let pos = 0;
  for (;;) {
    const eol = raw.indexOf('\r\n', pos);
    if (eol === -1) break;
    const size = Number.parseInt(
      raw.subarray(pos, eol).toString('ascii').split(';')[0]?.trim() ?? '',
      16,
    );
    if (!Number.isFinite(size) || size <= 0) break;
    pos = eol + 2;
    out.push(raw.subarray(pos, pos + size));
    pos += size + 2;
  }
  return Buffer.concat(out);
}

/** Rejection for an aborted request: the signal's reason, as native fetch does. */
function toAbortRejection(signal: AbortSignal | undefined): unknown {
  const reason = signal?.reason;
  if (reason !== undefined) return reason;
  return new DOMException('This operation was aborted', 'AbortError');
}

/** A socket failure during an abort reports the abort, not the torn socket. */
function abortedOr(error: Error, signal: AbortSignal | undefined): unknown {
  return signal?.aborted === true ? toAbortRejection(signal) : error;
}

/** Reads up to the blank line ending the head; `prefix` carries pipelined bytes. */
function readHead(
  sock: net.Socket,
  signal: AbortSignal | undefined,
  prefix: Buffer = Buffer.alloc(0),
): Promise<{ head: string; rest: Buffer }> {
  const at = prefix.indexOf('\r\n\r\n');
  if (at !== -1) {
    return Promise.resolve({
      head: prefix.subarray(0, at).toString('latin1'),
      rest: prefix.subarray(at + 4),
    });
  }
  signal?.throwIfAborted();
  return new Promise((resolve, reject) => {
    let acc = prefix;
    const onData = (chunk: Buffer): void => {
      acc = Buffer.concat([acc, chunk]);
      const idx = acc.indexOf('\r\n\r\n');
      if (idx !== -1) {
        cleanup();
        resolve({ head: acc.subarray(0, idx).toString('latin1'), rest: acc.subarray(idx + 4) });
        return;
      }
      settleIfExhausted();
    };
    const onError = (error: Error): void => {
      cleanup();
      reject(abortedOr(error, signal));
    };
    const onClose = (): void => {
      cleanup();
      reject(
        abortedOr(new Error('proxy-fetch: connection closed before the response arrived'), signal),
      );
    };
    const settleIfExhausted = (): void => {
      // `end`/`close` never re-fire, and a coalesced FIN lands on the next
      // tick — ahead of the microtask that attaches these listeners — so an
      // already-ended stream would hang the read. The state flags are the
      // truth then; buffered bytes still flow through `data` first.
      if (sock.readableLength > 0 || (!sock.readableEnded && !sock.destroyed)) return;
      cleanup();
      reject(
        abortedOr(new Error('proxy-fetch: connection closed before the response arrived'), signal),
      );
    };
    const cleanup = (): void => {
      sock.off('data', onData);
      sock.off('error', onError);
      sock.off('close', onClose);
    };
    sock.on('data', onData);
    sock.on('error', onError);
    sock.on('close', onClose);
    settleIfExhausted();
  });
}

interface BodyOptions {
  method: string;
  signal: AbortSignal | undefined;
}

/**
 * Reads the response body per its framing: Content-Length, chunked, or
 * close-delimited (every request sends `Connection: close`, so EOF always
 * ends the body). A short close resolves for HEAD — which by definition
 * receives no bytes — and fails otherwise, since partial bytes presented as
 * complete would corrupt the caller silently.
 */
function readBody(
  sock: net.Socket,
  headers: Headers,
  rest: Buffer,
  options: BodyOptions,
): Promise<Buffer> {
  const { method, signal } = options;
  signal?.throwIfAborted();
  const length = headers.get('content-length');
  const chunked = (headers.get('transfer-encoding') ?? '').toLowerCase().includes('chunked');
  if (!chunked && length !== null && /^\d+$/.test(length.trim())) {
    const need = Number(length.trim());
    const chunks = [rest];
    let have = rest.length;
    if (have >= need) return Promise.resolve(Buffer.concat(chunks).subarray(0, need));
    return new Promise<void>((resolve, reject) => {
      const onData = (chunk: Buffer): void => {
        chunks.push(chunk);
        have += chunk.length;
        if (have >= need) {
          cleanup();
          resolve();
          return;
        }
        settleIfExhausted();
      };
      const onError = (error: Error): void => {
        cleanup();
        reject(abortedOr(error, signal));
      };
      const onEnded = (): void => {
        cleanup();
        if (signal?.aborted === true) {
          reject(toAbortRejection(signal));
        } else if (method === 'HEAD') {
          resolve();
        } else {
          reject(
            new Error(`proxy-fetch: connection closed with ${need - have} of ${need} bytes left`),
          );
        }
      };
      // See readHead: an `end` that fired before these listeners attached
      // never re-fires, so consult the state flags as well as the events.
      const settleIfExhausted = (): void => {
        if (sock.readableLength > 0 || (!sock.readableEnded && !sock.destroyed)) return;
        onEnded();
      };
      const cleanup = (): void => {
        sock.off('data', onData);
        sock.off('error', onError);
        sock.off('end', onEnded);
        sock.off('close', onEnded);
      };
      sock.on('data', onData);
      sock.on('error', onError);
      sock.on('end', onEnded);
      sock.on('close', onEnded);
      settleIfExhausted();
    }).then(() => Buffer.concat(chunks).subarray(0, need));
  }
  const chunks = [rest];
  return new Promise<void>((resolve, reject) => {
    const onData = (chunk: Buffer): void => {
      chunks.push(chunk);
      settleIfExhausted();
    };
    const onError = (error: Error): void => {
      cleanup();
      reject(abortedOr(error, signal));
    };
    const onEnded = (): void => {
      cleanup();
      if (signal?.aborted === true) reject(toAbortRejection(signal));
      else resolve();
    };
    const settleIfExhausted = (): void => {
      if (sock.readableLength > 0 || (!sock.readableEnded && !sock.destroyed)) return;
      onEnded();
    };
    const cleanup = (): void => {
      sock.off('data', onData);
      sock.off('error', onError);
      sock.off('end', onEnded);
      sock.off('close', onEnded);
    };
    sock.on('data', onData);
    sock.on('error', onError);
    sock.on('end', onEnded);
    sock.on('close', onEnded);
    settleIfExhausted();
  }).then(() => {
    const raw = Buffer.concat(chunks);
    return chunked ? decodeChunked(raw) : raw;
  });
}

function dialTcp(host: string, port: number): Promise<net.Socket> {
  return new Promise((resolve, reject) => {
    const sock = net.connect(port, host);
    sock.once('connect', () => resolve(sock));
    sock.once('error', reject);
  });
}

function secureWithTls(sock: net.Socket, servername: string): Promise<net.Socket> {
  return new Promise((resolve, reject) => {
    const tlsSock = tls.connect({ socket: sock, servername }, () => resolve(tlsSock));
    tlsSock.once('error', (error: Error) => {
      tlsSock.destroy();
      reject(error);
    });
  });
}

/** Transport seam: real sockets in production, scripted ones in tests. */
export interface ProxyTransport {
  dial: (host: string, port: number) => Promise<net.Socket>;
  secure: (sock: net.Socket, servername: string) => Promise<net.Socket>;
}

const defaultTransport: ProxyTransport = { dial: dialTcp, secure: secureWithTls };

/** `Basic` credentials for `Proxy-Authorization`, when the proxy URL has them. */
function proxyCredentials(proxy: URL): string | undefined {
  if (proxy.username === '') return undefined;
  let userinfo: string;
  try {
    userinfo = `${decodeURIComponent(proxy.username)}:${decodeURIComponent(proxy.password)}`;
  } catch {
    userinfo = `${proxy.username}:${proxy.password}`;
  }
  return Buffer.from(userinfo).toString('base64');
}

/**
 * One request/response exchange through the proxy: CONNECT plus a TLS upgrade
 * for HTTPS, absolute-form for plain HTTP. Single-use sockets with
 * `Connection: close`, so framing never spans requests.
 */
async function singleHop(
  url: string,
  method: string,
  headers: HeadersInit,
  body: Buffer | undefined,
  signal: AbortSignal | undefined,
  proxy: URL,
  transport: ProxyTransport,
): Promise<{ status: number; headers: Headers; body: Buffer }> {
  signal?.throwIfAborted();
  const target = new URL(url);
  const isTls = target.protocol === 'https:';
  const targetPort = target.port !== '' ? Number(target.port) : isTls ? 443 : 80;
  const credentials = proxyCredentials(proxy);
  let sock = await transport.dial(proxy.hostname, Number(proxy.port) || 80);
  // The closure reads the current binding, so after the TLS upgrade it
  // destroys the new socket rather than the wrapped one.
  const onAbort = (): void => {
    sock.destroy();
  };
  if (signal !== undefined) {
    if (signal.aborted) {
      sock.destroy();
      throw toAbortRejection(signal);
    }
    signal.addEventListener('abort', onAbort, { once: true });
  }
  try {
    let path = `${target.pathname}${target.search}`;
    if (isTls) {
      let connect = `CONNECT ${target.hostname}:${targetPort} HTTP/1.1\r\n`;
      connect += `Host: ${target.hostname}:${targetPort}\r\n`;
      if (credentials !== undefined) connect += `Proxy-Authorization: Basic ${credentials}\r\n`;
      sock.write(`${connect}\r\n`);
      const { head } = await readHead(sock, signal);
      const { status } = parseStatusAndHeaders(head);
      if (status !== 200) {
        throw new Error(
          `proxy-fetch: CONNECT to ${target.hostname}:${targetPort} failed with ${status}`,
        );
      }
      sock = await transport.secure(sock, target.hostname);
    } else {
      path = url;
    }
    const out = new Headers(headers);
    out.set('Host', target.host);
    out.set('Connection', 'close');
    if (!isTls && credentials !== undefined && !out.has('Proxy-Authorization')) {
      out.set('Proxy-Authorization', `Basic ${credentials}`);
    }
    if (body !== undefined && !out.has('Content-Length')) {
      out.set('Content-Length', String(body.length));
    }
    let head = `${method} ${path} HTTP/1.1\r\n`;
    out.forEach((value, key) => {
      head += `${key}: ${value}\r\n`;
    });
    sock.write(`${head}\r\n`);
    if (body !== undefined) sock.write(body);
    // Interim 1xx responses carry no body; the final response follows.
    let response = await readHead(sock, signal);
    let parsed = parseStatusAndHeaders(response.head);
    while (parsed.status >= 100 && parsed.status < 200) {
      response = await readHead(sock, signal, response.rest);
      parsed = parseStatusAndHeaders(response.head);
    }
    const responseBody = await readBody(sock, parsed.headers, response.rest, { method, signal });
    return { status: parsed.status, headers: parsed.headers, body: responseBody };
  } finally {
    signal?.removeEventListener('abort', onAbort);
    sock.destroy();
  }
}

export interface ProxiedFetchOptions {
  proxy: URL;
  noProxy: readonly string[];
  nativeFetch: typeof fetch;
  transport?: ProxyTransport;
}

/**
 * Builds the proxying fetch. Redirects are followed like the native fetch
 * (method rewrite, credential stripping, `manual`/`error` policies); a hop
 * that lands outside the proxy — a NO_PROXY match or a non-HTTP(S) target —
 * is handed back to the native fetch with the redirect chain intact.
 */
export function createProxiedFetch(options: ProxiedFetchOptions): typeof fetch {
  const { proxy, noProxy, nativeFetch } = options;
  const transport = options.transport ?? defaultTransport;
  const proxied: typeof fetch = async (input, init = {}) => {
    let url: string;
    let method = 'GET';
    let headers: HeadersInit = {};
    let body: Buffer | undefined;
    let redirect: RequestRedirect = 'follow';
    let signal: AbortSignal | undefined;
    if (typeof input === 'string' || input instanceof URL) {
      url = input.toString();
    } else {
      const request = input as Request;
      if (typeof request?.url !== 'string') return nativeFetch(input, init);
      url = request.url;
      method = request.method !== '' ? request.method : method;
      headers = request.headers;
      redirect = request.redirect;
      signal = request.signal ?? undefined;
      if (init.body === undefined && request.body !== null && request.body !== undefined) {
        body = Buffer.from(await request.clone().arrayBuffer());
      }
    }
    method = (init.method ?? method).toUpperCase();
    if (init.headers !== undefined) headers = init.headers;
    if (init.body !== undefined) body = toBodyBuffer(init.body);
    if (init.redirect !== undefined) redirect = init.redirect;
    if (init.signal !== undefined) signal = init.signal ?? undefined;

    let parsed: URL;
    try {
      parsed = new URL(url);
    } catch {
      // Unparseable here is unparseable there; let native fetch raise it.
      return nativeFetch(input, init);
    }
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
      return nativeFetch(input, init);
    }
    if (isBypassed(parsed.hostname, noProxy)) return nativeFetch(input, init);

    let current = url;
    let currentParsed = parsed;
    let hops = 0;
    for (;;) {
      let hop: { status: number; headers: Headers; body: Buffer };
      try {
        hop = await singleHop(current, method, headers, body, signal, proxy, transport);
      } catch (error) {
        if (signal?.aborted === true) throw toAbortRejection(signal);
        if (error instanceof DOMException && error.name === 'AbortError') throw error;
        // Native fetch rejects transport failures as TypeError, and the graph
        // client retries on exactly that; a raw socket error would skip retry.
        throw new TypeError('fetch failed', { cause: error });
      }
      const location = hop.headers.get('location');
      const isRedirect = REDIRECT_STATUSES.has(hop.status) && location !== null;
      if (!isRedirect || redirect === 'manual') {
        const empty = isRedirect || NULL_BODY_STATUSES.has(hop.status);
        return new Response(empty ? null : hop.body, { status: hop.status, headers: hop.headers });
      }
      if (redirect === 'error') {
        throw new TypeError('proxy-fetch: redirect blocked by `redirect: "error"`');
      }
      hops += 1;
      if (hops > MAX_REDIRECTS) throw new TypeError('proxy-fetch: too many redirects');
      let next: URL;
      try {
        next = new URL(location, current);
      } catch {
        // Unparseable Location: native fetch returns the redirect as-is.
        return new Response(hop.body, { status: hop.status, headers: hop.headers });
      }
      if (next.protocol !== 'http:' && next.protocol !== 'https:') {
        return nativeFetch(next.toString(), { method, headers, body, redirect, signal });
      }
      if (isBypassed(next.hostname, noProxy)) {
        return nativeFetch(next.toString(), { method, headers, body, redirect, signal });
      }
      if (next.origin !== currentParsed.origin) {
        const stripped = new Headers(headers);
        stripped.delete('authorization');
        stripped.delete('proxy-authorization');
        headers = stripped;
      }
      // 301/302 rewrite POST only; 303 rewrites everything except GET and
      // HEAD; 307/308 never rewrite. The content headers go with a dropped body.
      if (
        (hop.status === 303 && method !== 'GET' && method !== 'HEAD') ||
        ((hop.status === 301 || hop.status === 302) && method === 'POST')
      ) {
        method = 'GET';
        body = undefined;
        const rewritten = new Headers(headers);
        rewritten.delete('content-length');
        rewritten.delete('content-type');
        headers = rewritten;
      }
      current = next.toString();
      currentParsed = next;
    }
  };
  return proxied;
}

/**
 * The fetch this module found at load: the restore point for `uninstall`, and
 * the delegate for everything the proxy must not touch. Captured unbound so a
 * round trip restores the exact original reference.
 */
const nativeFetch: typeof fetch = globalThis.fetch;

let installed = false;

/** True once the global fetch routes through the proxy. */
export function isProxyFetchInstalled(): boolean {
  return installed;
}

/**
 * Wraps the global fetch when a proxy is configured, else leaves it untouched.
 * Idempotent: a second call neither re-wraps nor re-logs. Returns whether
 * proxying is active afterwards.
 */
export function installProxyFetch(env: NodeJS.ProcessEnv = process.env): boolean {
  if (installed) return true;
  const config = resolveProxyConfig(env);
  if (config === undefined) {
    // A proxy was requested but no value was usable: say so once, naming the
    // variable but never its value, which may carry credentials.
    const unusable = PROXY_VARS.filter((name) => (env[name] ?? '').trim() !== '');
    if (unusable.length > 0) {
      logger.warn(
        `Ignoring malformed ${unusable.map((name) => `$${name}`).join(', ')}: ` +
          'not a valid proxy URL; continuing without a proxy.',
      );
    }
    return false;
  }
  globalThis.fetch = createProxiedFetch({
    proxy: config.url,
    noProxy: config.noProxy,
    nativeFetch,
  });
  installed = true;
  // Info, not debug: this runs at import time, before flags are parsed, so a
  // verbose-gated line would never print. Host and port only, never the
  // credentials the URL may carry.
  const port = config.url.port !== '' ? config.url.port : '80';
  logger.info(`Proxying fetch requests via ${config.url.hostname}:${port} (from $${config.source}).`);
  return true;
}

/** Restores the native fetch. The CLI never calls this; tests do. */
export function uninstallProxyFetch(): void {
  if (!installed) return;
  globalThis.fetch = nativeFetch;
  installed = false;
}

// The side effect `index.ts` imports first: MSAL, the graph client and the
// update check all resolve the global fetch per request, so wrapping it
// during module evaluation routes every one of them. Guarded so proxy setup
// can never take the server down, no matter what the environment holds.
try {
  installProxyFetch();
} catch {
  logger.warn('proxy-fetch: proxy setup failed; continuing without a proxy.');
}
