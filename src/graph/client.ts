/**
 * The Microsoft Graph HTTP client: one place that owns URL construction, token
 * attachment, retry, paging, host pinning and $batch chunking. Every tool module
 * goes through here, so the invariants below are load-bearing:
 *
 *  - URLs are built by `resolveGraphUrl`, never by string concatenation, and every
 *    absolute link we follow is re-checked against the configured Graph host.
 *  - Only GET is replayed after a transport failure. A POST/PATCH/PUT/DELETE is
 *    replayed only when Graph itself said it did not process the request.
 *  - A non-2xx always throws `GraphError`; a `$batch` entry's own status never does.
 */

import { setTimeout as sleep } from 'node:timers/promises';

import { GraphError } from '../contracts.js';
import type {
  AuthProvider,
  GraphBatchRequest,
  GraphBatchResponse,
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
  GraphVersion,
  HttpMethod,
  ServerConfig,
} from '../contracts.js';
import { graphBaseUrl } from '../config.js';
import { assertSameGraphHost, resolveGraphUrl } from '../util/paths.js';
import { isRetryableStatus, parseGraphError } from './errors.js';

/** Total tries for one logical request, i.e. the first attempt plus three retries. */
const MAX_ATTEMPTS = 4;
/** First backoff ceiling; doubles per retry. */
const BASE_BACKOFF_MS = 500;
/** Upper bound on a computed backoff. */
const MAX_BACKOFF_MS = 20_000;
/** Full jitter can draw ~0ms; a floor stops a retry storm from becoming a hot loop. */
const MIN_BACKOFF_MS = 100;
/**
 * The longest this client will sleep before one retry. `Retry-After` is honoured
 * exactly when it fits; when Graph asks for longer we decline to retry at all and
 * surface the 429 instead, because an MCP tool call that silently blocks for
 * minutes is worse for the caller than an error carrying a hint.
 */
const MAX_RETRY_WAIT_MS = 30_000;
/** Documented JSON batching limit. */
const BATCH_CHUNK_SIZE = 20;
/**
 * The only statuses that prove Graph rejected a request before acting on it, so
 * they are the only ones safe to replay for a mutating method.
 */
const NOT_PROCESSED_STATUSES: ReadonlySet<number> = new Set([429, 503, 504]);

/**
 * Resources where `ConsistencyLevel: eventual` is required for advanced queries.
 * Matched on the first path segment, so Outlook's `/me/contacts` is excluded while
 * the directory's `/contacts` is not.
 */
const DIRECTORY_SEGMENTS: ReadonlySet<string> = new Set([
  'users',
  'groups',
  'applications',
  'servicePrincipals',
  'devices',
  'directoryObjects',
  'directoryRoles',
  'directoryRoleTemplates',
  'administrativeUnits',
  'contacts',
  'organization',
]);

/**
 * `GraphRequestOptions` plus the header the advanced-query endpoints need. It is
 * detected automatically for `$count` and for directory `$search`/`endsWith`, and
 * can be forced by callers holding a reference to the concrete client.
 */
export interface GraphRequestOptionsInternal extends GraphRequestOptions {
  /** Sent verbatim as the `ConsistencyLevel` header. */
  consistencyLevel?: 'eventual';
}

/** The concrete client. Assignable to `GraphClient`; adds the internal request shape. */
export interface GraphClientInternal extends GraphClient {
  request<T = unknown>(opts: GraphRequestOptionsInternal): Promise<GraphResponse<T>>;
}

export function createGraphClient(opts: {
  auth: AuthProvider;
  config: ServerConfig;
}): GraphClientInternal {
  const { auth, config } = opts;
  const base = graphBaseUrl(config);

  function debug(message: string): void {
    // stdio transport: stdout carries the MCP framing, so diagnostics go to stderr.
    process.stderr.write(`[graph] ${message}\n`);
  }

  function resolveVersion(requested: GraphVersion | undefined): GraphVersion {
    const version = requested ?? config.graphVersion;
    if (version === 'beta' && !config.allowBeta) {
      throw new Error(
        'This call targets the Microsoft Graph beta endpoint, which is disabled. ' +
          'Start the server with --beta to enable beta, or use the v1.0 equivalent.',
      );
    }
    return version;
  }

  /**
   * One HTTP exchange with retries. Resolves only on 2xx; every other outcome
   * throws, so callers never have to re-check `status`.
   */
  async function execute(req: {
    url: URL;
    method: HttpMethod;
    scopes: string[];
    body?: unknown;
    headers?: Record<string, string> | undefined;
    consistencyLevel?: string | undefined;
    signal?: AbortSignal | undefined;
    errorPath: string;
  }): Promise<{ status: number; bodyText: string; headers: Headers }> {
    const hasBody = req.body !== undefined && req.body !== null;
    // A string body is sent verbatim. Graph's content endpoints
    // (`/drive/…/:/content`) store the request body as the file's bytes, so
    // JSON-encoding a string there would silently write a file wrapped in
    // quotes with escaped newlines. Everything else is JSON.
    const isRawBody = typeof req.body === 'string';
    const serialized = hasBody
      ? isRawBody
        ? (req.body as string)
        : JSON.stringify(req.body)
      : undefined;
    let retries = 0;
    let authRetried = false;

    for (;;) {
      req.signal?.throwIfAborted();
      const token = await auth.getToken(req.scopes);

      const headers = new Headers({ Accept: 'application/json' });
      if (req.consistencyLevel !== undefined) {
        headers.set('ConsistencyLevel', req.consistencyLevel);
      }
      for (const [key, value] of Object.entries(req.headers ?? {})) {
        headers.set(key, value);
      }
      // Set after the caller's headers: the credential is never overridable from
      // a tool argument.
      headers.set('Authorization', `Bearer ${token}`);
      // Content-Type does yield to the caller, because only the caller knows what
      // a raw body actually is. Absent one, a string defaults to text/plain and
      // any other value to JSON.
      if (serialized !== undefined && !headers.has('Content-Type')) {
        headers.set('Content-Type', isRawBody ? 'text/plain' : 'application/json');
      }

      if (config.verbose) debug(`${req.method} ${req.url.pathname}${req.url.search}`);

      let response: Response;
      try {
        response = await fetch(req.url, {
          method: req.method,
          headers,
          body: serialized,
          signal: req.signal,
        });
      } catch (error) {
        if (req.signal?.aborted === true) throw error;
        if (!isTransportError(error)) throw error;
        // A transport failure is no evidence about whether the server acted, so
        // only a safe method may be replayed.
        if (req.method !== 'GET' || retries >= MAX_ATTEMPTS - 1) throw error;
        retries += 1;
        const wait = backoffMs(retries);
        debug(
          `retry ${retries}/${MAX_ATTEMPTS - 1} after ${describeError(error)} on ${req.method} ${req.errorPath} in ${wait}ms`,
        );
        await pause(wait, req.signal);
        continue;
      }

      const bodyText = await response.text();
      if (response.ok) {
        return { status: response.status, bodyText, headers: response.headers };
      }

      if (response.status === 401 && !authRetried) {
        // Exactly one replay: the provider may hold a stale cached token and can
        // refresh it. A second 401 means the session is genuinely gone.
        authRetried = true;
        debug(`401 on ${req.method} ${req.errorPath}; re-acquiring token and retrying once`);
        continue;
      }

      const replayable =
        isRetryableStatus(response.status) &&
        (req.method === 'GET' || NOT_PROCESSED_STATUSES.has(response.status));
      if (replayable && retries < MAX_ATTEMPTS - 1) {
        const requested = parseRetryAfter(response.headers.get('retry-after'));
        const wait = requested ?? backoffMs(retries + 1);
        if (wait <= MAX_RETRY_WAIT_MS) {
          retries += 1;
          debug(
            `retry ${retries}/${MAX_ATTEMPTS - 1} after HTTP ${response.status} on ${req.method} ${req.errorPath} in ${wait}ms`,
          );
          await pause(wait, req.signal);
          continue;
        }
        debug(
          `not retrying ${req.method} ${req.errorPath}: Graph asked for a ${Math.round(wait / 1000)}s pause`,
        );
      }

      throw parseGraphError(response.status, req.errorPath, bodyText, response.headers);
    }
  }

  async function request<T = unknown>(
    options: GraphRequestOptionsInternal,
  ): Promise<GraphResponse<T>> {
    const version = resolveVersion(options.version);
    const method = options.method ?? 'GET';

    // A caller — or a model driving the generic tool — may fold the query into
    // the path. `normalizeGraphPath` treats '?' as an ordinary character, so it
    // would be percent-encoded into a segment; split it off instead.
    const separator = options.path.indexOf('?');
    const bare = separator === -1 ? options.path : options.path.slice(0, separator);
    const inlineQuery = separator === -1 ? '' : options.path.slice(separator + 1);

    const url = resolveGraphUrl(base, version, bare);
    for (const [key, value] of new URLSearchParams(inlineQuery)) {
      url.searchParams.append(key, value);
    }
    applyQuery(url, options.query);

    const consistencyLevel = resolveConsistencyLevel(bare, url, options.consistencyLevel);
    const errorPath = `${bare}${url.search}`;

    const first = await execute({
      url,
      method,
      scopes: options.scopes,
      body: options.body,
      headers: options.headers,
      consistencyLevel,
      signal: options.signal,
      errorPath,
    });

    const body = parseJsonBody(first.bodyText);
    let nextLink = readLink(body, '@odata.nextLink');
    let deltaLink = readLink(body, '@odata.deltaLink');

    const maxPages = Math.max(1, Math.trunc(options.maxPages ?? 1));
    const collection = readValueArray(body);
    if (maxPages <= 1 || collection === undefined || nextLink === undefined) {
      return buildResponse<T>(first.status, body, nextLink, deltaLink);
    }

    const merged = [...collection];
    let pages = 1;
    while (pages < maxPages && nextLink !== undefined) {
      const pageUrl = parseGraphLink(nextLink, config.graphHost);
      const page = await execute({
        url: pageUrl,
        method: 'GET',
        scopes: options.scopes,
        headers: options.headers,
        // The nextLink re-encodes the original query options, so whatever needed
        // advanced query support on page one still needs it here.
        consistencyLevel,
        signal: options.signal,
        errorPath,
      });
      const pageBody = parseJsonBody(page.bodyText);
      const pageValue = readValueArray(pageBody);
      if (pageValue === undefined) break;
      merged.push(...pageValue);
      nextLink = readLink(pageBody, '@odata.nextLink');
      deltaLink = readLink(pageBody, '@odata.deltaLink') ?? deltaLink;
      pages += 1;
    }

    const record = asRecord(body) ?? {};
    const mergedBody: Record<string, unknown> = { ...record, value: merged };
    delete mergedBody['@odata.nextLink'];
    delete mergedBody['@odata.deltaLink'];
    if (nextLink !== undefined) mergedBody['@odata.nextLink'] = nextLink;
    if (deltaLink !== undefined) mergedBody['@odata.deltaLink'] = deltaLink;

    return buildResponse<T>(first.status, mergedBody, nextLink, deltaLink);
  }

  async function follow<T = unknown>(link: string, scopes: string[]): Promise<GraphResponse<T>> {
    const url = parseGraphLink(link, config.graphHost);
    const result = await execute({
      url,
      method: 'GET',
      scopes,
      consistencyLevel: linkNeedsEventualConsistency(url) ? 'eventual' : undefined,
      errorPath: `${url.pathname}${url.search}`,
    });
    const body = parseJsonBody(result.bodyText);
    return buildResponse<T>(
      result.status,
      body,
      readLink(body, '@odata.nextLink'),
      readLink(body, '@odata.deltaLink'),
    );
  }

  async function batch(
    requests: GraphBatchRequest[],
    scopes: string[],
  ): Promise<GraphBatchResponse[]> {
    if (requests.length === 0) return [];
    const version = resolveVersion(undefined);
    const url = resolveGraphUrl(base, version, '/$batch');
    const results: GraphBatchResponse[] = [];

    // Sequential on purpose: Outlook dispatches at most four sub-requests per
    // mailbox in parallel, so overlapping chunks buys nothing and risks 429s.
    for (const group of chunk(requests, BATCH_CHUNK_SIZE)) {
      let responses = await sendBatch(group, scopes, url);

      const throttled = group.filter((entry) => responses.get(entry.id)?.status === 429);
      if (throttled.length > 0) {
        const wait = Math.min(maxRetryAfter(throttled, responses) ?? BASE_BACKOFF_MS, MAX_RETRY_WAIT_MS);
        debug(`batch: ${throttled.length} entries throttled, retrying once in ${wait}ms`);
        await pause(wait, undefined);
        // Dependencies left behind in the first batch cannot be expressed in the
        // retry, so drop them rather than provoking a 400 on the whole envelope.
        const retryIds = new Set(throttled.map((entry) => entry.id));
        const retryGroup = throttled.map((entry) => ({
          ...entry,
          dependsOn: entry.dependsOn?.filter((id) => retryIds.has(id)),
        }));
        const retried = await sendBatch(retryGroup, scopes, url);
        responses = new Map([...responses, ...retried]);
      }

      // Graph may answer out of order; hand the caller the order it asked for.
      const unmatched = new Map(responses);
      for (const entry of group) {
        const result = responses.get(entry.id);
        unmatched.delete(entry.id);
        results.push(result ?? missingBatchEntry(entry.id));
      }
      for (const stray of unmatched.values()) results.push(stray);
    }

    return results;
  }

  async function sendBatch(
    group: GraphBatchRequest[],
    scopes: string[],
    url: URL,
  ): Promise<Map<string, GraphBatchResponse>> {
    const payload = {
      requests: group.map((entry) => {
        const hasBody = entry.body !== undefined && entry.body !== null;
        // Graph rejects a batch entry that carries a body without this header.
        const headers = hasBody
          ? { 'Content-Type': 'application/json', ...entry.headers }
          : entry.headers;
        const item: Record<string, unknown> = {
          id: entry.id,
          method: entry.method,
          url: batchRelativeUrl(entry.url),
        };
        if (hasBody) item['body'] = entry.body;
        if (headers !== undefined) item['headers'] = headers;
        if (entry.dependsOn !== undefined && entry.dependsOn.length > 0) {
          item['dependsOn'] = entry.dependsOn;
        }
        return item;
      }),
    };

    const envelope = await execute({
      url,
      method: 'POST',
      scopes,
      body: payload,
      errorPath: '/$batch',
    });

    const parsed = asRecord(parseJsonBody(envelope.bodyText));
    const entries = parsed === undefined ? undefined : parsed['responses'];
    if (!Array.isArray(entries)) {
      throw new GraphError({
        message: 'Microsoft Graph accepted the $batch request but returned no responses array.',
        status: envelope.status,
        code: 'invalidBatchResponse',
        path: '/$batch',
        requestId: envelope.headers.get('request-id') ?? undefined,
        hint: 'Retry the calls individually; the batch envelope itself was malformed.',
      });
    }

    const byId = new Map<string, GraphBatchResponse>();
    for (const raw of entries) {
      const item = toBatchResponse(raw);
      if (item !== undefined) byId.set(item.id, item);
    }
    return byId;
  }

  return { request, batch, follow };
}

/** Returns `data.value` when it is an array, else an empty array. */
export function extractCollection<T>(data: unknown): T[] {
  const record = asRecord(data);
  const value = record?.['value'];
  return Array.isArray(value) ? (value as T[]) : [];
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function buildResponse<T>(
  status: number,
  data: unknown,
  nextLink: string | undefined,
  deltaLink: string | undefined,
): GraphResponse<T> {
  const response: GraphResponse<T> = { status, data: data as T };
  if (nextLink !== undefined) response.nextLink = nextLink;
  if (deltaLink !== undefined) response.deltaLink = deltaLink;
  return response;
}

function applyQuery(
  url: URL,
  query: Record<string, string | number | boolean | undefined> | undefined,
): void {
  if (query === undefined) return;
  for (const [key, value] of Object.entries(query)) {
    if (value === undefined) continue;
    // `set`, so an explicit query parameter wins over one inlined in the path.
    url.searchParams.set(key, String(value));
  }
}

/**
 * Advanced directory queries fail with a 400 unless `ConsistencyLevel: eventual`
 * is present. `$count` needs it everywhere it is supported; `$search` and
 * `endsWith` need it only on directory resources.
 */
function resolveConsistencyLevel(
  path: string,
  url: URL,
  explicit: string | undefined,
): string | undefined {
  if (explicit !== undefined) return explicit;

  const params = url.searchParams;
  const count = params.get('$count') ?? params.get('count');
  if (count !== null && count !== '' && count !== 'false') return 'eventual';

  if (!isDirectoryResource(path)) return undefined;
  if (params.has('$search')) return 'eventual';
  const filter = params.get('$filter');
  if (filter !== null && filter.toLowerCase().includes('endswith(')) return 'eventual';
  return undefined;
}

function isDirectoryResource(path: string): boolean {
  const segment = path.replace(/^\/+/, '').split('/', 1)[0] ?? '';
  return DIRECTORY_SEGMENTS.has(segment);
}

function linkNeedsEventualConsistency(url: URL): boolean {
  const count = url.searchParams.get('$count');
  if (count !== null && count !== 'false') return true;
  // The version prefix sits in front of the resource on an absolute link.
  const segments = url.pathname.split('/').filter((part) => part.length > 0);
  const resource = segments[1] ?? '';
  return DIRECTORY_SEGMENTS.has(resource) && url.searchParams.has('$search');
}

/** Parses an absolute Graph link and pins it to the configured host. */
function parseGraphLink(link: string, expectedHost: string): URL {
  let url: URL;
  try {
    url = new URL(link);
  } catch {
    throw new Error(
      `Not an absolute Microsoft Graph link: ${link}. Pass the @odata.nextLink or @odata.deltaLink exactly as Graph returned it.`,
    );
  }
  assertSameGraphHost(url, expectedHost);
  return url;
}

/** `$batch` entries address Graph with a version-relative URL. */
function batchRelativeUrl(url: string): string {
  if (/^[a-z][a-z0-9+.-]*:/i.test(url)) {
    throw new Error(
      `A $batch entry URL must be Graph-relative, e.g. /me/messages?$top=5, not an absolute URL (${url}).`,
    );
  }
  return url.startsWith('/') ? url : `/${url}`;
}

function missingBatchEntry(id: string): GraphBatchResponse {
  return {
    id,
    status: 502,
    body: {
      error: {
        code: 'missingBatchResponse',
        message: 'Microsoft Graph returned no $batch entry for this request id.',
      },
    },
  };
}

function parseJsonBody(bodyText: string): unknown {
  const trimmed = bodyText.trim();
  if (trimmed.length === 0) return null;
  try {
    return JSON.parse(trimmed);
  } catch {
    // A 2xx that is not JSON — a `$value` download, for instance. Hand back the text.
    return bodyText;
  }
}

function readValueArray(body: unknown): unknown[] | undefined {
  const record = asRecord(body);
  const value = record?.['value'];
  return Array.isArray(value) ? value : undefined;
}

function readLink(body: unknown, key: '@odata.nextLink' | '@odata.deltaLink'): string | undefined {
  const record = asRecord(body);
  const value = record?.[key];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function toBatchResponse(entry: unknown): GraphBatchResponse | undefined {
  const record = asRecord(entry);
  if (record === undefined) return undefined;
  const id = record['id'];
  const status = record['status'];
  if (typeof id !== 'string' || typeof status !== 'number') return undefined;
  const result: GraphBatchResponse = { id, status };
  if (record['body'] !== undefined) result.body = record['body'];
  const headers = toStringRecord(record['headers']);
  if (headers !== undefined) result.headers = headers;
  return result;
}

function toStringRecord(value: unknown): Record<string, string> | undefined {
  const record = asRecord(value);
  if (record === undefined) return undefined;
  const out: Record<string, string> = {};
  for (const [key, item] of Object.entries(record)) {
    if (typeof item === 'string') out[key] = item;
  }
  return out;
}

function maxRetryAfter(
  entries: GraphBatchRequest[],
  responses: Map<string, GraphBatchResponse>,
): number | undefined {
  let wait: number | undefined;
  for (const entry of entries) {
    const headers = responses.get(entry.id)?.headers;
    if (headers === undefined) continue;
    const raw = Object.entries(headers).find(([key]) => key.toLowerCase() === 'retry-after')?.[1];
    const parsed = parseRetryAfter(raw ?? null);
    if (parsed !== undefined && (wait === undefined || parsed > wait)) wait = parsed;
  }
  return wait;
}

function chunk<T>(items: readonly T[], size: number): T[][] {
  const groups: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    groups.push(items.slice(index, index + size));
  }
  return groups;
}

/** `Retry-After` is either delta-seconds or an HTTP-date. Both are honoured as given. */
export function parseRetryAfter(value: string | null, now: number = Date.now()): number | undefined {
  if (value === null) return undefined;
  const trimmed = value.trim();
  if (trimmed.length === 0) return undefined;

  if (/^\d+$/.test(trimmed)) return Number(trimmed) * 1000;

  const at = Date.parse(trimmed);
  if (Number.isNaN(at)) return undefined;
  return Math.max(0, at - now);
}

/** Full jitter: a uniform draw below an exponentially growing ceiling. */
function backoffMs(retry: number): number {
  const ceiling = Math.min(MAX_BACKOFF_MS, BASE_BACKOFF_MS * 2 ** (retry - 1));
  return Math.max(MIN_BACKOFF_MS, Math.floor(Math.random() * ceiling));
}

async function pause(ms: number, signal: AbortSignal | undefined): Promise<void> {
  if (ms <= 0) {
    signal?.throwIfAborted();
    return;
  }
  await sleep(ms, undefined, { signal });
}

function isTransportError(error: unknown): boolean {
  if (error instanceof TypeError) return true;
  return error instanceof Error && error.name === 'AbortError';
}

function describeError(error: unknown): string {
  return error instanceof Error ? `${error.name}: ${error.message}` : String(error);
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}
