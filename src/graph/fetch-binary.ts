/**
 * Byte-range reads of Graph content endpoints (`/driveItem/content`,
 * attachment `/$value`).
 *
 * Two facts shape this module. `/content` answers 302 to a pre-authenticated
 * storage URL, and the caller's `Range` must be re-issued on that Location
 * WITHOUT the Authorization header: the download URL already carries its own
 * credential, and forwarding the bearer cross-host would leak it to storage.
 * `fetch` with the default redirect policy forwards headers to the new host,
 * so the first request goes out with `redirect: 'manual'` and the follow-up
 * carries `Range` only — safe by construction, not by trusting header
 * filtering. Second, the Graph client decodes every 2xx as text, which is
 * wrong for Office and PDF bytes; this module reads `arrayBuffer` instead.
 */

import type { AuthProvider, GraphVersion, ServerConfig } from '../contracts.js';
import { graphBaseUrl } from '../config.js';
import { resolveGraphUrl } from '../util/paths.js';

export interface FetchRangeOptions {
  /** Graph-relative path, e.g. `/drives/{id}/items/{id}/content`. */
  path: string;
  scopes: string[];
  /** First byte wanted. Defaults to 0. */
  offset?: number;
  /** Most bytes to read. Defaults to 4 MiB. */
  byteBudget?: number;
  version?: GraphVersion;
  signal?: AbortSignal;
  /** Override for tests; defaults to the global fetch. */
  fetchImpl?: typeof fetch;
}

export interface FetchRangeResult {
  bytes: Uint8Array;
  /** Offset to pass back for the next window. */
  nextOffset: number;
  /** Full length from Content-Range, when the server sent one. */
  totalLength?: number | undefined;
  /** True when the server says more bytes remain past this window. */
  truncated: boolean;
}

const DEFAULT_BYTE_BUDGET = 4 * 1024 * 1024;
const MAX_REDIRECTS = 3;
const REDIRECT_STATUSES = new Set([301, 302, 303, 307, 308]);

/** `bytes 0-1023/2048` -> 2048. `bytes 0-1023/*` -> undefined. */
function totalFromContentRange(header: string | null): number | undefined {
  if (header === null) return undefined;
  const total = /\/(\d+)\s*$/.exec(header)?.[1];
  if (total === undefined) return undefined;
  const parsed = Number.parseInt(total, 10);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

export async function fetchContentRange(
  auth: AuthProvider,
  config: Pick<ServerConfig, 'graphHost' | 'graphVersion'>,
  options: FetchRangeOptions,
): Promise<FetchRangeResult> {
  const fetchImpl = options.fetchImpl ?? fetch;
  const offset = Math.max(0, Math.trunc(options.offset ?? 0));
  const byteBudget = Math.max(1, Math.trunc(options.byteBudget ?? DEFAULT_BYTE_BUDGET));
  const last = offset + byteBudget - 1;
  const range = `bytes=${offset}-${last}`;

  const token = await auth.getToken(options.scopes);
  const first = resolveGraphUrl(
    graphBaseUrl(config),
    options.version ?? config.graphVersion,
    options.path,
  );

  let response = await fetchImpl(first, {
    method: 'GET',
    headers: { Authorization: `Bearer ${token}`, Range: range, Accept: '*/*' },
    // Manual so the bearer never rides along to the storage host below.
    redirect: 'manual',
    signal: options.signal,
  });

  for (let hop = 0; hop < MAX_REDIRECTS && REDIRECT_STATUSES.has(response.status); hop += 1) {
    const location = response.headers.get('location');
    if (location === null || location.length === 0) {
      throw new Error(
        `Graph answered ${response.status} for ${options.path} without a Location to follow.`,
      );
    }
    // Deliberately no Authorization: the Location is a pre-authenticated URL.
    response = await fetchImpl(location, {
      method: 'GET',
      headers: { Range: range, Accept: '*/*' },
      redirect: 'follow',
      signal: options.signal,
    });
  }

  if (!response.ok) {
    throw new Error(
      `Download failed with HTTP ${response.status} for ${options.path}. ` +
        'The item may have moved, or the token may lack Files.Read.All for a shared file.',
    );
  }

  const bytes = new Uint8Array(await response.arrayBuffer());
  const totalLength = totalFromContentRange(response.headers.get('content-range'));
  const nextOffset = offset + bytes.length;
  return {
    bytes,
    nextOffset,
    totalLength,
    truncated: totalLength !== undefined && nextOffset < totalLength,
  };
}
