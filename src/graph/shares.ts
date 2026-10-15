/**
 * Sharing-link resolution: a `contentUrl` from a Teams message or a mail
 * reference attachment is a SharePoint/OneDrive sharing link, not a drive
 * address. Graph resolves it through `GET /shares/{token}/driveItem`, where
 * the token is the URL itself, base64url-encoded and prefixed with `u!`.
 *
 * This module owns that encoding plus the one request, so chat, teams, mail
 * and files all resolve links the same way.
 */

import type { GraphClient } from '../contracts.js';

export interface SharedDriveItem {
  driveId: string;
  itemId: string;
  name?: string | undefined;
  kind?: 'file' | 'folder' | undefined;
  size?: number | undefined;
  mimeType?: string | undefined;
  webUrl?: string | undefined;
}

interface ShareDriveItemPayload {
  id?: string | null;
  name?: string | null;
  size?: number | null;
  webUrl?: string | null;
  file?: { mimeType?: string | null } | null;
  folder?: { childCount?: number } | null;
  parentReference?: { driveId?: string | null } | null;
}

/**
 * Encodes a sharing URL as the `{token}` in `GET /shares/{token}/driveItem`:
 * the URL as UTF-8 bytes, base64-encoded, made URL-safe, padding trimmed,
 * prefixed with `u!`. Throws on anything that is not an absolute http(s) URL
 * rather than minting a token Graph would 400 on.
 */
export function shareUrlToToken(shareUrl: string): string {
  const raw = shareUrl.trim();
  if (raw.length === 0) {
    throw new Error('shareUrl must be a non-empty sharing link, e.g. a contentUrl from a Teams message attachment.');
  }
  let parsed: URL;
  try {
    parsed = new URL(raw);
  } catch {
    throw new Error(
      `Invalid shareUrl ${JSON.stringify(shareUrl)}: pass the full sharing link, e.g. "https://tenant.sharepoint.com/:w:/g/...".`,
    );
  }
  if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') {
    throw new Error(
      `Invalid shareUrl ${JSON.stringify(shareUrl)}: only http(s) sharing links resolve through /shares.`,
    );
  }
  const token = Buffer.from(raw, 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
  return `u!${token}`;
}

/** Inverse of `shareUrlToToken`, for tests and for echoing the link back. */
export function shareTokenToUrl(token: string): string {
  const body = token.startsWith('u!') ? token.slice(2) : token;
  const padded = body.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - (body.length % 4)) % 4);
  return Buffer.from(padded, 'base64').toString('utf8');
}

/**
 * Resolves one sharing link to the drive item behind it. Needs a token that
 * can read outside the caller's own drive — `Files.Read.All` in practice —
 * because shared links usually point at someone else's files.
 */
export async function resolveShare(
  graph: GraphClient,
  scopes: string[],
  shareUrl: string,
): Promise<SharedDriveItem> {
  const token = shareUrlToToken(shareUrl);
  const res = await graph.request<ShareDriveItemPayload>({
    // The token is base64url plus a `u!` prefix: no slash, no query shaping.
    path: `/shares/${token}/driveItem`,
    query: { $select: 'id,name,size,file,folder,webUrl,parentReference' },
    scopes,
  });
  const item = res.data ?? {};
  const driveId = item.parentReference?.driveId ?? undefined;
  const itemId = item.id ?? undefined;
  if (driveId === undefined || itemId === undefined) {
    throw new Error(
      'The sharing link resolved, but Graph returned a drive item without its drive id. ' +
        'The link may point at a deleted item.',
    );
  }
  const isFolder = item.folder !== null && item.folder !== undefined;
  return {
    driveId,
    itemId,
    name: item.name ?? undefined,
    kind: isFolder ? 'folder' : 'file',
    size: typeof item.size === 'number' ? item.size : undefined,
    mimeType: item.file?.mimeType ?? undefined,
    webUrl: item.webUrl ?? undefined,
  };
}
