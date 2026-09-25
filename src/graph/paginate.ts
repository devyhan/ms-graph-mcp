/**
 * A paced, budgeted walk over a paged Microsoft Graph collection.
 *
 * `GraphClient.request` can already follow `@odata.nextLink` on its own, but it
 * does so as fast as the network allows and hands back everything at once. A
 * Teams history walk cannot work that way:
 *
 *  - Teams allows one request per second per app per tenant on a given chat or
 *    channel, so the gap between requests is the binding constraint.
 *  - An MCP tool call has a wall clock to answer within, and its output is
 *    truncated near 25k tokens by the client, so a walk must be able to stop
 *    early and hand back a resumable cursor.
 *  - History is walked until some boundary — a date, an id already seen — which
 *    is a per-item decision the client layer cannot make.
 *
 * So this engine owns paging: every underlying call is `maxPages: 1`, and the
 * loop here does the pacing, the budgeting and the early exit.
 */

import { setTimeout as sleep } from 'node:timers/promises';

import type { GraphClient } from '../contracts.js';
import { extractCollection } from './client.js';

export interface PageWalkOptions<T> {
  path: string;
  query?: Record<string, string | number | boolean | undefined>;
  scopes: string[];
  maxItems: number;
  maxPages: number;
  /** Minimum gap between HTTP requests. Teams allows 1 rps per chat/channel. */
  minIntervalMs: number;
  /** Wall-clock budget for the whole walk. */
  budgetMs: number;
  /** Resume from a previous walk's nextLink instead of path+query. */
  cursor?: string | undefined;
  /** Drop an item from the result without ending the walk. */
  keep?: (item: T) => boolean;
  /** End the walk at this item; the item itself is NOT collected. */
  stop?: (item: T) => boolean;
  signal?: AbortSignal | undefined;
}

export type WalkEndReason =
  | 'complete' // the collection ran out
  | 'maxItems'
  | 'maxPages'
  | 'budget'
  | 'stopped'; // the stop() predicate fired

export interface PageWalkResult<T> {
  items: T[];
  pages: number;
  reason: WalkEndReason;
  /** Present when more remains. Feed back as options.cursor to continue. */
  nextLink?: string | undefined;
  elapsedMs: number;
  /** True when the walk was paced by minIntervalMs at least once. */
  throttled: boolean;
}

export function createPaginator(graph: GraphClient): {
  walk<T>(opts: PageWalkOptions<T>): Promise<PageWalkResult<T>>;
} {
  async function walk<T>(opts: PageWalkOptions<T>): Promise<PageWalkResult<T>> {
    const startedAt = performance.now();
    const items: T[] = [];
    let pages = 0;
    let throttled = false;
    let reason: WalkEndReason = 'complete';
    /** Start of the previous request: pacing measures gap-to-gap, not idle time. */
    let previousRequestAt: number | undefined;
    /** The `@odata.nextLink` of the most recently fetched page. */
    let pageNextLink: string | undefined;
    /** The link to follow next; `undefined` means "issue the initial request". */
    let link: string | undefined = opts.cursor;

    for (;;) {
      opts.signal?.throwIfAborted();

      if (previousRequestAt !== undefined) {
        // Measured from the start of the previous request, so a slow request
        // does not add to the gap. Re-checked after each sleep because a timer
        // may fire a shade early and this floor is a rate limit, not a hint.
        for (;;) {
          const wait = opts.minIntervalMs - (performance.now() - previousRequestAt);
          if (wait <= 0) break;
          throttled = true;
          await sleep(wait, undefined, { signal: opts.signal });
        }
      }

      previousRequestAt = performance.now();
      const response =
        link === undefined
          ? await graph.request<unknown>({
              path: opts.path,
              query: opts.query,
              scopes: opts.scopes,
              // Never more than one: auto-paging inside the client would bypass
              // the pacing, the budget and stop() all at once.
              maxPages: 1,
              signal: opts.signal,
            })
          : await graph.follow<unknown>(link, opts.scopes);
      pages += 1;
      pageNextLink = response.nextLink;

      let stopped = false;
      /** True when the item cap cut this page short, so items remain unread in it. */
      let truncated = false;
      for (const item of extractCollection<T>(response.data)) {
        // The cap is checked before the predicates: once full, the walk must not
        // consume another item, so a resume from this page's nextLink starts
        // exactly where this walk left the collection.
        if (items.length >= opts.maxItems) {
          truncated = true;
          break;
        }
        if (opts.stop?.(item) === true) {
          stopped = true;
          break;
        }
        if (opts.keep === undefined || opts.keep(item)) items.push(item);
      }

      if (stopped) {
        reason = 'stopped';
        break;
      }
      // Nothing left anywhere — not in this page, not behind a nextLink.
      if (!truncated && pageNextLink === undefined) {
        reason = 'complete';
        break;
      }
      if (items.length >= opts.maxItems) {
        reason = 'maxItems';
        break;
      }
      if (pages >= opts.maxPages) {
        reason = 'maxPages';
        break;
      }
      // Checked before the next request, never in the middle of one: a page we
      // already paid for is always collected before the budget can end the walk.
      if (performance.now() - startedAt >= opts.budgetMs) {
        reason = 'budget';
        break;
      }
      if (pageNextLink === undefined) {
        // Unreachable given the checks above, but re-entering the loop with no
        // link would silently re-request page one forever.
        reason = 'complete';
        break;
      }
      link = pageNextLink;
    }

    const result: PageWalkResult<T> = {
      items,
      pages,
      reason,
      elapsedMs: Math.round(performance.now() - startedAt),
      throttled,
    };
    // MCP is stateless, so this link is the entire resume story: without it the
    // caller cannot continue and would have to restart the walk from the top.
    // Note that after 'stopped' it points past the whole page the stopping item
    // was in — the caller asked to end there, and everything after that item is
    // by definition on the far side of the boundary it chose.
    if (reason !== 'complete' && pageNextLink !== undefined) result.nextLink = pageNextLink;
    return result;
  }

  return { walk };
}
