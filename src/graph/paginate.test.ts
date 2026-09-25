import assert from 'node:assert/strict';
import { test } from 'node:test';

import { GraphError } from '../contracts.js';
import type {
  GraphBatchResponse,
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
} from '../contracts.js';
import { createPaginator } from './paginate.js';
import type { PageWalkOptions } from './paginate.js';

interface Item {
  id: string;
}

function item(id: string): Item {
  return { id };
}

const LINK_1 = 'https://graph.microsoft.com/v1.0/chats/19:chat/messages?$skiptoken=p2';
const LINK_2 = 'https://graph.microsoft.com/v1.0/chats/19:chat/messages?$skiptoken=p3';
const LINK_3 = 'https://graph.microsoft.com/v1.0/chats/19:chat/messages?$skiptoken=p4';

interface StubPage {
  value: Item[];
  next?: string;
}

interface StubCall {
  kind: 'request' | 'follow';
  /** The path for a request, the followed link for a follow. */
  target: string;
  maxPages: number | undefined;
  /** Same clock the paginator paces on, so gap assertions are exact. */
  at: number;
}

type StubGraph = GraphClient & { readonly calls: StubCall[] };

/** Serves `pages` in order; a call past the end fails loudly rather than looping. */
function stubGraph(pages: StubPage[]): StubGraph {
  const calls: StubCall[] = [];
  let index = 0;

  function serve(call: Omit<StubCall, 'at'>): GraphResponse<unknown> {
    calls.push({ ...call, at: performance.now() });
    const page = pages[index];
    index += 1;
    if (page === undefined) {
      throw new Error(`stub exhausted: call ${calls.length} has no page to serve`);
    }
    const response: GraphResponse<unknown> = { status: 200, data: { value: page.value } };
    if (page.next !== undefined) response.nextLink = page.next;
    return response;
  }

  return {
    calls,
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      const page = serve({ kind: 'request', target: opts.path, maxPages: opts.maxPages });
      return page as GraphResponse<T>;
    },
    async follow<T = unknown>(link: string): Promise<GraphResponse<T>> {
      const page = serve({ kind: 'follow', target: link, maxPages: undefined });
      return page as GraphResponse<T>;
    },
    async batch(): Promise<GraphBatchResponse[]> {
      throw new Error('batch is not exercised by these tests');
    },
  };
}

function options(overrides: Partial<PageWalkOptions<Item>> = {}): PageWalkOptions<Item> {
  return {
    path: '/chats/19:chat/messages',
    scopes: ['Chat.Read'],
    maxItems: 100,
    maxPages: 10,
    minIntervalMs: 0,
    budgetMs: 10_000,
    ...overrides,
  };
}

function ids(items: Item[]): string[] {
  return items.map((entry) => entry.id);
}

function callAt(graph: StubGraph, index: number): StubCall {
  const call = graph.calls[index];
  assert.ok(call !== undefined, `expected a call at index ${index}`);
  return call;
}

test('walks multiple pages and concatenates in order', async () => {
  const graph = stubGraph([
    { value: [item('a'), item('b')], next: LINK_1 },
    { value: [item('c')], next: LINK_2 },
    { value: [item('d'), item('e')] },
  ]);

  const result = await createPaginator(graph).walk<Item>(options());

  assert.deepEqual(ids(result.items), ['a', 'b', 'c', 'd', 'e']);
  assert.equal(result.pages, 3);
  assert.equal(result.reason, 'complete');
  assert.equal(result.throttled, false);
  // The engine owns paging; the client must never auto-page underneath it.
  assert.equal(callAt(graph, 0).maxPages, 1);
  assert.equal(callAt(graph, 1).target, LINK_1);
  assert.equal(callAt(graph, 2).target, LINK_2);
});

test('stops at maxItems and returns a nextLink', async () => {
  const graph = stubGraph([
    { value: [item('a'), item('b')], next: LINK_1 },
    { value: [item('c'), item('d')], next: LINK_2 },
    { value: [item('e')] },
  ]);

  const result = await createPaginator(graph).walk<Item>(options({ maxItems: 3 }));

  assert.deepEqual(ids(result.items), ['a', 'b', 'c']);
  assert.equal(result.reason, 'maxItems');
  assert.equal(result.pages, 2);
  // The resume handle is the last fetched page's link, not the one after it.
  assert.equal(result.nextLink, LINK_2);
});

test('a cap landing exactly on the end of the collection reports complete', async () => {
  const graph = stubGraph([{ value: [item('a'), item('b')] }]);

  const result = await createPaginator(graph).walk<Item>(options({ maxItems: 2 }));

  assert.deepEqual(ids(result.items), ['a', 'b']);
  assert.equal(result.reason, 'complete');
  assert.equal(result.nextLink, undefined);
});

test('stop() excludes the stopping item and everything after it', async () => {
  const graph = stubGraph([
    { value: [item('a'), item('b')], next: LINK_1 },
    { value: [item('c'), item('halt'), item('d')], next: LINK_2 },
    { value: [item('e')] },
  ]);

  const result = await createPaginator(graph).walk<Item>(
    options({ stop: (entry) => entry.id === 'halt' }),
  );

  assert.deepEqual(ids(result.items), ['a', 'b', 'c']);
  assert.equal(result.reason, 'stopped');
  assert.equal(result.pages, 2);
  assert.equal(result.nextLink, LINK_2);
});

test('keep() filters items without ending the walk', async () => {
  const graph = stubGraph([
    { value: [item('keep-1'), item('drop-1')], next: LINK_1 },
    { value: [item('drop-2'), item('drop-3')], next: LINK_2 },
    { value: [item('keep-2')] },
  ]);

  const result = await createPaginator(graph).walk<Item>(
    options({ keep: (entry) => entry.id.startsWith('keep') }),
  );

  assert.deepEqual(ids(result.items), ['keep-1', 'keep-2']);
  assert.equal(result.pages, 3);
  assert.equal(result.reason, 'complete');
});

test('a collection that runs out reports complete with no nextLink', async () => {
  const graph = stubGraph([{ value: [item('a'), item('b')] }]);

  const result = await createPaginator(graph).walk<Item>(options());

  assert.equal(result.reason, 'complete');
  assert.equal(result.nextLink, undefined);
  assert.equal(result.pages, 1);
  assert.equal(graph.calls.length, 1);
});

test('maxPages caps the request count', async () => {
  const graph = stubGraph([
    { value: [item('a')], next: LINK_1 },
    { value: [item('b')], next: LINK_2 },
    { value: [item('c')], next: LINK_3 },
  ]);

  const result = await createPaginator(graph).walk<Item>(options({ maxPages: 2 }));

  assert.equal(graph.calls.length, 2);
  assert.equal(result.pages, 2);
  assert.deepEqual(ids(result.items), ['a', 'b']);
  assert.equal(result.reason, 'maxPages');
  assert.equal(result.nextLink, LINK_2);
});

test('minIntervalMs paces consecutive requests and reports throttled', async () => {
  const interval = 30;
  const graph = stubGraph([
    { value: [item('a')], next: LINK_1 },
    { value: [item('b')] },
  ]);

  const result = await createPaginator(graph).walk<Item>(options({ minIntervalMs: interval }));

  assert.equal(graph.calls.length, 2);
  const gap = callAt(graph, 1).at - callAt(graph, 0).at;
  assert.ok(gap >= interval, `expected at least ${interval}ms between requests, saw ${gap}ms`);
  assert.equal(result.throttled, true);
});

test('a cursor resumes through follow rather than request', async () => {
  const graph = stubGraph([
    { value: [item('c')], next: LINK_2 },
    { value: [item('d')] },
  ]);

  const result = await createPaginator(graph).walk<Item>(options({ cursor: LINK_1 }));

  assert.equal(callAt(graph, 0).kind, 'follow');
  assert.equal(callAt(graph, 0).target, LINK_1);
  assert.equal(callAt(graph, 1).target, LINK_2);
  assert.ok(!graph.calls.some((call) => call.kind === 'request'));
  assert.deepEqual(ids(result.items), ['c', 'd']);
  assert.equal(result.reason, 'complete');
});

test('an already exhausted budget still returns the first page', async () => {
  const graph = stubGraph([
    { value: [item('a'), item('b')], next: LINK_1 },
    { value: [item('c')] },
  ]);

  const result = await createPaginator(graph).walk<Item>(options({ budgetMs: 0 }));

  assert.equal(graph.calls.length, 1);
  assert.deepEqual(ids(result.items), ['a', 'b']);
  assert.equal(result.reason, 'budget');
  assert.equal(result.nextLink, LINK_1);
});

test('an aborted signal rejects the walk', async () => {
  const graph = stubGraph([{ value: [item('a')] }]);

  await assert.rejects(
    () => createPaginator(graph).walk<Item>(options({ signal: AbortSignal.abort() })),
    (error: unknown) => error instanceof Error && error.name === 'AbortError',
  );
  assert.equal(graph.calls.length, 0);
});

test('a GraphError propagates untouched', async () => {
  const failure = new GraphError({
    message: 'Too many requests',
    status: 429,
    code: 'activityLimitReached',
    path: '/chats/19:chat/messages',
  });
  const graph: GraphClient = {
    async request<T = unknown>(): Promise<GraphResponse<T>> {
      throw failure;
    },
    async follow<T = unknown>(): Promise<GraphResponse<T>> {
      throw failure;
    },
    async batch(): Promise<GraphBatchResponse[]> {
      throw new Error('batch is not exercised by these tests');
    },
  };

  await assert.rejects(
    () => createPaginator(graph).walk<Item>(options()),
    (error: unknown) => error === failure,
  );
});
