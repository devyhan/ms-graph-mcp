import assert from 'node:assert/strict';
import test from 'node:test';

import { scopeReadingOnlyAuth } from '../auth/index.js';
import type {
  GraphBatchRequest,
  GraphBatchResponse,
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
  ServerConfig,
  ToolDefinition,
} from '../contracts.js';
import { serializeResult } from '../util/truncate.js';
import { chatModule } from './chat.js';

/**
 * Handler-level tests for the `chat` group, mirroring mail.test.ts.
 *
 * Every Graph call is stubbed, so this needs no account and no network. The
 * stub records the requests it receives: the paging and tiering assertions
 * below read those recordings, not just the projected rows.
 */

const config: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['chat'],
  readOnly: true,
  graphVersion: 'v1.0',
  allowBeta: false,
  discovery: false,
  orgMode: false,
  maxOutputChars: 60_000,
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'auto',
  authPort: 0,
  allowGenericWrite: false,
};

const CHAT_A = '19:aaa@thread.v2';
const CHAT_B = '19:bbb@thread.v2';

function preview(
  id: string,
  created: string,
  from: string,
  body: string,
  contentType = 'text',
): Record<string, unknown> {
  return {
    id,
    createdDateTime: created,
    messageType: 'message',
    from: { user: { displayName: from } },
    body: { contentType, content: body },
  };
}

function chat(
  id: string,
  topic: string | null,
  chatType: string,
  lastMessagePreview?: Record<string, unknown> | null,
): Record<string, unknown> {
  return { id, topic, chatType, lastMessagePreview: lastMessagePreview ?? null };
}

function message(
  id: string,
  created: string,
  from: string,
  body: string,
): Record<string, unknown> {
  return {
    id,
    createdDateTime: created,
    messageType: 'message',
    from: { user: { displayName: from } },
    body: { contentType: 'text', content: body },
  };
}

interface StubOptions {
  chats?: Record<string, unknown>[];
  nextLink?: string;
  messagesByChat?: Record<string, Record<string, unknown>[]>;
}

type StubGraph = GraphClient & {
  readonly requests: GraphRequestOptions[];
  readonly batchUrls: string[];
};

function stubGraph(options: StubOptions = {}): StubGraph {
  const requests: GraphRequestOptions[] = [];
  const batchUrls: string[] = [];
  return {
    requests,
    batchUrls,
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      requests.push(opts);
      if (opts.path === '/me/chats') {
        const response: GraphResponse<T> = {
          status: 200,
          data: { value: options.chats ?? [] } as T,
        };
        if (options.nextLink !== undefined) response.nextLink = options.nextLink;
        return response;
      }
      throw new Error(`stub has no route for ${opts.path}`);
    },
    async batch(reqs: GraphBatchRequest[]): Promise<GraphBatchResponse[]> {
      return reqs.map((req) => {
        batchUrls.push(req.url);
        const id = decodeURIComponent(/\/chats\/([^/]+)\/messages/.exec(req.url)?.[1] ?? '');
        const rows = options.messagesByChat?.[id];
        if (rows === undefined) return { id: req.id, status: 404 };
        return { id: req.id, status: 200, body: { value: rows } };
      });
    },
    async follow<T = unknown>(): Promise<GraphResponse<T>> {
      throw new Error('stub does not page');
    },
  };
}

function setup(options: StubOptions = {}): { graph: StubGraph; tool: (name: string) => ToolDefinition } {
  const graph = stubGraph(options);
  const defs = chatModule.build({
    graph,
    config,
    auth: scopeReadingOnlyAuth(),
    consentScopes: () => [],
  });
  return {
    graph,
    tool: (name: string) => {
      const found = defs.find((entry) => entry.name === name);
      assert.ok(found !== undefined, `${name} should exist`);
      return found;
    },
  };
}

type Row = Record<string, unknown>;

type ListResult = { count: number; nextLink?: string; skipToken?: string; chats: Row[] };

type ActivityResult = {
  chatsScanned: number;
  count: number;
  range: { from?: string; to?: string };
  activity: Row[];
};

// ---------------------------------------------------------------------------
// chat_list_chats cursor paging
// ---------------------------------------------------------------------------

test('chat_list_chats sends the skipToken back as $skiptoken', async () => {
  const { graph, tool } = setup({ chats: [chat(CHAT_A, 'Alpha', 'group')] });
  await tool('chat_list_chats').handler({ top: 2, skipToken: 'tok-1' });

  assert.equal(graph.requests.length, 1);
  assert.equal(graph.requests[0]?.query?.['$skiptoken'], 'tok-1');
  assert.equal(graph.requests[0]?.query?.['$top'], 2);
});

test('chat_list_chats surfaces the next page token beside the link', async () => {
  const nextLink = 'https://graph.microsoft.com/v1.0/me/chats?$top=2&$skiptoken=abc123';
  const { tool } = setup({ chats: [], nextLink });
  const result = (await tool('chat_list_chats').handler({ top: 2 })) as unknown as ListResult;

  assert.equal(result.nextLink, nextLink);
  assert.equal(result.skipToken, 'abc123');
});

test('chat_list_chats reads the capitalised $skipToken spelling too', async () => {
  const nextLink = 'https://graph.microsoft.com/v1.0/me/chats?$top=2&$skipToken=xyz789';
  const { tool } = setup({ chats: [], nextLink });
  const result = (await tool('chat_list_chats').handler({})) as unknown as ListResult;

  assert.equal(result.skipToken, 'xyz789');
});

test('chat_list_chats omits the token when there is no next page', async () => {
  const { tool } = setup({ chats: [chat(CHAT_A, 'Alpha', 'group')] });
  const result = (await tool('chat_list_chats').handler({})) as unknown as ListResult;

  assert.equal(result.nextLink, undefined);
  assert.equal(result.skipToken, undefined);
  assert.equal(result.count, 1);
});

// ---------------------------------------------------------------------------
// chat_recent_activity Tier 1: one server-ordered request
// ---------------------------------------------------------------------------

const TIER1_CHATS = [
  chat(CHAT_A, 'Alpha', 'group', preview('m-a1', '2026-10-10T09:00:00Z', 'Ada', 'alpha latest')),
  chat(
    CHAT_B,
    null,
    'oneOnOne',
    preview('m-b1', '2026-10-09T08:00:00Z', 'Bob', '<p>hi <b>there</b></p>', 'html'),
  ),
];

test('Tier 1 is one request ordered by the expanded preview', async () => {
  const { graph, tool } = setup({ chats: TIER1_CHATS });
  const result = (await tool('chat_recent_activity').handler({})) as unknown as ActivityResult;

  assert.equal(graph.requests.length, 1);
  assert.equal(graph.requests[0]?.path, '/me/chats');
  assert.equal(graph.requests[0]?.query?.['$orderby'], 'lastMessagePreview/createdDateTime desc');
  assert.equal(graph.requests[0]?.query?.['$expand'], 'lastMessagePreview');
  assert.equal(graph.requests[0]?.query?.['$top'], 10);
  assert.equal(graph.batchUrls.length, 0);

  // Server order is kept: no client re-sort on Tier 1.
  assert.equal(result.chatsScanned, 2);
  assert.equal(result.count, 2);
  assert.deepEqual(
    result.activity.map((row) => row['id']),
    ['m-a1', 'm-b1'],
  );
  assert.deepEqual(result.range, { from: '2026-10-09T08:00:00Z', to: '2026-10-10T09:00:00Z' });
});

test('Tier 1 rows carry the chat beside the message', async () => {
  const { tool } = setup({ chats: TIER1_CHATS });
  const result = (await tool('chat_recent_activity').handler({})) as unknown as ActivityResult;

  assert.deepEqual(result.activity[0], {
    chatId: CHAT_A,
    chatTopic: 'Alpha',
    chatType: 'group',
    id: 'm-a1',
    createdDateTime: '2026-10-10T09:00:00Z',
    from: 'Ada',
    body: 'alpha latest',
    attachments: undefined,
    messageType: undefined,
    deleted: undefined,
  });
  // A one-to-one chat has no topic; the HTML preview arrives as plain text.
  assert.equal(result.activity[1]?.['chatId'], CHAT_B);
  assert.equal(result.activity[1]?.['chatTopic'], undefined);
  assert.equal(result.activity[1]?.['from'], 'Bob');
  assert.equal(result.activity[1]?.['body'], 'hi there');
});

test('a chat with no preview is scanned but contributes no row', async () => {
  const { tool } = setup({
    chats: [
      chat(CHAT_A, 'Alpha', 'group', preview('m-a1', '2026-10-10T09:00:00Z', 'Ada', 'alpha')),
      chat('19:quiet@thread.v2', 'Quiet', 'group'),
    ],
  });
  const result = (await tool('chat_recent_activity').handler({})) as unknown as ActivityResult;

  assert.equal(result.chatsScanned, 2);
  assert.equal(result.count, 1);
  assert.equal(result.activity[0]?.['chatId'], CHAT_A);
});

test('a system preview is filtered unless includeSystem is set', async () => {
  const system = chat(CHAT_A, 'Alpha', 'group', {
    id: 'm-sys',
    createdDateTime: '2026-10-10T09:00:00Z',
    messageType: 'systemEventMessage',
    from: null,
    body: { contentType: 'text', content: '' },
  });
  const { tool } = setup({ chats: [system] });

  const filtered = (await tool('chat_recent_activity').handler({})) as unknown as ActivityResult;
  assert.equal(filtered.count, 0);

  const kept = (await tool('chat_recent_activity').handler({
    includeSystem: true,
  })) as unknown as ActivityResult;
  assert.equal(kept.count, 1);
  assert.equal(kept.activity[0]?.['from'], null);
});

// ---------------------------------------------------------------------------
// chat_recent_activity Tier 2: one batched fetch, merged newest-first
// ---------------------------------------------------------------------------

const TIER2_MESSAGES: Record<string, Record<string, unknown>[]> = {
  [CHAT_A]: [
    message('a-new', '2026-10-10T10:00:00Z', 'Ada', 'a new'),
    message('a-old', '2026-10-10T08:00:00Z', 'Ada', 'a old'),
  ],
  [CHAT_B]: [
    message('b-new', '2026-10-10T09:00:00Z', 'Bob', 'b new'),
    message('b-old', '2026-10-10T07:00:00Z', 'Bob', 'b old'),
  ],
};

test('Tier 2 fetches each chat once and merges newest-first', async () => {
  const { graph, tool } = setup({
    chats: [chat(CHAT_A, 'Alpha', 'group'), chat(CHAT_B, null, 'oneOnOne')],
    messagesByChat: TIER2_MESSAGES,
  });
  const result = (await tool('chat_recent_activity').handler({
    messagesPerChat: 2,
  })) as unknown as ActivityResult;

  // One listing plus one batch of two entries.
  assert.equal(graph.requests.length, 1);
  assert.equal(graph.batchUrls.length, 2);
  for (const url of graph.batchUrls) {
    const [path, query] = url.split('?') as [string, string];
    assert.match(path, /^\/chats\/[^/]+\/messages$/);
    const params = new URLSearchParams(query);
    assert.equal(params.get('$top'), '2');
    assert.equal(params.get('$orderby'), 'createdDateTime desc');
  }
  assert.ok(
    graph.batchUrls.some((url) => url.startsWith(`/chats/${encodeURIComponent(CHAT_A)}/messages?`)),
    'chat A is fetched by its encoded id',
  );

  assert.equal(result.chatsScanned, 2);
  assert.equal(result.count, 4);
  assert.deepEqual(
    result.activity.map((row) => row['id']),
    ['a-new', 'b-new', 'a-old', 'b-old'],
  );
  assert.deepEqual(
    result.activity.map((row) => row['chatId']),
    [CHAT_A, CHAT_B, CHAT_A, CHAT_B],
  );
});

test('Tier 2 skips an unreadable chat without losing the others', async () => {
  const { tool } = setup({
    chats: [chat(CHAT_A, 'Alpha', 'group'), chat(CHAT_B, null, 'oneOnOne')],
    // No entry for CHAT_B: the stub answers 404 for it.
    messagesByChat: { [CHAT_A]: [message('a-new', '2026-10-10T10:00:00Z', 'Ada', 'a new')] },
  });
  const result = (await tool('chat_recent_activity').handler({
    messagesPerChat: 2,
  })) as unknown as ActivityResult;

  assert.equal(result.chatsScanned, 2);
  assert.equal(result.count, 1);
  assert.equal(result.activity[0]?.['id'], 'a-new');
});

// ---------------------------------------------------------------------------
// Shape and safety rails
// ---------------------------------------------------------------------------

test('a chat id that could escape its path segment is rejected', async () => {
  const { tool } = setup();
  await assert.rejects(
    () => tool('chat_list_messages').handler({ chatId: '19:aaa/bbb@thread.v2' }),
    /slash/,
  );
});

test('activity is the only array at the top level', async () => {
  const { tool } = setup({ chats: TIER1_CHATS });
  const result = (await tool('chat_recent_activity').handler({})) as unknown as Row;

  const arrays = Object.entries(result)
    .filter(([, value]) => Array.isArray(value))
    .map(([key]) => key);
  assert.deepEqual(arrays, ['activity']);
});

test('the result still parses as JSON when the serialiser has to cut it', async () => {
  // Same rule as mail_fetch_thread: serializeResult shortens the largest
  // top-level array, so a second array would risk a character-cut fragment.
  const { tool } = setup({
    chats: [chat(CHAT_A, 'Alpha', 'group'), chat(CHAT_B, null, 'oneOnOne')],
    messagesByChat: {
      [CHAT_A]: [message('a-new', '2026-10-10T10:00:00Z', 'Ada', `${'Long update. '.repeat(60)}`)],
      [CHAT_B]: [message('b-new', '2026-10-10T09:00:00Z', 'Bob', `${'Another update. '.repeat(60)}`)],
    },
  });
  const result = await tool('chat_recent_activity').handler({ messagesPerChat: 2 });

  const whole = serializeResult(result, 10_000_000);
  assert.equal(whole.truncated, false, 'the fixture must be under a huge cap');

  const cut = serializeResult(result, 500);
  assert.equal(cut.truncated, true, 'the fixture must be over a small one');
  const parsed = JSON.parse(cut.text) as { truncated?: { field?: string } };
  assert.equal(parsed.truncated?.field, 'activity', 'activity is what should shrink');
});
