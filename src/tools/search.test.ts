import assert from 'node:assert/strict';
import { scopeReadingOnlyAuth } from '../auth/index.js';
import test from 'node:test';

import type {
  GraphBatchResponse,
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
  ServerConfig,
  ToolDefinition,
} from '../contracts.js';
import { searchModule } from './search.js';

/**
 * These exist because a search that touched only chat demanded `Sites.Read.All`
 * and `Files.Read` as well, and failed with AADSTS65001 on a mailbox where an
 * administrator had never granted them. The tool never read those indexes; the
 * scopes rode along because they were bundled at the group level.
 *
 * Measured on the mailbox where it happened: the failure sent the caller down a
 * path of fetching whole chat rooms and scanning them by hand — 110,000
 * characters of tool output to answer something the search answers in 900.
 *
 * Every Graph call is stubbed, so this needs no account and no network.
 */

const config: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['search'],
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

interface Sent {
  entityTypes: string[];
  scopes: string[];
}

/** Records what each split asked for, and answers with one empty hit container. */
function stubGraph(): GraphClient & { readonly sent: Sent[] } {
  const sent: Sent[] = [];
  return {
    sent,
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      const body = opts.body as { requests?: Array<{ entityTypes?: string[] }> };
      sent.push({
        entityTypes: body?.requests?.[0]?.entityTypes ?? [],
        scopes: [...opts.scopes],
      });
      return {
        status: 200,
        data: { value: [{ hitsContainers: [{ hits: [], total: 0 }] }] } as T,
      };
    },
    async batch(): Promise<GraphBatchResponse[]> {
      throw new Error('search does not batch');
    },
    async follow<T = unknown>(): Promise<GraphResponse<T>> {
      throw new Error('search does not page by link');
    },
  };
}

function build(): { tool: ToolDefinition; graph: ReturnType<typeof stubGraph> } {
  const graph = stubGraph();
  const found = searchModule.build({ graph, config, auth: scopeReadingOnlyAuth(), consentScopes: () => [] }).find((t) => t.name === 'search_query');
  assert.ok(found !== undefined, 'search_query should exist');
  return { tool: found, graph };
}

function splitFor(sent: Sent[], entityType: string): Sent {
  const found = sent.find((s) => s.entityTypes.includes(entityType));
  assert.ok(found !== undefined, `expected a split covering ${entityType}`);
  return found;
}

test('a chat-only search asks for chat scopes and nothing else', async () => {
  // The regression: this split used to carry Sites.Read.All and Files.Read too,
  // so it failed for anyone without an administrator — for two indexes it never
  // reads.
  const { tool, graph } = build();
  await tool.handler({ query: 'workshop', entityTypes: ['chatMessage'] });

  assert.equal(graph.sent.length, 1);
  assert.deepEqual(splitFor(graph.sent, 'chatMessage').scopes, [
    'Chat.Read',
    'ChannelMessage.Read.All',
  ]);
});

test('chatMessage asks for ChannelMessage.Read.All, which Graph also requires', async () => {
  // Graph's own 403: "Access to ChatMessage in Graph API requires the following
  // permissions: Chat.Read or Chat.ReadWrite, ChannelMessage.Read.All."
  const { tool, graph } = build();
  await tool.handler({ query: 'workshop', entityTypes: ['chatMessage'] });
  assert.ok(splitFor(graph.sent, 'chatMessage').scopes.includes('ChannelMessage.Read.All'));
});

test('a mail-only search does not ask for SharePoint or chat', async () => {
  const { tool, graph } = build();
  await tool.handler({ query: 'invoice', entityTypes: ['message'] });

  assert.deepEqual(splitFor(graph.sent, 'message').scopes, ['Mail.Read']);
});

test('a file search asks for the file scopes only', async () => {
  const { tool, graph } = build();
  await tool.handler({ query: 'deck', entityTypes: ['driveItem', 'site'] });

  assert.deepEqual(splitFor(graph.sent, 'driveItem').scopes, ['Files.Read', 'Sites.Read.All']);
});

test('an event search adds Calendars.Read and only that', async () => {
  const { tool, graph } = build();
  await tool.handler({ query: 'standup', entityTypes: ['event'] });

  assert.deepEqual(splitFor(graph.sent, 'event').scopes, ['Calendars.Read']);
});

test('a mixed search splits, and neither half carries the other half’s scopes', async () => {
  // chatMessage cannot share a request with any other entity type, so this is
  // two calls. Before the fix both carried the same bundle.
  const { tool, graph } = build();
  await tool.handler({ query: 'release', entityTypes: ['message', 'chatMessage'] });

  assert.equal(graph.sent.length, 2);
  const mail = splitFor(graph.sent, 'message');
  const chat = splitFor(graph.sent, 'chatMessage');
  assert.ok(!mail.scopes.includes('Chat.Read'), 'the mail split must not ask for chat');
  assert.ok(!chat.scopes.includes('Mail.Read'), 'the chat split must not ask for mail');
  assert.ok(!mail.scopes.includes('Sites.Read.All'), 'nor for SharePoint it never reads');
});

test('the declared scopes stay user-consentable', async () => {
  // ChannelMessage.Read.All needs a tenant administrator and belongs to the
  // `teams` group. Declaring it here would put every personal install behind an
  // admin, because `search` is in the personal preset.
  const { tool } = build();
  assert.ok(!tool.scopes.includes('ChannelMessage.Read.All'));
  assert.ok(tool.scopes.includes('Mail.Read'));
  assert.ok(tool.scopes.includes('Chat.Read'));
});
