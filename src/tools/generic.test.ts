import assert from 'node:assert/strict';
import test from 'node:test';

import type {
  AuthProvider,
  DeviceLoginPrompt,
  DeviceLoginStatus,
  GraphBatchResponse,
  GraphClient,
  GraphResponse,
  ServerConfig,
  ToolDefinition,
} from '../contracts.js';
import { collectTools, exposedToolNames } from '../server.js';
import { consentScopesFor } from './groups.js';
import { scopeReadingOnlyAuth } from '../auth/index.js';
import { genericModule } from './generic.js';

/**
 * These exist because of a gap a user found rather than a bug a test caught.
 *
 * There was no way to sign in from inside a conversation. `login` was a CLI
 * subcommand, so someone installing this as a Claude Code plugin had to go and
 * find a terminal before the server could do anything at all — and nothing in
 * 102 tools said so. The browser flow genuinely cannot run here, because the MCP
 * transport owns stdio, but the device authorization grant never needed stdio.
 * Leaving it out was an oversight.
 *
 * So what is locked in below is mostly presence and shape: the tools exist, they
 * survive `--read-only`, they are reachable in discovery mode, and they never
 * return a token.
 */

const config: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['me', 'mail'],
  readOnly: false,
  graphVersion: 'v1.0',
  allowBeta: false,
  discovery: false,
  orgMode: false,
  maxOutputChars: 60_000,
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'device',
  authPort: 0,
  allowGenericWrite: false,
};

const PROMPT: DeviceLoginPrompt = {
  userCode: 'ABCD-EFGH',
  verificationUri: 'https://example.test/device',
  expiresAt: '2099-01-01T00:00:00.000Z',
  pollIntervalSeconds: 5,
};

function inertGraph(): GraphClient {
  const refuse = async (): Promise<never> => {
    throw new Error('these tests never reach Graph');
  };
  return {
    request: refuse as unknown as GraphClient['request'],
    batch: refuse as unknown as () => Promise<GraphBatchResponse[]>,
    follow: refuse as unknown as GraphClient['follow'],
  };
}

/** Records the scopes a sign-in asked for, and reports whatever status is set. */
function stubAuth(
  status: DeviceLoginStatus = { state: 'pending', prompt: PROMPT },
): AuthProvider & { readonly asked: string[][] } {
  const asked: string[][] = [];
  return {
    asked,
    async getToken() {
      throw new Error('not exercised');
    },
    async getAccount() {
      return null;
    },
    async login() {
      throw new Error('not exercised');
    },
    async logout() {
      // not exercised
    },
    async beginDeviceLogin(scopes: string[]) {
      asked.push([...scopes]);
      return PROMPT;
    },
    async deviceLoginStatus() {
      return status;
    },
  };
}

/**
 * Builds the module the way the server does, with the consent set the server
 * would compute — the union of the group catalogue and every tool's own extras,
 * which is what the CLI `login` asks for. The sign-in tool asking for anything
 * narrower is the bug these tests exist to hold shut.
 */
function toolsFor(
  over: Partial<ServerConfig> = {},
  auth: AuthProvider = stubAuth(),
): ToolDefinition[] {
  const merged = { ...config, ...over };
  let built: ToolDefinition[] = [];
  built = collectTools({
    graph: inertGraph(),
    config: merged,
    auth,
    consentScopes: () => consentScopesFor(merged, built),
  });
  return built;
}

function byName(tools: ToolDefinition[], name: string): ToolDefinition {
  const found = tools.find((tool) => tool.name === name);
  assert.ok(found !== undefined, `${name} should exist`);
  return found;
}

// ---------------------------------------------------------------------------
// Presence
// ---------------------------------------------------------------------------

test('the generic module offers a way to sign in', async () => {
  const names = toolsFor().map((tool) => tool.name);
  assert.ok(names.includes('auth_begin_login'));
  assert.ok(names.includes('auth_login_status'));
});

test('signing in survives --read-only', async () => {
  // --read-only exists to stop this server changing anything in the tenant, and
  // signing in changes nothing there. Marking it a write would leave every
  // read-only install permanently unable to authenticate.
  const tools = toolsFor({ readOnly: true });
  const begin = byName(tools, 'auth_begin_login');
  assert.notEqual(begin.write, true, 'auth_begin_login must not be flagged as a write');
  assert.notEqual(byName(tools, 'auth_login_status').write, true);
});

test('the served catalogue carries the sign-in tools', async () => {
  // One layer out from the module: whatever collectTools assembles for a real
  // session has to include them, because a caller with no session has nothing
  // else to call.
  const names = toolsFor().map((t) => t.name);
  assert.ok(names.includes('auth_begin_login'));
  assert.ok(names.includes('auth_login_status'));
});

test('discovery mode exposes signing in without a lookup', async () => {
  // In discovery mode the client sees a handful of tools rather than a hundred.
  // A caller with no session cannot reach a tool through call_tool that it has
  // not discovered, and signing in is the first thing such a caller needs — so
  // assert on what the server actually exposes, not on the catalogue it was
  // built from.
  const merged = { ...config, discovery: true };
  const names = exposedToolNames(toolsFor({ discovery: true }), merged);

  assert.ok(names.includes('auth_begin_login'), `exposed: ${names.join(', ')}`);
  assert.ok(names.includes('auth_login_status'));
  assert.ok(names.includes('discover_tools'), 'the discovery pair is still there');
  assert.ok(!names.includes('mail_list_messages'), 'and the catalogue is still hidden');
});

// ---------------------------------------------------------------------------
// Behaviour
// ---------------------------------------------------------------------------

test('starting a sign-in returns a code and a URL, and no token', async () => {
  const result = (await byName(toolsFor(), 'auth_begin_login').handler({})) as Record<string, unknown>;

  assert.equal(result['userCode'], PROMPT.userCode);
  assert.equal(result['verificationUri'], PROMPT.verificationUri);
  assert.ok(String(result['nextStep']).includes('auth_login_status'), 'it says what to call next');

  const serialized = JSON.stringify(result);
  for (const forbidden of ['accessToken', 'access_token', 'refreshToken', 'Bearer']) {
    assert.ok(!serialized.includes(forbidden), `${forbidden} must not be returned`);
  }
});

test('it asks for exactly what the CLI login asks for', async () => {
  // The defect: this tool used to request the group scopes alone while `login`
  // requested those PLUS the extras individual tools declare. A default install
  // that signed in from the conversation was then left with five tools —
  // me_get_mailbox_settings, calendar_get_schedule, calendar_find_meeting_times,
  // files_list_shared and search_query — returning 403 for ever, and nothing
  // suggested that signing in again would not help.
  const auth = stubAuth();
  const tools = toolsFor({ groups: ['me', 'mail', 'calendar', 'files', 'search'] }, auth);
  await byName(tools, 'auth_begin_login').handler({});

  const expected = consentScopesFor(
    { groups: ['me', 'mail', 'calendar', 'files', 'search'], readOnly: config.readOnly },
    tools,
  );
  assert.deepEqual([...(auth.asked[0] as string[])].sort(), expected);
  // The per-tool extras are the whole point, so name one explicitly: a future
  // refactor that went back to group scopes would still pass a set comparison
  // against its own wrong answer.
  assert.ok(
    (auth.asked[0] as string[]).includes('MailboxSettings.Read'),
    'a scope no group meta lists, declared by me_get_mailbox_settings',
  );
});

test('a narrower group list still means a smaller consent prompt', async () => {
  const narrow = stubAuth();
  await byName(toolsFor({ groups: ['me'] }, narrow), 'auth_begin_login').handler({});
  const wide = stubAuth();
  await byName(toolsFor({ groups: ['me', 'mail', 'calendar'] }, wide), 'auth_begin_login').handler({});

  assert.ok(narrow.asked[0] !== undefined && wide.asked[0] !== undefined);
  assert.ok((wide.asked[0] as string[]).length > (narrow.asked[0] as string[]).length);
  assert.ok(!(narrow.asked[0] as string[]).includes('Mail.Read'), 'a mail scope with no mail group');
});

test('status is passed through unchanged, in every state', async () => {
  for (const status of [
    { state: 'none' } as const,
    { state: 'pending', prompt: PROMPT } as const,
    { state: 'signedIn', account: 'someone@example.test' } as const,
    { state: 'expired', detail: 'timed out', prompt: PROMPT } as const,
    { state: 'failed', detail: 'declined' } as const,
  ]) {
    const tool = byName(toolsFor({}, stubAuth(status)), 'auth_login_status');
    assert.deepEqual(await tool.handler({}), status);
  }
});

test('neither sign-in tool asks for a Graph scope', async () => {
  // They talk to Entra, not to Graph. A scope declared here would be added to
  // every consent prompt for no reason.
  for (const name of ['auth_begin_login', 'auth_login_status']) {
    assert.deepEqual(byName(toolsFor(), name).scopes, [], `${name} needs no Graph scope`);
  }
});

// ---------------------------------------------------------------------------
// The refusing provider the CLI hands to a scope-only catalogue
// ---------------------------------------------------------------------------

test('the scope-reading provider refuses rather than pretending', async () => {
  // `status` and `permissions` build the catalogue only to read scopes off it.
  // Returning null from these would let a future caller quietly treat "no
  // session" and "no provider" as the same thing.
  const auth = scopeReadingOnlyAuth();
  await assert.rejects(() => auth.getToken(['User.Read']), /read tool scopes/);
  await assert.rejects(() => auth.beginDeviceLogin(['User.Read']), /read tool scopes/);
  await assert.rejects(() => auth.deviceLoginStatus(), /read tool scopes/);
});

test('a scope-only catalogue still lists the sign-in tools', async () => {
  // It must: `permissions` prints what a configuration requests, and a tool
  // missing from that output is a tool a reader does not know exists.
  const names = toolsFor({}, scopeReadingOnlyAuth()).map((tool) => tool.name);
  assert.ok(names.includes('auth_begin_login'));
});
