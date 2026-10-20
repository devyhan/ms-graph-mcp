import assert from 'node:assert/strict';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';

import type {
  AccountInfo,
  AuthProvider,
  GraphBatchResponse,
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
  IdentityRecord,
  ServerConfig,
  SessionHealth,
  ToolDefinition,
} from '../contracts.js';
import { writeIdentity } from '../auth/token-cache.js';
import { collectTools, exposedToolNames } from '../server.js';
import { MINIMAL_LOGIN_SCOPES, consentScopesFor, loginScopesFor } from './groups.js';
import { scopeReadingOnlyAuth } from '../auth/index.js';
import { genericModule } from './generic.js';

/**
 * These exist because of a gap a user found rather than a bug a test caught.
 *
 * There was no way to sign in from inside a conversation. `login` was a CLI
 * subcommand, so someone installing this as a Claude Code plugin had to go and
 * find a terminal before the server could do anything at all — and nothing in
 * 102 tools said so. An explicit tool call can launch the system browser while
 * MSAL receives the OAuth redirect over a loopback HTTP listener, without
 * writing anything to the MCP-owned stdout stream.
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
  output: 'json',
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'auto',
  authPort: 0,
  allowGenericWrite: false,
};

const ACCOUNT: AccountInfo = {
  username: 'someone@example.test',
  name: 'Someone',
  tenantId: 'tenant-id',
  homeAccountId: 'home-account-id',
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

/** Records the scopes and forced flow while keeping browser sign-in pending. */
function stubAuth(
  loginResult: Promise<AccountInfo> = new Promise(() => {}),
  account: AccountInfo | null = null,
  health: SessionHealth | Error | 'absent' = 'absent',
): AuthProvider & {
  readonly asked: string[][];
  readonly flows: Array<string | undefined>;
} {
  const asked: string[][] = [];
  const flows: Array<string | undefined> = [];
  return {
    asked,
    flows,
    async getToken() {
      throw new Error('not exercised');
    },
    async getAccount() {
      return account;
    },
    login(scopes, flow) {
      asked.push([...scopes]);
      flows.push(flow);
      return loginResult;
    },
    async logout() {
      // not exercised
    },
    async beginDeviceLogin() {
      throw new Error('device sign-in is not exercised');
    },
    async deviceLoginStatus() {
      return { state: 'none' as const };
    },
    // Omitted by default, like the scope-only provider: session_status must
    // report "not probed" rather than calling getToken when it is absent.
    ...(health === 'absent'
      ? {}
      : {
          async sessionHealth() {
            if (health instanceof Error) throw health;
            return health;
          },
        }),
  };
}

/**
 * Builds the module the way the server does, with the consent set the server
 * would compute — the union of the group catalogue and every tool's own extras,
 * which is what `session_status` reports and `login --full` asks for. A bare
 * sign-in asks for less than this by design (sign-in is incremental), and the
 * tests below hold that split shut from both sides.
 */
function toolsFor(
  over: Partial<ServerConfig> = {},
  auth: AuthProvider = stubAuth(),
  graph: GraphClient = inertGraph(),
): ToolDefinition[] {
  const merged = { ...config, ...over };
  let built: ToolDefinition[] = [];
  built = collectTools({
    graph,
    config: merged,
    auth,
    consentScopes: () => consentScopesFor(merged, built),
  });
  return built;
}

/** Answers GET /me with a canned profile; everything else is refused. */
function meGraph(me: Record<string, unknown>): GraphClient {
  const request = (async (): Promise<GraphResponse> => ({ status: 200, data: me }));
  const refuse = async (): Promise<never> => {
    throw new Error('only /me is stubbed');
  };
  return {
    request: request as unknown as GraphClient['request'],
    batch: refuse as unknown as () => Promise<GraphBatchResponse[]>,
    follow: refuse as unknown as GraphClient['follow'],
  };
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

test('starting a sign-in launches the browser flow without waiting, and returns no token', async () => {
  const result = (await byName(toolsFor(), 'auth_begin_login').handler({})) as Record<string, unknown>;

  assert.equal(result['state'], 'pending');
  assert.equal(result['flow'], 'browser');
  assert.ok(Date.parse(String(result['startedAt'])) > 0);
  assert.ok(String(result['nextStep']).includes('auth_login_status'), 'it says what to call next');

  const serialized = JSON.stringify(result);
  for (const forbidden of ['accessToken', 'access_token', 'refreshToken', 'Bearer']) {
    assert.ok(!serialized.includes(forbidden), `${forbidden} must not be returned`);
  }
});

test('starting again while the browser sign-in is pending reuses the attempt', async () => {
  const auth = stubAuth();
  const tools = toolsFor({}, auth);
  const begin = byName(tools, 'auth_begin_login');

  const first = await begin.handler({}) as Record<string, unknown>;
  const second = await begin.handler({}) as Record<string, unknown>;

  assert.equal(auth.asked.length, 1, 'only one interactive flow should be started');
  assert.equal(second['startedAt'], first['startedAt']);
});

test('a bare sign-in asks for the minimal set only', async () => {
  // Sign-in is incremental: the first login must not demand the whole
  // configuration up front, because one scope the tenant withholds from
  // ordinary users would then block every tool, including the allowed ones.
  // Consent grows as tools are first used instead.
  const auth = stubAuth();
  const tools = toolsFor({ groups: ['me', 'mail', 'calendar', 'files', 'search'] }, auth);
  const result = (await byName(tools, 'auth_begin_login').handler({})) as Record<string, unknown>;

  assert.deepEqual([...(auth.asked[0] as string[])].sort(), [...MINIMAL_LOGIN_SCOPES].sort());
  assert.equal(auth.flows[0], 'browser', 'the MCP tool must not fall back to device code');
  assert.ok(!(auth.asked[0] as string[]).includes('Mail.Read'), 'no group scope unasked');
  assert.ok(
    !(auth.asked[0] as string[]).includes('MailboxSettings.Read'),
    'no per-tool extra unasked',
  );
  assert.deepEqual(
    [...(result['scopesRequested'] as string[])].sort(),
    [...MINIMAL_LOGIN_SCOPES].sort(),
  );
});

test('named scopes are added to the minimal set', async () => {
  // The same rule the CLI `login --scopes` applies: minimal plus whatever the
  // caller named, so a repeat sign-in consents to just what was added.
  const auth = stubAuth();
  const tools = toolsFor({ groups: ['me', 'mail'] }, auth);
  await byName(tools, 'auth_begin_login').handler({ scopes: ['Mail.Read', ' Mail.Read '] });

  assert.deepEqual(auth.asked[0], loginScopesFor(['Mail.Read']));
  assert.ok((auth.asked[0] as string[]).includes('Mail.Read'));
  assert.ok(!(auth.asked[0] as string[]).includes('Chat.Read'), 'nothing beyond the ask');
});

test('mode full consents up front minus what an ordinary user cannot grant', async () => {
  // The middle ground between bare and ask-everything: the full consent set
  // for the enabled groups, minus the admin-consent and tenant-policy-blocked
  // scopes that would fail the whole prompt.
  const auth = stubAuth();
  const tools = toolsFor({ groups: ['me', 'mail', 'files'] }, auth);
  const result = (await byName(tools, 'auth_begin_login').handler({ mode: 'full' })) as Record<
    string,
    unknown
  >;

  const asked = auth.asked[0] as string[];
  assert.ok(asked.includes('User.Read'), 'the minimal set is included');
  assert.ok(asked.includes('offline_access'), 'the minimal set is included');
  assert.ok(asked.includes('Files.Read'), 'a grantable group scope is included');
  assert.ok(asked.includes('Mail.Send'), 'a grantable write scope is included');
  for (const blocked of ['Mail.Read', 'Mail.ReadWrite', 'MailboxSettings.Read', 'Files.Read.All']) {
    assert.ok(!asked.includes(blocked), `${blocked} must be left out of the upfront ask`);
  }
  assert.deepEqual([...asked].sort(), asked, 'sorted');
  assert.equal(auth.flows[0], 'browser');
  assert.deepEqual(result['scopesRequested'], asked);
});

test('mode full still honours explicitly named scopes', async () => {
  // Named means asked-for: silently dropping a scope the caller named would
  // report success for a consent it thinks it requested.
  const auth = stubAuth();
  const tools = toolsFor({ groups: ['me', 'mail'] }, auth);
  await byName(tools, 'auth_begin_login').handler({ mode: 'full', scopes: ['Chat.Read'] });

  assert.ok((auth.asked[0] as string[]).includes('Chat.Read'));
  assert.ok((auth.asked[0] as string[]).includes('Files.Read') === false);
});

test('mode bare behaves like the default', async () => {
  const auth = stubAuth();
  const tools = toolsFor({ groups: ['me', 'mail', 'files'] }, auth);
  await byName(tools, 'auth_begin_login').handler({ mode: 'bare' });

  assert.deepEqual([...(auth.asked[0] as string[])].sort(), [...MINIMAL_LOGIN_SCOPES].sort());
});

test('an empty scopes list is rejected rather than signed in bare', async () => {
  // Silently dropping to the minimal set would report success for a consent
  // the caller thinks it asked for. Zod rejects it before MSAL is reached.
  const auth = stubAuth();
  const tools = toolsFor({}, auth);
  await assert.rejects(() => byName(tools, 'auth_begin_login').handler({ scopes: [] }));
  assert.equal(auth.asked.length, 0, 'no interactive flow should have started');
});

test('a pending reuse reports the in-flight scopes, not the new ask', async () => {
  // The second call reuses the running attempt, so answering with the scopes
  // it named would describe a consent prompt that does not exist.
  const auth = stubAuth();
  const begin = byName(toolsFor({}, auth), 'auth_begin_login');

  const first = (await begin.handler({ scopes: ['Mail.Read'] })) as Record<string, unknown>;
  const second = (await begin.handler({ scopes: ['Chat.Read'] })) as Record<string, unknown>;

  assert.equal(auth.asked.length, 1, 'only one interactive flow should be started');
  assert.deepEqual(second['scopesRequested'], first['scopesRequested']);
  assert.deepEqual(second['scopesRequested'], loginScopesFor(['Mail.Read']));
});

test('status moves from none to pending to signed in', async () => {
  let finish: (account: AccountInfo) => void = () => {};
  const result = new Promise<AccountInfo>((resolve) => {
    finish = resolve;
  });
  const tools = toolsFor({}, stubAuth(result));
  const status = byName(tools, 'auth_login_status');

  assert.deepEqual(await status.handler({}), { state: 'none' });
  await byName(tools, 'auth_begin_login').handler({});
  assert.equal((await status.handler({}) as Record<string, unknown>)['state'], 'pending');

  finish(ACCOUNT);
  await new Promise((resolve) => setImmediate(resolve));
  const signedIn = await status.handler({}) as Record<string, unknown>;
  assert.equal(signedIn['state'], 'signedIn');
  assert.equal(signedIn['account'], ACCOUNT.username);
  assert.ok(Date.parse(String(signedIn['startedAt'])) > 0);
});

test('a browser sign-in failure is reported without leaking credentials', async () => {
  let fail: (error: Error) => void = () => {};
  const result = new Promise<AccountInfo>((_resolve, reject) => {
    fail = reject;
  });
  const tools = toolsFor({}, stubAuth(result));
  const status = byName(tools, 'auth_login_status');

  await byName(tools, 'auth_begin_login').handler({});
  fail(new Error('AADSTS53003: blocked by conditional access'));
  await new Promise((resolve) => setImmediate(resolve));

  const reported = await status.handler({});
  assert.equal((reported as Record<string, unknown>)['state'], 'failed');
  // Mapped, not raw: the Conditional Access guidance, still quoting the code.
  assert.match(String((reported as Record<string, unknown>)['detail']), /Conditional Access policy/);
  assert.match(String((reported as Record<string, unknown>)['detail']), /prefer the browser/);
  assert.match(String((reported as Record<string, unknown>)['detail']), /AADSTS53003/);
  assert.doesNotMatch(JSON.stringify(reported), /accessToken|refreshToken|Bearer/);
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

// ---------------------------------------------------------------------------
// session_status: the unified first-call probe
// ---------------------------------------------------------------------------

test('the catalogue offers a unified session probe', async () => {
  const probe = byName(toolsFor(), 'session_status');
  assert.notEqual(probe.write, true, 'session_status must not be flagged as a write');
  assert.deepEqual(probe.scopes, ['User.Read'], 'the live /me leg needs only User.Read');
  assert.ok(probe.description.includes('Call once first'));
});

test('discovery mode exposes the session probe without a lookup', async () => {
  // It is the first call an agent makes; hiding it behind discover_tools
  // would defeat the point.
  const merged = { ...config, discovery: true };
  const names = exposedToolNames(toolsFor({ discovery: true }), merged);
  assert.ok(names.includes('session_status'), `exposed: ${names.join(', ')}`);
});

test('session_status reports the cached account plus live reachability', async () => {
  const tools = toolsFor(
    {},
    stubAuth(new Promise(() => {}), ACCOUNT),
    meGraph({
      id: 'user-id',
      displayName: 'Someone',
      userPrincipalName: 'someone@example.test',
    }),
  );
  const result = (await byName(tools, 'session_status').handler({})) as Record<string, unknown>;

  const cached = result['cachedAccount'] as Record<string, unknown>;
  assert.equal(cached['username'], ACCOUNT.username);
  assert.equal(cached['tenantId'], ACCOUNT.tenantId);

  const live = result['liveAccount'] as Record<string, unknown>;
  assert.equal(live['signedIn'], true);
  assert.equal(live['displayName'], 'Someone');

  assert.equal((result['browserLogin'] as Record<string, unknown>)['state'], 'none');
  assert.ok((result['scopesRequested'] as string[]).includes('User.Read'));
  assert.doesNotMatch(JSON.stringify(result), /accessToken|refreshToken|Bearer/);
});

test('session_status survives a dead cache and a dead Graph', async () => {
  // A probe must not fail when the thing it probes is broken: each branch
  // reports its own failure and the rest still answers.
  const auth: AuthProvider = {
    ...stubAuth(),
    getAccount: async () => {
      throw new Error('cache locked');
    },
  };
  const result = (await byName(toolsFor({}, auth), 'session_status').handler({})) as Record<
    string,
    unknown
  >;

  assert.equal(result['cachedAccount'], null);
  assert.match(String(result['cacheNote']), /cache locked/);
  const live = result['liveAccount'] as Record<string, unknown>;
  assert.equal(live['signedIn'], false);
  assert.ok(typeof live['note'] === 'string');
  assert.ok(Array.isArray(result['scopesRequested']));
});

test('session_status reports whether the stored session still mints tokens', async () => {
  // The honest half of "signed in": a cached account whose refresh fails must
  // not read as healthy.
  const usable = toolsFor(
    {},
    stubAuth(new Promise(() => {}), ACCOUNT, { tokenUsable: true }),
    meGraph({ id: 'user-id' }),
  );
  const good = (await byName(usable, 'session_status').handler({})) as Record<string, unknown>;
  assert.equal(good['tokenUsable'], true);
  assert.equal(good['sessionNote'], undefined);

  const broken = toolsFor(
    {},
    stubAuth(new Promise(() => {}), ACCOUNT, {
      tokenUsable: false,
      note: 'Missing consent for Mail.Read; sign in again for just those scopes.',
    }),
    meGraph({ id: 'user-id' }),
  );
  const bad = (await byName(broken, 'session_status').handler({})) as Record<string, unknown>;
  assert.equal(bad['tokenUsable'], false);
  assert.match(String(bad['sessionNote']), /Missing consent for Mail\.Read/);
  // And the cached account is still reported alongside: the two halves agree
  // on what is wrong rather than one hiding the other.
  assert.equal((bad['cachedAccount'] as Record<string, unknown>)['username'], ACCOUNT.username);
  assert.doesNotMatch(JSON.stringify(bad), /accessToken|refreshToken|Bearer/);
});

test('session_status reports not probed when there is no session to probe', async () => {
  // The scope-only catalogue refuses getToken; the probe must not call it.
  const tools = toolsFor({}, scopeReadingOnlyAuth());
  const result = (await byName(tools, 'session_status').handler({})) as Record<string, unknown>;

  assert.equal(result['tokenUsable'], null);
  assert.match(String(result['sessionNote']), /not probed/);
  assert.match(String(result['cacheNote']), /read tool scopes/);
});

test('session_status survives a probe that throws', async () => {
  const tools = toolsFor(
    {},
    stubAuth(new Promise(() => {}), ACCOUNT, new Error('probe blew up')),
    meGraph({ id: 'user-id' }),
  );
  const result = (await byName(tools, 'session_status').handler({})) as Record<string, unknown>;

  assert.equal(result['tokenUsable'], false);
  assert.match(String(result['sessionNote']), /Could not probe the session/);
});

/** Answers GET /me with a canned profile while counting every request. */
function countingMeGraph(
  me: Record<string, unknown>,
): GraphClient & { readonly requests: GraphRequestOptions[] } {
  const requests: GraphRequestOptions[] = [];
  const refuse = async (): Promise<never> => {
    throw new Error('only /me is stubbed');
  };
  return {
    requests,
    request: (async (opts: GraphRequestOptions) => {
      requests.push(opts);
      return { status: 200, data: me };
    }) as unknown as GraphClient['request'],
    batch: refuse as unknown as () => Promise<GraphBatchResponse[]>,
    follow: refuse as unknown as GraphClient['follow'],
  };
}

function identityRecord(over: Partial<IdentityRecord> = {}): IdentityRecord {
  return {
    homeAccountId: ACCOUNT.homeAccountId,
    username: ACCOUNT.username,
    name: ACCOUNT.name,
    tenantId: ACCOUNT.tenantId,
    clientId: 'client-id',
    authority: 'https://login.microsoftonline.com/common',
    capturedAt: '2026-10-19T00:00:00.000Z',
    source: 'msal-account',
    ...over,
  };
}

test('session_status includes the cached identity without a Graph call for it', async () => {
  const cacheDir = await mkdtemp(join(tmpdir(), 'generic-identity-'));
  await writeIdentity(cacheDir, identityRecord());
  const graph = countingMeGraph({ id: 'user-id' });
  const tools = toolsFor(
    { cacheDir },
    stubAuth(new Promise(() => {}), ACCOUNT),
    graph,
  );
  const result = (await byName(tools, 'session_status').handler({})) as Record<string, unknown>;

  // Projected to what a reader can act on: no account id, client id, or authority.
  assert.deepEqual(result['cachedIdentity'], {
    username: ACCOUNT.username,
    name: ACCOUNT.name,
    tenantId: ACCOUNT.tenantId,
    capturedAt: '2026-10-19T00:00:00.000Z',
    source: 'msal-account',
  });
  // The only Graph call is the live /me leg: the identity leg read the disk.
  assert.equal(graph.requests.length, 1);
  assert.equal(graph.requests[0]?.path, '/me');
});

test('session_status reports a null cached identity when no snapshot exists', async () => {
  const cacheDir = await mkdtemp(join(tmpdir(), 'generic-identity-'));
  const graph = countingMeGraph({ id: 'user-id' });
  const tools = toolsFor(
    { cacheDir },
    stubAuth(new Promise(() => {}), ACCOUNT),
    graph,
  );
  const result = (await byName(tools, 'session_status').handler({})) as Record<string, unknown>;

  assert.equal(result['cachedIdentity'], null);
  assert.equal(graph.requests.length, 1);
});

test('session_status drops a snapshot naming a different account', async () => {
  // Another process signed in since the snapshot was taken: reporting it
  // would name a session that is gone.
  const cacheDir = await mkdtemp(join(tmpdir(), 'generic-identity-'));
  await writeIdentity(
    cacheDir,
    identityRecord({ homeAccountId: 'someone-else.uid.tid', username: 'someone-else@example.test' }),
  );
  const tools = toolsFor(
    { cacheDir },
    stubAuth(new Promise(() => {}), ACCOUNT),
    meGraph({ id: 'user-id' }),
  );
  const result = (await byName(tools, 'session_status').handler({})) as Record<string, unknown>;

  assert.equal(result['cachedIdentity'], null);
  assert.equal((result['cachedAccount'] as Record<string, unknown>)['username'], ACCOUNT.username);
});

test('login status falls back to the cached account when no browser sign-in was started', async () => {
  // The closure only knows about a sign-in started in this session. Without
  // the fallback, an existing CLI session reported "none" and the caller
  // signed in again for nothing.
  const tools = toolsFor({}, stubAuth(new Promise(() => {}), ACCOUNT));
  const reported = (await byName(tools, 'auth_login_status').handler({})) as Record<string, unknown>;
  assert.equal(reported['state'], 'signedIn');
  assert.equal(reported['account'], ACCOUNT.username);
});

test('an in-progress browser sign-in wins over the cached account', async () => {
  const tools = toolsFor({}, stubAuth(new Promise(() => {}), ACCOUNT));
  await byName(tools, 'auth_begin_login').handler({});
  const reported = (await byName(tools, 'auth_login_status').handler({})) as Record<string, unknown>;
  assert.equal(reported['state'], 'pending');
});

// ---------------------------------------------------------------------------
// graph_request scope narrowing
// ---------------------------------------------------------------------------

/** Records every request and answers `{}`: graph_request shape, not Graph data. */
function recordingGraph(): GraphClient & { readonly requests: GraphRequestOptions[] } {
  const requests: GraphRequestOptions[] = [];
  const refuse = async (): Promise<never> => {
    throw new Error('only request is stubbed');
  };
  return {
    requests,
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      requests.push(opts);
      return { status: 200, data: {} as T };
    },
    batch: refuse as unknown as () => Promise<GraphBatchResponse[]>,
    follow: refuse as unknown as GraphClient['follow'],
  };
}

test('graph_request declares the union but asks for the minimal scope', async () => {
  const graph = recordingGraph();
  const tools = toolsFor({ groups: ['me', 'mail', 'chat'] }, stubAuth(), graph);
  const def = byName(tools, 'graph_request');

  // Declared scopes are what consent is asked for, so the union stays — while
  // the token request narrows to what this path needs.
  assert.ok(def.scopes.includes('Mail.Read'));
  assert.ok(def.scopes.includes('Chat.Read'));
  assert.ok(def.scopes.length > 1);

  await def.handler({ path: '/me/chats' });

  assert.equal(graph.requests.length, 1);
  assert.deepEqual(graph.requests[0]?.scopes, ['Chat.Read']);
});

test('graph_request falls back to the union for an unknown path', async () => {
  const graph = recordingGraph();
  const tools = toolsFor({ groups: ['me', 'mail', 'chat'] }, stubAuth(), graph);
  const def = byName(tools, 'graph_request');

  await def.handler({ path: '/some/future/resource' });

  assert.equal(graph.requests.length, 1);
  assert.deepEqual(graph.requests[0]?.scopes, def.scopes);
});

test('an explicit scopes list overrides the path guess', async () => {
  const graph = recordingGraph();
  const tools = toolsFor({}, stubAuth(), graph);

  await byName(tools, 'graph_request').handler({ path: '/me/chats', scopes: ['Chat.ReadWrite'] });

  assert.equal(graph.requests.length, 1);
  assert.deepEqual(graph.requests[0]?.scopes, ['Chat.ReadWrite']);
});

// ---------------------------------------------------------------------------
// graph_list_permissions sidecar presence
// ---------------------------------------------------------------------------

test('graph_list_permissions reports the MIP decryptor sidecar state', async () => {
  const run = (tools: ToolDefinition[]) =>
    byName(tools, 'graph_list_permissions').handler({}) as Promise<Record<string, unknown>>;

  // Unconfigured: a bundled copy or a PATH hit still resolves, so either word
  // is honest and the assertion is the vocabulary, not the value.
  const unconfigured = await run(toolsFor());
  assert.ok(
    unconfigured['mipDecryptor'] === 'resolved' || unconfigured['mipDecryptor'] === 'absent',
  );

  const configured = await run(toolsFor({ mipDecryptor: '/opt/mip-decrypt' }));
  assert.equal(configured['mipDecryptor'], 'resolved');
});
