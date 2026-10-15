import assert from 'node:assert/strict';
import test from 'node:test';

import type { AuthProvider, GraphClient, ServerConfig } from './contracts.js';
import { InteractionRequiredError } from './contracts.js';
import {
  AGENT_START_PROMPT,
  SERVER_INSTRUCTIONS,
  agentStartPrompt,
  createServerFactory,
  signInMessage,
} from './server.js';

test('a tenant mismatch is a configuration problem, never a missing sign-in', () => {
  const text = signInMessage(
    new InteractionRequiredError('No usable Microsoft 365 session. AADSTS50194. Set --tenant-id.', [], true),
  );

  assert.match(text, /^Blocked by a server configuration problem/);
  assert.match(text, /do not offer `auth_begin_login`/);
  assert.match(text, /restart the server/);
  assert.match(text, /\nDetail: No usable Microsoft 365 session/);
});

test('the tenant mismatch wins over missing admin consent', () => {
  // Both can be set at once: the endpoint refusal is the error that actually
  // occurred, and admin consent is a bridge crossed only after it is fixed.
  const text = signInMessage(
    new InteractionRequiredError('AADSTS50194 on the refresh.', ['ChannelMessage.Read.All'], true),
  );

  assert.match(text, /^Blocked by a server configuration problem/);
  assert.doesNotMatch(text, /missing tenant admin consent/);
});

test('missing admin consent keeps its headline', () => {
  const text = signInMessage(
    new InteractionRequiredError('No usable Microsoft 365 session.', ['ChannelMessage.Read.All']),
  );

  assert.match(text, /^Blocked by missing tenant admin consent/);
});

test('a plain session failure still offers auth_begin_login', () => {
  const text = signInMessage(new InteractionRequiredError('No usable Microsoft 365 session.'));

  assert.match(text, /^Not signed in/);
  assert.match(text, /Call `auth_begin_login`/);
});

test('missing consent says the session is fine and signing in again is the fix', () => {
  const text = signInMessage(
    new InteractionRequiredError(
      'Blocked by missing consent for these scopes: Chat.Read.',
      [],
      false,
      true,
      false,
      false,
      false,
      ['Chat.Read'],
    ),
  );

  assert.match(text, /^Blocked by missing consent/);
  assert.match(text, /not by a missing sign-in/);
  assert.match(text, /Call `auth_begin_login`/);
  assert.match(text, /prompts only for what was added/);
  // Paste-ready, not a placeholder: the exact tool argument and CLI flag.
  assert.match(text, /Paste-ready: `auth_begin_login` with `\{"scopes": \["Chat\.Read"\]\}`/);
  assert.match(text, /`npx @devyhan\/ms-graph-mcp login --scopes Chat\.Read`/);
  assert.match(text, /\nDetail: Blocked by missing consent/);
});

test('missing consent without recorded scopes stays generic', () => {
  const text = signInMessage(
    new InteractionRequiredError(
      'Blocked by missing consent for these scopes: Chat.Read.',
      [],
      false,
      true,
    ),
  );
  // No scopes recorded, so no paste-ready line — but the headline still stands.
  assert.match(text, /^Blocked by missing consent/);
  assert.doesNotMatch(text, /Paste-ready/);
});

test('missing admin consent wins over missing consent', () => {
  // Both can be set at once: consent the tenant reserves to an administrator
  // is refused the same way, and re-running login cannot grant it.
  const text = signInMessage(
    new InteractionRequiredError('Blocked by missing consent.', ['ChannelMessage.Read.All'], false, true),
  );

  assert.match(text, /^Blocked by missing tenant admin consent/);
  assert.doesNotMatch(text, /^Blocked by missing consent —/m);
});

test('a Conditional Access refusal retries the browser, never the device code', () => {
  const text = signInMessage(
    new InteractionRequiredError('AADSTS53003 on the refresh.', [], false, false, true),
  );

  assert.match(text, /^Blocked by a Conditional Access policy/);
  assert.match(text, /not by a missing sign-in/);
  assert.match(text, /Call `auth_begin_login`/);
  assert.match(text, /commonly blocks it too/);
  assert.match(text, /\nDetail: AADSTS53003/);
});

test('a network failure retries the call, never the sign-in', () => {
  const text = signInMessage(
    new InteractionRequiredError('ENOTFOUND on the refresh.', [], false, false, false, true),
  );

  assert.match(text, /^Blocked by a network failure/);
  assert.match(text, /not by a missing sign-in/);
  assert.match(text, /do not offer `auth_begin_login`/);
  assert.match(text, /--use-env-proxy/);
});

test('a loopback timeout names the device flow as the way out', () => {
  const text = signInMessage(
    new InteractionRequiredError('loopback_server_timeout.', [], false, false, false, false, true),
  );

  assert.match(text, /^Blocked by a loopback redirect timeout/);
  assert.match(text, /not by a missing sign-in/);
  assert.match(text, /--auth-flow device/);
});

test('missing admin consent wins over Conditional Access and network failures', () => {
  // Every other fix fails until an administrator grants those scopes, so the
  // admin headline keeps its rank whatever else is set.
  const policy = signInMessage(
    new InteractionRequiredError('AADSTS53003.', ['ChannelMessage.Read.All'], false, false, true),
  );
  assert.match(policy, /^Blocked by missing tenant admin consent/);
  assert.doesNotMatch(policy, /^Blocked by a Conditional Access policy/m);

  const offline = signInMessage(
    new InteractionRequiredError('ENOTFOUND.', ['ChannelMessage.Read.All'], false, false, false, true),
  );
  assert.match(offline, /^Blocked by missing tenant admin consent/);
  assert.doesNotMatch(offline, /^Blocked by a network failure/m);
});

test('Conditional Access wins over missing consent', () => {
  // Both can be set at once; the policy verdict is the more specific refusal.
  const text = signInMessage(
    new InteractionRequiredError('AADSTS53003 and consent missing.', [], false, true, true),
  );

  assert.match(text, /^Blocked by a Conditional Access policy/);
  assert.doesNotMatch(text, /^Blocked by missing consent —/m);
});

test('the initialize instructions stay the exact onboarding paragraph', () => {
  assert.equal(
    SERVER_INSTRUCTIONS,
    'Call session_status once first. Sign in via auth_begin_login and auth_login_status; ' +
      'never ask for credentials. Orient via discover_tools. On errors read Hint/Blocked-by ' +
      'lines; re-login adds only missing scopes.',
  );
});

test('the agent_start prompt keeps its name, title and description', () => {
  assert.equal(AGENT_START_PROMPT.name, 'agent_start');
  assert.equal(AGENT_START_PROMPT.title, 'Agent start');
  assert.equal(
    AGENT_START_PROMPT.description,
    'Orients a new agent session: check session_status, sign in if needed via auth_begin_login ' +
      'plus auth_login_status, orient via discover_tools, and follow Hint and Blocked-by lines on errors.',
  );
});

test('the agent_start body walks the same onboarding in order', () => {
  const { messages } = agentStartPrompt();

  assert.equal(messages.length, 1);
  assert.equal(messages[0].role, 'user');
  assert.equal(messages[0].content.type, 'text');
  const text = messages[0].content.text;
  const order = [
    'session_status once first',
    'auth_begin_login',
    'auth_login_status',
    'Never ask for credentials',
    'discover_tools',
    'Hint and Blocked-by',
    'prompts only for what was added',
  ];
  let at = -1;
  for (const phrase of order) {
    const next = text.indexOf(phrase, at + 1);
    assert.ok(next > at, `expected ${JSON.stringify(phrase)} after offset ${at}`);
    at = next;
  }
});

test('the factory advertises instructions and registers agent_start', () => {
  // The factory only reads config and tools while registering, so Graph and
  // auth are stubs that fail if anything actually calls them.
  const unreachable = async (): Promise<never> => {
    throw new Error('the factory never calls Graph or auth while registering');
  };
  const graph = { request: unreachable, batch: unreachable, follow: unreachable } as GraphClient;
  const auth = {
    getToken: unreachable,
    getAccount: unreachable,
    login: unreachable,
    logout: unreachable,
    beginDeviceLogin: unreachable,
    deviceLoginStatus: unreachable,
  } as AuthProvider;
  const config: ServerConfig = {
    clientId: '00000000-0000-0000-0000-000000000000',
    tenantId: 'common',
    authority: 'https://login.microsoftonline.com/common',
    graphHost: 'graph.microsoft.com',
    groups: [],
    readOnly: false,
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

  const server = createServerFactory({ config, graph, auth, tools: [] })();

  // Private SDK state, read through a narrow cast: there is no public list of
  // registered prompts, and a test that cannot see the registration cannot
  // tell that the factory forgot it. Pinned against @modelcontextprotocol v2;
  // revisit if the SDK renames these fields.
  const registered = (
    server as unknown as { _registeredPrompts: Record<string, { description?: string }> }
  )._registeredPrompts;
  const inner = (
    server as unknown as { server: { _capabilities: Record<string, unknown>; _instructions?: string } }
  ).server;

  assert.ok('agent_start' in registered, 'expected agent_start among the registered prompts');
  assert.equal(registered['agent_start']?.description, AGENT_START_PROMPT.description);
  assert.ok('prompts' in inner._capabilities, 'expected the prompts capability to be declared');
  assert.equal(inner._instructions, SERVER_INSTRUCTIONS);
});
