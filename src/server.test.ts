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
    ),
  );

  assert.match(text, /^Blocked by missing consent/);
  assert.match(text, /not by a missing sign-in/);
  assert.match(text, /Call `auth_begin_login`/);
  assert.match(text, /prompts only for what was added/);
  assert.match(text, /\nDetail: Blocked by missing consent/);
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
