import assert from 'node:assert/strict';
import test from 'node:test';

import {
  AUTH_FLOWS,
  CLIENT_ID_SETUP_MESSAGE,
  CLOUDS,
  DEFAULT_AUTH_FLOW,
  DEFAULT_AUTH_PORT,
  DEFAULT_CLOUD,
  DEFAULT_MAX_OUTPUT_CHARS,
  SHIPPED_APP_DISCLOSURE,
  SHIPPED_CLIENT_ID,
  buildClientIdSetupMessage,
  clientIdGuidance,
  describeClientIdSource,
  graphBaseUrl,
  hasShippedClientId,
  isClientIdUnset,
  ConfigError,
  parseConfig,
  unsubstitutedPlaceholder,
  resolveClientId,
} from './config.js';
import { DEFAULT_GROUPS, GROUPS, GROUP_NAMES, scopesForGroups } from './tools/groups.js';

const NO_ENV: NodeJS.ProcessEnv = {};

/** Runs `fn` with stderr captured, so group warnings do not pollute test output. */
function withCapturedStderr<T>(fn: () => T): { result: T; stderr: string } {
  const original = process.stderr.write.bind(process.stderr);
  let captured = '';
  process.stderr.write = ((chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : Buffer.from(chunk).toString('utf8');
    return true;
  }) as typeof process.stderr.write;
  try {
    return { result: fn(), stderr: captured };
  } finally {
    process.stderr.write = original;
  }
}

test('bare invocation serves with the personal default groups', () => {
  const { config, command } = parseConfig([], NO_ENV);
  assert.deepEqual(command, { kind: 'serve' });
  assert.deepEqual(config.groups, DEFAULT_GROUPS);
  assert.equal(config.clientId, SHIPPED_CLIENT_ID);
  assert.equal(isClientIdUnset(config.clientId), !hasShippedClientId());
  assert.equal(config.tenantId, 'common');
  assert.equal(config.graphHost, 'graph.microsoft.com');
  assert.equal(config.authority, 'https://login.microsoftonline.com/common');
  assert.equal(config.maxOutputChars, DEFAULT_MAX_OUTPUT_CHARS);
  assert.equal(config.readOnly, false);
  assert.equal(config.orgMode, false);
  assert.equal(config.graphVersion, 'v1.0');
  assert.ok(config.cacheDir.includes('microsoft-graph-mcp'));
  assert.equal(config.authFlow, 'auto');
  assert.equal(config.authPort, 0);
});

test('--auth-flow defaults to auto and accepts every documented value', () => {
  assert.equal(DEFAULT_AUTH_FLOW, 'auto');
  assert.deepEqual(AUTH_FLOWS, ['auto', 'browser', 'device']);
  assert.equal(parseConfig([], NO_ENV).config.authFlow, 'auto');

  for (const flow of AUTH_FLOWS) {
    assert.equal(parseConfig(['--auth-flow', flow], NO_ENV).config.authFlow, flow, flow);
    assert.equal(parseConfig([`--auth-flow=${flow}`], NO_ENV).config.authFlow, flow, flow);
    assert.equal(
      parseConfig([], { MS365_MCP_AUTH_FLOW: flow }).config.authFlow,
      flow,
      `env ${flow}`,
    );
  }
});

test('the legacy interactive value still selects the browser flow', () => {
  assert.equal(parseConfig([], { MS365_MCP_AUTH_FLOW: 'interactive' }).config.authFlow, 'browser');
  assert.equal(parseConfig(['--auth-flow', 'INTERACTIVE'], NO_ENV).config.authFlow, 'browser');
});

test('an unknown --auth-flow value is rejected and the valid ones listed', () => {
  for (const argv of [['--auth-flow', 'popup'], ['--auth-flow=popup']]) {
    assert.throws(
      () => parseConfig(argv, NO_ENV),
      (error: unknown) => {
        assert.ok(error instanceof Error);
        assert.match(error.message, /Unknown auth flow "popup"/);
        assert.match(error.message, /auto/);
        assert.match(error.message, /browser/);
        assert.match(error.message, /device/);
        return true;
      },
      argv.join(' '),
    );
  }
  assert.throws(
    () => parseConfig([], { MS365_MCP_AUTH_FLOW: 'popup' }),
    /MS365_MCP_AUTH_FLOW.*Valid flows: auto, browser, device/s,
  );
});

test('--auth-port takes 0 or an unprivileged port and rejects the rest', () => {
  assert.equal(DEFAULT_AUTH_PORT, 0);
  assert.equal(parseConfig(['--auth-port', '0'], NO_ENV).config.authPort, 0);
  assert.equal(parseConfig(['--auth-port', '8400'], NO_ENV).config.authPort, 8400);
  assert.equal(parseConfig(['--auth-port=1024'], NO_ENV).config.authPort, 1024);
  assert.equal(parseConfig(['--auth-port', '65535'], NO_ENV).config.authPort, 65535);
  assert.equal(parseConfig([], { MS365_MCP_AUTH_PORT: '8400' }).config.authPort, 8400);

  assert.throws(() => parseConfig(['--auth-port', '80'], NO_ENV), /privileged port/);
  assert.throws(() => parseConfig(['--auth-port', '1023'], NO_ENV), /privileged port/);
  assert.throws(() => parseConfig(['--auth-port', 'abc'], NO_ENV), /must be a whole number/);
  assert.throws(() => parseConfig(['--auth-port', '70000'], NO_ENV), /1024-65535/);
  assert.throws(() => parseConfig(['--auth-port', '8400.5'], NO_ENV), /must be a whole number/);
  assert.throws(() => parseConfig([], { MS365_MCP_AUTH_PORT: '80' }), /MS365_MCP_AUTH_PORT/);
});

test('auth flags on the command line beat the environment', () => {
  const env: NodeJS.ProcessEnv = { MS365_MCP_AUTH_FLOW: 'device', MS365_MCP_AUTH_PORT: '8400' };
  const { config } = parseConfig(['--auth-flow', 'browser', '--auth-port', '9001'], env);
  assert.equal(config.authFlow, 'browser');
  assert.equal(config.authPort, 9001);

  // Nothing on the command line leaves the environment in charge.
  const fromEnvOnly = parseConfig([], env).config;
  assert.equal(fromEnvOnly.authFlow, 'device');
  assert.equal(fromEnvOnly.authPort, 8400);
});

test('help and the setup walkthrough document both flows and the redirect URI', () => {
  const help = parseConfig(['--help'], NO_ENV).command;
  assert.equal(help.kind, 'help');
  const text = help.kind === 'help' ? help.text : '';
  assert.match(text, /--auth-flow/);
  assert.match(text, /--auth-port/);
  assert.match(text, /MS365_MCP_AUTH_FLOW/);
  assert.match(text, /MS365_MCP_AUTH_PORT/);
  assert.match(text, /both are OAuth 2\.0/);
  assert.match(text, /PKCE/);
  assert.match(text, /[Dd]evice authorization grant/);

  assert.match(CLIENT_ID_SETUP_MESSAGE, /Mobile and desktop applications/);
  assert.match(CLIENT_ID_SETUP_MESSAGE, /http:\/\/localhost/);
  assert.match(CLIENT_ID_SETUP_MESSAGE, /Allow public client flows/);
  assert.match(CLIENT_ID_SETUP_MESSAGE, /--auth-port PORT/);
});

test('subcommands and env fallbacks are honoured', () => {
  const { config, command } = parseConfig(['login'], {
    MS365_MCP_CLIENT_ID: 'abc-123',
    MS365_MCP_READ_ONLY: 'true',
    MS365_MCP_MAX_OUTPUT_CHARS: '1234',
  });
  assert.deepEqual(command, { kind: 'login' });
  assert.equal(config.clientId, 'abc-123');
  assert.equal(config.readOnly, true);
  assert.equal(config.maxOutputChars, 1234);
  assert.equal(isClientIdUnset(config.clientId), false);
});

test('--preset all drops admin-consent groups unless --org-mode is set', () => {
  const adminGroups = GROUP_NAMES.filter((name) => GROUPS[name]?.requiresAdminConsent === true);
  assert.deepEqual(adminGroups, ['teams', 'directory', 'intune']);

  const { result, stderr } = withCapturedStderr(() => parseConfig(['--preset', 'all'], NO_ENV));
  for (const name of adminGroups) {
    assert.ok(!result.config.groups.includes(name), `${name} must be dropped`);
  }
  assert.equal(result.config.groups.length, GROUP_NAMES.length - adminGroups.length);
  assert.match(stderr, /--org-mode/);

  const withOrg = parseConfig(['--preset', 'all', '--org-mode'], NO_ENV);
  assert.deepEqual(withOrg.config.groups, GROUP_NAMES);
  assert.equal(withOrg.config.orgMode, true);
});

test('unknown flags are rejected and the valid ones listed', () => {
  assert.throws(
    () => parseConfig(['--not-a-flag'], NO_ENV),
    (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /Unknown option "--not-a-flag"/);
      assert.match(error.message, /--client-id/);
      assert.match(error.message, /--read-only/);
      return true;
    },
  );
});

test('a value flag without a value is an error', () => {
  assert.throws(() => parseConfig(['--client-id'], NO_ENV), /requires a value/);
  assert.throws(() => parseConfig(['--tenant-id', '--verbose'], NO_ENV), /requires a value/);
});

test('unknown groups are rejected and the valid ones listed', () => {
  assert.throws(
    () => parseConfig(['--groups', 'mail,nope'], NO_ENV),
    (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.match(error.message, /Unknown tool group "nope"/);
      assert.match(error.message, /sharepoint/);
      return true;
    },
  );
  assert.throws(() => parseConfig(['--preset', 'nope'], NO_ENV), /Unknown preset "nope"/);
});

test('cloud names map to the right Graph host and authority', () => {
  const cases: Array<[string, string, string]> = [
    ['global', 'graph.microsoft.com', 'https://login.microsoftonline.com/common'],
    ['usgov', 'graph.microsoft.us', 'https://login.microsoftonline.us/common'],
    ['usgovdod', 'dod-graph.microsoft.us', 'https://login.microsoftonline.us/common'],
    ['china', 'microsoftgraph.chinacloudapi.cn', 'https://login.partner.microsoftonline.cn/common'],
  ];
  for (const [cloud, graphHost, authority] of cases) {
    const { config } = parseConfig(['--cloud', cloud], NO_ENV);
    assert.equal(config.graphHost, graphHost, cloud);
    assert.equal(config.authority, authority, cloud);
    assert.equal(graphBaseUrl(config), `https://${graphHost}`);
    assert.ok(!config.graphHost.includes('://'), 'graphHost must be a bare hostname');
  }

  const { config } = parseConfig(['--cloud=china', '--tenant-id=contoso.onmicrosoft.com'], NO_ENV);
  assert.equal(config.authority, 'https://login.partner.microsoftonline.cn/contoso.onmicrosoft.com');
  assert.equal(Object.keys(CLOUDS).length, 4);

  assert.throws(() => parseConfig(['--cloud', 'mars'], NO_ENV), /Unknown cloud "mars"/);
});

test('--read-only removes write scopes but keeps the baseline ones', () => {
  const writable = parseConfig(['--groups', 'mail,calendar'], NO_ENV);
  const writeScopes = scopesForGroups(writable.config.groups, writable.config.readOnly);
  assert.ok(writeScopes.includes('Mail.Send'));
  assert.ok(writeScopes.includes('Calendars.ReadWrite'));

  const readOnly = parseConfig(['--groups', 'mail,calendar', '--read-only'], NO_ENV);
  const readScopes = scopesForGroups(readOnly.config.groups, readOnly.config.readOnly);
  assert.equal(readOnly.config.readOnly, true);
  assert.deepEqual(readScopes, ['Calendars.Read', 'Mail.Read', 'User.Read', 'offline_access']);
  assert.ok(!readScopes.some((scope) => scope.includes('ReadWrite')));
  assert.deepEqual([...readScopes].sort(), readScopes, 'scopes must be sorted');
});

test('help and version short-circuit without touching the rest of argv', () => {
  const help = parseConfig(['--help', '--groups', 'bogus'], NO_ENV);
  assert.equal(help.command.kind, 'help');
  assert.ok(help.command.kind === 'help' && help.command.text.includes('--client-id'));
  assert.ok(help.command.kind === 'help' && help.command.text.includes('sharepoint'));

  assert.equal(parseConfig(['--version'], NO_ENV).command.kind, 'version');
  assert.equal(parseConfig(['status'], NO_ENV).command.kind, 'status');
  assert.equal(parseConfig(['permissions'], NO_ENV).command.kind, 'permissions');
  assert.equal(parseConfig(['logout'], NO_ENV).command.kind, 'logout');
  assert.throws(() => parseConfig(['frobnicate'], NO_ENV), /Unknown command "frobnicate"/);
});

// ---------------------------------------------------------------------------
// The shipped application
// ---------------------------------------------------------------------------

/** A stand-in for a real registration, so the shipped branches stay testable. */
const FAKE_SHIPPED = 'a9bac4c3-0000-0000-0000-000000000001';

/**
 * The one test that pins today's state. When the project owner registers the
 * shared application and fills in SHIPPED_CLIENT_ID, this test flips to
 * expecting the 'shipped' source; every other test here is written to survive
 * that swap untouched.
 */
test('no application ships yet, so a bare invocation resolves to nothing', () => {
  assert.equal(SHIPPED_CLIENT_ID, '');
  assert.equal(hasShippedClientId(), false);

  const parsed = parseConfig([], NO_ENV);
  assert.equal(parsed.config.clientId, '');
  assert.equal(parsed.clientIdSource, 'unset');
  assert.equal(parsed.clientIdReason, undefined);
  assert.ok(isClientIdUnset(parsed.config.clientId));
});

test('the all-zeros placeholder is gone and cannot drift back', async () => {
  // A syntactically valid placeholder reaches Entra and fails with
  // AADSTS700038 only after the password prompt; an empty default fails here.
  const module: Record<string, unknown> = await import('./config.js');
  assert.equal(module['DEFAULT_CLIENT_ID'], undefined);
  assert.equal(module['isPlaceholderClientId'], undefined);
  assert.notEqual(SHIPPED_CLIENT_ID, '00000000-0000-0000-0000-000000000000');
});

test('--client-id beats the environment variable', () => {
  const env: NodeJS.ProcessEnv = { MS365_MCP_CLIENT_ID: 'from-env' };
  const parsed = parseConfig(['--client-id', 'from-flag'], env);
  assert.equal(parsed.config.clientId, 'from-flag');
  assert.equal(parsed.clientIdSource, 'flag');

  const inline = parseConfig(['--client-id=from-flag'], env);
  assert.equal(inline.config.clientId, 'from-flag');
  assert.equal(inline.clientIdSource, 'flag');
});

test('the environment variable alone is reported as the env source', () => {
  const parsed = parseConfig([], { MS365_MCP_CLIENT_ID: 'from-env' });
  assert.equal(parsed.config.clientId, 'from-env');
  assert.equal(parsed.clientIdSource, 'env');
});

test('resolveClientId walks flag, env, shipped, unset in that order', () => {
  assert.deepEqual(
    resolveClientId({ flag: 'f', env: 'e', shipped: FAKE_SHIPPED, cloud: DEFAULT_CLOUD }),
    { clientId: 'f', source: 'flag' },
  );
  assert.deepEqual(resolveClientId({ env: 'e', shipped: FAKE_SHIPPED, cloud: DEFAULT_CLOUD }), {
    clientId: 'e',
    source: 'env',
  });
  assert.deepEqual(resolveClientId({ shipped: FAKE_SHIPPED, cloud: DEFAULT_CLOUD }), {
    clientId: FAKE_SHIPPED,
    source: 'shipped',
  });
  assert.deepEqual(resolveClientId({ shipped: '', cloud: DEFAULT_CLOUD }), {
    clientId: '',
    source: 'unset',
  });

  // Whitespace-only values are not applications.
  assert.equal(resolveClientId({ flag: '  ', shipped: '', cloud: DEFAULT_CLOUD }).source, 'unset');
  assert.equal(
    resolveClientId({ flag: '  ', env: ' e ', shipped: '', cloud: DEFAULT_CLOUD }).clientId,
    'e',
  );
});

test('a sovereign cloud never falls back to the shipped application', () => {
  for (const cloud of ['usgov', 'usgovdod', 'china']) {
    const resolution = resolveClientId({ shipped: FAKE_SHIPPED, cloud });
    assert.equal(resolution.source, 'unset', cloud);
    assert.equal(resolution.clientId, '', cloud);
    assert.match(String(resolution.reason), new RegExp(`--cloud ${cloud}`), cloud);
    assert.match(String(resolution.reason), /global\s+cloud only/, cloud);

    // An explicit ID is still honoured there: only the shipped default is refused.
    const explicit = resolveClientId({ flag: 'mine', shipped: FAKE_SHIPPED, cloud });
    assert.equal(explicit.source, 'flag', cloud);
  }
});

test('--cloud china with no client ID is unset, never shipped', () => {
  const parsed = parseConfig(['--cloud', 'china'], NO_ENV);
  assert.equal(parsed.clientIdSource, 'unset');
  assert.notEqual(parsed.clientIdSource, 'shipped');
  assert.equal(parsed.config.clientId, '');

  const withId = parseConfig(['--cloud', 'china', '--client-id', 'mine'], NO_ENV);
  assert.equal(withId.clientIdSource, 'flag');
  assert.equal(withId.config.clientId, 'mine');
});

test('every client-ID source has a phrase naming the application', () => {
  assert.equal(describeClientIdSource('flag'), '--client-id');
  assert.equal(describeClientIdSource('env'), 'MS365_MCP_CLIENT_ID');
  assert.match(describeClientIdSource('shipped'), /shipped with this package/);
  assert.match(describeClientIdSource('unset'), /not configured/);
});

test('the guidance leads with the refusal reason when there is one', () => {
  assert.equal(clientIdGuidance(), CLIENT_ID_SETUP_MESSAGE);
  assert.equal(clientIdGuidance('   '), CLIENT_ID_SETUP_MESSAGE);

  const reason = String(resolveClientId({ shipped: FAKE_SHIPPED, cloud: 'china' }).reason);
  const guidance = clientIdGuidance(reason);
  assert.ok(guidance.startsWith(reason));
  assert.ok(guidance.endsWith(CLIENT_ID_SETUP_MESSAGE));
});

test('the walkthrough reframes itself once an application ships', () => {
  // Without a default, registering is the only way in.
  const bare = buildClientIdSetupMessage('');
  assert.match(bare, /ships no/);

  // With one, the walkthrough becomes the fallback and says so first.
  const shipped = buildClientIdSetupMessage(FAKE_SHIPPED);
  assert.match(shipped, /You normally do not need this/);
  assert.match(shipped, /shared multi-tenant/);
  assert.match(shipped, /sovereign cloud/);
  assert.ok(!shipped.includes('ships no'));

  // The exported message is whichever variant this build calls for.
  assert.equal(CLIENT_ID_SETUP_MESSAGE, buildClientIdSetupMessage(SHIPPED_CLIENT_ID));

  // Both variants keep the actual steps.
  for (const text of [bare, shipped]) {
    assert.match(text, /To register your own application/);
    assert.match(text, /Mobile and desktop applications/);
    assert.match(text, /Allow public client flows/);
    assert.match(text, /--auth-port PORT/);
    assert.match(text, /requestedAccessTokenVersion = 2/);
    assert.match(text, /--tenant-id consumers/);
  }
});

test('the disclosure names the consent the shared application still needs', () => {
  assert.match(SHIPPED_APP_DISCLOSURE, /enterprise\s+application/);
  assert.match(SHIPPED_APP_DISCLOSURE, /administrator/);
  assert.match(SHIPPED_APP_DISCLOSURE, /Mail\.Read/);
  assert.match(SHIPPED_APP_DISCLOSURE, /--client-id/);
});

test('help tells the user which application sign-in will use', () => {
  const help = parseConfig(['--help'], NO_ENV).command;
  const text = help.kind === 'help' ? help.text : '';
  assert.match(text, /--client-id/);
  assert.match(text, /MS365_MCP_CLIENT_ID/);

  if (hasShippedClientId()) {
    // Registering is optional, and the examples must not imply otherwise.
    assert.match(text, /application shipped with this package/);
    assert.match(text, /consent step/);
    assert.ok(!text.includes('microsoft-graph-mcp login --client-id <id>'));
  } else {
    assert.match(text, /ships no default application/);
    assert.match(text, /microsoft-graph-mcp login --client-id <id>/);
  }
});

// ---------------------------------------------------------------------------
// Unsubstituted placeholders
//
// A Claude Code plugin manifest interpolates ${user_config.key} before the
// value reaches us. When it does not — a skipped setting, a mistyped key, a
// host that does not support the syntax — the literal text arrives instead.
// The tenant is the dangerous one: it is concatenated into the authority URL,
// so it fails later and somewhere else.
// ---------------------------------------------------------------------------

test('an unsubstituted client-id placeholder is refused, not used', () => {
  const { config, clientIdSource, clientIdReason } = parseConfig([], {
    MS365_MCP_CLIENT_ID: '${user_config.client_id}',
  });
  assert.equal(config.clientId, '');
  assert.equal(clientIdSource, 'unset');
  assert.match(String(clientIdReason), /nothing replaced it/);
});

test('an unsubstituted tenant placeholder throws rather than reaching the authority', () => {
  assert.throws(
    () => parseConfig([], { MS365_MCP_TENANT_ID: '${user_config.tenant_id}' }),
    (err) => err instanceof ConfigError && /nothing replaced it/.test(err.message),
  );
});

test('an empty value is absent, not a placeholder failure', () => {
  const { config, clientIdSource } = parseConfig([], {
    MS365_MCP_CLIENT_ID: '',
    MS365_MCP_TENANT_ID: '',
  });
  assert.equal(clientIdSource, 'unset');
  // The server's own default, not the empty string, and never a literal.
  assert.equal(config.tenantId, 'common');
  assert.equal(config.authority, 'https://login.microsoftonline.com/common');
});

test('a real value that merely contains braces is not mistaken for a placeholder', () => {
  const { config, clientIdSource } = parseConfig(
    ['--client-id', '11111111-2222-3333-4444-555555555555'],
    {},
  );
  assert.equal(clientIdSource, 'flag');
  assert.equal(config.clientId, '11111111-2222-3333-4444-555555555555');
});

test('unsubstitutedPlaceholder matches only a whole ${...} token', () => {
  assert.equal(unsubstitutedPlaceholder('${user_config.client_id}'), 'user_config.client_id');
  assert.equal(unsubstitutedPlaceholder('  ${FOO}  '), 'FOO');
  assert.equal(unsubstitutedPlaceholder('${}'), '');
  // Not placeholders: a real id, a partial, or text that merely embeds one.
  assert.equal(unsubstitutedPlaceholder('11111111-2222-3333-4444-555555555555'), null);
  assert.equal(unsubstitutedPlaceholder('${FOO'), null);
  assert.equal(unsubstitutedPlaceholder('prefix-${FOO}'), null);
  assert.equal(unsubstitutedPlaceholder(''), null);
});
