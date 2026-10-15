import assert from 'node:assert/strict';
import test from 'node:test';

import type {
  AccountInfo as MsalAccountInfo,
  AuthenticationResult,
  DeviceCodeRequest,
  InteractiveRequest,
  SilentFlowRequest,
} from '@azure/msal-node';
import type { AuthFlow, ServerConfig } from '../contracts.js';
import { InteractionRequiredError } from '../contracts.js';
import {
  appUnavailableMessage,
  createAuthProvider,
  describeSignInTarget,
  deviceCodeRejectedMessage,
  isAppUnavailable,
  isConditionalAccess,
  isLoopbackTimeout,
  isNetworkFailure,
} from './provider.js';
import type { MsalPublicClient } from './provider.js';
import type { PreflightChecks, PreflightResult } from './preflight.js';

const SCOPES = ['User.Read'];

/** No test here touches the cache directory, but logout() would. */
const BASE_CONFIG: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['mail'],
  readOnly: false,
  graphVersion: 'v1.0',
  allowBeta: false,
  discovery: false,
  orgMode: false,
  maxOutputChars: 25_000,
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'auto',
  authPort: 0,
  allowGenericWrite: false,
};

function configWith(overrides: { authFlow?: AuthFlow; authPort?: number }): ServerConfig {
  return { ...BASE_CONFIG, ...overrides };
}

const ACCOUNT: MsalAccountInfo = {
  homeAccountId: 'home-account-id',
  environment: 'login.microsoftonline.com',
  tenantId: 'tenant-id',
  username: 'user@example.com',
  localAccountId: 'local-account-id',
  name: 'Example User',
};

function authResult(accessToken = 'access-token'): AuthenticationResult {
  return {
    authority: BASE_CONFIG.authority,
    uniqueId: ACCOUNT.localAccountId,
    tenantId: ACCOUNT.tenantId,
    scopes: [...SCOPES],
    account: ACCOUNT,
    idToken: 'id-token',
    idTokenClaims: {},
    accessToken,
    fromCache: false,
    expiresOn: new Date(Date.now() + 3_600_000),
    tokenType: 'Bearer',
    correlationId: 'correlation-id',
  };
}

interface FakeClient extends MsalPublicClient {
  readonly interactive: InteractiveRequest[];
  readonly device: DeviceCodeRequest[];
  readonly silent: SilentFlowRequest[];
}

/** A stubbed PublicClientApplication. Nothing here reaches the network. */
function fakeClient(
  opts: {
    interactiveError?: Error;
    deviceError?: Error;
    silentError?: Error;
    accounts?: MsalAccountInfo[];
    /**
     * Handed to `deviceCodeCallback` exactly as MSAL would: synchronously,
     * outside any try/catch, so an error thrown by the callback propagates out
     * of `acquireTokenByDeviceCode` unchanged.
     */
    deviceCodeResponse?: Parameters<DeviceCodeRequest['deviceCodeCallback']>[0];
  } = {},
): FakeClient {
  const interactive: InteractiveRequest[] = [];
  const device: DeviceCodeRequest[] = [];
  const silent: SilentFlowRequest[] = [];
  const accounts = opts.accounts ?? [];

  return {
    interactive,
    device,
    silent,
    async acquireTokenSilent(request) {
      silent.push(request);
      if (opts.silentError) throw opts.silentError;
      return authResult();
    },
    async acquireTokenInteractive(request) {
      interactive.push(request);
      if (opts.interactiveError) throw opts.interactiveError;
      return authResult();
    },
    async acquireTokenByDeviceCode(request) {
      device.push(request);
      if (opts.deviceError) throw opts.deviceError;
      if (opts.deviceCodeResponse) request.deviceCodeCallback(opts.deviceCodeResponse);
      return authResult();
    },
    getTokenCache() {
      return {
        async getAllAccounts() {
          return [...accounts];
        },
        async removeAccount() {
          // not exercised here
        },
      };
    },
    clearCache() {
      // not exercised here
    },
  };
}

/** Keeps the sign-in banners and fallback warnings out of the test output. */
async function withCapturedStderr<T>(fn: () => Promise<T>): Promise<{ result: T; stderr: string }> {
  const original = process.stderr.write.bind(process.stderr);
  let captured = '';
  process.stderr.write = ((chunk: string | Uint8Array): boolean => {
    captured += typeof chunk === 'string' ? chunk : Buffer.from(chunk).toString('utf8');
    return true;
  }) as typeof process.stderr.write;
  try {
    return { result: await fn(), stderr: captured };
  } finally {
    process.stderr.write = original;
  }
}

/**
 * A preflight that reports everything usable, so the login tests below
 * exercise MSAL behaviour rather than the machine they run on. Without it a
 * headless CI box would short-circuit every browser attempt before MSAL is
 * reached, and the fallback assertions would depend on DISPLAY.
 */
function preflightStub(
  overrides: {
    browser?: PreflightResult;
    loopback?: PreflightResult;
    reach?: PreflightResult;
  } = {},
): PreflightChecks {
  return {
    async canOpenBrowser() {
      return overrides.browser ?? { verdict: 'possible', reason: 'stubbed: a browser can open' };
    },
    async canBindLoopback() {
      return overrides.loopback ?? { verdict: 'possible', reason: 'stubbed: loopback binds' };
    },
    async canReachEntra() {
      return overrides.reach ?? { verdict: 'possible', reason: 'stubbed: Entra answers' };
    },
  };
}

test('authFlow device never opens a browser', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'device' }),
    client,
    preflight: preflightStub(),
  });

  const { result: account } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(account.username, ACCOUNT.username);
  assert.equal(client.interactive.length, 0);
  assert.equal(client.device.length, 1);
  assert.deepEqual(client.device[0]?.scopes, ['User.Read', 'offline_access']);
});

test('authFlow browser propagates the failure instead of falling back', async () => {
  const client = fakeClient({ interactiveError: new Error('xdg-open ENOENT') });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'browser' }),
    client,
    preflight: preflightStub(),
  });

  const { stderr } = await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /xdg-open ENOENT/);
        assert.match(err.message, /--auth-flow device/);
        assert.match(err.message, /http:\/\/localhost\b/);
        return true;
      },
    );
  });

  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 0, 'browser mode must not fall back');
  assert.doesNotMatch(stderr, /Falling back/);
});

test('the failure message names the pinned redirect URI when authPort is set', async () => {
  const client = fakeClient({ interactiveError: new Error('port busy') });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'browser', authPort: 4141 }),
    client,
    preflight: preflightStub(),
  });

  await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /http:\/\/localhost:4141/);
        return true;
      },
    );
  });
});

test('authFlow auto falls back to device code and says why', async () => {
  const client = fakeClient({ interactiveError: new Error('no display') });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub(),
  });

  const { result: account, stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(account.username, ACCOUNT.username);
  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 1);
  assert.match(stderr, /Browser sign-in failed: no display/);
  assert.match(stderr, /Falling back to device code/);
});

test('authFlow auto stays on the browser when it succeeds', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub(),
  });

  await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 0);
});

test('a caller can force browser sign-in without changing the configured flow', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'device' }),
    client,
    preflight: preflightStub(),
  });

  await withCapturedStderr(() => provider.login(SCOPES, 'browser'));

  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 0);
});

test('preferredPort is sent only when authPort is greater than zero', async () => {
  const pinned = fakeClient();
  await withCapturedStderr(() =>
    createAuthProvider({
      config: configWith({ authFlow: 'browser', authPort: 4141 }),
      client: pinned,
      preflight: preflightStub(),
    }).login(SCOPES),
  );
  assert.equal(pinned.interactive[0]?.preferredPort, 4141);

  const ephemeral = fakeClient();
  await withCapturedStderr(() =>
    createAuthProvider({
      config: configWith({ authFlow: 'browser', authPort: 0 }),
      client: ephemeral,
      preflight: preflightStub(),
    }).login(SCOPES),
  );
  const request = ephemeral.interactive[0];
  assert.ok(request);
  assert.equal('preferredPort' in request, false, 'MSAL must pick the port itself');
});

test('the browser flow sends self-contained success and error pages', async () => {
  const client = fakeClient();
  await withCapturedStderr(() =>
    createAuthProvider({
      config: configWith({ authFlow: 'browser' }),
      client,
      preflight: preflightStub(),
    }).login(SCOPES),
  );

  const request = client.interactive[0];
  assert.ok(request);
  for (const template of [request.successTemplate, request.errorTemplate]) {
    assert.ok(template, 'both templates must be set');
    assert.ok(template.length < 1200, 'templates stay small');
    assert.doesNotMatch(template, /<script|https?:\/\//, 'no scripts and no external resources');
    assert.match(template, /prefers-color-scheme/);
  }
  assert.match(String(request.errorTemplate), /--auth-flow device/);
});

test('getToken never triggers an interactive flow', async () => {
  const signedIn = fakeClient({ accounts: [ACCOUNT] });
  const token = await createAuthProvider({ config: BASE_CONFIG, client: signedIn }).getToken(SCOPES);
  assert.equal(token, 'access-token');
  assert.equal(signedIn.interactive.length, 0);
  assert.equal(signedIn.device.length, 0);
  assert.deepEqual(signedIn.silent[0]?.scopes, ['User.Read', 'offline_access']);

  const signedOut = fakeClient();
  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client: signedOut }).getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      return true;
    },
  );
  assert.equal(signedOut.interactive.length, 0);
  assert.equal(signedOut.device.length, 0);
});

test('a single-tenant application on a tenantless endpoint fails the refresh with a fix, not a sign-in loop', async () => {
  const client = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error(
      "AADSTS50194: Application '00000000-0000-0000-0000-000000000000' is not configured as a multi-tenant application. Usage of the /common endpoint is not supported for such applications.",
    ),
  });

  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client }).getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.equal(err.tenantMismatch, true);
      assert.equal(err.adminScopes.length, 0);
      assert.match(err.message, /AADSTS50194/);
      assert.match(err.message, /single-tenant/);
      assert.match(err.message, /will NOT fix/);
      assert.match(err.message, /--tenant-id/);
      assert.match(err.message, /MS365_MCP_TENANT_ID/);
      assert.match(err.message, /restart the server/);
      assert.match(err.message, /Scopes requested: User\.Read/);
      return true;
    },
  );
  assert.equal(client.interactive.length, 0);
  assert.equal(client.device.length, 0);
});

test('any other silent-refresh failure still asks for a sign-in', async () => {
  const client = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error('interaction_required: the refresh token expired'),
  });

  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client }).getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.equal(err.tenantMismatch, false);
      assert.equal(err.consentMissing, false);
      assert.match(err.message, /A sign-in is needed/);
      assert.doesNotMatch(err.message, /will NOT fix/);
      return true;
    },
  );
});

test('a silent refresh refused for missing consent names the scopes, not the session', async () => {
  const client = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error(
      'AADSTS65001: The user or admin has not consented to use the application.',
    ),
  });

  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client }).getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.equal(err.consentMissing, true);
      assert.equal(err.tenantMismatch, false);
      assert.equal(err.adminScopes.length, 0);
      assert.match(err.message, /Blocked by missing consent for these scopes: User\.Read\./);
      assert.match(err.message, /auth_begin_login` with `\{"scopes": \["User\.Read"\]\}/);
      assert.match(err.message, /login --scopes User\.Read/);
      assert.match(err.message, /prompts only for what was added/);
      return true;
    },
  );
  assert.equal(client.interactive.length, 0);
  assert.equal(client.device.length, 0);
});

test('the legacy MS365_MCP_AUTH_FLOW=interactive value still selects the browser', async () => {
  const previous = process.env['MS365_MCP_AUTH_FLOW'];
  process.env['MS365_MCP_AUTH_FLOW'] = 'interactive';
  try {
    const client = fakeClient();
    // authFlow cast away to model a config object built before the field existed.
    const config = { ...BASE_CONFIG, authFlow: undefined as unknown as AuthFlow };
    await withCapturedStderr(() =>
      createAuthProvider({ config, client, preflight: preflightStub() }).login(SCOPES),
    );
    assert.equal(client.interactive.length, 1);
    assert.equal(client.device.length, 0);
  } finally {
    if (previous === undefined) delete process.env['MS365_MCP_AUTH_FLOW'];
    else process.env['MS365_MCP_AUTH_FLOW'] = previous;
  }
});

// ---------------------------------------------------------------------------
// Failures about the application itself
//
// These arrive once a client ID ships with the package: the user configured
// nothing and did nothing wrong locally, so a message about firewalls and
// redirect URIs would send them looking in entirely the wrong place.
// ---------------------------------------------------------------------------

const APP_UNAVAILABLE_CASES: ReadonlyArray<{
  readonly name: string;
  readonly error: string;
  readonly expected: readonly RegExp[];
}> = [
  {
    name: 'AADSTS700016 blames the registration, not the browser',
    error:
      "AADSTS700016: Application with identifier '11111111-1111-1111-1111-111111111111' was not found in the directory 'Contoso'.\r\nTrace ID: trace\r\nCorrelation ID: correlation",
    expected: [/AADSTS700016/, /multi-tenant/, /enterprise application/, /--client-id/],
  },
  {
    name: 'AADSTS700038 says no application was configured',
    error: 'AADSTS700038: The provided value is not a valid application identifier.',
    expected: [/AADSTS700038/, /no client ID was configured/, /--client-id/, /MS365_MCP_CLIENT_ID/],
  },
  {
    name: 'AADSTS50020 names the consumers authority',
    error:
      "AADSTS50020: User account from identity provider 'live.com' does not exist in tenant 'Contoso'.",
    expected: [/AADSTS50020/, /identity provider/, /--tenant-id consumers/],
  },
  {
    name: 'AADSTS7000112 reports an application disabled in the tenant',
    error: 'AADSTS7000112: Application is disabled.',
    expected: [/AADSTS7000112/, /disabled/, /Enterprise applications/, /--client-id/],
  },
  {
    name: 'AADSTS700054 reports a disabled application too',
    error: 'AADSTS700054: response_type id_token is not enabled for the application.',
    expected: [/AADSTS700054/, /disabled/, /--client-id/],
  },
  {
    name: 'AADSTS9002332 explains a work-account-only audience',
    error: 'AADSTS9002332: Application is configured for use by Microsoft Entra users only.',
    expected: [/AADSTS9002332/, /work or school/, /--tenant-id consumers/],
  },
  {
    name: 'AADSTS50194 names the owning tenant for a single-tenant application',
    error:
      "AADSTS50194: Application '00000000-0000-0000-0000-000000000000' is not configured as a multi-tenant application. Usage of the /common endpoint is not supported for such applications.",
    expected: [/AADSTS50194/, /single-tenant/, /--tenant-id/, /MS365_MCP_TENANT_ID/],
  },
];

for (const testCase of APP_UNAVAILABLE_CASES) {
  test(testCase.name, async () => {
    const client = fakeClient({ interactiveError: new Error(testCase.error) });
    const provider = createAuthProvider({
      config: configWith({ authFlow: 'auto' }),
      clientIdSource: 'shipped',
      client,
      preflight: preflightStub(),
    });

    const { stderr } = await withCapturedStderr(async () => {
      await assert.rejects(
        () => provider.login(SCOPES),
        (err: unknown) => {
          assert.ok(err instanceof Error);
          for (const pattern of testCase.expected) assert.match(err.message, pattern);
          assert.match(err.message, /^Sign-in failed/);
          assert.match(err.message, /\nLikely cause: /);
          assert.match(err.message, /\nFix: /);
          // Guidance only: no stack, and none of MSAL's trace metadata.
          assert.doesNotMatch(err.message, /\n\s+at /);
          assert.doesNotMatch(err.message, /Correlation ID/);
          return true;
        },
      );
    });

    assert.equal(client.device.length, 0, 'the device flow would be refused identically');
    assert.match(stderr, /\[auth\]/, 'the stack belongs on stderr');
  });
}

test('an app-availability failure never falls back to the device code', async () => {
  const client = fakeClient({
    interactiveError: new Error("AADSTS700016: Application with identifier 'x' was not found."),
  });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    clientIdSource: 'shipped',
    client,
    preflight: preflightStub(),
  });

  const { stderr } = await withCapturedStderr(async () => {
    await assert.rejects(() => provider.login(SCOPES));
  });

  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 0);
  assert.doesNotMatch(stderr, /Falling back/);
});

test('the device flow explains an app-availability failure the same way', async () => {
  const client = fakeClient({ deviceError: new Error('AADSTS7000112: Application is disabled.') });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'device' }),
    clientIdSource: 'shipped',
    client,
    preflight: preflightStub(),
  });

  await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /AADSTS7000112/);
        assert.match(err.message, /Fix: /);
        return true;
      },
    );
  });
});

test('a browser failure still falls back and still blames the browser', async () => {
  const client = fakeClient({ interactiveError: new Error('spawn xdg-open ENOENT') });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub(),
  });

  const { stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(client.device.length, 1, 'transport failures keep the fallback');
  assert.match(stderr, /Falling back to device code/);
});

test('isAppUnavailable matches whole codes only', () => {
  assert.equal(isAppUnavailable(new Error('AADSTS700016: not found in the directory')), true);
  assert.equal(isAppUnavailable('AADSTS9002332: Entra users only'), true, 'non-Error input');
  assert.equal(
    isAppUnavailable(new Error('AADSTS50194: not configured as a multi-tenant application')),
    true,
  );
  assert.equal(isAppUnavailable(new Error('spawn xdg-open ENOENT')), false);
  assert.equal(isAppUnavailable(new Error('AADSTS65004: the user declined consent')), false);
  assert.equal(
    isAppUnavailable(new Error('AADSTS500200: an unrelated code')),
    false,
    'a longer code must not match a shorter one',
  );
  assert.equal(
    isAppUnavailable(new Error('AADSTS501940: an unrelated code')),
    false,
    'a longer code must not match a shorter one',
  );
});

test('appUnavailableMessage still gives guidance for an unclassified refusal', () => {
  const message = appUnavailableMessage(new Error('something opaque'), BASE_CONFIG, 'shipped');
  assert.match(message, /^Sign-in failed: /);
  assert.match(message, /Likely cause: /);
  assert.match(message, /Fix: .*--client-id/);
});

// ---------------------------------------------------------------------------
// Which application is about to be granted access
// ---------------------------------------------------------------------------

test('an empty client ID throws before any MSAL call', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({
    config: { ...BASE_CONFIG, clientId: '   ' },
    clientIdSource: 'unset',
    client,
  });

  const { stderr } = await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /No Entra application \(client\) ID is configured/);
        assert.match(err.message, /--client-id/);
        assert.match(err.message, /MS365_MCP_CLIENT_ID/);
        return true;
      },
    );
  });

  assert.equal(client.interactive.length, 0, 'nothing may reach Entra');
  assert.equal(client.device.length, 0);
  assert.equal(stderr, '', 'not even a sign-in banner');
});

test('describeSignInTarget names the authority and where the app came from', () => {
  assert.equal(
    describeSignInTarget(BASE_CONFIG, 'flag'),
    'Signing in at https://login.microsoftonline.com/common with application 00000000-0000-0000-0000-000000000000 (from --client-id).',
  );
  assert.match(describeSignInTarget(BASE_CONFIG, 'env'), /\(from MS365_MCP_CLIENT_ID\)/);
  assert.match(describeSignInTarget(BASE_CONFIG, 'shipped'), /shipped default/);

  const anonymous = describeSignInTarget(BASE_CONFIG);
  assert.match(anonymous, /^Signing in at https:\/\/login\.microsoftonline\.com\/common with /);
  assert.doesNotMatch(anonymous, /\(from/);
  assert.doesNotMatch(anonymous, /\n/, 'one line');

  assert.match(
    describeSignInTarget({ ...BASE_CONFIG, clientId: '' }, 'unset'),
    /no application \(client\) ID/,
  );
});

test('login prints the sign-in target before opening the browser', async () => {
  const client = fakeClient();
  const { stderr } = await withCapturedStderr(() =>
    createAuthProvider({
      config: configWith({ authFlow: 'browser' }),
      clientIdSource: 'shipped',
      client,
      preflight: preflightStub(),
    }).login(SCOPES),
  );

  const target = stderr.indexOf('Signing in at ');
  const opening = stderr.indexOf('Opening your browser');
  assert.ok(target >= 0, 'the target line is printed');
  assert.ok(opening > target, 'and it comes before the browser opens');
  assert.match(stderr, /shipped default/);
});

test('the device flow names the application too', async () => {
  const client = fakeClient();
  const { stderr } = await withCapturedStderr(() =>
    createAuthProvider({
      config: configWith({ authFlow: 'device' }),
      clientIdSource: 'flag',
      client,
      preflight: preflightStub(),
    }).login(SCOPES),
  );

  assert.match(stderr, /Signing in at https:\/\/login\.microsoftonline\.com\/common/);
  assert.match(stderr, /\(from --client-id\)/);
});

// ---------------------------------------------------------------------------
// The device code endpoint refusing to issue a code
// ---------------------------------------------------------------------------

/**
 * Reproduces what `@azure/msal-node` 6.0.0 actually hands the callback when
 * Entra rejects the application: it destructures only the success fields out
 * of the response body, so `error` and `error_description` — and with them the
 * AADSTS code — are gone, and every field arrives `undefined`.
 */
const REFUSED_DEVICE_CODE = {
  userCode: undefined,
  deviceCode: undefined,
  verificationUri: undefined,
  expiresIn: undefined,
  interval: undefined,
  message: undefined,
} as unknown as Parameters<DeviceCodeRequest['deviceCodeCallback']>[0];

test('a refused device code stops instead of telling the user to open "undefined"', async () => {
  const client = fakeClient({ deviceCodeResponse: REFUSED_DEVICE_CODE });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'device' }),
    clientIdSource: 'flag',
    client,
    preflight: preflightStub(),
  });

  const { stderr } = await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        // Left unhandled this surfaces as `post_request_failed: invalid_grant`
        // one poll later, which names neither the cause nor the fix.
        assert.doesNotMatch(err.message, /invalid_grant/);
        assert.match(err.message, /issued no device code/);
        assert.match(err.message, /--client-id/);
        assert.match(err.message, /--tenant-id consumers/);
        return true;
      },
    );
  });

  assert.doesNotMatch(stderr, /undefined/, 'no "Open undefined" instruction may be printed');
  assert.doesNotMatch(stderr, /Enter the code/);
});

test('auto does not fall back to a browser after the device code is refused', async () => {
  // The reverse direction of the browser-first rule: a refusal is about the
  // application, so retrying anywhere else only fails the same way.
  const client = fakeClient({ deviceCodeResponse: REFUSED_DEVICE_CODE });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'device' }),
    clientIdSource: 'shipped',
    client,
    preflight: preflightStub(),
  });

  await withCapturedStderr(async () => {
    await assert.rejects(() => provider.login(SCOPES));
  });

  assert.equal(client.device.length, 1);
  assert.equal(client.interactive.length, 0);
});

test('the refusal message names the app and its origin but invents no AADSTS code', () => {
  const message = deviceCodeRejectedMessage(
    { clientId: 'abc-123', tenantId: 'consumers' },
    'shipped',
  );
  assert.match(message, /abc-123/);
  assert.match(message, /shipped default/);
  assert.match(message, /consumers/);
  // The real code was discarded upstream. Quoting one the user cannot find in
  // their sign-in logs is worse than admitting it is unavailable, and a bare
  // AADSTS number here would also make `isAppUnavailable` re-wrap this message
  // in the generic one on the way out.
  assert.doesNotMatch(message, /AADSTS\d/);
  assert.equal(isAppUnavailable(new Error(message)), false);
});

// ---------------------------------------------------------------------------
// Device sign-in, started from a tool call
// ---------------------------------------------------------------------------

/**
 * The thing these guard is a shape, not a detail: `beginDeviceLogin` must return
 * when the CODE exists, not when the sign-in finishes. A tool call cannot sit
 * for the minutes a person takes, and before these tools existed there was no
 * way to authenticate from inside a conversation at all — `login` was a CLI
 * subcommand, so anyone installing this as a plugin had to go and find a
 * terminal first.
 *
 * So the stub below hands back a promise the test settles by hand. A stub that
 * resolves immediately would pass whether or not the implementation waited.
 */
function deferredDeviceClient(
  response: Parameters<DeviceCodeRequest['deviceCodeCallback']>[0],
  opts: { accounts?: MsalAccountInfo[] } = {},
): FakeClient & { settle: (result: 'ok' | Error) => void; starts: () => number } {
  const device: DeviceCodeRequest[] = [];
  let finish: (value: 'ok' | Error) => void = () => {};
  const gate = new Promise<'ok' | Error>((resolve) => {
    finish = resolve;
  });

  return {
    interactive: [],
    device,
    silent: [],
    starts: () => device.length,
    settle: (result: 'ok' | Error) => finish(result),
    async acquireTokenSilent() {
      return authResult();
    },
    async acquireTokenInteractive() {
      return authResult();
    },
    async acquireTokenByDeviceCode(request: DeviceCodeRequest) {
      device.push(request);
      request.deviceCodeCallback(response);
      const outcome = await gate;
      if (outcome !== 'ok') throw outcome;
      return authResult();
    },
    getTokenCache() {
      return {
        async getAllAccounts() {
          return [...(opts.accounts ?? [])];
        },
        async removeAccount() {
          // not exercised here
        },
      };
    },
  } as unknown as FakeClient & { settle: (result: 'ok' | Error) => void; starts: () => number };
}

function deviceCode(
  over: Partial<{ userCode: string; verificationUri: string; expiresIn: number; interval: number }> = {},
): Parameters<DeviceCodeRequest['deviceCodeCallback']>[0] {
  return {
    userCode: 'ABCD-EFGH',
    deviceCode: 'device-code-value',
    verificationUri: 'https://example.test/device',
    expiresIn: 900,
    interval: 5,
    message: 'Open the page and enter the code.',
    ...over,
  } as unknown as Parameters<DeviceCodeRequest['deviceCodeCallback']>[0];
}

const CODE = deviceCode();

test('starting a sign-in returns the code without waiting for the person', async () => {
  const client = deferredDeviceClient(CODE);
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  // If this awaited the flow it would hang here: the stub never settles.
  const prompt = await provider.beginDeviceLogin(SCOPES);

  assert.equal(prompt.userCode, 'ABCD-EFGH');
  assert.equal(prompt.verificationUri, 'https://example.test/device');
  assert.equal(prompt.pollIntervalSeconds, 5);
  assert.ok(Date.parse(prompt.expiresAt) > Date.now(), 'expiry is an instant in the future');
});

test('asking again while a code is live returns that code, not a second one', async () => {
  // A second flow would hand out a new code and silently invalidate the one the
  // person is in the middle of typing.
  const client = deferredDeviceClient(CODE);
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  const first = await provider.beginDeviceLogin(SCOPES);
  const second = await provider.beginDeviceLogin(SCOPES);

  assert.deepEqual(second, first);
  assert.equal(client.starts(), 1, 'only one device-code flow should have been started');
});

test('status moves from pending to signed in, and names the account', async () => {
  const client = deferredDeviceClient(CODE, { accounts: [ACCOUNT] });
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  assert.deepEqual(await provider.deviceLoginStatus(), { state: 'none' });

  await provider.beginDeviceLogin(SCOPES);
  const pending = await provider.deviceLoginStatus();
  assert.equal(pending.state, 'pending');
  assert.equal(pending.prompt?.userCode, 'ABCD-EFGH', 'the code is repeated so it can be shown again');

  client.settle('ok');
  await new Promise((r) => setImmediate(r));

  const done = await provider.deviceLoginStatus();
  assert.equal(done.state, 'signedIn');
  assert.equal(done.account, ACCOUNT.username);
});

test('a refused sign-in reports failed, with mapped guidance', async () => {
  const client = deferredDeviceClient(CODE);
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  await provider.beginDeviceLogin(SCOPES);
  client.settle(new Error('AADSTS65004: user declined the consent'));
  await new Promise((r) => setImmediate(r));

  const status = await provider.deviceLoginStatus();
  assert.equal(status.state, 'failed');
  // Mapped, not raw: the consent guidance, still quoting the code it mapped.
  assert.match(String(status.detail), /refused at the consent screen/);
  assert.match(String(status.detail), /Ask for less/);
  assert.match(String(status.detail), /AADSTS65004/);
});

test('status never carries a token, in any state', async () => {
  // SECURITY.md puts "the model must never see a token" in scope, and a status
  // field is exactly where one leaks by accident.
  const client = deferredDeviceClient(CODE, { accounts: [ACCOUNT] });
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  const seen: unknown[] = [await provider.deviceLoginStatus()];
  await provider.beginDeviceLogin(SCOPES);
  seen.push(await provider.deviceLoginStatus());
  client.settle('ok');
  await new Promise((r) => setImmediate(r));
  seen.push(await provider.deviceLoginStatus());

  const serialized = JSON.stringify(seen);
  for (const forbidden of ['accessToken', 'access_token', 'idToken', 'refreshToken', 'Bearer']) {
    assert.ok(!serialized.includes(forbidden), `${forbidden} must not appear in a status`);
  }
  assert.ok(!serialized.includes(authResult().accessToken), 'nor the token value itself');
});

test('an expired code is reported as expired rather than pending forever', async () => {
  const client = deferredDeviceClient(deviceCode({ expiresIn: 0 }));
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  await provider.beginDeviceLogin(SCOPES);
  const status = await provider.deviceLoginStatus();

  assert.equal(status.state, 'expired');
  assert.match(String(status.detail), /Start another sign-in/);
});

test('an expired code is not reused when a new sign-in is asked for', async () => {
  const client = deferredDeviceClient(deviceCode({ expiresIn: 0 }));
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  await provider.beginDeviceLogin(SCOPES);
  await provider.beginDeviceLogin(SCOPES);

  assert.equal(client.starts(), 2, 'a dead code must be replaced, not handed out again');
});

test('a registration that cannot issue device codes says so rather than hanging', async () => {
  // MSAL calls the callback synchronously and rethrows what it throws, so this
  // is the last point at which a refusal is still distinguishable from a token
  // endpoint failure minutes later.
  const client = deferredDeviceClient(deviceCode({ userCode: '', verificationUri: '' }));
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  await assert.rejects(() => provider.beginDeviceLogin(SCOPES), /public client|device code|Allow public/i);
});

test('a failure before any code is issued rejects instead of waiting for ever', async () => {
  // Offline, a proxy or TLS failure, a tenant that does not resolve: MSAL
  // rejects before deviceCodeCallback ever fires, so nothing resolves the
  // promise the caller is holding. It used to wait for ever, which is worse
  // than the error it was hiding, and a tool call cannot be cancelled.
  const client = {
    async acquireTokenSilent() {
      throw new Error('not exercised');
    },
    async acquireTokenInteractive() {
      throw new Error('not exercised');
    },
    async acquireTokenByDeviceCode() {
      throw new Error('getaddrinfo ENOTFOUND login.microsoftonline.com');
    },
    getTokenCache() {
      return {
        async getAllAccounts() {
          return [];
        },
        async removeAccount() {
          // not exercised
        },
      };
    },
  } as unknown as FakeClient;
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  await assert.rejects(
    () => provider.beginDeviceLogin(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof Error);
      // Mapped, not raw: the network guidance, still quoting the failure it mapped.
      assert.match(err.message, /Could not reach Microsoft Entra/);
      assert.match(err.message, /will NOT fix this while the network is down/);
      assert.match(err.message, /--use-env-proxy/);
      assert.match(err.message, /ENOTFOUND/);
      return true;
    },
  );
  // And it does not leave a dead flow behind claiming to be pending.
  assert.deepEqual(await provider.deviceLoginStatus(), { state: 'none' });
});

test('two sign-ins started at once share one code', async () => {
  // `pending` is written from inside the callback, so a guard that reads it has
  // nothing to read while the first request is still in the air. Two calls would
  // each start a flow and the code the first caller was already shown would be
  // dead on arrival.
  let release: () => void = () => {};
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const started: unknown[] = [];
  const client = {
    async acquireTokenSilent() {
      throw new Error('not exercised');
    },
    async acquireTokenInteractive() {
      throw new Error('not exercised');
    },
    async acquireTokenByDeviceCode(request: DeviceCodeRequest) {
      started.push(request);
      await gate;
      request.deviceCodeCallback(deviceCode());
      return authResult();
    },
    getTokenCache() {
      return {
        async getAllAccounts() {
          return [ACCOUNT];
        },
        async removeAccount() {
          // not exercised
        },
      };
    },
  } as unknown as FakeClient;
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  const both = Promise.all([provider.beginDeviceLogin(SCOPES), provider.beginDeviceLogin(SCOPES)]);
  release();
  const [first, second] = await both;

  assert.equal(started.length, 1, 'one flow, not two');
  assert.deepEqual(second, first);
});

// ---------------------------------------------------------------------------
// Silent-refresh failures the fallthrough used to wrap raw
// ---------------------------------------------------------------------------

test('a Conditional Access block on the refresh names the browser retry, not a dead session', async () => {
  const client = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error(
      'AADSTS53003: Access has been blocked by Conditional Access policies. The access policy does not allow token issuance.',
    ),
  });

  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client }).getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.equal(err.conditionalAccess, true);
      assert.equal(err.consentMissing, false);
      assert.equal(err.tenantMismatch, false);
      assert.deepEqual([...err.scopes], SCOPES);
      assert.match(err.message, /Conditional Access policy \(AADSTS53003\)/);
      assert.match(err.message, /auth_begin_login/);
      assert.match(err.message, /device-code flow is commonly blocked/);
      assert.match(err.message, /Scopes requested: User\.Read/);
      return true;
    },
  );
  assert.equal(client.interactive.length, 0);
  assert.equal(client.device.length, 0);
});

test('a network failure on the refresh blames the route, not the session', async () => {
  const client = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error('getaddrinfo ENOTFOUND login.microsoftonline.com'),
  });

  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client }).getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.equal(err.networkFailure, true);
      assert.deepEqual([...err.scopes], SCOPES);
      assert.match(err.message, /Could not reach Microsoft Entra/);
      assert.match(err.message, /will NOT fix this while the network is down/);
      assert.match(err.message, /--use-env-proxy/);
      assert.match(err.message, /ENOTFOUND/);
      return true;
    },
  );
});

test('a loopback timeout on the refresh points at the device flow', async () => {
  const client = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error('loopback_server_timeout: the redirect was never served'),
  });

  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client }).getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.equal(err.loopbackTimeout, true);
      assert.match(err.message, /timed out waiting for the loopback redirect/);
      assert.match(err.message, /--auth-flow device/);
      return true;
    },
  );
});

test('every silent-refresh failure records the scopes it asked for', async () => {
  const client = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error('interaction_required: the refresh token expired'),
  });

  await assert.rejects(
    () => createAuthProvider({ config: BASE_CONFIG, client }).getToken(['Mail.Read', 'User.Read']),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.deepEqual([...err.scopes], ['Mail.Read', 'User.Read']);
      return true;
    },
  );
});

test('the classifiers match whole failures only', () => {
  assert.equal(isConditionalAccess(new Error('AADSTS53003: blocked by policy')), true);
  assert.equal(isConditionalAccess(new Error('AADSTS530030: an unrelated code')), false);
  assert.equal(
    isNetworkFailure(new Error('getaddrinfo ENOTFOUND login.microsoftonline.com')),
    true,
  );
  assert.equal(isNetworkFailure(new Error('fetch failed')), true);
  assert.equal(isNetworkFailure(new Error('spawn xdg-open ENOENT')), false);
  assert.equal(isNetworkFailure(new Error('no display')), false);
  assert.equal(
    isLoopbackTimeout(new Error('loopback_server_timeout: no redirect arrived')),
    true,
  );
  assert.equal(isLoopbackTimeout(new Error('spawn xdg-open ENOENT')), false);
});

// ---------------------------------------------------------------------------
// sessionHealth: the honest half of "signed in"
// ---------------------------------------------------------------------------

test('sessionHealth reports usable when the silent refresh succeeds', async () => {
  const client = fakeClient({ accounts: [ACCOUNT] });
  const provider = createAuthProvider({ config: BASE_CONFIG, client });

  const health = await provider.sessionHealth?.();
  assert.deepEqual(health, { tokenUsable: true });
  assert.deepEqual(client.silent[0]?.scopes, ['User.Read', 'offline_access']);
  assert.equal(client.interactive.length, 0);
  assert.equal(client.device.length, 0);
});

test('sessionHealth reports unusable without signing out the cached account', async () => {
  // Unlike getToken the probe must not evict what it probed: a transient probe
  // failure must not turn the next data call into a sign-out.
  const accountList = [ACCOUNT];
  let silentError: Error | null = null;
  const silent: unknown[] = [];
  const client = {
    async acquireTokenSilent(request: unknown) {
      silent.push(request);
      if (silentError) throw silentError;
      return authResult();
    },
    async acquireTokenInteractive(): Promise<never> {
      throw new Error('not exercised');
    },
    async acquireTokenByDeviceCode(): Promise<never> {
      throw new Error('not exercised');
    },
    getTokenCache() {
      return {
        async getAllAccounts() {
          return [...accountList];
        },
        async removeAccount() {
          // not exercised
        },
      };
    },
    clearCache() {
      // not exercised
    },
  } as unknown as FakeClient;
  const provider = createAuthProvider({ config: BASE_CONFIG, client });

  // A first success populates the shared cached account.
  assert.equal(await provider.getToken(SCOPES), 'access-token');

  // The probe fails against a transient refusal...
  silentError = new Error('AADSTS65001: The user or admin has not consented to use the application.');
  const health = await provider.sessionHealth?.(['Mail.Read']);
  assert.equal(health?.tokenUsable, false);
  assert.match(String(health?.note), /Missing consent for Mail\.Read/);
  assert.match(String(health?.note), /AADSTS65001/);

  // ...but the shared account survives it: empty the source and the next call
  // still attempts a silent refresh from cache instead of reporting no account.
  accountList.length = 0;
  await assert.rejects(
    () => provider.getToken(SCOPES),
    (err: unknown) => {
      assert.ok(err instanceof InteractionRequiredError);
      assert.equal(err.consentMissing, true);
      return true;
    },
  );
  assert.equal(silent.length, 3, 'probe plus both data calls attempted silent refresh');
});

test('sessionHealth classifies Conditional Access and network failures', async () => {
  const blocked = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error('AADSTS53003: Access has been blocked by Conditional Access policies.'),
  });
  const blockedHealth = await createAuthProvider({ config: BASE_CONFIG, client: blocked }).sessionHealth?.();
  assert.equal(blockedHealth?.tokenUsable, false);
  assert.match(String(blockedHealth?.note), /Conditional Access blocked the silent refresh/);
  assert.match(String(blockedHealth?.note), /sign in again in the browser/);

  const offline = fakeClient({
    accounts: [ACCOUNT],
    silentError: new Error('getaddrinfo ENOTFOUND login.microsoftonline.com'),
  });
  const offlineHealth = await createAuthProvider({ config: BASE_CONFIG, client: offline }).sessionHealth?.();
  assert.equal(offlineHealth?.tokenUsable, false);
  assert.match(String(offlineHealth?.note), /unreachable/);
});

test('sessionHealth with no cached account reports not signed in', async () => {
  const client = fakeClient();
  const health = await createAuthProvider({ config: BASE_CONFIG, client }).sessionHealth?.();
  assert.equal(health?.tokenUsable, false);
  assert.match(String(health?.note), /not signed in/);
  assert.equal(client.silent.length, 0, 'no account, no silent attempt');
});

// ---------------------------------------------------------------------------
// Login failures the catches used to surface raw
// ---------------------------------------------------------------------------

test('a Conditional Access block in auto falls back with a warning, not silently', async () => {
  const client = fakeClient({
    interactiveError: new Error('AADSTS53003: Access has been blocked by Conditional Access policies.'),
  });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub(),
  });

  const { stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 1);
  assert.match(stderr, /blocked by a Conditional Access policy/);
  assert.match(stderr, /commonly blocks too/);
});

test('a Conditional Access block in explicit browser mode stops with guidance', async () => {
  const client = fakeClient({
    interactiveError: new Error('AADSTS53003: Access has been blocked by Conditional Access policies.'),
  });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'browser' }),
    client,
    preflight: preflightStub(),
  });

  await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /Conditional Access policy \(AADSTS53003\)/);
        assert.match(err.message, /device-code flow is commonly blocked/);
        return true;
      },
    );
  });
  assert.equal(client.device.length, 0, 'explicit browser never falls back');
});

test('a network failure during login never falls back into the same failure', async () => {
  const client = fakeClient({
    interactiveError: new Error('getaddrinfo ENOTFOUND login.microsoftonline.com'),
  });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub(),
  });

  await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /Could not reach Microsoft Entra/);
        assert.match(err.message, /--use-env-proxy/);
        return true;
      },
    );
  });
  assert.equal(client.device.length, 0, 'both grants need the same network');
});

test('a loopback timeout falls back to the device code, which needs no listener', async () => {
  const client = fakeClient({
    interactiveError: new Error('loopback_server_timeout: the redirect was never served'),
  });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub(),
  });

  const { stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 1);
  assert.match(stderr, /timed out waiting for the loopback redirect/);
  assert.match(stderr, /needs no loopback listener/);
});

test('the device grant maps policy and transport failures instead of surfacing them raw', async () => {
  const blocked = fakeClient({
    deviceError: new Error('AADSTS53003: Access has been blocked by Conditional Access policies.'),
  });
  await withCapturedStderr(async () => {
    await assert.rejects(
      () =>
        createAuthProvider({
          config: configWith({ authFlow: 'device' }),
          client: blocked,
          preflight: preflightStub(),
        }).login(SCOPES),
      /Conditional Access policy \(AADSTS53003\)/,
    );
  });

  const offline = fakeClient({
    deviceError: new Error('getaddrinfo ENOTFOUND login.microsoftonline.com'),
  });
  await withCapturedStderr(async () => {
    await assert.rejects(
      () =>
        createAuthProvider({
          config: configWith({ authFlow: 'device' }),
          client: offline,
          preflight: preflightStub(),
        }).login(SCOPES),
      /Could not reach Microsoft Entra/,
    );
  });
});

// ---------------------------------------------------------------------------
// Pre-sign-in environment checks
// ---------------------------------------------------------------------------

test('auto skips the browser entirely when it verifiably cannot work', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub({
      browser: { verdict: 'impossible', reason: 'no desktop session (stubbed)' },
    }),
  });

  const { result: account, stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(account.username, ACCOUNT.username);
  assert.equal(client.interactive.length, 0, 'no doomed browser attempt');
  assert.equal(client.device.length, 1);
  assert.match(stderr, /Browser sign-in is unavailable \(no desktop session \(stubbed\)\)/);
  assert.match(stderr, /using device code directly/);
});

test('explicit browser mode fails fast with the browser guidance when preflight rules it out', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'browser' }),
    client,
    preflight: preflightStub({
      loopback: { verdict: 'impossible', reason: 'loopback bind failed (stubbed)' },
    }),
  });

  await withCapturedStderr(async () => {
    await assert.rejects(
      () => provider.login(SCOPES),
      (err: unknown) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /^Browser sign-in failed: pre-sign-in check: /);
        assert.match(err.message, /loopback bind failed \(stubbed\)/);
        assert.match(err.message, /--auth-flow device/);
        return true;
      },
    );
  });
  assert.equal(client.interactive.length, 0, 'MSAL is never reached');
  assert.equal(client.device.length, 0, 'explicit browser never falls back');
});

test('an unknown preflight keeps the current behaviour', async () => {
  const client = fakeClient({ interactiveError: new Error('no display') });
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'auto' }),
    client,
    preflight: preflightStub({
      browser: { verdict: 'unknown', reason: 'a display is set but no browser was checked' },
      loopback: { verdict: 'unknown', reason: 'the bind check itself failed' },
    }),
  });

  const { stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(client.interactive.length, 1, 'the attempt goes ahead as before');
  assert.equal(client.device.length, 1);
  assert.match(stderr, /Browser sign-in failed: no display/);
  assert.match(stderr, /Falling back to device code/);
});

test('an unreachable Entra warns but never stops the sign-in', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({
    config: configWith({ authFlow: 'device' }),
    client,
    preflight: preflightStub({
      reach: { verdict: 'impossible', reason: 'no answer from Microsoft Entra (stubbed)' },
    }),
  });

  const { result: account, stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(account.username, ACCOUNT.username);
  assert.match(stderr, /no answer from Microsoft Entra \(stubbed\); sign-in will likely fail/);
});
