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
} from './provider.js';
import type { MsalPublicClient } from './provider.js';

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

test('authFlow device never opens a browser', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({ config: configWith({ authFlow: 'device' }), client });

  const { result: account } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(account.username, ACCOUNT.username);
  assert.equal(client.interactive.length, 0);
  assert.equal(client.device.length, 1);
  assert.deepEqual(client.device[0]?.scopes, ['User.Read', 'offline_access']);
});

test('authFlow browser propagates the failure instead of falling back', async () => {
  const client = fakeClient({ interactiveError: new Error('xdg-open ENOENT') });
  const provider = createAuthProvider({ config: configWith({ authFlow: 'browser' }), client });

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
  const provider = createAuthProvider({ config: configWith({ authFlow: 'auto' }), client });

  const { result: account, stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(account.username, ACCOUNT.username);
  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 1);
  assert.match(stderr, /Browser sign-in failed: no display/);
  assert.match(stderr, /Falling back to device code/);
});

test('authFlow auto stays on the browser when it succeeds', async () => {
  const client = fakeClient();
  const provider = createAuthProvider({ config: configWith({ authFlow: 'auto' }), client });

  await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(client.interactive.length, 1);
  assert.equal(client.device.length, 0);
});

test('preferredPort is sent only when authPort is greater than zero', async () => {
  const pinned = fakeClient();
  await withCapturedStderr(() =>
    createAuthProvider({
      config: configWith({ authFlow: 'browser', authPort: 4141 }),
      client: pinned,
    }).login(SCOPES),
  );
  assert.equal(pinned.interactive[0]?.preferredPort, 4141);

  const ephemeral = fakeClient();
  await withCapturedStderr(() =>
    createAuthProvider({
      config: configWith({ authFlow: 'browser', authPort: 0 }),
      client: ephemeral,
    }).login(SCOPES),
  );
  const request = ephemeral.interactive[0];
  assert.ok(request);
  assert.equal('preferredPort' in request, false, 'MSAL must pick the port itself');
});

test('the browser flow sends self-contained success and error pages', async () => {
  const client = fakeClient();
  await withCapturedStderr(() =>
    createAuthProvider({ config: configWith({ authFlow: 'browser' }), client }).login(SCOPES),
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

test('the legacy MS365_MCP_AUTH_FLOW=interactive value still selects the browser', async () => {
  const previous = process.env['MS365_MCP_AUTH_FLOW'];
  process.env['MS365_MCP_AUTH_FLOW'] = 'interactive';
  try {
    const client = fakeClient();
    // authFlow cast away to model a config object built before the field existed.
    const config = { ...BASE_CONFIG, authFlow: undefined as unknown as AuthFlow };
    await withCapturedStderr(() => createAuthProvider({ config, client }).login(SCOPES));
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
];

for (const testCase of APP_UNAVAILABLE_CASES) {
  test(testCase.name, async () => {
    const client = fakeClient({ interactiveError: new Error(testCase.error) });
    const provider = createAuthProvider({
      config: configWith({ authFlow: 'auto' }),
      clientIdSource: 'shipped',
      client,
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
  const provider = createAuthProvider({ config: configWith({ authFlow: 'auto' }), client });

  const { stderr } = await withCapturedStderr(() => provider.login(SCOPES));

  assert.equal(client.device.length, 1, 'transport failures keep the fallback');
  assert.match(stderr, /Falling back to device code/);
});

test('isAppUnavailable matches whole codes only', () => {
  assert.equal(isAppUnavailable(new Error('AADSTS700016: not found in the directory')), true);
  assert.equal(isAppUnavailable('AADSTS9002332: Entra users only'), true, 'non-Error input');
  assert.equal(isAppUnavailable(new Error('spawn xdg-open ENOENT')), false);
  assert.equal(isAppUnavailable(new Error('AADSTS65004: the user declined consent')), false);
  assert.equal(
    isAppUnavailable(new Error('AADSTS500200: an unrelated code')),
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
