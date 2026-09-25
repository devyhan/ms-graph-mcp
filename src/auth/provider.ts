/**
 * Delegated-permission authentication backed by MSAL Node.
 *
 * This is a local stdio MCP server, so per the MCP 2026-07-28 authorization spec
 * it does not implement MCP OAuth; it obtains user credentials itself as a
 * public client and presents the resulting bearer token to Microsoft Graph.
 *
 * Both interactive flows are OAuth 2.0: `browser` is the authorization code
 * grant with PKCE on a loopback redirect, `device` is the device authorization
 * grant. `auto` prefers the browser and falls back to the device code.
 */

import { spawn } from 'node:child_process';
import { PublicClientApplication } from '@azure/msal-node';
import type {
  AccountInfo as MsalAccountInfo,
  AuthenticationResult,
  DeviceCodeRequest,
  InteractiveRequest,
  SilentFlowRequest,
} from '@azure/msal-node';
import type { AccountInfo, AuthFlow, AuthProvider, ServerConfig } from '../contracts.js';
import { InteractionRequiredError } from '../contracts.js';
import { clearCache, createCachePlugin } from './token-cache.js';

/**
 * msal-node does not re-export `DeviceCodeResponse` from its entry point, and
 * `@azure/msal-common` is only a transitive dependency, so the shape is taken
 * from the callback signature we are handed.
 */
type DeviceCodeResponse = Parameters<DeviceCodeRequest['deviceCodeCallback']>[0];

/**
 * Compatibility shim only. `config.authFlow` is the source of truth: the config
 * layer already maps this variable and the `--auth-flow` flag, including the
 * legacy `interactive` alias. This is read here solely so a provider built with
 * a hand-rolled config that predates `authFlow` still honours the old variable.
 */
const AUTH_FLOW_ENV = 'MS365_MCP_AUTH_FLOW';

/** MSAL only issues a refresh token when this scope is requested. */
const OFFLINE_ACCESS = 'offline_access';

function log(message: string): void {
  // stdout carries the MCP JSON-RPC stream; everything human-facing goes here.
  process.stderr.write(`${message}\n`);
}

/** Diagnostics, prefixed the same way as the token cache's warnings. */
function warn(message: string): void {
  process.stderr.write(`[auth] ${message}\n`);
}

function describeError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

/**
 * Delegated Graph scopes that the default Microsoft-managed tenant consent
 * policy (`microsoft-user-default-recommended`, the default for new tenants
 * since August 2025) withholds from ordinary users, even though the permissions
 * reference marks every one of them `AdminConsentRequired: No`.
 *
 * This drives an explanatory message only, never a decision to skip a scope:
 * Microsoft states the managed policy is updated on their schedule without any
 * tenant-side change, so treating this list as authoritative would rot. The
 * authority is always what Entra actually answers.
 */
const TENANT_CONSENT_BLOCKED: ReadonlySet<string> = new Set([
  'Calendars.Read',
  'Calendars.ReadBasic',
  'Calendars.ReadWrite',
  'Calendars.Read.Shared',
  'Calendars.ReadWrite.Shared',
  'Chat.Read',
  'Chat.ReadWrite',
  'Contacts.ReadWrite',
  'Contacts.Read.Shared',
  'Contacts.ReadWrite.Shared',
  'Files.Read.All',
  'Files.ReadWrite.All',
  'Mail.Read',
  'Mail.ReadBasic',
  'Mail.ReadWrite',
  'Mail.Read.Shared',
  'Mail.ReadWrite.Shared',
  'MailboxSettings.Read',
  'MailboxSettings.ReadWrite',
  'OnlineMeetings.Read',
  'OnlineMeetings.ReadWrite',
  'People.Read',
  'Sites.Read.All',
  'Sites.ReadWrite.All',
  'Tasks.Read',
  'Tasks.Read.Shared',
  'Tasks.ReadWrite',
  'Tasks.ReadWrite.Shared',
]);

function toAccountInfo(account: MsalAccountInfo): AccountInfo {
  return {
    username: account.username,
    name: account.name,
    tenantId: account.tenantId,
    homeAccountId: account.homeAccountId,
  };
}

/** Adds `offline_access` without disturbing the caller's ordering or duplicates. */
function withRefreshScope(scopes: string[]): string[] {
  const hasOffline = scopes.some((s) => s.toLowerCase() === OFFLINE_ACCESS);
  return hasOffline ? [...scopes] : [...scopes, OFFLINE_ACCESS];
}

/** Renders an account for status output and error messages. */
export function describeAccount(a: AccountInfo | null): string {
  if (!a) return 'not signed in';
  const who = a.name && a.name !== a.username ? `${a.name} <${a.username}>` : a.username;
  return a.tenantId ? `${who} (tenant ${a.tenantId})` : who;
}

/**
 * Accepts the three documented values plus the legacy `interactive`, which
 * shipped as an undocumented env-only opt-in before the browser flow had a
 * name. Returns null for anything unrecognised so the caller can decide.
 */
function normaliseAuthFlow(value: string | undefined): AuthFlow | null {
  switch (value?.trim().toLowerCase()) {
    case 'auto':
      return 'auto';
    case 'browser':
    case 'interactive':
      return 'browser';
    case 'device':
      return 'device';
    default:
      return null;
  }
}

/** Resolves the flow from config, falling back to the legacy env var, then `auto`. */
export function resolveAuthFlow(config: Pick<ServerConfig, 'authFlow'>): AuthFlow {
  const fromConfig = normaliseAuthFlow(config.authFlow);
  return fromConfig ?? normaliseAuthFlow(process.env[AUTH_FLOW_ENV]) ?? 'auto';
}

/**
 * The redirect URI a tenant has to register for the browser flow. Entra accepts
 * bare `http://localhost` as a wildcard over every loopback port; a tenant that
 * pins an exact port needs `--auth-port` set to the same number.
 */
function redirectUriFor(authPort: number): string {
  return authPort > 0 ? `http://localhost:${authPort}` : 'http://localhost';
}

function signInRequiredMessage(scopes: string[], detail?: string): string {
  const parts = [
    'No usable Microsoft 365 session.',
    'Run the server\'s `login` command in a terminal (for example `npx ms-graph-mcp login`), complete the sign-in, then retry.',
    `Scopes requested: ${scopes.length > 0 ? scopes.join(', ') : '(none)'}.`,
  ];
  if (detail) parts.push(`Underlying error: ${detail}`);
  return parts.join(' ');
}

/**
 * What a stuck user reads. Every line names a cause they can act on, because
 * the raw MSAL error is usually a bare `ENOENT: xdg-open` or a redirect
 * mismatch code that says nothing about which knob to turn.
 */
/**
 * True when Entra rejected the redirect URI rather than the browser failing.
 *
 *   AADSTS900971 the registration has no reply address at all
 *   AADSTS50011  the redirect URI does not match one on the registration
 *
 * Worth separating from the generic browser failure because the fix is exact
 * and, unlike most browser problems, the device-code flow sidesteps it entirely
 * — that grant uses no redirect URI, so it works against a registration that
 * has none.
 */
function isRedirectUriProblem(err: unknown): boolean {
  const text = describeError(err);
  return /AADSTS900971(?!\d)/.test(text) || /AADSTS50011(?!\d)/.test(text);
}

function redirectUriMessage(authPort: number, detail: string): string {
  return [
    `Browser sign-in failed: ${detail}`,
    '',
    'The Entra app registration has no redirect URI the browser flow can use.',
    'That grant has to hand the authorization code back to a loopback address,',
    'so a registration without one cannot complete it.',
    '',
    'Two ways forward:',
    `  - Add ${redirectUriFor(authPort)} to the registration, under`,
    '    Authentication > Add a platform > Mobile and desktop applications.',
    '  - Or sign in with --auth-flow device, which uses no redirect URI at all',
    '    and works against this registration as it stands.',
  ].join('\n');
}

function browserFailureMessage(authPort: number, detail: string): string {
  const port = authPort > 0 ? ` (--auth-port ${authPort})` : '';
  return [
    `Browser sign-in failed: ${detail}`,
    'Likely causes:',
    '  - No desktop session. Over SSH, in a container, or in CI there is no browser to open; use --auth-flow device.',
    `  - The loopback redirect could not be served${port}. A firewall, a VPN, or another process on the port will do that.`,
    `  - The Entra app registration has no matching redirect URI. Add ${redirectUriFor(authPort)} under "Mobile and desktop applications".`,
    'To sign in from another device instead, retry with --auth-flow device.',
  ].join('\n');
}

/**
 * Where the application (client) ID in play came from.
 *
 * A plain parameter rather than a `ServerConfig` field: the config layer is
 * what resolved it, `index.ts` passes the answer through, and nothing in this
 * module tries to infer it. Once a shared multi-tenant app ID ships as the
 * default, "whose app am I granting access to" stops being self-evident, so it
 * has to be printable.
 *
 * The literals match `ClientIdSource` in `config.ts` exactly, so the resolution
 * assigns straight across without this module importing the CLI parser.
 */
export type ClientIdSource = 'flag' | 'env' | 'shipped' | 'unset';

/** The variable the config layer reads a client ID from. */
const CLIENT_ID_ENV = 'MS365_MCP_CLIENT_ID';

/**
 * ` (from --client-id)` and friends, or an empty string when there is nothing
 * worth saying — no source given, or no application resolved at all.
 */
function clientIdOrigin(source: ClientIdSource | undefined): string {
  switch (source) {
    case 'flag':
      return ' (from --client-id)';
    case 'env':
      return ` (from ${CLIENT_ID_ENV})`;
    case 'shipped':
      return " (from this project's shipped default)";
    default:
      return '';
  }
}

/**
 * One line naming what the user is about to sign in to, printed to stderr
 * before the browser opens. Names the authority, because that is what decides
 * whether a personal account is even eligible, and the app plus its origin, so
 * a user granting consent can see whether the app is theirs or ours.
 */
export function describeSignInTarget(
  config: Pick<ServerConfig, 'clientId' | 'authority'>,
  source?: ClientIdSource,
): string {
  const clientId = config.clientId.trim();
  const app = clientId === '' ? 'no application (client) ID' : `application ${clientId}`;
  return `Signing in at ${config.authority} with ${app}${clientIdOrigin(source)}.`;
}

/**
 * Shown instead of a sign-in attempt when nothing was configured at all.
 * Sending an empty client ID to Entra only earns an AADSTS700038 round trip.
 */
export function missingClientIdMessage(): string {
  return [
    'No Entra application (client) ID is configured, so there is nothing to sign in to.',
    'Likely cause: no --client-id was passed and no default is compiled into this build.',
    `Fix: pass --client-id <application id> from your own Entra app registration, or set ${CLIENT_ID_ENV}. A shared multi-tenant application ID ships as the default in a future release; until then the ID has to be your own.`,
  ].join('\n');
}

/** How an app-level rejection went, which decides what the user is told. */
type AppUnavailableKind =
  | 'not-in-tenant'
  | 'not-an-app-id'
  | 'wrong-account-type'
  | 'disabled'
  | 'entra-users-only';

interface AppUnavailableCase {
  /** The Entra error code, quoted back so the user can search for it. */
  readonly code: string;
  readonly pattern: RegExp;
  readonly kind: AppUnavailableKind;
}

/**
 * Entra rejections that are about the APPLICATION, not about the browser leg.
 *
 * These matter more now that a client ID can be one this project shipped rather
 * than one the user registered: the user did nothing wrong locally, so a
 * message about firewalls and redirect URIs would send them hunting in the
 * wrong place entirely.
 *
 *   AADSTS700016   the app is not in this directory: registered single-tenant,
 *                  blocked by the tenant, or its enterprise application deleted
 *   AADSTS700038   not a valid application identifier — nothing was configured
 *   AADSTS50020    the account comes from a different identity provider than
 *                  the app accepts (personal account vs work-only app, or vice
 *                  versa)
 *   AADSTS7000112  the enterprise application is disabled in the tenant
 *   AADSTS700054   the app is disabled for this sign-in
 *   AADSTS9002332  the app takes Microsoft Entra users only; a personal
 *                  Microsoft account tried it
 *
 * `(?!\d)` so a shorter code never matches inside a longer one.
 */
const APP_UNAVAILABLE: readonly AppUnavailableCase[] = [
  { code: 'AADSTS700016', pattern: /AADSTS700016(?!\d)/, kind: 'not-in-tenant' },
  { code: 'AADSTS700038', pattern: /AADSTS700038(?!\d)/, kind: 'not-an-app-id' },
  { code: 'AADSTS50020', pattern: /AADSTS50020(?!\d)/, kind: 'wrong-account-type' },
  { code: 'AADSTS7000112', pattern: /AADSTS7000112(?!\d)/, kind: 'disabled' },
  { code: 'AADSTS700054', pattern: /AADSTS700054(?!\d)/, kind: 'disabled' },
  { code: 'AADSTS9002332', pattern: /AADSTS9002332(?!\d)/, kind: 'entra-users-only' },
];

function matchAppUnavailable(err: unknown): AppUnavailableCase | null {
  const text = describeError(err);
  return APP_UNAVAILABLE.find((entry) => entry.pattern.test(text)) ?? null;
}

/**
 * True when the app registration itself is what Entra refused.
 *
 * Like a consent decision and unlike a transport failure, every one of these
 * reappears identically in the device-code flow — the app is just as absent or
 * as disabled there — so `auto` must not fall back after one.
 */
export function isAppUnavailable(err: unknown): boolean {
  return matchAppUnavailable(err) !== null;
}

/**
 * MSAL folds a trace ID, a correlation ID and a timestamp into the message on
 * their own lines. The first line is the sentence a user can act on.
 */
function firstLine(detail: string): string {
  const [first] = detail.split(/\r?\n/);
  return (first ?? detail).trim();
}

/**
 * What happened, why, and the exact flag that fixes it. The stack goes to
 * stderr separately; none of it belongs in front of the user.
 */
export function appUnavailableMessage(
  err: unknown,
  config: Pick<ServerConfig, 'clientId' | 'tenantId'>,
  source?: ClientIdSource,
): string {
  const clientId = config.clientId.trim() === '' ? '(none configured)' : config.clientId;
  const app = `application ${clientId}${clientIdOrigin(source)}`;
  const ownApp = `pass --client-id <application id> from your own Entra registration (or set ${CLIENT_ID_ENV})`;
  const reported = `Entra reported: ${firstLine(describeError(err))}`;

  const matched = matchAppUnavailable(err);
  if (!matched) {
    return [
      `Sign-in failed: Microsoft Entra would not accept ${app}.`,
      'Likely cause: the app registration is not usable from this tenant or with this account.',
      `Fix: ${ownApp}.`,
      reported,
    ].join('\n');
  }

  const lines: readonly string[] = ((): readonly string[] => {
    switch (matched.kind) {
      case 'not-in-tenant':
        return [
          `Sign-in failed (${matched.code}): Microsoft Entra found no ${app} in tenant ${config.tenantId}.`,
          'Likely cause: that app is registered single-tenant rather than multi-tenant (a bug if it is the one this project ships), the tenant has blocked it, or an administrator deleted its enterprise application.',
          `Fix: ${ownApp}.`,
        ];
      case 'not-an-app-id':
        return [
          `Sign-in failed (${matched.code}): Microsoft Entra rejected "${clientId}" as an application identifier.`,
          'Likely cause: no client ID was configured at all, so what went out was a placeholder rather than a real registration.',
          `Fix: ${ownApp}.`,
        ];
      case 'wrong-account-type':
        return [
          `Sign-in failed (${matched.code}): that account comes from a different identity provider than ${app} accepts.`,
          'Likely cause: a personal Microsoft account signing in to a work-or-school-only app, or a work or school account signing in to a consumer-only app.',
          `Fix: pass --tenant-id consumers for a personal Microsoft account, or sign in with a work or school account for a work app (this run used --tenant-id ${config.tenantId}).`,
        ];
      case 'disabled':
        return [
          `Sign-in failed (${matched.code}): ${app} is disabled in tenant ${config.tenantId}.`,
          'Likely cause: an administrator disabled its enterprise application, or blocked that service principal from sign-in.',
          `Fix: ask an administrator to re-enable it under Enterprise applications > Properties > "Enabled for users to sign-in", or ${ownApp}.`,
        ];
      case 'entra-users-only':
        return [
          `Sign-in failed (${matched.code}): ${app} accepts Microsoft Entra work or school users only.`,
          'Likely cause: you signed in with a personal Microsoft account against an app whose signInAudience excludes them.',
          `Fix: sign in with a work or school account, or ${ownApp} for an app that allows personal Microsoft accounts, and add --tenant-id consumers.`,
        ];
    }
  })();

  return [...lines, reported].join('\n');
}

/**
 * Pages MSAL serves on its loopback listener once the redirect lands.
 *
 * Self-contained by necessity: the loopback server only ever returns this one
 * string, so an external stylesheet, font, or script would simply 404. MSAL
 * substitutes nothing into either template (it swaps the whole default body for
 * this string, see LoopbackClient), so both are constants with no interpolation
 * and therefore nothing to escape, and the error page points at the terminal,
 * which is where the actual MSAL error is printed.
 */
const PAGE_STYLE =
  '<style>:root{color-scheme:light dark;--bg:#fff;--fg:#111;--muted:#555;--line:#d8d8d8}' +
  '@media(prefers-color-scheme:dark){:root{--bg:#161616;--fg:#f2f2f2;--muted:#a6a6a6;--line:#3a3a3a}}' +
  'body{margin:0;min-height:100vh;display:grid;place-items:center;background:var(--bg);color:var(--fg);' +
  'font:16px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}' +
  'main{max-width:30rem;padding:2rem;border:1px solid var(--line);border-radius:12px}' +
  'h1{margin:0 0 .75rem;font-size:1.2rem}p{margin:.5rem 0;color:var(--muted)}' +
  'code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}</style>';

const SUCCESS_TEMPLATE =
  '<!doctype html><html lang="en"><meta charset="utf-8"><title>Signed in</title>' +
  PAGE_STYLE +
  '<main><h1>Signed in to Microsoft 365</h1>' +
  '<p>Your account is cached locally on this machine, in this server\'s own encrypted token cache.</p>' +
  '<p>You can close this tab and go back to the terminal.</p></main>';

const ERROR_TEMPLATE =
  '<!doctype html><html lang="en"><meta charset="utf-8"><title>Sign-in failed</title>' +
  PAGE_STYLE +
  '<main><h1>Sign-in failed</h1>' +
  '<p>Microsoft 365 sign-in did not complete. Nothing was cached.</p>' +
  '<p>Close this tab and read the terminal: it prints the error and what to do about it.</p>' +
  '<p>To sign in from another device instead, retry with <code>--auth-flow device</code>.</p></main>';

/**
 * Opens the system browser without pulling in a dependency. Detached with
 * `stdio: 'ignore'` so the child can never inherit and pollute our stdout.
 */
async function openBrowser(url: string): Promise<void> {
  const parsed = new URL(url);
  if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') {
    throw new Error(`refusing to open a non-http(s) URL: ${parsed.protocol}`);
  }

  const [command, args] =
    process.platform === 'darwin'
      ? (['open', [url]] as const)
      : process.platform === 'win32'
        ? (['cmd', ['/c', 'start', '', url]] as const)
        : (['xdg-open', [url]] as const);

  await new Promise<void>((resolve, reject) => {
    const child = spawn(command, [...args], { detached: true, stdio: 'ignore' });
    child.once('error', reject);
    child.once('spawn', () => {
      child.unref();
      resolve();
    });
  });
}

function announceBrowser(): void {
  log('');
  log('Opening your browser to sign in to Microsoft 365...');
  log('If no window appears, sign in from another device with --auth-flow device.');
  log('');
}

function announceDeviceCode(response: DeviceCodeResponse): void {
  log('');
  log('To sign in to Microsoft 365:');
  log(`  1. Open ${response.verificationUri}`);
  log(`  2. Enter the code ${response.userCode}`);
  log('Waiting for you to finish signing in...');
  log('');
}

/**
 * True when the device code endpoint refused rather than issued.
 *
 * `@azure/msal-node` 6.0.0 destructures only the success fields out of that
 * endpoint's JSON body (`DeviceCodeClient.executePostRequestToDeviceCodeEndpoint`)
 * and drops `error` and `error_description` on the floor, so a refusal reaches
 * the callback as a response whose every field is `undefined`. The AADSTS code
 * is gone by then and cannot be recovered from MSAL.
 */
function isDeviceCodeMissing(response: DeviceCodeResponse): boolean {
  return typeof response.userCode !== 'string' || typeof response.verificationUri !== 'string';
}

/**
 * What to say when no device code was issued.
 *
 * Left unhandled this is the worst message in the program: the user is told to
 * "Open undefined" and "Enter the code undefined", MSAL then polls the token
 * endpoint with `device_code=undefined`, and what finally surfaces is
 * `post_request_failed: invalid_grant` — which names neither the cause nor the
 * fix. Every realistic cause is about the application, so this reads like the
 * other app-level failures.
 *
 * Deliberately quotes no AADSTS number: the real one was discarded upstream,
 * and inventing one the user cannot find in their sign-in logs is worse than
 * omitting it. It also keeps `matchAppUnavailable` from re-wrapping this
 * message in the generic one on the way out.
 */
export function deviceCodeRejectedMessage(
  config: Pick<ServerConfig, 'clientId' | 'tenantId'>,
  source?: ClientIdSource,
): string {
  const clientId = config.clientId.trim() === '' ? '(none configured)' : config.clientId;
  return [
    `Sign-in failed: Microsoft Entra issued no device code for application ${clientId}${clientIdOrigin(source)}.`,
    `Likely cause: that application is not usable from tenant ${config.tenantId} — never registered, registered single-tenant rather than multi-tenant, or blocked there. The device code endpoint also refuses --tenant-id common and organizations outright when it cannot identify a tenant from the application.`,
    `Fix: pass --client-id <application id> from your own Entra registration (or set ${CLIENT_ID_ENV}), and for a personal Microsoft account add --tenant-id consumers.`,
    'No trace ID is available: the endpoint\'s error body is discarded by @azure/msal-node before this process can read it. The tenant\'s Entra sign-in logs have the real code.',
  ].join('\n');
}

/**
 * The slice of `PublicClientApplication` this provider uses. Declaring it keeps
 * the provider testable with a stub instead of a live MSAL client; production
 * code passes nothing and gets the real one.
 */
export interface MsalPublicClient {
  acquireTokenSilent(request: SilentFlowRequest): Promise<AuthenticationResult>;
  acquireTokenInteractive(request: InteractiveRequest): Promise<AuthenticationResult>;
  acquireTokenByDeviceCode(request: DeviceCodeRequest): Promise<AuthenticationResult | null>;
  getTokenCache(): {
    getAllAccounts(): Promise<MsalAccountInfo[]>;
    removeAccount(account: MsalAccountInfo): Promise<void>;
  };
  clearCache(): void;
}

export function createAuthProvider(opts: {
  config: ServerConfig;
  /**
   * Where `config.clientId` came from, for the line `login` prints before it
   * sends the user to Entra. Optional: omitted, the line just names the app.
   */
  clientIdSource?: ClientIdSource;
  /** Test seam. Omit in production to get a real `PublicClientApplication`. */
  client?: MsalPublicClient;
}): AuthProvider {
  const { config, clientIdSource } = opts;

  const pca: MsalPublicClient =
    opts.client ??
    new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        // Never hard-coded: personal Microsoft accounts need a `consumers`
        // authority because refresh tokens minted through `common` began being
        // rejected at first refresh in June 2026. The config layer decides.
        authority: config.authority,
      },
      cache: { cachePlugin: createCachePlugin(config.cacheDir) },
    });

  // Avoids a cache file read on every single token request. Dropped whenever a
  // silent acquisition fails, so a sign-out from another process self-heals.
  let cachedAccount: MsalAccountInfo | null = null;

  async function resolveAccount(force = false): Promise<MsalAccountInfo | null> {
    if (cachedAccount && !force) return cachedAccount;
    const accounts = await pca.getTokenCache().getAllAccounts();
    cachedAccount = accounts[0] ?? null;
    return cachedAccount;
  }

  async function accountFromResult(result: AuthenticationResult | null): Promise<AccountInfo> {
    if (result?.account) {
      cachedAccount = result.account;
      return toAccountInfo(result.account);
    }
    // Some flows return a result without an account; the cache still has it.
    const account = await resolveAccount(true);
    if (!account) {
      throw new Error('Sign-in completed but no account was written to the token cache.');
    }
    return toAccountInfo(account);
  }

  /**
   * True when Entra refused on consent grounds rather than because the browser
   * leg failed. Every one of these reappears identically in the device-code
   * flow, so falling back to it only makes the user refuse twice.
   *
   *   AADSTS65004  the user declined the consent screen
   *   AADSTS65001  no consent recorded for this app or resource
   *   AADSTS90093  the scopes need an administrator
   *   AADSTS900941 admin consent required for the requested permissions
   */
  function isConsentDecision(err: unknown): boolean {
    const text = describeError(err);
    return (
      /AADSTS(65004|65001|90093|900941)/.test(text) ||
      /\baccess_denied\b/.test(text) ||
      /\bconsent_required\b/.test(text)
    );
  }

  /** What to tell someone whose consent screen was refused. */
  function consentFailureMessage(scopes: string[], detail: string): string {
    const blocked = scopes.filter((s) => TENANT_CONSENT_BLOCKED.has(s));
    const lines = [
      `Sign-in was refused at the consent screen: ${detail}`,
      '',
      `The app requested ${scopes.length} delegated permissions at once.`,
      'Nothing was granted and no session was created.',
      '',
      'What usually fixes this:',
      '  - Ask for less. Sign in with a narrower --groups set and widen it later;',
      '    Entra prompts again only for the permissions you add.',
      '  - Add --read-only, which drops every write permission from the request.',
    ];
    if (blocked.length > 0) {
      lines.push(
        '',
        `${blocked.length} of the requested permissions cannot be granted by an ordinary`,
        'user under the default Microsoft-managed tenant consent policy, however small',
        'the request is:',
        ...blocked.map((s) => `    ${s}`),
        '',
        'An administrator grants those once, in Entra under the app registration:',
        '  API permissions > Grant admin consent.',
      );
    }
    return lines.join('\n');
  }

  /**
   * The stack is diagnostics, not user guidance: it names MSAL internals and
   * says nothing about which flag to pass. stderr keeps it available without
   * putting it in front of the person who is stuck.
   */
  function logStack(err: unknown): void {
    if (err instanceof Error && err.stack) warn(err.stack);
  }

  /**
   * Rethrows an app-level rejection as guidance the user can act on. Returns
   * false when the error was something else, so the caller keeps its own
   * handling for browser and consent failures.
   */
  function throwIfAppUnavailable(err: unknown): void {
    if (!isAppUnavailable(err)) return;
    logStack(err);
    throw new Error(appUnavailableMessage(err, config, clientIdSource), { cause: err });
  }

  /** Authorization code + PKCE on a loopback redirect, run by MSAL itself. */
  async function loginByBrowser(scopes: string[]): Promise<AuthenticationResult> {
    const request: InteractiveRequest = {
      scopes: withRefreshScope(scopes),
      openBrowser,
      successTemplate: SUCCESS_TEMPLATE,
      errorTemplate: ERROR_TEMPLATE,
    };
    // Omitted entirely at 0 so MSAL asks the OS for a free port. When set, MSAL
    // still falls back to a random port if the preferred one is busy, which
    // breaks a tenant that pinned an exact redirect URI. The port actually used
    // is not reported back on the result and MSAL runs the listener itself, so
    // there is no cheap way to detect that here; the failure surfaces as a
    // redirect-URI mismatch, which browserFailureMessage names explicitly.
    if (config.authPort > 0) request.preferredPort = config.authPort;

    announceBrowser();
    return pca.acquireTokenInteractive(request);
  }

  /** Device authorization grant. Never opens a browser on this machine. */
  async function loginByDeviceCode(scopes: string[]): Promise<AuthenticationResult | null> {
    try {
      return await pca.acquireTokenByDeviceCode({
        scopes: withRefreshScope(scopes),
        deviceCodeCallback: (response) => {
          // MSAL invokes this synchronously and rethrows unchanged, so this is
          // the last point at which we still know a refusal happened rather
          // than a token-endpoint failure minutes later.
          if (isDeviceCodeMissing(response)) {
            throw new Error(deviceCodeRejectedMessage(config, clientIdSource));
          }
          announceDeviceCode(response);
        },
      });
    } catch (err) {
      // Reached both as the chosen flow and as the fallback. An app that is
      // absent or disabled is absent or disabled here too, so the same
      // guidance applies rather than a raw AADSTS code.
      throwIfAppUnavailable(err);
      throw err;
    }
  }

  return {
    async getToken(scopes: string[]): Promise<string> {
      const account = await resolveAccount();
      if (!account) {
        // Deliberately no interactive fallback: the MCP transport owns stdio and
        // a browser or prompt opened mid-request would corrupt the JSON-RPC
        // stream. The caller must run `login` out of band.
        throw new InteractionRequiredError(signInRequiredMessage(scopes));
      }

      try {
        const result = await pca.acquireTokenSilent({
          account,
          scopes: withRefreshScope(scopes),
        });
        if (!result.accessToken) {
          throw new Error('MSAL returned an empty access token');
        }
        return result.accessToken;
      } catch (err) {
        // Expired or revoked refresh token, a new scope needing consent, or a
        // conditional-access challenge all land here and all need interaction.
        cachedAccount = null;
        throw new InteractionRequiredError(signInRequiredMessage(scopes, describeError(err)));
      }
    },

    async getAccount(): Promise<AccountInfo | null> {
      const account = await resolveAccount(true);
      return account ? toAccountInfo(account) : null;
    },

    async login(scopes: string[]): Promise<AccountInfo> {
      // Before MSAL, before the network: an empty client ID can only come back
      // as AADSTS700038 several seconds later, saying less than this does.
      if (config.clientId.trim() === '') {
        throw new Error(missingClientIdMessage());
      }

      const flow = resolveAuthFlow(config);
      // Which app is about to be granted access, while the user can still stop.
      log(describeSignInTarget(config, clientIdSource));

      if (flow === 'device') {
        return accountFromResult(await loginByDeviceCode(scopes));
      }

      let result: AuthenticationResult;
      try {
        result = await loginByBrowser(scopes);
      } catch (err) {
        const detail = describeError(err);
        // Checked before anything else, and before the `browser` branch: this
        // is not a browser failure, and blaming firewalls or redirect URIs
        // would send the user hunting in the wrong place. The device-code flow
        // would be refused identically, so `auto` must not fall back either.
        // Order matters, and it is not the obvious one. Classify WHAT Entra
        // refused before deciding what to do about the flow: a consent refusal
        // reported as "check your firewall and redirect URI" sends the user
        // hunting through the portal for a problem that is not there. Both of
        // these checks therefore run ahead of the `flow === 'browser'` branch.
        throwIfAppUnavailable(err);
        if (isConsentDecision(err)) {
          // Identical in either flow — the device-code grant shows the same
          // consent screen — so there is nothing to fall back to.
          throw new Error(consentFailureMessage(scopes, detail), { cause: err });
        }
        // A missing or mismatched reply address IS browser-only: the
        // device-code grant uses no redirect URI and succeeds against the same
        // registration. Worth naming precisely, since "add this exact URI"
        // beats a list of things it might have been.
        if (isRedirectUriProblem(err)) {
          if (flow === 'browser') {
            throw new Error(redirectUriMessage(config.authPort, detail), { cause: err });
          }
          warn(`Browser sign-in failed: ${detail}`);
          warn(
            `The app registration has no usable redirect URI. Add ${redirectUriFor(config.authPort)} under "Mobile and desktop applications" to enable the browser flow; falling back to device code, which needs none.`,
          );
          return accountFromResult(await loginByDeviceCode(scopes));
        }
        if (flow === 'browser') {
          // Explicitly asked for the browser and the browser leg itself failed:
          // quietly doing something else is worse than stopping.
          throw new Error(browserFailureMessage(config.authPort, detail), { cause: err });
        }
        warn(`Browser sign-in failed: ${detail}`);
        warn(
          'Falling back to device code. Pass --auth-flow device to skip the browser attempt, or --auth-flow browser to fail instead of falling back.',
        );
        return accountFromResult(await loginByDeviceCode(scopes));
      }

      return accountFromResult(result);
    },

    async logout(): Promise<void> {
      const cache = pca.getTokenCache();
      const accounts = await cache.getAllAccounts();
      for (const account of accounts) {
        await cache.removeAccount(account);
      }
      cachedAccount = null;
      pca.clearCache();
      // Removing accounts rewrites the cache file; this deletes it and the key.
      await clearCache(config.cacheDir);
    },
  };
}
