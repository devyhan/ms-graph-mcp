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
import type {
  AccountInfo,
  AuthFlow,
  AuthProvider,
  DeviceLoginPrompt,
  DeviceLoginStatus,
  ServerConfig,
  SessionHealth,
} from '../contracts.js';
import {
  ADMIN_CONSENT_SCOPES,
  TENANT_CONSENT_BLOCKED,
  groupOwningScope,
} from '../tools/groups.js';
import { InteractionRequiredError } from '../contracts.js';
import { defaultPreflight } from './preflight.js';
import type { PreflightChecks } from './preflight.js';
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

/** Scopes in `scopes` that no amount of signing in will produce. */
function adminOnly(scopes: readonly string[]): string[] {
  return scopes.filter((scope) => ADMIN_CONSENT_SCOPES.has(scope));
}

/**
 * What a blocked caller reads.
 *
 * The important branch is the admin one. Telling someone to sign in again for a
 * scope their tenant reserves to an administrator sends them round a loop that
 * cannot terminate, and this server used to do exactly that — a chat search
 * failed for a missing `ChannelMessage.Read.All` and the advice was "run
 * login", which changed nothing however many times it was followed.
 */
function signInRequiredMessage(scopes: string[], detail?: string): string {
  const admin = adminOnly(scopes);
  const parts = ['No usable Microsoft 365 session.'];

  if (admin.length > 0) {
    const owners = [...new Set(admin.map((scope) => groupOwningScope(scope)).filter(Boolean))];
    parts.push(
      `Signing in again will NOT fix this: ${admin.join(', ')} ${admin.length === 1 ? 'is a scope' : 'are scopes'} only a Microsoft 365 tenant administrator can consent to for this application.`,
      `Ask an administrator to grant ${admin.join(' ')} to this application, then sign in again with \`--org-mode\` and the owning tool ${owners.length === 1 ? 'group' : 'groups'} enabled (${owners.join(', ')}) — \`npx @devyhan/ms-graph-mcp login --org-mode --groups …\`.`,
    );
  } else {
    // No terminal instruction here. This string is nested inside the tool result
    // as `Detail:`, under a headline that already says to call auth_begin_login,
    // and two different instructions in one response is how a reader ends up
    // following the wrong one. It also used to name the pre-rename package.
    parts.push(
      'A sign-in is needed. From a tool call, `auth_begin_login` starts one; from a terminal, `npx @devyhan/ms-graph-mcp login` does.',
    );
  }

  parts.push(`Scopes requested: ${scopes.length > 0 ? scopes.join(', ') : '(none)'}.`);
  if (detail) parts.push(`Underlying error: ${detail}`);
  return parts.join(' ');
}

/**
 * What a caller reads when the stored session cannot be refreshed because the
 * server signs in through an endpoint its application cannot use (AADSTS50194):
 * a single-tenant registration with the `/common` or `/organizations`
 * authority.
 *
 * Like the admin branch of `signInRequiredMessage` and unlike its sign-in
 * branch, this must say that signing in again changes nothing: the same
 * endpoint refuses every grant identically, so `auth_begin_login` — or another
 * CLI `login` — fails the same way. The fix is configuration, and the cached
 * session covers the corrected endpoint, so retrying comes before signing in.
 */
function tenantMismatchMessage(
  scopes: string[],
  config: Pick<ServerConfig, 'clientId' | 'tenantId'>,
  detail: string,
): string {
  const clientId = config.clientId.trim() === '' ? '(none configured)' : config.clientId;
  return [
    'No usable Microsoft 365 session.',
    `Signing in again will NOT fix this: application ${clientId} is registered single-tenant, but this server signs in through the /${config.tenantId} endpoint (AADSTS50194), which Entra refuses for such applications.`,
    `Set the tenant that owns the registration — --tenant-id <tenant-id> or ${TENANT_ID_ENV} — and restart the server; retry the call before signing in again, because the cached session is reused once the endpoint matches.`,
    `Scopes requested: ${scopes.length > 0 ? scopes.join(', ') : '(none)'}.`,
    `Underlying error: ${detail}`,
  ].join(' ');
}

/**
 * What a caller reads when the silent refresh fails for consent rather than
 * for a dead session: the session is fine, Entra just has no consent on record
 * for the scopes asked for — typically AADSTS65001 for a tool used the first
 * time since sign-in. Unlike the tenant mismatch above, signing in again is
 * the fix, and Entra prompts only for what was added, so this names the scopes
 * instead of the session, with both ways to consent to just them, each quoted
 * exactly as it is typed — no placeholder the reader has to fill in.
 */
export function consentRequiredMessage(scopes: string[]): string {
  const list = scopes.length > 0 ? scopes.join(', ') : '(none)';
  const toolArg = `{"scopes": [${scopes.map((s) => JSON.stringify(s)).join(', ')}]}`;
  return (
    `Blocked by missing consent for these scopes: ${list}. ` +
    'Sign in again for just these scopes — from a tool call, `auth_begin_login` with ' +
    `\`${toolArg}\`; from a terminal, \`npx @devyhan/ms-graph-mcp login --scopes ` +
    `${scopes.join(',')}` +
    '` — and Entra prompts only for what was added.'
  );
}

/**
 * What a caller reads when a Conditional Access policy blocked the refresh
 * (AADSTS53003): the session is not dead, but the policy demands something a
 * silent grant cannot do — MFA, a compliant or hybrid-joined device, an
 * approved client app. Only an interactive browser sign-in can complete those
 * challenges, so this retries the browser rather than the session, and warns
 * that the device-code grant is commonly stopped by the same policy.
 */
export function conditionalAccessMessage(scopes: string[], detail: string): string {
  const list = scopes.length > 0 ? scopes.join(', ') : '(none)';
  return [
    'Blocked by a Conditional Access policy (AADSTS53003): Microsoft Entra refused the request, not the session.',
    'The policy needs something this attempt did not satisfy — typically MFA, a compliant device, or an approved app.',
    'Sign in again in the browser — `auth_begin_login` from a tool call, `npx @devyhan/ms-graph-mcp login` from a terminal — which can complete interactive challenges such as MFA.',
    'The device-code flow is commonly blocked by the same policy, so prefer the browser here.',
    `Scopes requested: ${list}.`,
    `Underlying error: ${detail}`,
  ].join(' ');
}

/**
 * What a caller reads when the request never reached Entra at all: DNS, proxy,
 * TLS, or simply offline. Signing in again changes nothing while the route is
 * down — both grants need the same network — so this says to fix the route and
 * retry the call, and names the proxy behavior: the server honors proxy
 * environment variables automatically, so a failure here means the proxy
 * itself (or the route to it) is at fault, not a missing flag.
 */
export function networkFailureMessage(scopes: string[], detail: string): string {
  const list = scopes.length > 0 ? scopes.join(', ') : '(none)';
  return [
    'Could not reach Microsoft Entra: the request failed before any sign-in state mattered.',
    'Signing in again will NOT fix this while the network is down — check the connection and retry the call first.',
    'Behind a proxy, the server honors proxy environment variables automatically (HTTPS_PROXY/ALL_PROXY plus NO_PROXY); if this persists, check the proxy itself is reachable.',
    `Scopes requested: ${list}.`,
    `Underlying error: ${detail}`,
  ].join(' ');
}

/**
 * What a caller reads when the browser flow's loopback redirect could not be
 * served and MSAL gave up waiting (`loopback_server_timeout`). A firewall, a
 * VPN that swallows loopback, or the pinned `--auth-port` taken by another
 * process all do that. The device-code grant listens on nothing, so it is the
 * way out rather than another browser attempt.
 */
export function loopbackTimeoutMessage(authPort: number, detail: string): string {
  const port = authPort > 0 ? ` (--auth-port ${authPort})` : '';
  return [
    `Browser sign-in timed out waiting for the loopback redirect: ${detail}`,
    'Likely causes:',
    `  - The loopback listener could not be served${port}. A firewall, a VPN, or another process on the port will do that.`,
    '  - The sign-in was never finished in the browser window that opened.',
    'To sign in from another device instead, retry with --auth-flow device.',
  ].join('\n');
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
export function isConsentDecision(err: unknown): boolean {
  const text = describeError(err);
  return (
    /AADSTS(65004|65001|90093|900941)/.test(text) ||
    /\baccess_denied\b/.test(text) ||
    /\bconsent_required\b/.test(text)
  );
}

/**
 * True when a Conditional Access policy blocked the grant (AADSTS53003).
 * Unlike a consent refusal this is not answered by asking for less: the
 * policy wants an interactive challenge, so the browser — not the device
 * code — is the retry.
 */
export function isConditionalAccess(err: unknown): boolean {
  return /AADSTS53003(?!\d)/.test(describeError(err));
}

/**
 * True when the request never reached Entra: DNS, connection, TLS, or a fetch
 * that failed outright. Deliberately narrow — a browser-launch failure such
 * as `spawn xdg-open ENOENT` is local, not network, and must keep falling
 * back to the device code instead of being told to check the proxy.
 */
export function isNetworkFailure(err: unknown): boolean {
  const text = describeError(err);
  return (
    /\b(ENOTFOUND|EAI_AGAIN|ECONNREFUSED|ECONNRESET|ECONNABORTED|ETIMEDOUT|ENETUNREACH|EHOSTUNREACH|ENETDOWN|EHOSTDOWN|EPROTO|ERR_SOCKET_CLOSED)\b/.test(
      text,
    ) ||
    /\bgetaddrinfo\b/.test(text) ||
    /fetch failed/i.test(text) ||
    /network (request failed|error|unreachable|timeout|is unreachable)/i.test(text) ||
    /unable to resolve/i.test(text)
  );
}

/**
 * True when the browser flow's loopback redirect was never served and MSAL
 * gave up waiting. Browser-only by construction: the device-code grant
 * listens on nothing, so it sidesteps this one entirely.
 */
export function isLoopbackTimeout(err: unknown): boolean {
  const text = describeError(err);
  return (
    /loopback_server_timeout/.test(text) ||
    (/loopback/i.test(text) && /timed?\s?out/i.test(text)) ||
    (/timed?\s?out/i.test(text) && /redirect/i.test(text))
  );
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

/** What to tell someone whose consent screen was refused. */
export function consentFailureMessage(scopes: string[], detail: string): string {
  const blocked = scopes.filter((s) => TENANT_CONSENT_BLOCKED.has(s));
  const lines = [
    `Sign-in was refused at the consent screen: ${detail}`,
    '',
    `The app requested ${scopes.length} delegated permissions at once.`,
    'Nothing was granted and no session was created.',
    '',
    'What usually fixes this:',
    '  - Ask for less. A bare `login` asks only for the minimal set; consent to',
    '    more with `login --scopes …` as tools need them, and Entra prompts again',
    '    only for the permissions you add.',
    '  - Add --read-only, which drops every write permission from a --full request.',
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
 * Maps a sign-in failure to the guidance its cause deserves, or returns the
 * raw message when the cause is unrecognised.
 *
 * The provider's own catches classify inline so they can also decide about
 * fallback; this is for the paths that only record a failure — the
 * `auth_login_status` detail and the device-login status — where there is no
 * fallback to decide, only a reader to inform. Unknown stays raw: inventing
 * guidance for an error never seen is worse than quoting it.
 */
export function mapLoginFailure(
  scopes: string[],
  err: unknown,
  config: Pick<ServerConfig, 'clientId' | 'tenantId'>,
  authPort = 0,
  source?: ClientIdSource,
): string {
  const detail = describeError(err);
  if (isAppUnavailable(err)) return appUnavailableMessage(err, config, source);
  if (isConsentDecision(err)) return consentFailureMessage(scopes, detail);
  if (isConditionalAccess(err)) return conditionalAccessMessage(scopes, detail);
  if (isNetworkFailure(err)) return networkFailureMessage(scopes, detail);
  if (isLoopbackTimeout(err)) return loopbackTimeoutMessage(authPort, detail);
  if (isRedirectUriProblem(err)) return redirectUriMessage(authPort, detail);
  return detail;
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

/** The variable the config layer reads a tenant ID from. */
const TENANT_ID_ENV = 'MS365_MCP_TENANT_ID';

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
  | 'entra-users-only'
  | 'single-tenant-common';

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
 *   AADSTS50194    the app is registered single-tenant but the authority is
 *                  a tenantless endpoint (`/common` or `/organizations`),
 *                  which Entra refuses for such applications
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
  { code: 'AADSTS50194', pattern: /AADSTS50194(?!\d)/, kind: 'single-tenant-common' },
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
      case 'single-tenant-common':
        return [
          `Sign-in failed (${matched.code}): ${app} is registered single-tenant, but this run signs in through the /${config.tenantId} endpoint, which Entra refuses for such applications.`,
          'Likely cause: the registration\'s "Supported account types" is "Accounts in this organizational directory only" — the portal default for a new registration — while the server uses a tenantless authority.',
          `Fix: sign in to the tenant that owns the registration with --tenant-id <tenant-id> (or set ${TENANT_ID_ENV}), and sign in again. From an MCP client, set the tenant in the server configuration and restart the server.`,
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
  // Emptiness counts, not just absence. A response carrying `userCode: ''` is a
  // response with no code in it, and telling someone to "enter the code " is the
  // failure this function exists to stop — checking only the type let that
  // through.
  const present = (value: unknown): boolean => typeof value === 'string' && value.trim() !== '';
  return !present(response.userCode) || !present(response.verificationUri);
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

/** How long the session-health probe waits for MSAL before calling it unusable. */
export const SESSION_HEALTH_TIMEOUT_MS = 8000;

/** Rejects with `Error('timed out …')` when `promise` outlasts `ms`. Never throws itself. */
function withTimeout<T>(promise: Promise<T>, ms: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_resolve, reject) => {
    // Ref'd on purpose: an unref'd timer lets the event loop drain first, and
    // the timeout that never fires is no timeout at all.
    timer = setTimeout(() => reject(new Error(`timed out after ${ms}ms`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

/**
 * One line saying why the stored session cannot mint tokens, classifying by
 * code the same way `getToken` does. A status note, not guidance: the reader
 * is a status report, and the full message lives on the error itself.
 */
function healthNote(scopes: string[], err: unknown): string {
  const detail = firstLine(describeError(err));
  if (matchAppUnavailable(err)?.kind === 'single-tenant-common') {
    return `Single-tenant application on a tenantless endpoint (AADSTS50194); set the owning tenant: ${detail}`;
  }
  if (isConditionalAccess(err)) {
    return `Conditional Access blocked the silent refresh (AADSTS53003); sign in again in the browser: ${detail}`;
  }
  if (isConsentDecision(err)) {
    return `Missing consent for ${scopes.length > 0 ? scopes.join(', ') : '(none)'}; sign in again for just those scopes: ${detail}`;
  }
  if (isNetworkFailure(err)) {
    return `Microsoft Entra was unreachable; check the network and retry: ${detail}`;
  }
  if (isLoopbackTimeout(err)) {
    return `The browser redirect timed out; retry with --auth-flow device: ${detail}`;
  }
  return `The stored session cannot be refreshed: ${detail}`;
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
  /**
   * Pre-sign-in environment checks. Omit in production to probe the real
   * environment; tests inject fakes so no check touches the network or the
   * process environment.
   */
  preflight?: PreflightChecks;
}): AuthProvider {
  const { config, clientIdSource } = opts;
  const preflight: PreflightChecks = opts.preflight ?? defaultPreflight;

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
      // guidance applies rather than a raw AADSTS code — and so do the
      // policy and transport failures, which the device grant meets head-on.
      throwIfAppUnavailable(err);
      throwIfKnownGrantFailure(scopes, err);
      throw err;
    }
  }

  /**
   * Rethrows a policy or transport failure as guidance the user can act on.
   * Shared by the blocking device grant and the tool-started one: both meet
   * the same endpoint refusals, and both used to surface them raw.
   */
  function throwIfKnownGrantFailure(scopes: string[], err: unknown): void {
    const detail = describeError(err);
    if (isConsentDecision(err)) {
      logStack(err);
      throw new Error(consentFailureMessage(scopes, detail), { cause: err });
    }
    if (isConditionalAccess(err)) {
      logStack(err);
      throw new Error(conditionalAccessMessage(scopes, detail), { cause: err });
    }
    if (isNetworkFailure(err)) {
      logStack(err);
      throw new Error(networkFailureMessage(scopes, detail), { cause: err });
    }
  }

  /**
   * A device-code sign-in in flight, kept because a tool call cannot block for
   * the minutes a person needs to finish one.
   *
   * MSAL does the polling: `acquireTokenByDeviceCode` returns a promise that
   * settles when the person finishes, and the code arrives earlier through the
   * callback. So the tool that starts a sign-in awaits the CODE, not the
   * PROMISE, and the tool that reports progress reads the flags this object
   * carries. Nothing here holds a token.
   */
  let pending:
    | {
        prompt: DeviceLoginPrompt;
        state: 'pending' | 'signedIn' | 'failed';
        detail?: string;
        /** The scopes the in-flight attempt asks for, so a failure can map them into guidance. */
        scopes: string[];
      }
    | null = null;

  /**
   * A start that has not yet produced a code.
   *
   * Two tool calls can arrive before MSAL has said anything, and `pending` is
   * only written from inside the callback, so a guard that reads it would not
   * have fired yet: both calls would start a flow and the code the first caller
   * was already shown would be dead. Sharing the in-flight promise is what makes
   * "asking twice returns the same code" true rather than merely usually true.
   */
  let starting: Promise<DeviceLoginPrompt> | null = null;

  async function beginDeviceLogin(scopes: string[]): Promise<DeviceLoginPrompt> {
    // A live code is reused rather than replaced. Starting a fresh flow would
    // hand out a second code and silently invalidate the one being typed.
    if (pending?.state === 'pending' && Date.parse(pending.prompt.expiresAt) > Date.now()) {
      return pending.prompt;
    }
    if (starting !== null) return starting;

    let deliverPrompt: (prompt: DeviceLoginPrompt) => void = () => {};
    let failPrompt: (error: unknown) => void = () => {};
    const prompted = new Promise<DeviceLoginPrompt>((resolve, reject) => {
      deliverPrompt = resolve;
      failPrompt = reject;
    });

    const entry: {
      prompt: DeviceLoginPrompt;
      state: 'pending' | 'signedIn' | 'failed';
      detail?: string;
      scopes: string[];
    } = {
      // Replaced the moment the callback fires; `prompted` is what the caller
      // waits on, so this placeholder is never observable.
      prompt: { userCode: '', verificationUri: '', expiresAt: '', pollIntervalSeconds: 5 },
      state: 'pending',
      scopes: [...scopes],
    };

    // Set BEFORE the call, not after it. MSAL may invoke deviceCodeCallback
    // synchronously, and the callback clears this; assigning afterwards would
    // overwrite that clear with a stale promise and the guard would then hold
    // for the life of the flow rather than until a code exists.
    starting = prompted;

    void pca
      .acquireTokenByDeviceCode({
        scopes: withRefreshScope(scopes),
        deviceCodeCallback: (response) => {
          if (isDeviceCodeMissing(response)) {
            const error = new Error(deviceCodeRejectedMessage(config, clientIdSource));
            failPrompt(error);
            throw error;
          }
          const prompt: DeviceLoginPrompt = {
            userCode: response.userCode,
            verificationUri: response.verificationUri,
            expiresAt: new Date(Date.now() + response.expiresIn * 1000).toISOString(),
            pollIntervalSeconds: Math.max(1, response.interval),
          };
          entry.prompt = prompt;
          pending = entry;
          deliverPrompt(prompt);
          // The guard covers exactly the window before a code exists. Past this
          // point the expiry check above decides whether to reuse, and holding
          // it any longer would stop a dead code ever being replaced.
          starting = null;
        },
      })
      .then(
        (result) => {
          if (result === null) {
            entry.state = 'failed';
            entry.detail = 'Microsoft Entra returned no result for the device code.';
          } else {
            entry.state = 'signedIn';
            // The account cache is keyed on a session that did not exist a
            // moment ago, so force the next read to go to MSAL.
            cachedAccount = null;
          }
        },
        (err: unknown) => {
          entry.state = 'failed';
          // Mapped, not raw: a bare AADSTS code in a status field says nothing
          // about which knob to turn. Unknown failures still quote verbatim.
          entry.detail = mapLoginFailure(entry.scopes, err, config, config.authPort, clientIdSource);
          // The caller may still be waiting for a code that is never coming:
          // MSAL rejects before the callback for anything that stops the device
          // authorization request itself — no network, a proxy or TLS failure, a
          // tenant that does not resolve. Without this the tool call waits for
          // ever, which is worse than the error it is hiding. Rejecting an
          // already-resolved promise is a no-op, so the normal path is
          // unaffected. The rejection stays raw here: the outer `beginDeviceLogin`
          // maps it on the way out, once, with the same guidance.
          failPrompt(err);
          starting = null;
          // Nothing arrived, so there is no code to report a state for. Leaving
          // `pending` set would answer "pending" about a flow that is dead.
          if (pending !== entry) pending = null;
        },
      );

    return prompted;
  }

  return {
    async beginDeviceLogin(scopes: string[]): Promise<DeviceLoginPrompt> {
      try {
        return await beginDeviceLogin(scopes);
      } catch (err) {
        throwIfAppUnavailable(err);
        throwIfKnownGrantFailure(scopes, err);
        throw err;
      }
    },

    async deviceLoginStatus(): Promise<DeviceLoginStatus> {
      if (pending === null) return { state: 'none' };
      if (pending.state === 'signedIn') {
        const account = await resolveAccount(true);
        return { state: 'signedIn', account: account?.username ?? undefined };
      }
      if (pending.state === 'failed') {
        return { state: 'failed', detail: pending.detail };
      }
      if (Date.parse(pending.prompt.expiresAt) <= Date.now()) {
        return {
          state: 'expired',
          detail: 'The code timed out. Start another sign-in to get a fresh one.',
          prompt: pending.prompt,
        };
      }
      return { state: 'pending', prompt: pending.prompt };
    },

    async getToken(scopes: string[]): Promise<string> {
      const account = await resolveAccount();
      if (!account) {
        // Deliberately no interactive fallback: a data request must not
        // unexpectedly open a browser and stay pending while a person signs in.
        // The caller must explicitly use auth_begin_login or the CLI command.
        throw new InteractionRequiredError(signInRequiredMessage(scopes), adminOnly(scopes));
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
        // Except the endpoint refusal: signing in again through the same
        // tenantless endpoint fails identically, so this names the tenant fix
        // and marks the error for the tool-result headline. Only this kind is
        // special-cased — a wrong-account-type refusal, for example, can still
        // be fixed by signing in with a different account.
        if (matchAppUnavailable(err)?.kind === 'single-tenant-common') {
          throw new InteractionRequiredError(
            tenantMismatchMessage(scopes, config, describeError(err)),
            adminOnly(scopes),
            true,
            false,
            false,
            false,
            false,
            scopes,
          );
        }
        // A Conditional Access refusal is not a dead session either: the
        // refresh token is fine, but the policy demands an interactive
        // challenge a silent grant cannot complete. Checked before consent —
        // a policy verdict is the more specific refusal when both match.
        if (isConditionalAccess(err)) {
          throw new InteractionRequiredError(
            conditionalAccessMessage(scopes, describeError(err)),
            adminOnly(scopes),
            false,
            false,
            true,
            false,
            false,
            scopes,
          );
        }
        // A consent refusal is not a dead session either: the refresh token is
        // fine, Entra just has no consent on record for the scopes asked for.
        // Signing in again for just those scopes prompts only for what was
        // added, so this names the scopes instead of the session. The admin
        // filter still applies: consent the tenant reserves to an administrator
        // is refused the same way, and the headline must say so.
        if (isConsentDecision(err)) {
          throw new InteractionRequiredError(
            consentRequiredMessage(scopes),
            adminOnly(scopes),
            false,
            true,
            false,
            false,
            false,
            scopes,
          );
        }
        // The request never reached Entra: nothing about the session mattered,
        // so the fix is the route, not another sign-in.
        if (isNetworkFailure(err)) {
          throw new InteractionRequiredError(
            networkFailureMessage(scopes, describeError(err)),
            adminOnly(scopes),
            false,
            false,
            false,
            true,
            false,
            scopes,
          );
        }
        // A stale loopback error surfacing on the silent path (a cached
        // failure, a proxy that answers for localhost): the browser leg is
        // what timed out, not the session.
        if (isLoopbackTimeout(err)) {
          throw new InteractionRequiredError(
            loopbackTimeoutMessage(config.authPort, describeError(err)),
            adminOnly(scopes),
            false,
            false,
            false,
            false,
            true,
            scopes,
          );
        }
        throw new InteractionRequiredError(
          signInRequiredMessage(scopes, describeError(err)),
          adminOnly(scopes),
          false,
          false,
          false,
          false,
          false,
          scopes,
        );
      }
    },

    async sessionHealth(scopes: string[] = ['User.Read']): Promise<SessionHealth> {
      // Status only: whether the stored session still mints tokens. Unlike
      // getToken this never clears the shared cached account — a probe that
      // evicts what it probes would turn one transient failure into a
      // sign-out — and it never throws for an unusable session.
      let accounts: MsalAccountInfo[];
      try {
        accounts = await pca.getTokenCache().getAllAccounts();
      } catch (error) {
        return {
          tokenUsable: false,
          note: `Could not read the token cache: ${describeError(error)}`,
        };
      }
      const account = accounts[0] ?? cachedAccount;
      if (!account) {
        return { tokenUsable: false, note: 'No cached account: not signed in.' };
      }
      try {
        const result = await withTimeout(
          pca.acquireTokenSilent({ account, scopes: withRefreshScope(scopes) }),
          SESSION_HEALTH_TIMEOUT_MS,
        );
        if (!result.accessToken) {
          return { tokenUsable: false, note: 'MSAL returned an empty access token.' };
        }
        return { tokenUsable: true };
      } catch (err) {
        return { tokenUsable: false, note: healthNote(scopes, err) };
      }
    },

    async getAccount(): Promise<AccountInfo | null> {
      const account = await resolveAccount(true);
      return account ? toAccountInfo(account) : null;
    },

    async login(scopes: string[], flowOverride?: AuthFlow): Promise<AccountInfo> {
      // Before MSAL, before the network: an empty client ID can only come back
      // as AADSTS700038 several seconds later, saying less than this does.
      if (config.clientId.trim() === '') {
        throw new Error(missingClientIdMessage());
      }

      const flow = flowOverride ?? resolveAuthFlow(config);
      // Which app is about to be granted access, while the user can still stop.
      log(describeSignInTarget(config, clientIdSource));

      // Before MSAL opens anything: when the browser verifiably cannot work —
      // no desktop session, no loopback listener — `auto` goes straight to the
      // device code instead of failing through the browser first, and an
      // explicit `browser` stops with the browser guidance instead of an MSAL
      // timeout minutes later. `unknown` keeps the current behaviour: the
      // attempt goes ahead as before.
      if (flow === 'auto' || flow === 'browser') {
        let blockers: string[] = [];
        try {
          const [browser, loopback] = await Promise.all([
            preflight.canOpenBrowser(),
            preflight.canBindLoopback(config.authPort),
          ]);
          blockers = [browser, loopback]
            .filter((result) => result.verdict === 'impossible')
            .map((result) => result.reason);
        } catch {
          blockers = [];
        }
        if (blockers.length > 0) {
          const why = blockers.join('; ');
          if (flow === 'browser') {
            throw new Error(browserFailureMessage(config.authPort, `pre-sign-in check: ${why}`));
          }
          warn(`Browser sign-in is unavailable (${why}); using device code directly.`);
          warn('Pass --auth-flow browser to fail instead of using device code.');
          return accountFromResult(await loginByDeviceCode(scopes));
        }
      }

      // Warn-only, every flow: a dead route fails both grants identically, and
      // hearing it up front beats one timeout per grant. Never stops anything.
      try {
        const reach = await preflight.canReachEntra();
        if (reach.verdict === 'impossible') {
          warn(`${reach.reason}; sign-in will likely fail. Check the network and proxy, then retry.`);
        }
      } catch {
        // The probe itself must never break a sign-in.
      }

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
        // hunting through the portal for a problem that is not there. All of
        // these checks therefore run ahead of the `flow === 'browser'` branch.
        throwIfAppUnavailable(err);
        if (isConsentDecision(err)) {
          // Identical in either flow — the device-code grant shows the same
          // consent screen — so there is nothing to fall back to.
          throw new Error(consentFailureMessage(scopes, detail), { cause: err });
        }
        if (isConditionalAccess(err)) {
          // A policy block is about the grant, not the browser leg: the
          // device grant meets the same policy, so `auto` falls back with a
          // warning rather than failing — the policy may still let the device
          // grant through, but commonly does not.
          if (flow === 'browser') {
            throw new Error(conditionalAccessMessage(scopes, detail), { cause: err });
          }
          warn(`Browser sign-in was blocked by a Conditional Access policy: ${detail}`);
          warn(
            'Falling back to device code, which the same policy commonly blocks too; prefer finishing the required challenge (MFA, compliant device) in the browser.',
          );
          return accountFromResult(await loginByDeviceCode(scopes));
        }
        if (isNetworkFailure(err)) {
          // Both grants need the same network, so `auto` must not fall back
          // into a second identical failure.
          logStack(err);
          throw new Error(networkFailureMessage(scopes, detail), { cause: err });
        }
        if (isLoopbackTimeout(err)) {
          // Browser-only, like the reply-address problem below: the device
          // grant listens on nothing, so it sidesteps this one entirely.
          if (flow === 'browser') {
            throw new Error(loopbackTimeoutMessage(config.authPort, detail), { cause: err });
          }
          warn(`Browser sign-in timed out waiting for the loopback redirect: ${detail}`);
          warn('Falling back to device code, which needs no loopback listener.');
          return accountFromResult(await loginByDeviceCode(scopes));
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
