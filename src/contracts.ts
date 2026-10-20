/**
 * Shared contracts for the Microsoft Graph MCP server.
 *
 * Every module in this project builds against these types. Nothing here has a
 * runtime dependency on the MCP SDK, so tool modules stay testable in isolation.
 */

import type { z } from 'zod';

/** Which Microsoft Graph endpoint a request targets. */
export type GraphVersion = 'v1.0' | 'beta';

/** HTTP methods this server is willing to send to Microsoft Graph. */
export type HttpMethod = 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';

// ---------------------------------------------------------------------------
// Authentication
// ---------------------------------------------------------------------------

/** Details about the currently signed-in principal. */
export interface AccountInfo {
  username: string;
  name?: string;
  tenantId?: string;
  homeAccountId: string;
}

/**
 * Login-time snapshot of the signed-in identity, cached beside the token
 * cache so status output can name the session without a Graph call.
 *
 * Stored as JSON encrypted with the same 32-byte key as the MSAL blob, never
 * inside that blob: it must survive token-cache wipes and die with key
 * rotation. `capturedAt` is an ISO-8601 instant; `source` says whether the
 * snapshot came from the MSAL account or a Graph `/me` call.
 */
export interface IdentityRecord {
  homeAccountId: string;
  username: string;
  name?: string;
  tenantId?: string;
  clientId: string;
  authority: string;
  capturedAt: string;
  source: 'msal-account' | 'graph-me';
}

/**
 * Acquires Microsoft Graph access tokens for delegated scopes.
 *
 * Implementations cache tokens and refresh silently. `getToken` must never
 * block on interactive input once a session exists.
 */
export interface AuthProvider {
  /** Returns a bearer token valid for every scope in `scopes`. */
  getToken(scopes: string[]): Promise<string>;
  /** The signed-in account, or null when no session exists. */
  getAccount(): Promise<AccountInfo | null>;
  /**
   * Starts an interactive sign-in and resolves once a session exists.
   *
   * `flow` is normally omitted so the configured CLI flow is used. Callers
   * that expose a particular user experience, such as the browser-opening MCP
   * sign-in tool, may force that flow without changing the server-wide config.
   */
  login(scopes: string[], flow?: AuthFlow): Promise<AccountInfo>;
  /** Clears the cached session. */
  logout(): Promise<void>;
  /**
   * Starts a device-code sign-in and returns as soon as there is a code to show,
   * without waiting for the person to use it.
   *
   * This remains available for headless callers. The normal MCP sign-in tool
   * starts the browser flow in the background instead, while the device grant
   * returns the code and URL as ordinary tool output.
   */
  beginDeviceLogin(scopes: string[]): Promise<DeviceLoginPrompt>;
  /** Where a sign-in started by `beginDeviceLogin` has got to. Never a token. */
  deviceLoginStatus(): Promise<DeviceLoginStatus>;
  /**
   * Probes whether the stored session can still mint tokens, without needing
   * interaction and without disturbing the cached account. Status only — never
   * a token, and never throws for an unusable session.
   *
   * Optional so scope-only catalogues (see `scopeReadingOnlyAuth`) need not
   * pretend to have a session to probe: callers that find it absent must report
   * "not probed" rather than calling `getToken`, which refuses there.
   */
  sessionHealth?(scopes?: string[]): Promise<SessionHealth>;
}

/**
 * Whether the stored session can still be refreshed.
 *
 * `tokenUsable` answers the question; `note` says why not (or what was not
 * checked) in one line a status report can quote. Never carries a token.
 */
export interface SessionHealth {
  tokenUsable: boolean;
  note?: string | undefined;
}

/** What a person needs in order to finish a device-code sign-in. */
export interface DeviceLoginPrompt {
  /** The code to type at `verificationUri`. */
  userCode: string;
  verificationUri: string;
  /** When this code stops working, as an ISO-8601 instant. */
  expiresAt: string;
  /** Seconds MSAL waits between polls, so a caller can pace its own checks. */
  pollIntervalSeconds: number;
}

/**
 * The state of a device-code sign-in.
 *
 * `account` is the signed-in identity, never a token: SECURITY.md puts "the
 * model must never see one" in scope, and a status field is exactly where one
 * would leak by accident.
 */
export interface DeviceLoginStatus {
  state: 'none' | 'pending' | 'signedIn' | 'expired' | 'failed';
  account?: string | undefined;
  /** Present on `failed`, and on `expired` to say what to do next. */
  detail?: string | undefined;
  /** Present while `pending`, so a caller can show the code again. */
  prompt?: DeviceLoginPrompt | undefined;
}

/** Raised when a token cannot be acquired without user interaction. */
export class InteractionRequiredError extends Error {
  override readonly name = 'InteractionRequiredError';
  /**
   * Scopes among those requested that only a tenant administrator can grant.
   * Present when signing in again cannot fix the failure, so the caller can say
   * so instead of sending the user round a loop that will not terminate.
   */
  readonly adminScopes: readonly string[];
  /**
   * True when the server signs in through an endpoint its application cannot
   * use — today, a single-tenant registration with the `/common` or
   * `/organizations` authority (AADSTS50194). Signing in again fails
   * identically; the message names the tenant to configure instead.
   */
  readonly tenantMismatch: boolean;
  /**
   * True when the session is fine but Entra has no consent on record for the
   * scopes asked for (AADSTS65001 and friends). Signing in again for just
   * those scopes is the fix; Entra prompts only for what was added.
   */
  readonly consentMissing: boolean;
  /**
   * True when a Conditional Access policy blocked the request (AADSTS53003).
   * The session is not dead: the policy demands something a silent refresh
   * cannot do (MFA, a compliant device), so signing in again in the browser —
   * which can complete those challenges — is the fix, while the device-code
   * flow is commonly blocked by the same policy.
   */
  readonly conditionalAccess: boolean;
  /**
   * True when the request never reached Entra (DNS, proxy, TLS, offline).
   * Signing in again changes nothing while the network is down; retrying the
   * call once it is back is the fix.
   */
  readonly networkFailure: boolean;
  /**
   * True when the browser flow's loopback redirect could not be served and
   * timed out. The device-code flow needs no loopback listener, so it is the
   * fix rather than another browser attempt.
   */
  readonly loopbackTimeout: boolean;
  /**
   * The scopes the failed call asked for, so a headline can echo them back in
   * paste-ready form instead of gesturing at "the scopes the detail names".
   * Empty when the raiser did not record them; headlines then stay generic.
   */
  readonly scopes: readonly string[];
  constructor(
    message: string,
    adminScopes: readonly string[] = [],
    tenantMismatch = false,
    consentMissing = false,
    conditionalAccess = false,
    networkFailure = false,
    loopbackTimeout = false,
    scopes: readonly string[] = [],
  ) {
    super(message);
    this.adminScopes = adminScopes;
    this.tenantMismatch = tenantMismatch;
    this.consentMissing = consentMissing;
    this.conditionalAccess = conditionalAccess;
    this.networkFailure = networkFailure;
    this.loopbackTimeout = loopbackTimeout;
    this.scopes = scopes;
  }
}

/**
 * Which interactive sign-in flow `login` uses. Both are OAuth 2.0.
 *
 * - `browser` — authorization code with PKCE (RFC 7636) against a loopback
 *   redirect. MSAL opens the system browser and catches the redirect on
 *   127.0.0.1. Best experience, but needs a desktop session and a free port.
 * - `device`  — device authorization grant (RFC 8628). Prints a code and a URL
 *   to sign in from any other device. Works over SSH and in containers.
 * - `auto`    — try `browser`, fall back to `device` when no browser or no
 *   loopback port is available. The default.
 */
export type AuthFlow = 'auto' | 'browser' | 'device';

// ---------------------------------------------------------------------------
// Microsoft Graph client
// ---------------------------------------------------------------------------

/** A single Microsoft Graph request. */
export interface GraphRequestOptions {
  /** Graph-relative path beginning with `/`, e.g. `/me/messages`. */
  path: string;
  method?: HttpMethod;
  version?: GraphVersion;
  /** Query parameters. `undefined` values are dropped. */
  query?: Record<string, string | number | boolean | undefined>;
  body?: unknown;
  headers?: Record<string, string>;
  /** Delegated scopes required for this call. */
  scopes: string[];
  /**
   * How many pages of a collection to follow via `@odata.nextLink`.
   * Defaults to 1, meaning no automatic paging.
   */
  maxPages?: number;
  signal?: AbortSignal;
}

/** The outcome of a Graph request. */
export interface GraphResponse<T = unknown> {
  status: number;
  data: T;
  /** Opaque cursor for the next page, when the collection has more pages. */
  nextLink?: string;
  /** Opaque cursor for a future delta query, when the resource supports delta. */
  deltaLink?: string;
}

/** One entry in a `$batch` payload. */
export interface GraphBatchRequest {
  id: string;
  method: HttpMethod;
  /** Graph-relative URL including query string, e.g. `/me/messages?$top=5`. */
  url: string;
  body?: unknown;
  headers?: Record<string, string>;
  dependsOn?: string[];
}

/** One entry in a `$batch` response. */
export interface GraphBatchResponse {
  id: string;
  status: number;
  body?: unknown;
  headers?: Record<string, string>;
}

/** Talks to Microsoft Graph with retry, paging, and host pinning. */
export interface GraphClient {
  request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>>;
  /**
   * Sends up to 20 requests as one `$batch`. Larger arrays are split into
   * consecutive batches automatically. Each entry carries its own status.
   */
  batch(requests: GraphBatchRequest[], scopes: string[]): Promise<GraphBatchResponse[]>;
  /** Follows an absolute `@odata.nextLink` or `@odata.deltaLink` URL. */
  follow<T = unknown>(link: string, scopes: string[]): Promise<GraphResponse<T>>;
}

/** A Microsoft Graph error surfaced to the caller. */
export class GraphError extends Error {
  override readonly name = 'GraphError';
  readonly status: number;
  readonly code: string;
  readonly requestId: string | undefined;
  readonly path: string;
  readonly hint: string | undefined;

  constructor(init: {
    message: string;
    status: number;
    code: string;
    path: string;
    requestId?: string | undefined;
    hint?: string | undefined;
  }) {
    super(init.message);
    this.status = init.status;
    this.code = init.code;
    this.path = init.path;
    this.requestId = init.requestId;
    this.hint = init.hint;
  }
}

// ---------------------------------------------------------------------------
// Tool layer
// ---------------------------------------------------------------------------

/** Metadata describing a logical group of tools and what it needs consent for. */
export interface ToolGroupMeta {
  /** Stable identifier used by `--groups` and `--preset`, e.g. `mail`. */
  name: string;
  title: string;
  description: string;
  /** Delegated scopes the group's read tools need. */
  readScopes: string[];
  /** Extra delegated scopes the group's write tools need. */
  writeScopes: string[];
  /**
   * True when any scope in this group requires tenant admin consent.
   * Groups flagged here are excluded unless `--org-mode` is set.
   */
  requiresAdminConsent: boolean;
}

/** A single callable tool, independent of how it gets registered. */
export interface ToolDefinition {
  /** Wire name, e.g. `mail_list_messages`. Lowercase with underscores. */
  name: string;
  title: string;
  /** Shown to the model. State what it returns and any important limits. */
  description: string;
  /** Must be a `z.object({...})`. */
  inputSchema: z.ZodType;
  /** True when the tool mutates data. Suppressed under `--read-only`. */
  write?: boolean;
  /**
   * Opt-in identity attribution: when true, the registration layer prepends a
   * one-line rendering of the login-time cached identity to this tool's
   * serialized output. Defaults to off; an absent snapshot means no line,
   * never an error. Tools whose excerpts travel without their query context
   * opt in so a reader can tell which session produced them without a Graph
   * call. Nothing else about the output shape changes.
   */
  includeIdentity?: boolean;
  /** Delegated scopes this specific tool needs. */
  scopes: string[];
  /** The group this tool belongs to, for discovery and filtering. */
  group: string;
  /**
   * Runs the tool. Return any JSON-serialisable value; the registration layer
   * serialises and truncates it. Throw `GraphError` for Graph failures.
   */
  handler: (args: Record<string, unknown>) => Promise<unknown>;
}

/** Everything a tool module needs to build its definitions. */
export interface ToolDeps {
  graph: GraphClient;
  config: ServerConfig;
  /**
   * Required, not optional. The sign-in tools live in the generic module and are
   * the only way to authenticate from inside a conversation; if this were
   * optional, forgetting to pass it would remove them silently and the install
   * would look fine until someone tried to sign in. That is the bug this field
   * exists to have fixed, so the type refuses to let it come back. The two CLI
   * paths that build the catalogue only to read its scopes pass a provider that
   * throws if anything actually calls it.
   */
  auth: AuthProvider;
  /**
   * Every scope this configuration may ask consent for, including the extras
   * individual tools declare beyond their group's meta.
   *
   * A function rather than a value because the catalogue is what it is computed
   * from, and the catalogue is what this object is being passed to build. The
   * server resolves the cycle by letting the closure read a list it fills in
   * immediately afterwards; by the time anything calls this, it is complete.
   *
   * Required for the same reason `auth` is: `session_status` reports this as
   * the consent checklist, and a default that quietly reported less left five
   * tools permanently returning 403 with nothing naming the missing scopes. An
   * ordinary sign-in asks for less than this — the minimal set plus the scopes
   * the caller names — and consent grows from there.
   */
  consentScopes: () => string[];
}

/** One domain area of Microsoft 365, e.g. mail or calendar. */
export interface ToolModule {
  group: ToolGroupMeta;
  build(deps: ToolDeps): ToolDefinition[];
}

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

/**
 * Serialisation used for tool result text.
 *
 * - `json` — indented JSON. The default.
 * - `toon` — compact TOON-style tables: headers once plus bare rows.
 */
export type OutputFormat = 'json' | 'toon';

/** Fully resolved runtime configuration. */
export interface ServerConfig {
  /** Entra application (client) ID. */
  clientId: string;
  /** Tenant ID, or `common` / `organizations` / `consumers`. */
  tenantId: string;
  /** Full authority URL, derived from tenantId and the cloud. */
  authority: string;
  /**
   * Graph HOSTNAME for the selected cloud, e.g. `graph.microsoft.com`.
   *
   * Bare hostname, no scheme and no trailing slash: `assertSameGraphHost`
   * compares it against `URL.hostname`. Use `graphBaseUrl(config)` from
   * `config.ts` when you need `https://<host>`.
   */
  graphHost: string;
  /** Names of the enabled tool groups. */
  groups: string[];
  /**
   * Requested groups dropped for needing tenant admin consent while orgMode
   * is off. Optional so existing configs stay valid; `status` reports it so
   * the Groups line names what was skipped and how to enable it.
   */
  droppedGroups?: string[];
  /** Suppress every write tool. */
  readOnly: boolean;
  /** Endpoint used when a tool does not specify one. */
  graphVersion: GraphVersion;
  /** Allow tools and the generic caller to target the beta endpoint. */
  allowBeta: boolean;
  /** Expose `discover_tools`/`call_tool` instead of every domain tool. */
  discovery: boolean;
  /** Enable groups that require tenant admin consent. */
  orgMode: boolean;
  /** Hard cap on serialised tool output, in characters. */
  maxOutputChars: number;
  /** Serialisation used for tool result text. */
  output: OutputFormat;
  /** Log Graph requests to stderr. */
  verbose: boolean;
  /** Directory holding the encrypted token cache. */
  cacheDir: string;
  /** Which interactive sign-in flow `login` uses. */
  authFlow: AuthFlow;
  /**
   * Fixed loopback port for the browser flow's redirect URI, or 0 to let the OS
   * pick. Set this when the tenant registers an exact `http://localhost:PORT`
   * redirect instead of the bare `http://localhost` wildcard.
   */
  authPort: number;
  /** Expose the generic `graph_request` write escape hatch. */
  allowGenericWrite: boolean;
  /**
   * Explicit path to the MIP decryptor sidecar (`--mip-decryptor`). Absent
   * when neither the flag nor its environment fallback was given; the
   * launcher then still tries its bundled and PATH guesses, so absence here
   * means "unconfigured", not "unavailable".
   */
  mipDecryptor?: string;
}
