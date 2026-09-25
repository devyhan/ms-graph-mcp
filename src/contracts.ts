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
  /** Starts an interactive sign-in. Resolves once a session exists. */
  login(scopes: string[]): Promise<AccountInfo>;
  /** Clears the cached session. */
  logout(): Promise<void>;
}

/** Raised when a token cannot be acquired without user interaction. */
export class InteractionRequiredError extends Error {
  override readonly name = 'InteractionRequiredError';
  constructor(message: string) {
    super(message);
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
}

/** One domain area of Microsoft 365, e.g. mail or calendar. */
export interface ToolModule {
  group: ToolGroupMeta;
  build(deps: ToolDeps): ToolDefinition[];
}

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

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
}
