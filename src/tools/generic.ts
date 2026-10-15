/**
 * The `generic` group: an escape hatch for Graph endpoints no domain tool
 * covers, plus two introspection tools that help a model use it correctly.
 *
 * This group is deliberately absent from `GROUPS`. It is never selectable with
 * `--groups`; the server appends it unconditionally, so its meta is built here.
 */

import { z } from 'zod';

import type {
  AccountInfo,
  GraphVersion,
  HttpMethod,
  ServerConfig,
  ToolDefinition,
  ToolDeps,
  ToolGroupMeta,
  ToolModule,
} from '../contracts.js';
import {
  ADMIN_CONSENT_SCOPES,
  GROUPS,
  excludeUpfrontBlocked,
  loginScopesFor,
  scopesForGroups,
} from './groups.js';
import { mapLoginFailure } from '../auth/provider.js';
import { normalizeGraphPath } from '../util/paths.js';
import { guessScopesForPath } from '../util/scopes.js';

/**
 * Its scopes are the floor, not the ceiling: `graph_request` runs with the
 * union of every enabled group's scopes (see `genericScopes`), so it reuses
 * consent the user has already granted rather than asking for more.
 */
const GROUP: ToolGroupMeta = {
  name: 'generic',
  title: 'Generic Graph access',
  description: 'Direct Microsoft Graph calls and permission introspection.',
  readScopes: ['User.Read'],
  writeScopes: [],
  requiresAdminConsent: false,
};

const WRITE_METHODS: HttpMethod[] = ['POST', 'PATCH', 'PUT', 'DELETE'];

/** Above this the model is better served by a `$filter`; five pages is already ~5000 rows. */
const MAX_PAGES = 5;

/**
 * The scope set `graph_request` and `graph_schema` run with.
 *
 * `offline_access` is dropped: it is a token-lifetime scope MSAL adds for
 * itself, and it means nothing to a Graph resource. Write scopes are excluded
 * under `--read-only` because the tool refuses write methods there anyway, and
 * asking for consent this server will never use is not least privilege.
 */
function genericScopes(config: ServerConfig): string[] {
  const scopes = scopesForGroups(config.groups, config.readOnly);
  return scopes.filter((scope) => scope.toLowerCase() !== 'offline_access');
}

const requestInput = z.object({
  path: z
    .string()
    .describe(
      'Graph-relative path beginning with "/", e.g. "/me/messages" or ' +
        '"/users/alice@contoso.com/drive/root/children". Never a full URL: the host is fixed ' +
        'and an absolute URL is rejected. A "?query" suffix is accepted and merged with the ' +
        '`query` argument, but passing options in `query` is clearer.',
    ),
  method: z
    .enum(['GET', 'POST', 'PATCH', 'PUT', 'DELETE'])
    .optional()
    .describe(
      'HTTP method. Defaults to GET. The write methods are refused unless the server was ' +
        'started with --allow-generic-write, and always under --read-only.',
    ),
  version: z
    .enum(['v1.0', 'beta'])
    .optional()
    .describe(
      'Graph endpoint. Defaults to the server default (v1.0). "beta" is refused unless the ' +
        'server was started with --beta; beta resources can change without notice.',
    ),
  query: z
    .record(z.string(), z.string())
    .optional()
    .describe(
      'Query string parameters, e.g. {"$select":"id,subject","$top":"10"}. Values are sent ' +
        'verbatim, so OData operators belong here rather than spliced into `path`.',
    ),
  body: z
    .unknown()
    .optional()
    .describe(
      'Request body for a write method, as JSON. Ignored for GET. A string is sent as raw ' +
        'bytes (which is what Graph content endpoints expect); anything else is JSON-encoded.',
    ),
  maxPages: z
    .number()
    .int()
    .min(1)
    .max(MAX_PAGES)
    .optional()
    .describe(
      `How many pages of a collection to follow via @odata.nextLink, 1-${MAX_PAGES}. Defaults ` +
        'to 1 (no paging). Pages are merged into a single value array.',
    ),
  scopes: z
    .array(z.string().min(1))
    .min(1)
    .optional()
    .describe(
      'Scopes to request a token for. Defaults to the minimal scope guessed from `path` ' +
        '(Mail.Read for a mail path, Chat.Read for a chat path, and so on), falling back to ' +
        'the union of every enabled group when the path matches nothing known. Pass an ' +
        'explicit list only when the endpoint needs a scope the guess does not cover.',
    ),
});

const schemaInput = z.object({
  entityPath: z
    .string()
    .describe(
      'Graph-relative path to a collection or a single entity, e.g. "/me/messages", ' +
        '"/me/drive/root/children" or "/me". Must begin with "/".',
    ),
  version: z
    .enum(['v1.0', 'beta'])
    .optional()
    .describe('Graph endpoint to probe. Defaults to the server default (v1.0).'),
});

const permissionsInput = z.object({});

const beginLoginInput = z.object({
  scopes: z
    .array(z.string().min(1))
    .min(1)
    .optional()
    .describe(
      'Delegated scopes to consent to on top of the minimal sign-in set, e.g. ["Mail.Read"]. ' +
        'Omit it for the minimal set alone. When a tool fails for missing consent, pass the ' +
        'scopes its error names: Entra prompts only for what was added.',
    ),
  mode: z
    .enum(['bare', 'full'])
    .optional()
    .describe(
      'How much to consent to. Omit it (or pass "bare") for the minimal sign-in set plus ' +
        '`scopes`. Pass "full" to consent up front to everything this configuration may need ' +
        'except what an ordinary user cannot grant on their own: admin-consent and ' +
        'tenant-policy-blocked scopes are left out rather than failing the whole prompt.',
    ),
});

/** Splits an optional inline query string off a caller-supplied path. */
function splitPath(raw: string): { path: string; inline: Record<string, string> } {
  const mark = raw.indexOf('?');
  const bare = mark === -1 ? raw : raw.slice(0, mark);
  const inline: Record<string, string> = {};
  if (mark !== -1) {
    for (const [key, value] of new URLSearchParams(raw.slice(mark + 1))) {
      inline[key] = value;
    }
  }
  // SECURITY: the only path that reaches the client. `resolveGraphUrl` checks
  // it again on the way out, and the client pins the host; all three stay.
  return { path: normalizeGraphPath(bare), inline };
}

function resolveVersion(
  requested: GraphVersion | undefined,
  config: ServerConfig,
): GraphVersion {
  const version = requested ?? config.graphVersion;
  if (version === 'beta' && !config.allowBeta) {
    throw new Error(
      'The Microsoft Graph beta endpoint is disabled. Restart the server with --beta ' +
        '(or MS365_MCP_ALLOW_BETA=1) to allow it, or call the v1.0 equivalent instead.',
    );
  }
  return version;
}

function assertMethodAllowed(method: HttpMethod, config: ServerConfig): void {
  if (!WRITE_METHODS.includes(method)) return;

  if (config.readOnly) {
    throw new Error(
      `${method} is refused: the server is running in --read-only mode, which suppresses every ` +
        'write. Restart it without --read-only to allow writes.',
    );
  }
  throw new Error(
    `${method} is refused: the generic caller is read-only unless the server is started with ` +
      '--allow-generic-write (or MS365_MCP_ALLOW_GENERIC_WRITE=1). Use a purpose-built tool ' +
      'for this change if one exists, or ask the operator to enable the flag.',
  );
}

/** The JSON type of a value, at the granularity a model can act on. */
function inferType(value: unknown): string {
  if (value === null) return 'null';
  if (Array.isArray(value)) {
    const first = value[0];
    return value.length === 0 ? 'array (empty)' : `array of ${inferType(first)}`;
  }
  switch (typeof value) {
    case 'string':
      return 'string';
    case 'number':
      return Number.isInteger(value) ? 'integer' : 'number';
    case 'boolean':
      return 'boolean';
    case 'object':
      return 'object';
    default:
      return typeof value;
  }
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

function describeProperties(entity: Record<string, unknown>): Array<{
  name: string;
  type: string;
}> {
  return Object.keys(entity)
    .filter((key) => !key.startsWith('@'))
    .sort()
    .map((name) => ({ name, type: inferType(entity[name]) }));
}

export const genericModule: ToolModule = {
  group: GROUP,
  build({ graph, config, auth, consentScopes }: ToolDeps): ToolDefinition[] {
    const scopes = genericScopes(config);
    const canWrite = config.allowGenericWrite && !config.readOnly;

    // Sign-in, in the two tools it has to be.
    //
    // There was no way to authenticate from inside a conversation until now:
    // `login` is a CLI subcommand, so anyone who installed this as a plugin had
    // to find a terminal before the server could do anything at all. An explicit
    // tool call can safely launch the browser because the OAuth callback travels
    // over a loopback HTTP listener, while MCP continues to own stdout.
    //
    // Two tools rather than one because a single blocking call would sit for the
    // minutes a person takes to finish and hit the client's timeout first.
    type BrowserLogin = {
      state: 'pending' | 'signedIn' | 'failed';
      startedAt: string;
      /** The scopes the in-flight attempt asks for, so a reuse reports them, not the new ask. */
      scopes: string[];
      account?: string;
      detail?: string;
    };
    let browserLogin: BrowserLogin | null = null;

    function startedResponse(entry: BrowserLogin): Record<string, unknown> {
      return {
        state: entry.state,
        flow: 'browser',
        startedAt: entry.startedAt,
        nextStep:
          'The system browser has been asked to open Microsoft 365 sign-in on this machine. ' +
          'Tell the user to finish signing in there, then call auth_login_status. Do not ask ' +
          'for credentials or an authorization code in the conversation.',
        scopesRequested: entry.scopes,
      };
    }

    /**
     * The `mode: 'full'` ask: the full consent set from the server closure —
     * already read-only aware and restricted to the enabled groups — minus
     * what an ordinary user cannot grant, plus anything the caller named.
     * The same set `upfrontScopesFor` computes from the catalogue.
     */
    function fullLoginScopes(extra?: readonly string[]): string[] {
      const scopes = new Set(excludeUpfrontBlocked(consentScopes()));
      for (const raw of extra ?? []) {
        const scope = raw.trim();
        if (scope !== '') scopes.add(scope);
      }
      return [...scopes].sort();
    }

    const authTools: ToolDefinition[] = [
      {
        name: 'auth_begin_login',
        title: 'Start signing in to Microsoft 365',
        // Not a write. --read-only exists to stop this server changing anything
        // in the tenant, and signing in changes nothing there; a read-only
        // install that could never authenticate would just be broken.
        description:
          'Starts OAuth 2.0 browser sign-in to Microsoft 365 using authorization code + PKCE. ' +
          'The system browser opens on this machine and the tool returns immediately while the ' +
          'user signs in there; call auth_login_status afterward to see whether they finished. ' +
          'Calling it again while a sign-in is pending reuses the same attempt instead of opening ' +
          'another browser window. Use it when a tool reports that there is no session, and ' +
          'again with `scopes` when a tool reports missing consent: sign-in is incremental, so ' +
          'a bare call asks only for the minimal set and each repeat consents to just what was ' +
          'added. Pass `mode: "full"` to consent up front to the rest instead, minus what an ' +
          'ordinary user cannot grant. Never ask the user to paste credentials or an ' +
          'authorization code into the conversation. No token is ever returned.',
        inputSchema: beginLoginInput,
        scopes: [],
        group: GROUP.name,
        handler: async (args) => {
          if (browserLogin?.state === 'pending') return startedResponse(browserLogin);

          // The minimal set plus whatever the caller named — the same rule the
          // CLI `login` applies. Consent grows one sign-in at a time instead of
          // demanding the whole configuration up front, unless the caller asked
          // for `mode: 'full'`: the full consent set minus what an ordinary
          // user cannot grant, so one prompt covers the rest without putting an
          // ungrantable scope in charge of the entire sign-in.
          const parsed = beginLoginInput.parse(args);
          const wanted =
            parsed.mode === 'full'
              ? fullLoginScopes(parsed.scopes)
              : loginScopesFor(parsed.scopes);
          const entry: BrowserLogin = {
            state: 'pending',
            startedAt: new Date().toISOString(),
            scopes: wanted,
          };
          browserLogin = entry;

          // Intentionally not awaited: the tool must return before the user has
          // spent minutes in the browser. MSAL owns the one-shot loopback
          // listener and this continuation only records a token-free status.
          void auth.login(wanted, 'browser').then(
            (account) => {
              entry.state = 'signedIn';
              entry.account = account.username;
            },
            (error: unknown) => {
              entry.state = 'failed';
              // Mapped, not raw: a bare AADSTS code in a status field says
              // nothing about which knob to turn. Unknown stays verbatim.
              entry.detail = mapLoginFailure(entry.scopes, error, config, config.authPort);
            },
          );

          return startedResponse(entry);
        },
      },
      {
        name: 'auth_login_status',
        title: 'Check the sign-in',
        description:
          'Reports where the browser sign-in started by auth_begin_login has got to: "pending" ' +
          '(the user has not finished), "signedIn" (with the account), "failed" (with the ' +
          'reason), or "none" (no sign-in was started). When no browser sign-in was started ' +
          'in this session it falls back to the cached account, so an existing CLI session ' +
          'reports "signedIn" rather than "none". Never returns a token. Poll it rather ' +
          'than waiting: this returns immediately either way.',
        inputSchema: z.object({}),
        scopes: [],
        group: GROUP.name,
        handler: async () => {
          // The closure only knows about a sign-in started in this session.
          // Someone who signed in from the CLI (or an earlier session) has a
          // cached account and no closure entry; reporting "none" would send
          // them to sign in again for nothing. The MSAL cache is the
          // authority, so ask it — and a cache read that fails leaves the
          // closure state (or "none") standing rather than failing a probe.
          if (browserLogin !== null) return browserLogin;
          let cached: AccountInfo | null = null;
          try {
            cached = await auth.getAccount();
          } catch {
            cached = null;
          }
          if (cached === null) return { state: 'none' as const };
          return { state: 'signedIn' as const, account: cached.username, source: 'cache' as const };
        },
      },
    ];

    return [
      ...authTools,
      {
        name: 'session_status',
        title: 'Check the session and consent',
        description:
          'Unified session and consent probe. Returns cached account (Entra ID, tenant, ' +
          'consented scopes) plus whether the stored session still mints tokens and live ' +
          'GET /me reachability. Call once first; on consent errors sign in again with ' +
          'just the missing scopes.',
        inputSchema: z.object({}),
        scopes: [...GROUP.readScopes],
        group: GROUP.name,
        handler: async () => {
          // A probe that fails to probe is useless: each branch is caught so a
          // dead cache read still reports live reachability and vice versa. The
          // same reason graph_list_permissions never lets /me hide the map.
          let cachedAccount: AccountInfo | null = null;
          let cacheNote: string | undefined;
          try {
            cachedAccount = await auth.getAccount();
          } catch (error) {
            cacheNote =
              'Could not read the cached account: ' +
              (error instanceof Error ? error.message : String(error));
          }
          // Whether the cached session still mints tokens — the honest half of
          // "signed in". A cached account whose refresh fails is exactly the
          // state that used to report healthy while every call failed.
          // Absent on scope-only catalogues, which must not call getToken here:
          // that provider refuses, and a probe must not fail while probing.
          let tokenUsable: boolean | null = null;
          let sessionNote: string | undefined;
          if (typeof auth.sessionHealth === 'function') {
            try {
              const health = await auth.sessionHealth();
              tokenUsable = health.tokenUsable;
              sessionNote = health.note;
            } catch (error) {
              tokenUsable = false;
              sessionNote =
                'Could not probe the session: ' +
                (error instanceof Error ? error.message : String(error));
            }
          } else {
            sessionNote =
              'Token refresh was not probed: this catalogue was built to read scopes, not to use a session.';
          }
          return {
            cachedAccount,
            ...(cacheNote === undefined ? {} : { cacheNote }),
            tokenUsable,
            ...(sessionNote === undefined ? {} : { sessionNote }),
            liveAccount: await describeSignedInAccount(graph),
            browserLogin: browserLogin ?? { state: 'none' as const },
            scopesRequested: consentScopes(),
            note:
              'scopesRequested is the full set this configuration may ask consent for, not ' +
              'proof it was granted; consent is recorded per application in Microsoft Entra ' +
              'ID. Sign-in is incremental: a bare login asks only for the minimal set, and ' +
              'when a call fails for missing consent, signing in again with just those ' +
              'scopes prompts only for what was added.',
          };
        },
      },
      {
        name: 'graph_request',
        title: 'Call Microsoft Graph directly',
        description:
          'Sends an arbitrary request to Microsoft Graph and returns the raw JSON response. Use ' +
          'this only when no purpose-built tool covers the endpoint: the dedicated tools project ' +
          'and trim their results, while this one returns everything Graph sends, which is often ' +
          'far more text. The host is fixed to this tenant\'s Graph endpoint and `path` must be ' +
          'relative, so an absolute URL is rejected. GET is always available; the write methods ' +
          'require --allow-generic-write. Pass `maxPages` to follow @odata.nextLink, and prefer ' +
          '$select and $top in `query` to keep the response small. The call authenticates with ' +
          'the minimal scope guessed from `path`, so a path whose scope is granted succeeds even ' +
          'when an unrelated scope is missing; pass `scopes` to override the guess.',
        inputSchema: requestInput,
        // A read-only server still gets the GET hatch; only the write capability
        // is what makes this tool a write tool.
        write: canWrite,
        scopes,
        group: GROUP.name,
        handler: async (args) => {
          const parsed = requestInput.parse(args);
          const method: HttpMethod = parsed.method ?? 'GET';
          assertMethodAllowed(method, config);
          const version = resolveVersion(parsed.version, config);
          const { path, inline } = splitPath(parsed.path);

          // An explicit `query` entry wins: the caller wrote it as a structured
          // argument, whereas the inline form is a convenience.
          const query: Record<string, string> = { ...inline, ...(parsed.query ?? {}) };

          // The declared scopes stay the union — that is what consent is asked
          // for — but the token request narrows to what this path needs. A
          // union that includes an unconsented scope (an admin-consent one the
          // user could never get, or a group widened since the last login)
          // would otherwise fail a call its own scope covers. Narrowing is
          // free where everything is consented: a silent token request for a
          // subset of what was granted succeeds exactly as before.
          const callScopes = parsed.scopes ?? guessScopesForPath(path) ?? scopes;

          const res = await graph.request({
            path,
            method,
            version,
            query,
            ...(method === 'GET' ? {} : { body: parsed.body }),
            scopes: callScopes,
            maxPages: parsed.maxPages ?? 1,
          });

          return {
            status: res.status,
            path,
            version,
            method,
            nextLink: res.nextLink,
            deltaLink: res.deltaLink,
            data: res.data,
          };
        },
      },
      {
        name: 'graph_schema',
        title: 'Inspect a Graph resource shape',
        description:
          'Fetches one item from a Graph path ($top=1) and reports the property names it carries ' +
          'with their JSON types, plus the @odata.context that names the resource type. Use it ' +
          'before graph_request to learn what $select and $filter can reference. It reports the ' +
          'shape of one real item, so properties that happen to be null on that item are absent ' +
          'from the list, and an empty collection yields no properties at all. Read-only.',
        inputSchema: schemaInput,
        scopes,
        group: GROUP.name,
        handler: async (args) => {
          const parsed = schemaInput.parse(args);
          const version = resolveVersion(parsed.version, config);
          const { path, inline } = splitPath(parsed.entityPath);

          const res = await graph.request({
            path,
            method: 'GET',
            version,
            query: { ...inline, $top: '1' },
            scopes,
          });

          const body = asRecord(res.data);
          const context = body?.['@odata.context'];
          const odataContext = typeof context === 'string' ? context : undefined;
          const collection = body?.['value'];

          if (Array.isArray(collection)) {
            const first = asRecord(collection[0]);
            if (first === undefined) {
              return {
                path,
                version,
                kind: 'collection' as const,
                odataContext,
                properties: [],
                note:
                  `The collection at ${path} is empty, so there is no item to infer a shape ` +
                  'from. Try a path that holds data, or consult the Microsoft Graph reference ' +
                  'for this resource type.',
              };
            }
            return {
              path,
              version,
              kind: 'collection' as const,
              odataContext,
              properties: describeProperties(first),
              note: 'Inferred from the first item; properties null on that item are omitted.',
            };
          }

          if (body === undefined) {
            return {
              path,
              version,
              kind: 'unknown' as const,
              odataContext,
              properties: [],
              note: `${path} did not return a JSON object, so no property list could be inferred.`,
            };
          }

          return {
            path,
            version,
            kind: 'entity' as const,
            odataContext,
            properties: describeProperties(body),
            note: 'Properties null on this entity are omitted.',
          };
        },
      },
      {
        name: 'graph_list_permissions',
        title: 'List granted permissions',
        description:
          'Reports how this server is configured: which tool groups are enabled, the delegated ' +
          'Microsoft Graph scopes each one uses, which of them need a tenant administrator to ' +
          'consent, whether write tools are suppressed, and the signed-in account. Call it when ' +
          'a request fails with 403 or "insufficient privileges" to see whether the needed area ' +
          'is switched on at all. Reads only the signed-in account from Graph.',
        inputSchema: permissionsInput,
        scopes: [...GROUP.readScopes],
        group: GROUP.name,
        handler: async () => {
          const groups = config.groups.map((name) => {
            const meta = GROUPS[name];
            if (meta === undefined) {
              return { name, title: name, unknown: true as const };
            }
            return {
              name: meta.name,
              title: meta.title,
              description: meta.description,
              readScopes: meta.readScopes,
              // Under --read-only the write scopes are neither requested nor
              // usable, so reporting them would misdescribe the session.
              writeScopes: config.readOnly ? [] : meta.writeScopes,
              requiresAdminConsent: meta.requiresAdminConsent,
            };
          });

          // Per scope, not only per group. A group can be user-consentable
          // overall and still carry one scope an administrator must grant, and
          // the group flag alone sends the user to argue for the wrong thing.
          const requested = genericScopes(config);
          const needsAdmin = requested.filter((scope) => ADMIN_CONSENT_SCOPES.has(scope));

          const account = await describeSignedInAccount(graph);

          return {
            account,
            groups,
            effectiveScopes: genericScopes(config),
            readOnly: config.readOnly,
            orgMode: config.orgMode,
            discovery: config.discovery,
            graphVersion: config.graphVersion,
            allowBeta: config.allowBeta,
            allowGenericWrite: config.allowGenericWrite,
            adminConsentRequired: config.groups.filter(
              (name) => GROUPS[name]?.requiresAdminConsent === true,
            ),
            adminConsentScopes: needsAdmin,
            // The paragraph to hand the person who can actually unblock this.
            // Without it the model tends to suggest signing in again, which for
            // these scopes changes nothing however often it is repeated.
            howToGetAdminConsent:
              needsAdmin.length === 0
                ? undefined
                : {
                    summary:
                      `${needsAdmin.length} requested scope(s) sit in tool groups that need Microsoft 365 ` +
                      'tenant administrator consent. A few are user-consentable on their own, but they ' +
                      'arrive in one consent prompt with scopes that are not, and a prompt is approved ' +
                      'or refused as a whole — so signing in again will not obtain them.',
                    askAnAdministratorFor: needsAdmin,
                    application: config.clientId === '' ? undefined : config.clientId,
                    tenant: config.tenantId,
                    where:
                      'Entra admin center → Identity → Applications → App registrations → this ' +
                      'application → API permissions → Grant admin consent.',
                    thenRun:
                      'After consent is granted, run the server\'s `login` command again with ' +
                      '`--org-mode` and the owning tool group enabled.',
                  },
            note:
              'A scope listed here is what this server asks for, not proof it was granted. ' +
              'Consent is recorded per application in Microsoft Entra ID.',
          };
        },
      },
    ];
  },
};

interface SignedInAccount {
  signedIn: boolean;
  displayName?: string;
  userPrincipalName?: string;
  id?: string;
  note?: string;
}

/**
 * Failing to read `/me` must not hide the permission map — that map is exactly
 * what a caller needs when the sign-in is the thing that is broken.
 */
async function describeSignedInAccount(
  graph: ToolDeps['graph'],
): Promise<SignedInAccount> {
  try {
    const res = await graph.request<Record<string, unknown>>({
      path: '/me',
      method: 'GET',
      query: { $select: 'id,displayName,userPrincipalName' },
      scopes: ['User.Read'],
    });
    const me = asRecord(res.data) ?? {};
    const account: SignedInAccount = { signedIn: true };
    const id = me['id'];
    const displayName = me['displayName'];
    const upn = me['userPrincipalName'];
    if (typeof id === 'string') account.id = id;
    if (typeof displayName === 'string') account.displayName = displayName;
    if (typeof upn === 'string') account.userPrincipalName = upn;
    return account;
  } catch (error) {
    return {
      signedIn: false,
      note: `Could not read the signed-in account: ${
        error instanceof Error ? error.message : String(error)
      }`,
    };
  }
}
