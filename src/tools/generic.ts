/**
 * The `generic` group: an escape hatch for Graph endpoints no domain tool
 * covers, plus two introspection tools that help a model use it correctly.
 *
 * This group is deliberately absent from `GROUPS`. It is never selectable with
 * `--groups`; the server appends it unconditionally, so its meta is built here.
 */

import { z } from 'zod';

import type {
  GraphVersion,
  HttpMethod,
  ServerConfig,
  ToolDefinition,
  ToolDeps,
  ToolGroupMeta,
  ToolModule,
} from '../contracts.js';
import { GROUPS, scopesForGroups } from './groups.js';
import { normalizeGraphPath } from '../util/paths.js';

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
  build({ graph, config }: ToolDeps): ToolDefinition[] {
    const scopes = genericScopes(config);
    const canWrite = config.allowGenericWrite && !config.readOnly;

    return [
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
          '$select and $top in `query` to keep the response small.',
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

          const res = await graph.request({
            path,
            method,
            version,
            query,
            ...(method === 'GET' ? {} : { body: parsed.body }),
            scopes,
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
