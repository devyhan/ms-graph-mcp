/**
 * Microsoft Search tools.
 *
 * `/search/query` is a POST that only reads: the KQL query string, its entity
 * types and the paging window do not fit in a URL, so Graph takes them as a
 * body. Nothing here mutates, hence no `write` flag on the definition.
 */

import { z } from 'zod';
import { GraphError } from '../contracts.js';
import type { ToolDefinition, ToolDeps, ToolGroupMeta, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { stripHtml, truncateText } from '../util/truncate.js';

// The group catalogue is a static literal; `search` is always present.
const GROUP: ToolGroupMeta = GROUPS['search']!;

/**
 * What each index actually needs, so a call asks for that and nothing else.
 *
 * This used to send all three group read scopes with every search, on the
 * reasoning that enabling the group consents them together anyway. That holds
 * only where the whole group was granted. `Sites.Read.All` needs an
 * administrator in most tenants, and a user who could not get one — the user
 * this server exists for — then found a chatMessage-only search failing on
 * AADSTS65001 for two scopes it never touches. Measured on one such mailbox:
 * the failure cost 110,000 characters of chat history fetched and scanned by
 * hand to answer a question the search would have answered in 900.
 *
 * Narrowing is free where everything is consented: a silent token request for a
 * subset of what was granted succeeds exactly as before.
 */
const ENTITY_SCOPES: Record<EntityType, readonly string[]> = {
  message: ['Mail.Read'],
  event: ['Calendars.Read'],
  driveItem: ['Files.Read', 'Sites.Read.All'],
  listItem: ['Sites.Read.All'],
  site: ['Sites.Read.All'],
  // Graph wants both, and says so itself on a 403: "Access to ChatMessage in
  // Graph API requires the following permissions: Chat.Read or Chat.ReadWrite,
  // ChannelMessage.Read.All." Asking for only the first produced a token the
  // search endpoint then refused, which reads as a permissions problem with no
  // permission left to fix.
  chatMessage: ['Chat.Read', 'ChannelMessage.Read.All'],
};

/** The union of what one bucket's entity types need, in a stable order. */
function scopesFor(entityTypes: readonly EntityType[]): string[] {
  const out: string[] = [];
  for (const type of entityTypes) {
    for (const scope of ENTITY_SCOPES[type]) if (!out.includes(scope)) out.push(scope);
  }
  return out;
}

/**
 * Declared on the definition, which is what a login asks consent for.
 *
 * `ChannelMessage.Read.All` is deliberately excluded even though a chatMessage
 * search needs it. It is an administrator-consent scope owned by the `teams`
 * group, and pulling it in here would put every personal install behind an
 * administrator it does not otherwise need — `search` is in the personal preset.
 * Searching chat therefore means enabling `teams` too, which is already the gate
 * for that permission.
 */
const ADMIN_ONLY_SCOPE = 'ChannelMessage.Read.All';

const SEARCH_SCOPES: string[] = [...new Set(Object.values(ENTITY_SCOPES).flat())].filter(
  (scope) => scope !== ADMIN_ONLY_SCOPE,
);

const SEARCH_PATH = '/search/query';

const ENTITY_TYPES = ['message', 'event', 'driveItem', 'listItem', 'site', 'chatMessage'] as const;
type EntityType = (typeof ENTITY_TYPES)[number];

const DEFAULT_ENTITY_TYPES: EntityType[] = ['message', 'driveItem'];

// ---------------------------------------------------------------------------
// Graph payload shapes (only the fields this tool projects)
// ---------------------------------------------------------------------------

interface SearchHit {
  hitId?: string;
  rank?: number;
  summary?: string;
  resource?: Record<string, unknown> | null;
}

interface HitsContainer {
  hits?: SearchHit[] | null;
  total?: number;
  moreResultsAvailable?: boolean;
}

interface SearchResponseEntry {
  hitsContainers?: HitsContainer[] | null;
}

interface SearchRequestEntry {
  entityTypes: EntityType[];
  query: { queryString: string };
  from: number;
  size: number;
  fields?: string[];
}

/** One split of the caller's request, with the scopes that split needs. */
interface Bucket {
  entityTypes: EntityType[];
  scopes: string[];
  /** chatMessage rejects `fields`, so the projection falls back to defaults. */
  supportsFields: boolean;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function compact<T extends Record<string, unknown>>(value: T): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value)) {
    if (v !== undefined && v !== null) out[key] = v;
  }
  return out;
}

function unique<T>(values: readonly T[]): T[] {
  return [...new Set(values)];
}

/**
 * Graph accepts exactly one searchRequest per call, and `chatMessage` cannot
 * share a request with any other entity type, so a mixed ask becomes two calls
 * whose hits are merged back together.
 */
function splitIntoBuckets(entityTypes: EntityType[]): Bucket[] {
  const others = entityTypes.filter((type) => type !== 'chatMessage');
  const buckets: Bucket[] = [];

  if (others.length > 0) {
    buckets.push({ entityTypes: others, scopes: scopesFor(others), supportsFields: true });
  }

  if (entityTypes.includes('chatMessage')) {
    buckets.push({
      entityTypes: ['chatMessage'],
      scopes: scopesFor(['chatMessage']),
      supportsFields: false,
    });
  }

  return buckets;
}

function readString(source: Record<string, unknown>, key: string): string | undefined {
  const value = source[key];
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

/** `#microsoft.graph.driveItem` -> `driveItem`. */
function readEntityType(resource: Record<string, unknown>): string | undefined {
  const odataType = readString(resource, '@odata.type');
  if (odataType === undefined) return undefined;
  const dot = odataType.lastIndexOf('.');
  return dot === -1 ? odataType.replace(/^#/, '') : odataType.slice(dot + 1);
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

/**
 * Mail models the sender as `from.emailAddress`, Teams as an identitySet with a
 * `user`, so one reader covers both rather than branching per entity type.
 */
function readSender(value: unknown): string | undefined {
  const from = asRecord(value);
  if (from === undefined) return undefined;

  const email = asRecord(from['emailAddress']);
  if (email !== undefined) {
    const address = readString(email, 'address');
    const name = readString(email, 'name');
    if (address !== undefined) return name === undefined ? address : `${name} <${address}>`;
    return name;
  }

  const user = asRecord(from['user']);
  if (user !== undefined) return readString(user, 'displayName') ?? readString(user, 'id');

  return undefined;
}

/** Outlook events carry `start` as a dateTimeTimeZone complex type. */
function readDateTimeTimeZone(value: unknown): string | undefined {
  const slot = asRecord(value);
  if (slot === undefined) return undefined;
  const dateTime = readString(slot, 'dateTime');
  if (dateTime === undefined) return undefined;
  const timeZone = readString(slot, 'timeZone');
  return timeZone === undefined ? dateTime : `${dateTime} ${timeZone}`;
}

/**
 * Summaries come back with `<c0>` hit-highlight tags around the matched terms.
 * They are markup, not content, so they are stripped before the text reaches
 * the model.
 */
function readSummary(summary: string | undefined, limit: number): string | undefined {
  if (typeof summary !== 'string' || summary.length === 0) return undefined;
  const text = stripHtml(summary).trim();
  return text === '' ? undefined : truncateText(text, limit);
}

function projectHit(
  hit: SearchHit,
  fields: string[] | undefined,
  summaryChars: number,
): Record<string, unknown> {
  const resource = asRecord(hit.resource) ?? {};

  const projected = compact({
    rank: hit.rank,
    type: readEntityType(resource),
    id: readString(resource, 'id') ?? (typeof hit.hitId === 'string' ? hit.hitId : undefined),
    name:
      readString(resource, 'subject') ??
      readString(resource, 'name') ??
      readString(resource, 'displayName') ??
      readString(resource, 'title'),
    // Mail spells the browser link `webLink`; everything else uses `webUrl`.
    webUrl: readString(resource, 'webUrl') ?? readString(resource, 'webLink'),
    lastModifiedDateTime:
      readString(resource, 'lastModifiedDateTime') ??
      readString(resource, 'receivedDateTime') ??
      readString(resource, 'createdDateTime'),
    from: readSender(resource['from']),
    start: readDateTimeTimeZone(resource['start']),
    summary: readSummary(hit.summary, summaryChars),
  });

  // Anything the caller asked for by name is copied through verbatim, since the
  // projection above cannot know what a custom managed property means.
  if (fields !== undefined) {
    for (const field of fields) {
      if (projected[field] !== undefined) continue;
      const value = resource[field];
      if (value !== undefined && value !== null) projected[field] = value;
    }
  }

  return projected;
}

function describeError(error: unknown): Record<string, unknown> {
  if (error instanceof GraphError) {
    return compact({
      status: error.status,
      code: error.code,
      message: error.message,
      hint: error.hint,
    });
  }
  return { message: error instanceof Error ? error.message : String(error) };
}

// ---------------------------------------------------------------------------
// Input schema
// ---------------------------------------------------------------------------

const searchQuerySchema = z.object({
  query: z
    .string()
    .min(1)
    .describe(
      'The search terms, in Microsoft Search KQL. Free text works ("quarterly budget"), as do property restrictions such as "subject:budget", "from:sara@contoso.com", "filetype:pptx" and "lastModifiedTime>=2026-01-01", combined with AND/OR/NOT.',
    ),
  entityTypes: z
    .array(z.enum(ENTITY_TYPES))
    .min(1)
    .max(ENTITY_TYPES.length)
    .default(DEFAULT_ENTITY_TYPES)
    .describe(
      "Which kinds of content to search. Defaults to ['message','driveItem'] (Outlook mail plus OneDrive/SharePoint files).",
    ),
  from: z
    .number()
    .int()
    .min(0)
    .max(500)
    .default(0)
    .describe('Zero-based offset of the first hit to return; use it to page. Defaults to 0.'),
  size: z
    .number()
    .int()
    .min(1)
    .max(25)
    .default(10)
    .describe('Maximum hits to return per entity-type group. Defaults to 10, capped at 25.'),
  fields: z
    .array(z.string().regex(/^[A-Za-z][A-Za-z0-9_.]{0,63}$/))
    .min(1)
    .max(20)
    .optional()
    .describe(
      'Extra resource properties to retrieve and include on each hit, for example ["from","importance"] or a SharePoint managed property. Ignored for chatMessage, which does not support field selection.',
    ),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const searchModule: ToolModule = {
  group: GROUP,

  build({ graph, config }: ToolDeps): ToolDefinition[] {
    // Summaries are the only unbounded field and there can be one per hit, so
    // give them a slice of the output budget instead of a hard-coded guess.
    const summaryChars = Math.min(400, Math.max(120, Math.floor(config.maxOutputChars / 80)));

    return [
      {
        name: 'search_query',
        title: 'Search Microsoft 365',
        group: GROUP.name,
        scopes: SEARCH_SCOPES,
        description:
          "Runs a relevance-ranked Microsoft Search query across Microsoft 365 and returns a flattened list of hits, each with its rank, a text summary and the resource's id, name or subject, web URL and last-modified time. Defaults to 10 hits per entity-type group over ['message','driveItem']; page with `from`. Traps: chatMessage cannot be combined with other entity types, so a mixed request is split into two calls and merged, and its hits ignore `fields`; each split asks only for the scopes its own index needs — Sites.Read.All for driveItem, listItem and site, Mail.Read for message, Calendars.Read for event, and BOTH Chat.Read and ChannelMessage.Read.All for chatMessage. That last pair is the one to watch: ChannelMessage.Read.All needs tenant admin consent and is carried by the `teams` group, so a chat search needs `teams` enabled as well as `chat`, and without it Graph refuses the chatMessage split with a 403 naming both permissions while every other split still returns; results are ranked, not sorted by date, and freshly changed items may not be indexed yet. This is read-only despite using POST.",
        inputSchema: searchQuerySchema,
        handler: async (args) => {
          const { query, entityTypes, from, size, fields } = searchQuerySchema.parse(args);
          const requested = unique(entityTypes);
          const buckets = splitIntoBuckets(requested);

          const hits: Record<string, unknown>[] = [];
          const failures: Record<string, unknown>[] = [];
          let total = 0;
          let moreResultsAvailable = false;
          let firstError: unknown;

          for (const bucket of buckets) {
            const entry: SearchRequestEntry = {
              entityTypes: bucket.entityTypes,
              query: { queryString: query },
              from,
              size,
            };
            if (fields !== undefined && bucket.supportsFields) entry.fields = fields;

            let data: unknown;
            try {
              const res = await graph.request<unknown>({
                path: SEARCH_PATH,
                method: 'POST',
                body: { requests: [entry] },
                scopes: bucket.scopes,
              });
              data = res.data;
            } catch (error) {
              // A split is this tool's own doing, invisible to the caller, so a
              // half-failed search reports which half failed instead of losing
              // the hits that did come back. When nothing succeeds the original
              // error is rethrown below, status and code intact.
              if (buckets.length === 1) throw error;
              if (firstError === undefined) firstError = error;
              failures.push({ entityTypes: bucket.entityTypes, error: describeError(error) });
              continue;
            }

            for (const entryResponse of extractCollection<SearchResponseEntry>(data)) {
              for (const container of entryResponse.hitsContainers ?? []) {
                if (typeof container.total === 'number') total += container.total;
                if (container.moreResultsAvailable === true) moreResultsAvailable = true;
                for (const hit of container.hits ?? []) {
                  hits.push(projectHit(hit, bucket.supportsFields ? fields : undefined, summaryChars));
                }
              }
            }
          }

          if (failures.length === buckets.length && firstError !== undefined) throw firstError;

          // Hits stay in bucket order with the rank Graph assigned inside each
          // request; ranks are per-request, so a merged global ordering would be
          // invented rather than measured.
          return compact({
            query,
            entityTypes: requested,
            count: hits.length,
            total,
            from,
            size,
            // Search pages by offset, not by an @odata.nextLink cursor.
            nextFrom: moreResultsAvailable ? from + size : undefined,
            moreResultsAvailable,
            hits,
            partialFailures: failures.length > 0 ? failures : undefined,
          });
        },
      },
    ];
  },
};
