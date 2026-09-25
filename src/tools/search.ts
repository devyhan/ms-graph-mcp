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
 * All three group read scopes travel with every search call. Microsoft Search
 * federates Exchange, OneDrive and SharePoint behind one endpoint and selects
 * the index from `entityTypes`, and the three scopes are consented together the
 * moment the group is enabled, so narrowing the token per entity type would buy
 * nothing but extra ways to fail.
 */
const BASE_SCOPES: string[] = [...GROUP.readScopes];

/**
 * Two entity types reach indexes the group meta does not cover: Graph refuses
 * `event` without Calendars.Read and `chatMessage` without Chat.Read. Both are
 * user-consentable, and both are added to a call only when the caller actually
 * asks for that entity type, so an ordinary mail search never widens the token.
 */
const EVENT_SCOPE = 'Calendars.Read';
const CHAT_MESSAGE_SCOPE = 'Chat.Read';

/** Declared on the definition: the widest set any single invocation can need. */
const SEARCH_SCOPES: string[] = [...BASE_SCOPES, EVENT_SCOPE, CHAT_MESSAGE_SCOPE];

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
    const scopes = others.includes('event') ? [...BASE_SCOPES, EVENT_SCOPE] : BASE_SCOPES;
    buckets.push({ entityTypes: others, scopes, supportsFields: true });
  }

  if (entityTypes.includes('chatMessage')) {
    buckets.push({
      entityTypes: ['chatMessage'],
      scopes: [...BASE_SCOPES, CHAT_MESSAGE_SCOPE],
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
          "Runs a relevance-ranked Microsoft Search query across Microsoft 365 and returns a flattened list of hits, each with its rank, a text summary and the resource's id, name or subject, web URL and last-modified time. Defaults to 10 hits per entity-type group over ['message','driveItem']; page with `from`. Traps: chatMessage cannot be combined with other entity types, so a mixed request is split into two calls and merged, and its hits ignore `fields`; searching driveItem, listItem or site needs Sites.Read.All, event needs Calendars.Read and chatMessage needs Chat.Read, so enable the matching tool group or the split for that type fails on its own while the rest still return; results are ranked, not sorted by date, and freshly changed items may not be indexed yet. This is read-only despite using POST.",
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
