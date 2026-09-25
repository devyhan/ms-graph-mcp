/**
 * SharePoint tools: sites, document libraries, and lists.
 *
 * Graph addresses a site two incompatible ways and the difference leaks into
 * every child path. An opaque `hostname,siteGuid,webGuid` id composes with a
 * plain slash (`/sites/{id}/drives`), while the human-readable
 * `contoso.sharepoint.com:/sites/Team` form has to be terminated with a second
 * colon before any child segment (`/sites/contoso.sharepoint.com:/sites/Team:/drives`).
 * Getting that wrong yields a 400 that reads like a bad id, so `siteRef` is the
 * single place in this file that knows the rule.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolGroupMeta, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { stripHtml, truncateText } from '../util/truncate.js';

// The group catalogue is a static literal; `sharepoint` is always present.
const GROUP: ToolGroupMeta = GROUPS['sharepoint']!;
const READ_SCOPES: string[] = [...GROUP.readScopes];
const WRITE_SCOPES: string[] = [...GROUP.readScopes, ...GROUP.writeScopes];

// ---------------------------------------------------------------------------
// Graph payload shapes (only the fields these tools project)
// ---------------------------------------------------------------------------

interface IdentitySet {
  user?: { displayName?: string | null; email?: string | null } | null;
  application?: { displayName?: string | null } | null;
}

interface GraphSite {
  id?: string;
  name?: string | null;
  displayName?: string | null;
  description?: string | null;
  webUrl?: string | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  isPersonalSite?: boolean | null;
  root?: Record<string, unknown> | null;
  siteCollection?: { hostname?: string | null } | null;
}

interface GraphDrive {
  id?: string;
  name?: string | null;
  description?: string | null;
  driveType?: string | null;
  webUrl?: string | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  quota?: { used?: number | null; total?: number | null; remaining?: number | null } | null;
}

interface GraphDriveItem {
  id?: string;
  name?: string | null;
  size?: number | null;
  webUrl?: string | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  folder?: { childCount?: number | null } | null;
  file?: { mimeType?: string | null } | null;
  package?: { type?: string | null } | null;
  lastModifiedBy?: IdentitySet | null;
  parentReference?: { driveId?: string | null; id?: string | null; path?: string | null } | null;
}

interface GraphColumn {
  id?: string;
  name?: string | null;
  displayName?: string | null;
  description?: string | null;
  required?: boolean | null;
  readOnly?: boolean | null;
  hidden?: boolean | null;
  indexed?: boolean | null;
}

interface GraphList {
  id?: string;
  name?: string | null;
  displayName?: string | null;
  description?: string | null;
  webUrl?: string | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  list?: { template?: string | null; hidden?: boolean | null; contentTypesEnabled?: boolean | null } | null;
  system?: Record<string, unknown> | null;
  columns?: GraphColumn[] | null;
}

interface GraphListItem {
  id?: string;
  webUrl?: string | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  createdBy?: IdentitySet | null;
  lastModifiedBy?: IdentitySet | null;
  contentType?: { name?: string | null } | null;
  fields?: unknown;
}

// ---------------------------------------------------------------------------
// Site addressing
// ---------------------------------------------------------------------------

/**
 * `contoso.sharepoint.com` — the host half of the readable site form. At least
 * one dot is required so a pasted `https://contoso.sharepoint.com/sites/X`
 * fails loudly instead of being read as the host `https` with a path after it.
 */
const SITE_HOSTNAME = /^[A-Za-z0-9][A-Za-z0-9-]*(?:\.[A-Za-z0-9-]+)+$/;

/**
 * The opaque form: `root`, a bare GUID, or `hostname,siteGuid,webGuid`. Commas
 * are deliberately left unencoded — Graph's own ids carry them, and percent-
 * encoding a separator Graph put there itself invites a 400.
 */
const SITE_OPAQUE_ID = /^[A-Za-z0-9][A-Za-z0-9.,_-]{0,511}$/;

/** How a site is spliced into a Graph path, and how children hang off it. */
interface SiteRef {
  /** Path to the site resource itself. */
  self: string;
  /** Path to a child collection or item below the site. */
  child(suffix: string): string;
}

function siteRef(raw: string): SiteRef {
  const value = raw.trim();
  if (value.length === 0) {
    throw new Error('siteId must not be empty.');
  }

  const colon = value.indexOf(':');
  if (colon === -1) {
    if (!SITE_OPAQUE_ID.test(value)) {
      throw new Error(
        `Invalid siteId ${JSON.stringify(raw)}. Expected the composite id from sharepoint_search_sites ` +
          '(e.g. "contoso.sharepoint.com,8a2f…,1c7b…"), the literal "root", a hostname, or the ' +
          '"contoso.sharepoint.com:/sites/Marketing" form.',
      );
    }
    const base = `/sites/${value}`;
    return { self: base, child: (suffix) => `${base}/${suffix}` };
  }

  const host = value.slice(0, colon);
  if (!SITE_HOSTNAME.test(host)) {
    throw new Error(
      `Invalid SharePoint hostname ${JSON.stringify(host)} in siteId. Expected something like ` +
        '"contoso.sharepoint.com:/sites/Marketing" — a bare hostname and server-relative path, not a full URL.',
    );
  }

  // A caller who copied the id out of a Graph URL may include the terminating
  // colon; the trailing one is ours to add, so drop theirs.
  const rest = value.slice(colon + 1).replace(/:+$/, '');
  const segments = rest.split('/').filter((s) => s.length > 0);
  for (const segment of segments) {
    if (segment === '.' || segment === '..' || segment.includes(':')) {
      throw new Error(`Invalid server-relative site path segment ${JSON.stringify(segment)} in siteId.`);
    }
  }

  // `contoso.sharepoint.com:/` addresses the root site, which has no path half.
  if (segments.length === 0) {
    const base = `/sites/${host}`;
    return { self: base, child: (suffix) => `${base}/${suffix}` };
  }

  const base = `/sites/${host}:/${segments.map(encodeURIComponent).join('/')}`;
  return { self: base, child: (suffix) => `${base}:/${suffix}` };
}

/**
 * Drive, list and item ids are opaque and routinely carry `!`, `+`, `/` and `=`,
 * any of which would otherwise change the shape of the URL.
 */
function seg(id: string, label: string): string {
  const value = id.trim();
  if (value.length === 0) throw new Error(`${label} must not be empty.`);
  return encodeURIComponent(value);
}

/** A drive-relative folder path, encoded segment by segment for the `root:/…:/` form. */
function encodeItemPath(raw: string, label: string): string {
  const segments = raw.split('/').filter((s) => s.length > 0);
  if (segments.length === 0) throw new Error(`${label} must name at least one folder.`);
  for (const segment of segments) {
    if (segment === '.' || segment === '..') {
      throw new Error(`${label} must not contain a "." or ".." segment.`);
    }
  }
  return segments.map(encodeURIComponent).join('/');
}

// ---------------------------------------------------------------------------
// Column names
// ---------------------------------------------------------------------------

/**
 * SharePoint internal column names are letters, digits and underscores — spaces
 * and punctuation arrive already escaped as `_x0020_`. Validating here matters
 * because these names are spliced into `$expand=fields($select=…)` rather than
 * passed as a standalone query value.
 */
const COLUMN_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,127}$/;

function parseColumnList(csv: string | undefined, label: string): string[] | undefined {
  if (csv === undefined) return undefined;
  const names = csv
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  if (names.length === 0) return undefined;

  for (const name of names) {
    if (!COLUMN_NAME.test(name)) {
      throw new Error(
        `Invalid column name ${JSON.stringify(name)} in ${label}. Use internal names as returned by ` +
          'sharepoint_get_list with includeColumns (letters, digits and underscores only, e.g. "Due_x0020_Date").',
      );
    }
  }
  return names;
}

function fieldsExpand(columns: string[] | undefined): string {
  return columns === undefined ? 'fields' : `fields($select=${columns.join(',')})`;
}

// ---------------------------------------------------------------------------
// Projection helpers
// ---------------------------------------------------------------------------

function compact<T extends Record<string, unknown>>(value: T): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value)) {
    if (v !== undefined && v !== null) out[key] = v;
  }
  return out;
}

function identityName(v: IdentitySet | null | undefined): string | undefined {
  if (v === null || v === undefined) return undefined;
  const name = v.user?.displayName ?? v.application?.displayName;
  const email = v.user?.email;
  if (typeof name === 'string' && name.length > 0) {
    return typeof email === 'string' && email.length > 0 && email !== name ? `${name} <${email}>` : name;
  }
  return typeof email === 'string' && email.length > 0 ? email : undefined;
}

/**
 * Columns SharePoint adds for its own web UI. They carry no information a model
 * can act on and, on a wide list, crowd out the columns that do.
 */
const PRESENTATION_COLUMNS = new Set([
  'AppAuthor',
  'AppEditor',
  'DocIcon',
  'Edit',
  'FolderChildCount',
  'ItemChildCount',
  'LinkTitle',
  'LinkTitleNoMenu',
  '_ComplianceFlags',
  '_ComplianceTag',
  '_ComplianceTagUserId',
  '_ComplianceTagWrittenTime',
  '_UIVersionString',
]);

const HTML_TAG = /<[a-z][a-z0-9]*\b[^>]*>/i;

/**
 * Flattens a `fieldValueSet` into something bounded. Multi-line SharePoint
 * columns store rich text, so an unfiltered item can be mostly markup.
 */
function projectFields(raw: unknown, maxChars: number, explicitColumns: boolean): Record<string, unknown> | undefined {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return undefined;

  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    if (key.startsWith('@odata.')) continue;
    if (!explicitColumns && PRESENTATION_COLUMNS.has(key)) continue;
    if (value === null || value === undefined) continue;

    out[key] =
      typeof value === 'string'
        ? truncateText(HTML_TAG.test(value) ? stripHtml(value) : value, maxChars)
        : value;
  }

  return Object.keys(out).length > 0 ? out : undefined;
}

function projectSite(site: GraphSite): Record<string, unknown> {
  return compact({
    id: site.id,
    displayName: site.displayName ?? site.name,
    name: site.name,
    description: site.description,
    webUrl: site.webUrl,
    hostname: site.siteCollection?.hostname,
    isRootSite: site.root === null || site.root === undefined ? undefined : true,
    isPersonalSite: site.isPersonalSite === true ? true : undefined,
    createdDateTime: site.createdDateTime,
    lastModifiedDateTime: site.lastModifiedDateTime,
  });
}

function projectDrive(drive: GraphDrive): Record<string, unknown> {
  return compact({
    id: drive.id,
    name: drive.name,
    description: drive.description,
    driveType: drive.driveType,
    webUrl: drive.webUrl,
    usedBytes: drive.quota?.used,
    lastModifiedDateTime: drive.lastModifiedDateTime,
  });
}

function projectDriveItem(item: GraphDriveItem): Record<string, unknown> {
  const kind = item.folder !== null && item.folder !== undefined ? 'folder' : 'file';
  return compact({
    id: item.id,
    name: item.name,
    kind: item.package !== null && item.package !== undefined ? `package:${item.package.type ?? 'unknown'}` : kind,
    size: item.size,
    childCount: item.folder?.childCount,
    mimeType: item.file?.mimeType,
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    lastModifiedBy: identityName(item.lastModifiedBy),
  });
}

/**
 * Which facet a column carries is how SharePoint encodes its type; there is no
 * scalar `type` property to read.
 */
const COLUMN_TYPE_FACETS = [
  'text',
  'number',
  'currency',
  'dateTime',
  'boolean',
  'choice',
  'lookup',
  'personOrGroup',
  'hyperlinkOrPicture',
  'calculated',
  'geolocation',
  'term',
  'termSet',
  'thumbnail',
  'contentApprovalStatus',
] as const;

function projectColumn(column: GraphColumn): Record<string, unknown> {
  const raw = column as unknown as Record<string, unknown>;

  let type: string | undefined;
  let choices: unknown;
  for (const facet of COLUMN_TYPE_FACETS) {
    const value = raw[facet];
    if (value === undefined || value === null) continue;
    type = facet;
    if (facet === 'choice' && typeof value === 'object' && !Array.isArray(value)) {
      choices = (value as Record<string, unknown>)['choices'];
    }
    break;
  }

  return compact({
    name: column.name,
    displayName: column.displayName,
    type,
    choices,
    required: column.required === true ? true : undefined,
    readOnly: column.readOnly === true ? true : undefined,
    indexed: column.indexed === true ? true : undefined,
  });
}

function projectList(list: GraphList): Record<string, unknown> {
  return compact({
    id: list.id,
    displayName: list.displayName ?? list.name,
    name: list.name,
    description: list.description,
    template: list.list?.template,
    hidden: list.list?.hidden === true ? true : undefined,
    isSystemList: list.system === null || list.system === undefined ? undefined : true,
    webUrl: list.webUrl,
    createdDateTime: list.createdDateTime,
    lastModifiedDateTime: list.lastModifiedDateTime,
  });
}

function projectListItem(
  item: GraphListItem,
  maxChars: number,
  explicitColumns: boolean,
): Record<string, unknown> {
  return compact({
    id: item.id,
    contentType: item.contentType?.name,
    createdDateTime: item.createdDateTime,
    createdBy: identityName(item.createdBy),
    lastModifiedDateTime: item.lastModifiedDateTime,
    lastModifiedBy: identityName(item.lastModifiedBy),
    webUrl: item.webUrl,
    fields: projectFields(item.fields, maxChars, explicitColumns),
  });
}

/**
 * `$filter` and `$orderby` on list items only work against indexed columns.
 * This header downgrades SharePoint's hard refusal to a best-effort query,
 * which is the difference between "always 400" and "works on normal lists".
 */
const HONOR_NON_INDEXED: Record<string, string> = {
  Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly',
};

/** Keeps a driveItem list from carrying download URLs, cTags and hashes. */
const DRIVE_ITEM_SELECT =
  'id,name,size,webUrl,createdDateTime,lastModifiedDateTime,folder,file,package,lastModifiedBy';

// ---------------------------------------------------------------------------
// Schemas
// ---------------------------------------------------------------------------

const siteIdField = z
  .string()
  .min(1)
  .describe(
    'Site identifier: the composite id from sharepoint_search_sites (e.g. "contoso.sharepoint.com,8a2f…,1c7b…"), ' +
      'the literal "root" for the tenant root site, a bare hostname, or the readable "contoso.sharepoint.com:/sites/Marketing" form.',
  );

const listIdField = z
  .string()
  .min(1)
  .describe('List id (a GUID) or list name from sharepoint_list_lists.');

const searchSitesSchema = z.object({
  query: z
    .string()
    .min(1)
    .describe('Text matched against site titles and URLs. Pass "*" to list every site the user can reach.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(25)
    .describe('Maximum sites to return in one page. Defaults to 25.'),
});

const getSiteSchema = z.object({ siteId: siteIdField });

const listDrivesSchema = z.object({
  siteId: siteIdField,
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(25)
    .describe('Maximum document libraries to return. Defaults to 25.'),
});

const listDriveItemsSchema = z.object({
  siteId: siteIdField,
  driveId: z.string().min(1).describe('Document library (drive) id from sharepoint_list_drives.'),
  itemId: z
    .string()
    .min(1)
    .optional()
    .describe('Folder item id to list. Omit to list the library root. Mutually exclusive with folderPath.'),
  folderPath: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Library-relative folder path to list, e.g. "General/Reports". Mutually exclusive with itemId.',
    ),
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(50)
    .describe('Maximum children to return in one page. Defaults to 50.'),
});

const listListsSchema = z.object({
  siteId: siteIdField,
  includeHidden: z
    .boolean()
    .default(false)
    .describe('Include hidden system lists such as "Form Templates" and workflow history. Defaults to false.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(50)
    .describe('Maximum lists to return in one page. Defaults to 50.'),
});

const getListSchema = z.object({
  siteId: siteIdField,
  listId: listIdField,
  includeColumns: z
    .boolean()
    .default(false)
    .describe(
      'Also return every column with its internal name, type and required flag. Those internal names are what the create and update tools expect.',
    ),
});

const listListItemsSchema = z.object({
  siteId: siteIdField,
  listId: listIdField,
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(25)
    .describe('Maximum items to return in one page. Defaults to 25.'),
  filter: z
    .string()
    .min(1)
    .optional()
    .describe(
      "OData $filter over column values, prefixed with `fields/` — e.g. \"fields/Status eq 'Open'\" or 'fields/Priority gt 2'. Column names are internal names.",
    ),
  fields: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Comma-separated internal column names to return, e.g. "Title,Status,Due_x0020_Date". Strongly recommended: without it every column is expanded and the result is often truncated.',
    ),
});

const getListItemSchema = z.object({
  siteId: siteIdField,
  listId: listIdField,
  itemId: z.string().min(1).describe('List item id (the integer id SharePoint shows as "ID", as a string).'),
  fields: z
    .string()
    .min(1)
    .optional()
    .describe('Comma-separated internal column names to return. Omit for every column.'),
});

const fieldsBody = z
  .record(z.string(), z.unknown())
  .describe(
    'Column values keyed by internal column name, e.g. {"Title":"Q3 review","Status":"Open"}. Lookup and person columns are written as "{ColumnName}LookupId" with the target\'s integer lookup id, never a display name.',
  );

const createListItemSchema = z.object({
  siteId: siteIdField,
  listId: listIdField,
  fields: fieldsBody,
});

const updateListItemSchema = z.object({
  siteId: siteIdField,
  listId: listIdField,
  itemId: z.string().min(1).describe('List item id to update.'),
  fields: fieldsBody,
});

const deleteListItemSchema = z.object({
  siteId: siteIdField,
  listId: listIdField,
  itemId: z.string().min(1).describe('List item id to delete.'),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const sharepointModule: ToolModule = {
  group: GROUP,

  build({ graph, config }: ToolDeps): ToolDefinition[] {
    // One rich-text column must not be able to consume the whole output budget.
    const fieldChars = Math.min(2000, Math.max(200, Math.floor(config.maxOutputChars / 20)));

    return [
      {
        name: 'sharepoint_search_sites',
        title: 'Search SharePoint sites',
        description:
          'Searches SharePoint sites and returns id, display name, description and web URL for each match, plus a count and a nextLink when more pages exist. Returns 25 sites by default. The id it returns is the composite value every other sharepoint tool accepts. Traps: this is Graph\'s plain `search` parameter, not `$search`, so it matches only site titles and URLs — never document contents — and it supports no $filter, no date ranges and no sorting. Personal OneDrive sites are excluded. Pass "*" to enumerate every site the signed-in user can reach.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: searchSitesSchema,
        handler: async (args) => {
          const a = searchSitesSchema.parse(args);

          const res = await graph.request<unknown>({
            path: '/sites',
            method: 'GET',
            query: { search: a.query, $top: a.top },
            scopes: READ_SCOPES,
          });

          const sites = extractCollection<GraphSite>(res.data);
          return {
            count: sites.length,
            items: sites.map(projectSite),
            nextLink: res.nextLink,
          };
        },
      },

      {
        name: 'sharepoint_get_site',
        title: 'Get a SharePoint site',
        description:
          'Returns one site: its composite id, display name, description, web URL, hostname and timestamps. Accepts the composite id from sharepoint_search_sites, the literal "root" for the tenant root site, a bare hostname (which resolves to that tenant\'s root site), or the readable "contoso.sharepoint.com:/sites/Marketing" form. Use it to turn a SharePoint URL a user pasted into the id the other sharepoint tools need.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: getSiteSchema,
        handler: async (args) => {
          const a = getSiteSchema.parse(args);

          const res = await graph.request<GraphSite>({
            path: siteRef(a.siteId).self,
            method: 'GET',
            scopes: READ_SCOPES,
          });

          return projectSite(res.data ?? {});
        },
      },

      {
        name: 'sharepoint_list_drives',
        title: 'List site document libraries',
        description:
          'Lists the document libraries (drives) on a site with id, name, description, drive type and web URL, plus a count and a nextLink when more pages exist. Returns 25 libraries by default. Most sites have one library called "Documents"; pass the returned drive id to sharepoint_list_drive_items. Reading or writing file content is handled by the files tools, not here.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: listDrivesSchema,
        handler: async (args) => {
          const a = listDrivesSchema.parse(args);
          const site = siteRef(a.siteId);

          const res = await graph.request<unknown>({
            path: site.child('drives'),
            method: 'GET',
            query: { $top: a.top },
            scopes: READ_SCOPES,
          });

          const drives = extractCollection<GraphDrive>(res.data);
          return {
            count: drives.length,
            items: drives.map(projectDrive),
            nextLink: res.nextLink,
          };
        },
      },

      {
        name: 'sharepoint_list_drive_items',
        title: 'List document library items',
        description:
          'Lists the children of a document-library folder: name, whether it is a file or a folder, size, child count, MIME type, last modified time and author, and the web URL, plus a count and a nextLink when more pages exist. Defaults to the library root; pass itemId or folderPath (not both) to descend into a folder. Returns 50 children by default. Pre-authenticated download URLs are deliberately omitted — use the files tools to read content.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: listDriveItemsSchema,
        handler: async (args) => {
          const a = listDriveItemsSchema.parse(args);

          if (a.itemId !== undefined && a.folderPath !== undefined) {
            throw new Error('Pass either itemId or folderPath, not both.');
          }

          const site = siteRef(a.siteId);
          const drive = site.child(`drives/${seg(a.driveId, 'driveId')}`);
          let path: string;
          if (a.itemId !== undefined) {
            path = `${drive}/items/${seg(a.itemId, 'itemId')}/children`;
          } else if (a.folderPath !== undefined) {
            path = `${drive}/root:/${encodeItemPath(a.folderPath, 'folderPath')}:/children`;
          } else {
            path = `${drive}/root/children`;
          }

          const res = await graph.request<unknown>({
            path,
            method: 'GET',
            query: { $top: a.top, $select: DRIVE_ITEM_SELECT },
            scopes: READ_SCOPES,
          });

          const items = extractCollection<GraphDriveItem>(res.data);
          return {
            count: items.length,
            items: items.map(projectDriveItem),
            nextLink: res.nextLink,
          };
        },
      },

      {
        name: 'sharepoint_list_lists',
        title: 'List SharePoint lists',
        description:
          'Lists the SharePoint lists on a site with id, display name, description, template and web URL, plus a count and a nextLink when more pages exist. Returns 50 lists by default. Note that document libraries are lists too and appear here with template "documentLibrary" — use sharepoint_list_drives to browse their files. Hidden system lists are omitted unless includeHidden is set.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: listListsSchema,
        handler: async (args) => {
          const a = listListsSchema.parse(args);
          const site = siteRef(a.siteId);

          const res = await graph.request<unknown>({
            path: site.child('lists'),
            method: 'GET',
            query: { $top: a.top },
            scopes: READ_SCOPES,
          });

          const lists = extractCollection<GraphList>(res.data);
          // Graph has no server-side filter for the `hidden` facet on lists, so
          // the page is filtered here; a filtered page can come back short.
          const visible = a.includeHidden ? lists : lists.filter((l) => l.list?.hidden !== true);

          return {
            count: visible.length,
            hiddenOmitted: a.includeHidden ? undefined : lists.length - visible.length,
            items: visible.map(projectList),
            nextLink: res.nextLink,
          };
        },
      },

      {
        name: 'sharepoint_get_list',
        title: 'Get a SharePoint list',
        description:
          'Returns one list: id, display name, description, template, web URL and timestamps. Set includeColumns to also get every column with its internal name, type, choice values and required flag. Those internal names are what sharepoint_create_list_item and sharepoint_update_list_item expect and they usually differ from what the UI shows — a column displayed as "Due Date" is internally "Due_x0020_Date". The columns list also marks which columns are indexed, and only indexed columns filter reliably.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: getListSchema,
        handler: async (args) => {
          const a = getListSchema.parse(args);
          const site = siteRef(a.siteId);

          const res = await graph.request<GraphList>({
            path: site.child(`lists/${seg(a.listId, 'listId')}`),
            method: 'GET',
            query: a.includeColumns ? { $expand: 'columns' } : undefined,
            scopes: READ_SCOPES,
          });

          const list = res.data ?? {};
          const out = projectList(list);
          if (a.includeColumns) {
            const columns = list.columns ?? [];
            out['columns'] = columns.filter((c) => c.hidden !== true).map(projectColumn);
          }
          return out;
        },
      },

      {
        name: 'sharepoint_list_list_items',
        title: 'List SharePoint list items',
        description:
          'Lists items in a SharePoint list with their column values, plus a count and a nextLink when more pages exist. Returns 25 items by default. Pass `fields` with the internal column names you need — without it every column is expanded and the output is easily truncated; purely presentational columns (LinkTitle, DocIcon, …) are dropped unless you name them. Traps: a `filter` must reference columns as `fields/InternalName`, and SharePoint refuses filters on non-indexed columns, so this tool sends the HonorNonIndexedQueriesWarningMayFailRandomly header to let those run anyway — on a large list they can still fail intermittently. Use sharepoint_get_list with includeColumns to learn the internal names and which columns are indexed.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: listListItemsSchema,
        handler: async (args) => {
          const a = listListItemsSchema.parse(args);
          const site = siteRef(a.siteId);
          const columns = parseColumnList(a.fields, 'fields');

          const res = await graph.request<unknown>({
            path: site.child(`lists/${seg(a.listId, 'listId')}/items`),
            method: 'GET',
            query: { $top: a.top, $expand: fieldsExpand(columns), $filter: a.filter },
            headers: a.filter === undefined ? undefined : HONOR_NON_INDEXED,
            scopes: READ_SCOPES,
          });

          const items = extractCollection<GraphListItem>(res.data);
          return {
            count: items.length,
            items: items.map((item) => projectListItem(item, fieldChars, columns !== undefined)),
            nextLink: res.nextLink,
          };
        },
      },

      {
        name: 'sharepoint_get_list_item',
        title: 'Get a SharePoint list item',
        description:
          'Returns one list item with its column values, creator, last editor, timestamps and web URL. Pass `fields` to limit which columns come back. The itemId is SharePoint\'s integer "ID" column as a string, not a GUID; sharepoint_list_list_items returns it.',
        group: 'sharepoint',
        scopes: READ_SCOPES,
        inputSchema: getListItemSchema,
        handler: async (args) => {
          const a = getListItemSchema.parse(args);
          const site = siteRef(a.siteId);
          const columns = parseColumnList(a.fields, 'fields');

          const res = await graph.request<GraphListItem>({
            path: site.child(`lists/${seg(a.listId, 'listId')}/items/${seg(a.itemId, 'itemId')}`),
            method: 'GET',
            query: { $expand: fieldsExpand(columns) },
            scopes: READ_SCOPES,
          });

          return projectListItem(res.data ?? {}, fieldChars, columns !== undefined);
        },
      },

      {
        name: 'sharepoint_create_list_item',
        title: 'Create a SharePoint list item',
        description:
          'Creates an item in a SharePoint list and returns its new id, web URL and stored column values. `fields` is keyed by internal column name — call sharepoint_get_list with includeColumns first, because the internal name rarely matches the displayed one. Most lists require Title. Traps: read-only columns (ID, Created, Author, Modified) are rejected; lookup and person columns are written as "{ColumnName}LookupId" with the target\'s integer lookup id rather than a name; a choice column only accepts one of its defined choices unless fill-in is enabled.',
        group: 'sharepoint',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: createListItemSchema,
        handler: async (args) => {
          const a = createListItemSchema.parse(args);
          if (Object.keys(a.fields).length === 0) {
            throw new Error('fields must contain at least one column value.');
          }

          const site = siteRef(a.siteId);
          const res = await graph.request<GraphListItem>({
            path: site.child(`lists/${seg(a.listId, 'listId')}/items`),
            method: 'POST',
            body: { fields: a.fields },
            scopes: WRITE_SCOPES,
          });

          return {
            created: true,
            listId: a.listId,
            item: projectListItem(res.data ?? {}, fieldChars, true),
          };
        },
      },

      {
        name: 'sharepoint_update_list_item',
        title: 'Update a SharePoint list item',
        description:
          'Updates column values on an existing list item and returns the stored values. This is a partial update against the item\'s field set: columns you do not pass are left alone, and passing null clears a column. Uses internal column names, same as sharepoint_create_list_item. Read-only columns (ID, Created, Author, Modified) are rejected by SharePoint. No etag is needed — a concurrent edit is overwritten silently, so read the item first if that matters.',
        group: 'sharepoint',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: updateListItemSchema,
        handler: async (args) => {
          const a = updateListItemSchema.parse(args);
          if (Object.keys(a.fields).length === 0) {
            throw new Error('fields must contain at least one column value.');
          }

          const site = siteRef(a.siteId);
          const res = await graph.request<unknown>({
            path: site.child(`lists/${seg(a.listId, 'listId')}/items/${seg(a.itemId, 'itemId')}/fields`),
            method: 'PATCH',
            body: a.fields,
            scopes: WRITE_SCOPES,
          });

          return {
            updated: true,
            listId: a.listId,
            itemId: a.itemId,
            // PATCH on /fields answers with the fieldValueSet itself, not a listItem.
            fields: projectFields(res.data, fieldChars, true),
          };
        },
      },

      {
        name: 'sharepoint_delete_list_item',
        title: 'Delete a SharePoint list item',
        description:
          'Deletes one item from a SharePoint list. The item goes to the site recycle bin rather than being destroyed, but this server cannot restore it — confirm with the user first. Returns a confirmation with the deleted id.',
        group: 'sharepoint',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: deleteListItemSchema,
        handler: async (args) => {
          const a = deleteListItemSchema.parse(args);
          const site = siteRef(a.siteId);

          const res = await graph.request<unknown>({
            path: site.child(`lists/${seg(a.listId, 'listId')}/items/${seg(a.itemId, 'itemId')}`),
            method: 'DELETE',
            scopes: WRITE_SCOPES,
          });

          return { deleted: true, listId: a.listId, itemId: a.itemId, status: res.status };
        },
      },
    ];
  },
};
