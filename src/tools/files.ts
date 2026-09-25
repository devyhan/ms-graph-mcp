/**
 * OneDrive file tools for the signed-in user.
 *
 * Every path here is rooted at `/me/drive`, so these tools only ever see the
 * caller's own OneDrive. Items surfaced by `files_list_recent` and
 * `files_list_shared` can live in *someone else's* drive; those come back as
 * stubs carrying a `remoteItem`, and the projection below lifts the remote ids
 * out so the caller can at least identify what it found.
 */

import { z } from 'zod';
import type { ToolDefinition, ToolDeps, ToolGroupMeta, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { escapeODataString } from '../util/odata.js';
import { truncateText } from '../util/truncate.js';

// The group catalogue is a static literal; `files` is always present.
const GROUP: ToolGroupMeta = GROUPS['files']!;
const READ_SCOPES: string[] = [...GROUP.readScopes];
const WRITE_SCOPES: string[] = [...GROUP.readScopes, ...GROUP.writeScopes];

/**
 * `sharedWithMe` reads items that live in other users' drives, which `Files.Read`
 * (own drive only) does not cover — Graph answers 403 without a `.All` scope.
 * It is still user-consentable, so this does not make the group admin-only.
 */
const SHARED_READ_SCOPES: string[] = [...GROUP.readScopes, 'Files.Read.All'];

const DRIVE = '/me/drive';

// ---------------------------------------------------------------------------
// Graph payload shapes (only the fields these tools project)
// ---------------------------------------------------------------------------

interface IdentitySet {
  user?: { displayName?: string; email?: string } | null;
  application?: { displayName?: string } | null;
}

interface FileFacet {
  mimeType?: string;
}

interface FolderFacet {
  childCount?: number;
}

interface ParentReference {
  driveId?: string;
  id?: string;
  path?: string;
}

interface RemoteItem {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  file?: FileFacet | null;
  folder?: FolderFacet | null;
  parentReference?: ParentReference | null;
  lastModifiedDateTime?: string;
}

interface DriveItem {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  eTag?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: FileFacet | null;
  folder?: FolderFacet | null;
  package?: { type?: string } | null;
  parentReference?: ParentReference | null;
  remoteItem?: RemoteItem | null;
  createdBy?: IdentitySet | null;
  lastModifiedBy?: IdentitySet | null;
  shared?: { scope?: string } | null;
}

interface SharingLink {
  type?: string;
  scope?: string;
  webUrl?: string;
  preventsDownload?: boolean;
}

interface Permission {
  id?: string;
  roles?: string[];
  link?: SharingLink | null;
  expirationDateTime?: string;
  hasPassword?: boolean;
}

// ---------------------------------------------------------------------------
// Path building
// ---------------------------------------------------------------------------

/** Drive item ids are opaque and carry `!` on personal OneDrive. */
function seg(id: string): string {
  // These three encode to `%2f`, `%5c` and `%25`, which normalizeGraphPath
  // refuses as traversal smuggling — an error that blames the path rather than
  // the id. Real OneDrive ids never contain them, so say what actually went
  // wrong: the caller most likely passed a file path where an id belongs.
  const bad = /[/\\%]/.exec(id);
  if (bad !== null) {
    throw new Error(
      `Drive item id ${JSON.stringify(id)} contains ${JSON.stringify(bad[0])}, which no OneDrive item id does. Pass an id from a files_* listing, or use the path argument to address an item by name.`,
    );
  }
  return encodeURIComponent(id);
}

// Characters OneDrive rejects in item names. ':' additionally terminates the
// `root:/path:` addressing grammar, and '\' is refused outright by
// normalizeGraphPath. Spaces are legal in OneDrive names and stay legal here.
const FORBIDDEN_NAME_CHARS = /["*:<>?\\|]|[\u0000-\u001f\u007f]/;

/**
 * `normalizeGraphPath` — applied by the Graph client to every outgoing path —
 * rejects the escapes `%25`, `%2f`, `%2e` and `%5c` outright, because for the
 * generic caller those are traversal smuggling. A literal `%` in a file name
 * encodes to exactly `%25`, so path addressing simply cannot express such a
 * name. Fail here with an actionable message rather than letting the client
 * raise a confusing security error.
 */
function assertAddressable(value: string, what: string): void {
  if (value.includes('%')) {
    throw new Error(
      `${what} contains "%", which cannot be expressed in a OneDrive path URL. Address the item by itemId instead.`,
    );
  }
  const bad = FORBIDDEN_NAME_CHARS.exec(value);
  if (bad !== null) {
    throw new Error(
      `${what} contains the character ${JSON.stringify(bad[0])}, which OneDrive does not allow in item names.`,
    );
  }
}

/**
 * A search term rides inside an OData function literal in the URL *path*, so it
 * meets the same `normalizeGraphPath` guard: '/', '\' and '%' would encode to
 * `%2f`, `%5c` and `%25`, which it rejects as traversal smuggling. None of the
 * three is meaningful to OneDrive's full-text index — it tokenises on them
 * anyway — so fold them to spaces instead of refusing the search. ':' is left
 * alone: it is legal later in a path and carries meaning in query syntax.
 */
function sanitizeSearchTerm(query: string): string {
  return query
    .replace(/[/\\%\u0000-\u001f\u007f]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Splits a user-supplied drive path into encoded segments. Returns an empty
 * array for the root, so callers can pick the shorter `/me/drive/root` form.
 */
function drivePathSegments(raw: string): string[] {
  const parts = raw
    .split('/')
    .map((part) => part.trim())
    .filter((part) => part.length > 0);

  const encoded: string[] = [];
  for (const part of parts) {
    if (part === '.' || part === '..') {
      throw new Error(
        `Relative segment ${JSON.stringify(part)} is not allowed in a OneDrive path; give the full path from the drive root.`,
      );
    }
    assertAddressable(part, `Path segment ${JSON.stringify(part)}`);
    encoded.push(encodeURIComponent(part));
  }
  return encoded;
}

/**
 * Resolves the two mutually exclusive addressing modes to a single Graph path.
 * `suffix` is appended after the item, e.g. `/children` or `/content`.
 */
function itemPath(itemId: string | undefined, path: string | undefined, suffix: string): string {
  if (itemId !== undefined && path !== undefined) {
    throw new Error('Pass either itemId or path, not both: they address the same item two different ways.');
  }

  if (itemId !== undefined) return `${DRIVE}/items/${seg(itemId)}${suffix}`;

  if (path !== undefined) {
    const segments = drivePathSegments(path);
    if (segments.length > 0) {
      // The closing ':' terminates the path only when more URL follows;
      // `/me/drive/root:/a/b.txt` is the documented form when nothing does.
      const joined = segments.join('/');
      return suffix === '' ? `${DRIVE}/root:/${joined}` : `${DRIVE}/root:/${joined}:${suffix}`;
    }
  }

  return `${DRIVE}/root${suffix}`;
}

// ---------------------------------------------------------------------------
// Projection
// ---------------------------------------------------------------------------

function compact<T extends Record<string, unknown>>(value: T): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value)) {
    if (v !== undefined && v !== null) out[key] = v;
  }
  return out;
}

function displayName(identity: IdentitySet | null | undefined): string | undefined {
  const name = identity?.user?.displayName;
  return typeof name === 'string' && name.length > 0 ? name : undefined;
}

function safeDecode(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/** `parentReference.path` is `/drive/root:/Documents`; the prefix is noise. */
function folderPath(parent: ParentReference | null | undefined): string | undefined {
  const raw = parent?.path;
  if (typeof raw !== 'string' || raw.length === 0) return undefined;
  const marker = raw.indexOf('root:');
  const tail = marker === -1 ? raw : raw.slice(marker + 'root:'.length);
  const decoded = safeDecode(tail);
  return decoded === '' ? '/' : decoded;
}

function isFolder(item: { folder?: FolderFacet | null }): boolean {
  return item.folder !== null && item.folder !== undefined;
}

function projectItem(item: DriveItem): Record<string, unknown> {
  const remote = item.remoteItem;
  const file = item.file ?? remote?.file;
  const folder = item.folder ?? remote?.folder;

  return compact({
    id: item.id,
    name: item.name ?? remote?.name,
    kind: folder !== null && folder !== undefined ? 'folder' : 'file',
    size: item.size ?? remote?.size,
    mimeType: file?.mimeType,
    childCount: folder?.childCount,
    lastModifiedDateTime: item.lastModifiedDateTime ?? remote?.lastModifiedDateTime,
    lastModifiedBy: displayName(item.lastModifiedBy),
    folderPath: folderPath(item.parentReference),
    webUrl: item.webUrl ?? remote?.webUrl,
    // sharedWithMe and recent return stubs: the real item lives in another
    // drive and is only addressable through these two ids.
    remoteDriveId: remote?.parentReference?.driveId,
    remoteItemId: remote?.id,
  });
}

function projectItemDetail(item: DriveItem): Record<string, unknown> {
  const pkg = item.package;
  // `@microsoft.graph.downloadUrl` is deliberately never returned: it is a
  // pre-authenticated URL that grants the file's contents to anyone holding it.
  return compact({
    ...projectItem(item),
    kind: isFolder(item) ? 'folder' : pkg !== null && pkg !== undefined ? 'package' : 'file',
    packageType: pkg?.type,
    createdDateTime: item.createdDateTime,
    createdBy: displayName(item.createdBy),
    parentItemId: item.parentReference?.id,
    driveId: item.parentReference?.driveId,
    sharedScope: item.shared?.scope,
    eTag: item.eTag,
  });
}

function collectionResult(
  items: Record<string, unknown>[],
  key: string,
  nextLink: string | undefined,
): Record<string, unknown> {
  return compact({ count: items.length, [key]: items, nextLink });
}

// ---------------------------------------------------------------------------
// Text-file detection
// ---------------------------------------------------------------------------

const TEXT_EXTENSIONS = new Set([
  'txt', 'text', 'md', 'markdown', 'rst', 'adoc', 'log', 'csv', 'tsv', 'json', 'jsonl', 'ndjson',
  'xml', 'yaml', 'yml', 'toml', 'ini', 'cfg', 'conf', 'env', 'properties', 'tf', 'tfvars',
  'html', 'htm', 'css', 'scss', 'sass', 'less', 'svg', 'vue', 'svelte',
  'js', 'mjs', 'cjs', 'jsx', 'ts', 'tsx', 'mts', 'cts', 'py', 'rb', 'go', 'rs', 'java', 'kt', 'kts',
  'c', 'h', 'cpp', 'cc', 'cxx', 'hpp', 'cs', 'swift', 'php', 'pl', 'lua', 'r', 'scala', 'dart',
  'sh', 'bash', 'zsh', 'fish', 'ps1', 'bat', 'cmd', 'sql', 'graphql', 'gql', 'proto',
  'dockerfile', 'gitignore', 'gitattributes', 'editorconfig', 'patch', 'diff',
  'srt', 'vtt', 'ics', 'vcf', 'plist', 'gradle', 'cmake', 'mk',
]);

const BINARY_EXTENSIONS = new Set([
  'pdf', 'doc', 'docx', 'dot', 'dotx', 'xls', 'xlsx', 'xlsb', 'ppt', 'pptx', 'pps', 'ppsx',
  'one', 'onepkg', 'onetoc2', 'vsd', 'vsdx', 'pub', 'accdb', 'mdb', 'odt', 'ods', 'odp', 'rtf',
  'zip', '7z', 'rar', 'gz', 'tgz', 'bz2', 'xz', 'tar', 'jar', 'war', 'iso', 'dmg', 'pkg', 'msi',
  'exe', 'dll', 'so', 'dylib', 'bin', 'class', 'pyc', 'wasm', 'lib',
  'png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff', 'webp', 'heic', 'heif', 'ico', 'psd', 'ai',
  'eps', 'raw', 'cr2', 'nef', 'arw', 'dng', 'avif',
  'mp3', 'wav', 'flac', 'aac', 'ogg', 'oga', 'm4a', 'wma', 'mp4', 'mov', 'avi', 'mkv', 'wmv',
  'webm', 'm4v', 'mpg', 'mpeg',
  'ttf', 'otf', 'woff', 'woff2', 'eot', 'db', 'sqlite', 'sqlite3', 'pst', 'ost',
]);

const TEXT_MIME_TYPES = new Set([
  'application/json', 'application/ld+json', 'application/x-ndjson', 'application/csv',
  'application/xml', 'application/xhtml+xml', 'application/javascript', 'application/x-javascript',
  'application/ecmascript', 'application/yaml', 'application/x-yaml', 'application/toml',
  'application/sql', 'application/x-sh', 'application/x-shellscript', 'application/graphql',
  'image/svg+xml',
]);

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  if (dot <= 0 || dot === name.length - 1) return '';
  return name.slice(dot + 1).toLowerCase();
}

/**
 * Extension wins over `file.mimeType`, because OneDrive derives the mime type
 * from the extension and gets developer files badly wrong — `.ts` is reported
 * as `video/mp2t`, and anything unrecognised becomes
 * `application/octet-stream`. Trusting the mime type alone would refuse to read
 * most source files.
 *
 * Returns the reason to refuse, or undefined when the file looks like text.
 */
function textRefusalReason(name: string, mimeType: string | undefined): string | undefined {
  const ext = extensionOf(name);

  if (ext !== '' && BINARY_EXTENSIONS.has(ext)) {
    return `"${name}" is a binary .${ext} file. This tool only reads text-like files; use files_create_link or the item's webUrl to open it in its native app.`;
  }
  if (ext !== '' && TEXT_EXTENSIONS.has(ext)) return undefined;

  const mime = (mimeType ?? '').toLowerCase().split(';')[0]?.trim() ?? '';
  if (mime.startsWith('text/') || TEXT_MIME_TYPES.has(mime)) return undefined;

  const seen = mime === '' ? 'no reported content type' : `content type ${mime}`;
  const extNote = ext === '' ? ' and no file extension' : ` and an unrecognised .${ext} extension`;
  return `"${name}" has ${seen}${extNote}, so it is not known to hold text. This tool refuses to read binary content.`;
}

/**
 * The Graph client owns response decoding, and `/content` is the one endpoint
 * here that does not return JSON. Accept whatever shape it hands back rather
 * than assuming one.
 */
function coerceText(data: unknown): string {
  if (typeof data === 'string') return data;
  if (data === null || data === undefined) return '';
  if (data instanceof Uint8Array) return new TextDecoder().decode(data);
  if (data instanceof ArrayBuffer) return new TextDecoder().decode(new Uint8Array(data));
  // A .json or .xml file may already have been parsed; re-serialising is much
  // closer to the file's real contents than "[object Object]".
  try {
    return JSON.stringify(data, null, 2) ?? String(data);
  } catch {
    return String(data);
  }
}

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const ITEM_ID_DESC = 'Drive item id, as returned by any files_* listing tool. Mutually exclusive with path.';
const PATH_DESC =
  'Path from the OneDrive root, e.g. "Documents/Reports/q3.docx". Leading slashes are ignored. Mutually exclusive with itemId.';

const listChildrenSchema = z.object({
  itemId: z
    .string()
    .min(1)
    .optional()
    .describe(`Id of the folder to list. ${ITEM_ID_DESC} Omit both to list the drive root.`),
  path: z
    .string()
    .optional()
    .describe(`Path of the folder to list, e.g. "Documents/Reports". ${PATH_DESC} Omit both to list the drive root.`),
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(50)
    .describe('Maximum number of children to return in this page. Defaults to 50.'),
  orderby: z
    .enum([
      'name asc',
      'name desc',
      'lastModifiedDateTime asc',
      'lastModifiedDateTime desc',
      'size asc',
      'size desc',
    ])
    .optional()
    .describe(
      'Sort order for the children. Graph defaults to name ascending. OneDrive for Business (SharePoint-backed) drives silently ignore some sorts, notably by size.',
    ),
});

const getItemSchema = z.object({
  itemId: z.string().min(1).optional().describe(ITEM_ID_DESC),
  path: z.string().optional().describe(PATH_DESC),
});

const searchSchema = z.object({
  query: z
    .string()
    .min(1)
    .describe('Free text matched against file names and indexed file contents across the whole drive.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(25)
    .describe('Maximum number of matches to return in this page. Defaults to 25.'),
});

const listRecentSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(25)
    .describe('Maximum number of recent items to return. Defaults to 25.'),
});

const listSharedSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(25)
    .describe(
      'Maximum number of shared items to return. Defaults to 25. sharedWithMe applies this loosely and may return a different number.',
    ),
});

const readTextSchema = z.object({
  itemId: z
    .string()
    .min(1)
    .describe('Id of the file to read, as returned by files_list_children or files_search.'),
  maxChars: z
    .number()
    .int()
    .min(200)
    .max(200_000)
    .default(20_000)
    .describe('Maximum characters of file content to return. Defaults to 20000; longer files are truncated with a marker.'),
});

const createFolderSchema = z.object({
  name: z
    .string()
    .min(1)
    .describe('Name of the new folder. An existing folder of the same name does not cause a failure; the new one gets a numeric suffix.'),
  parentItemId: z
    .string()
    .min(1)
    .optional()
    .describe('Id of the parent folder. Mutually exclusive with parentPath. Omit both to create in the drive root.'),
  parentPath: z
    .string()
    .optional()
    .describe(
      'Path of the parent folder from the drive root, e.g. "Documents/Reports". Mutually exclusive with parentItemId. Omit both to create in the drive root.',
    ),
});

const uploadTextSchema = z.object({
  path: z
    .string()
    .min(1)
    .describe(
      'Destination path from the OneDrive root including the file name, e.g. "Notes/meeting-2026-09-03.md". Missing parent folders are created automatically.',
    ),
  content: z.string().describe('The full text to write. This replaces the file contents; there is no append mode.'),
});

const deleteSchema = z.object({
  itemId: z
    .string()
    .min(1)
    .describe('Id of the file or folder to delete. Deleting a folder deletes everything inside it.'),
});

const createLinkSchema = z.object({
  itemId: z.string().min(1).describe('Id of the file or folder to create a sharing link for.'),
  type: z
    .enum(['view', 'edit'])
    .default('view')
    .describe('Link permission: view for read-only, edit for read-write. Defaults to view.'),
  scope: z
    .enum(['anonymous', 'organization'])
    .default('organization')
    .describe(
      'Who can use the link. organization requires the recipient to sign in to the same tenant; anonymous works for anyone holding the URL and is disabled by policy in many tenants.',
    ),
});

// Simple PUT upload is documented up to 250 MB, but 4 MB is the boundary that
// holds everywhere; past it Graph wants a resumable upload session, which this
// tool does not implement.
const MAX_UPLOAD_BYTES = 4 * 1024 * 1024;

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const filesModule: ToolModule = {
  group: GROUP,

  build({ graph, config }: ToolDeps): ToolDefinition[] {
    // File content is the only unbounded field here, so cap it against the
    // server's own output budget instead of trusting the caller's maxChars.
    const contentCeiling = Math.max(2_000, Math.floor(config.maxOutputChars * 0.8));

    return [
      {
        name: 'files_list_children',
        title: 'List OneDrive folder contents',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          "Lists the direct children of a folder in the signed-in user's OneDrive, returning id, name, kind (file or folder), size, mime type, child count, last-modified time and webUrl per entry, plus count and nextLink when more pages exist. Address the folder by itemId or by path; omit both for the drive root. Page size defaults to 50. This is not recursive: use files_search to look through the whole drive.",
        inputSchema: listChildrenSchema,
        handler: async (args) => {
          const { itemId, path, top, orderby } = listChildrenSchema.parse(args);
          const res = await graph.request({
            path: itemPath(itemId, path, '/children'),
            query: {
              $top: top,
              $orderby: orderby,
              $select: 'id,name,size,file,folder,webUrl,lastModifiedDateTime,lastModifiedBy,parentReference',
            },
            scopes: READ_SCOPES,
          });
          const items = extractCollection<DriveItem>(res.data).map(projectItem);
          return collectionResult(items, 'items', res.nextLink);
        },
      },

      {
        name: 'files_get_item',
        title: 'Get a OneDrive item',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          'Reads the metadata of one OneDrive file or folder: name, kind, size, mime type, created and last-modified times and authors, parent folder path and ids, and webUrl. Address it by itemId or by path, not both. The pre-authenticated download URL is deliberately not returned; use files_read_text for text contents or files_create_link to share the item.',
        inputSchema: getItemSchema,
        handler: async (args) => {
          const { itemId, path } = getItemSchema.parse(args);
          if (itemId === undefined && path === undefined) {
            throw new Error('Give either itemId or path to identify the item.');
          }
          const res = await graph.request<DriveItem>({
            path: itemPath(itemId, path, ''),
            scopes: READ_SCOPES,
          });
          return projectItemDetail(res.data);
        },
      },

      {
        name: 'files_search',
        title: 'Search OneDrive',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          "Searches the whole of the signed-in user's OneDrive for free text, matching file names and indexed file contents, and returns the same compact fields as files_list_children plus count and nextLink. Page size defaults to 25. This endpoint accepts only a search term: it supports no date range, no file-type filter and no sort order, so narrow the results yourself or list a folder with files_list_children instead. Recently uploaded files are missing until the service has indexed them.",
        inputSchema: searchSchema,
        handler: async (args) => {
          const { query, top } = searchSchema.parse(args);
          // The term sits inside an OData function literal in the *path*: strip
          // what cannot ride there, double the single quotes, then percent-encode.
          const cleaned = sanitizeSearchTerm(query);
          if (cleaned === '') {
            throw new Error(`Search query ${JSON.stringify(query)} has no searchable characters left once "/", "\\" and "%" are removed.`);
          }
          const term = encodeURIComponent(escapeODataString(cleaned));
          const res = await graph.request({
            path: `${DRIVE}/root/search(q='${term}')`,
            query: { $top: top },
            scopes: READ_SCOPES,
          });
          const items = extractCollection<DriveItem>(res.data).map(projectItem);
          return collectionResult(items, 'items', res.nextLink);
        },
      },

      {
        name: 'files_list_recent',
        title: 'List recent OneDrive files',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          "Lists files the signed-in user recently opened or edited, most recent first, with count and nextLink. Defaults to 25 items. Entries can come from other people's drives or from SharePoint; those carry remoteDriveId and remoteItemId, and their plain id will not resolve against your own drive.",
        inputSchema: listRecentSchema,
        handler: async (args) => {
          const { top } = listRecentSchema.parse(args);
          const res = await graph.request({
            path: `${DRIVE}/recent`,
            query: { $top: top },
            scopes: READ_SCOPES,
          });
          const items = extractCollection<DriveItem>(res.data).map(projectItem);
          return collectionResult(items, 'items', res.nextLink);
        },
      },

      {
        name: 'files_list_shared',
        title: 'List files shared with me',
        group: GROUP.name,
        // Files.Read covers the caller's own drive only; these items are not in it.
        scopes: SHARED_READ_SCOPES,
        description:
          "Lists files and folders other people have shared with the signed-in user, with count and nextLink. Defaults to 25 items. These entries are stubs pointing into the owner's drive: use remoteDriveId together with remoteItemId to identify one, or open its webUrl. The other files_* tools address only your own OneDrive, so they cannot read these items by id. Needs the Files.Read.All scope, which a user can consent to without an administrator.",
        inputSchema: listSharedSchema,
        handler: async (args) => {
          const { top } = listSharedSchema.parse(args);
          const res = await graph.request({
            path: `${DRIVE}/sharedWithMe`,
            query: { $top: top },
            scopes: SHARED_READ_SCOPES,
          });
          const items = extractCollection<DriveItem>(res.data).map(projectItem);
          return collectionResult(items, 'items', res.nextLink);
        },
      },

      {
        name: 'files_read_text',
        title: 'Read a text file from OneDrive',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          'Returns the contents of a text-like OneDrive file (plain text, Markdown, CSV, JSON, XML, source code, config) as a string, truncated to maxChars, which defaults to 20000. Reads the file metadata first and refuses binary formats — PDFs, images, Office documents, archives, media — with an explanation, because their bytes are not readable as text. Word, Excel and PowerPoint files are zip containers and cannot be read here.',
        inputSchema: readTextSchema,
        handler: async (args) => {
          const { itemId, maxChars } = readTextSchema.parse(args);
          const limit = Math.min(maxChars, contentCeiling);

          const meta = await graph.request<DriveItem>({
            path: `${DRIVE}/items/${seg(itemId)}`,
            query: { $select: 'id,name,size,file,folder,package,webUrl,lastModifiedDateTime' },
            scopes: READ_SCOPES,
          });
          const item = meta.data;
          const name = item.name ?? itemId;

          if (isFolder(item)) {
            throw new Error(`"${name}" is a folder, not a file. Use files_list_children to see what is inside it.`);
          }
          const pkg = item.package;
          if (pkg !== null && pkg !== undefined) {
            throw new Error(
              `"${name}" is a ${pkg.type ?? 'package'} package (such as a OneNote notebook), which has no readable text stream.`,
            );
          }
          const refusal = textRefusalReason(name, item.file?.mimeType);
          if (refusal !== undefined) throw new Error(refusal);

          // Ask for only as much as could survive truncation. Graph redirects
          // /content to a storage URL that honours Range; if the header is
          // ignored, the client-side truncation below still bounds the result.
          const size = item.size;
          const byteBudget = limit * 4 + 1024;
          const headers: Record<string, string> =
            typeof size === 'number' && size > byteBudget ? { Range: `bytes=0-${byteBudget - 1}` } : {};

          const content = await graph.request({
            path: `${DRIVE}/items/${seg(itemId)}/content`,
            headers,
            scopes: READ_SCOPES,
          });

          return compact({
            id: item.id,
            name: item.name,
            mimeType: item.file?.mimeType,
            size,
            lastModifiedDateTime: item.lastModifiedDateTime,
            webUrl: item.webUrl,
            content: truncateText(coerceText(content.data), limit),
          });
        },
      },

      {
        name: 'files_create_folder',
        title: 'Create a OneDrive folder',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          "Creates a folder in the signed-in user's OneDrive and returns the created item. Give the parent as parentItemId or parentPath, or omit both for the drive root. A name collision does not fail: Graph renames the new folder with a numeric suffix, so read the returned name rather than assuming it.",
        inputSchema: createFolderSchema,
        handler: async (args) => {
          const { name, parentItemId, parentPath } = createFolderSchema.parse(args);
          if (name.includes('/')) {
            throw new Error('Folder name must not contain "/". Use parentPath for the location and name for the folder itself.');
          }
          assertAddressable(name, 'Folder name');

          const res = await graph.request<DriveItem>({
            path: itemPath(parentItemId, parentPath, '/children'),
            method: 'POST',
            body: {
              name,
              folder: {},
              '@microsoft.graph.conflictBehavior': 'rename',
            },
            scopes: WRITE_SCOPES,
          });
          return projectItemDetail(res.data);
        },
      },

      {
        name: 'files_upload_text',
        title: 'Upload a text file to OneDrive',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          "Writes a text file to the given path in the signed-in user's OneDrive and returns the stored item. Missing parent folders are created automatically. An existing file at that path is REPLACED, not merged or appended, so read it with files_read_text first if you mean to edit it. Text only, up to 4 MB; larger or binary uploads need a resumable upload session, which this server does not expose.",
        inputSchema: uploadTextSchema,
        handler: async (args) => {
          const { path, content } = uploadTextSchema.parse(args);

          const segments = drivePathSegments(path);
          if (segments.length === 0) {
            throw new Error('path must include a file name, e.g. "Notes/summary.md".');
          }

          const bytes = new TextEncoder().encode(content).length;
          if (bytes > MAX_UPLOAD_BYTES) {
            throw new Error(
              `Content is ${bytes} bytes; the simple upload used here is limited to ${MAX_UPLOAD_BYTES} bytes. Split the file or upload it another way.`,
            );
          }

          // `/content` stores the request body as the file's bytes verbatim, so
          // this is the one call in this module that must NOT be JSON-encoded.
          // The Graph client passes a string body through untouched and lets an
          // explicit Content-Type stand; do not "normalise" this to an object,
          // or the file gains surrounding quotes and literal backslash-n.
          const res = await graph.request<DriveItem>({
            path: `${DRIVE}/root:/${segments.join('/')}:/content`,
            method: 'PUT',
            body: content,
            headers: { 'Content-Type': 'text/plain; charset=utf-8' },
            scopes: WRITE_SCOPES,
          });
          return projectItemDetail(res.data);
        },
      },

      {
        name: 'files_delete',
        title: 'Delete a OneDrive item',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          "Deletes a file or folder from the signed-in user's OneDrive. Deleting a folder deletes everything inside it. The item goes to the OneDrive recycle bin rather than being destroyed, but this server has no tool to restore it, so confirm the id with files_get_item first.",
        inputSchema: deleteSchema,
        handler: async (args) => {
          const { itemId } = deleteSchema.parse(args);
          await graph.request({
            path: `${DRIVE}/items/${seg(itemId)}`,
            method: 'DELETE',
            scopes: WRITE_SCOPES,
          });
          return { deleted: true, itemId, note: 'Moved to the OneDrive recycle bin.' };
        },
      },

      {
        name: 'files_create_link',
        title: 'Create a OneDrive sharing link',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          'Creates a sharing link for a OneDrive file or folder and returns its URL, permission type and scope. Defaults to a view link scoped to the organization. Anonymous links are blocked by policy in many tenants, which surfaces as a 403: retry with scope organization. Calling this twice with the same type and scope returns the existing link rather than creating a second one.',
        inputSchema: createLinkSchema,
        handler: async (args) => {
          const { itemId, type, scope } = createLinkSchema.parse(args);
          const res = await graph.request<Permission>({
            path: `${DRIVE}/items/${seg(itemId)}/createLink`,
            method: 'POST',
            body: { type, scope },
            scopes: WRITE_SCOPES,
          });
          const permission = res.data;
          return compact({
            permissionId: permission.id,
            type: permission.link?.type,
            scope: permission.link?.scope,
            webUrl: permission.link?.webUrl,
            roles: permission.roles,
            expirationDateTime: permission.expirationDateTime,
            hasPassword: permission.hasPassword,
          });
        },
      },
    ];
  },
};
