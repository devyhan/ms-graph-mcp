/**
 * The `mail` group: Outlook messages, folders, attachments, and sending.
 *
 * Mail is the group most likely to blow the model's context window. A single
 * HTML message body routinely runs past 100 KB, and `contentBytes` on one PDF
 * attachment is larger than every other tool result in a session combined.
 * Every read here therefore projects Graph's payload down to a fixed set of
 * fields and never returns a body unless the caller asked for one.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { buildQuery, escapeODataString, isoDate, quoteSearch } from '../util/odata.js';
import { stripHtml, truncateText } from '../util/truncate.js';

const GROUP = GROUPS['mail']!;

/**
 * Copied, not aliased: a tool's `scopes` array feeds the consent set the login
 * path builds, and handing out the live `GROUPS` arrays would let a caller that
 * sorts or dedupes in place mutate the group catalogue.
 */
const READ_SCOPES = [...GROUP.readScopes];

/**
 * `Mail.Send` is spelled out here because it is the one write scope in this
 * group that most write tools must NOT ask for. Graph treats sending as a
 * separate permission from mutation: `Mail.ReadWrite` creates, edits, moves and
 * deletes messages but cannot put one on the wire. Mailbox edits that never
 * send anything stay on the narrower set.
 */
const MUTATE_SCOPES = [
  ...GROUP.readScopes,
  ...GROUP.writeScopes.filter((scope) => scope !== 'Mail.Send'),
];

/** Tools that actually transmit a message: mutation scopes plus `Mail.Send`. */
const SEND_SCOPES = [...GROUP.readScopes, ...GROUP.writeScopes];

// ---------------------------------------------------------------------------
// Graph payload shapes
// ---------------------------------------------------------------------------

interface EmailAddress {
  name?: string | null;
  address?: string | null;
}

interface Recipient {
  emailAddress?: EmailAddress | null;
}

interface ItemBody {
  contentType?: string | null;
  content?: string | null;
}

/** The `message` properties this module reads. Extra keys survive for `select`. */
interface GraphMessage {
  id?: string | null;
  subject?: string | null;
  from?: Recipient | null;
  sender?: Recipient | null;
  toRecipients?: Recipient[] | null;
  ccRecipients?: Recipient[] | null;
  receivedDateTime?: string | null;
  sentDateTime?: string | null;
  createdDateTime?: string | null;
  isRead?: boolean | null;
  isDraft?: boolean | null;
  hasAttachments?: boolean | null;
  importance?: string | null;
  conversationId?: string | null;
  parentFolderId?: string | null;
  bodyPreview?: string | null;
  body?: ItemBody | null;
  webLink?: string | null;
  [key: string]: unknown;
}

interface GraphMailFolder {
  id?: string | null;
  displayName?: string | null;
  parentFolderId?: string | null;
  childFolderCount?: number | null;
  unreadItemCount?: number | null;
  totalItemCount?: number | null;
  childFolders?: GraphMailFolder[] | null;
}

interface GraphAttachment {
  id?: string | null;
  name?: string | null;
  contentType?: string | null;
  size?: number | null;
  isInline?: boolean | null;
  lastModifiedDateTime?: string | null;
  '@odata.type'?: string | null;
}

// ---------------------------------------------------------------------------
// Shared validation
// ---------------------------------------------------------------------------

/**
 * `$select` and `$orderby` values reach Graph as raw query values. Every
 * selectable property on `message` is a bare identifier, so anything else is
 * either a caller mistake or an attempt to smuggle extra OData into the query
 * string. `filter` is deliberately exempt — it exists to carry OData.
 */
const SELECT_FIELD = /^[A-Za-z][A-Za-z0-9]*$/;
const ORDERBY_CLAUSE =
  /^[A-Za-z][A-Za-z0-9]*(\/[A-Za-z][A-Za-z0-9]*)*( (asc|desc))?(\s*,\s*[A-Za-z][A-Za-z0-9]*(\/[A-Za-z][A-Za-z0-9]*)*( (asc|desc))?)*$/;

const EMAIL = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/;

const emailArg = z.string().trim().min(3).max(320).regex(EMAIL, 'must be a single email address');

const selectArg = z
  .array(z.string().min(1).regex(SELECT_FIELD, 'must be a bare Microsoft Graph property name'))
  .min(1)
  .max(30);

const DEFAULT_MESSAGE_SELECT = [
  'id',
  'subject',
  'from',
  'toRecipients',
  'receivedDateTime',
  'isRead',
  'hasAttachments',
  'bodyPreview',
  'webLink',
];

/** Properties `summarizeMessage` already emits, so `select` extras never duplicate them. */
const SUMMARY_FIELDS = new Set([...DEFAULT_MESSAGE_SELECT, 'sender']);

const DETAIL_SELECT = [
  'id',
  'subject',
  'from',
  'toRecipients',
  'ccRecipients',
  'receivedDateTime',
  'sentDateTime',
  'isRead',
  'isDraft',
  'hasAttachments',
  'importance',
  'conversationId',
  'parentFolderId',
  'webLink',
  'bodyPreview',
];

// ---------------------------------------------------------------------------
// Projection helpers
// ---------------------------------------------------------------------------

/**
 * Collapses Graph's three-level recipient object to one string. `toRecipients`
 * on a ten-message page is otherwise ~60 lines of nested JSON saying nothing the
 * model could not read from `"Ada Lovelace <ada@example.com>"`.
 */
function formatAddress(recipient: Recipient | null | undefined): string | undefined {
  const emailAddress = recipient?.emailAddress;
  if (emailAddress === null || emailAddress === undefined) return undefined;

  const address = typeof emailAddress.address === 'string' ? emailAddress.address.trim() : '';
  const name = typeof emailAddress.name === 'string' ? emailAddress.name.trim() : '';

  if (address === '') return name === '' ? undefined : name;
  // Graph repeats the address as the display name when the sender set no name.
  if (name === '' || name.toLowerCase() === address.toLowerCase()) return address;
  return `${name} <${address}>`;
}

function formatAddresses(list: Recipient[] | null | undefined): string[] | undefined {
  if (!Array.isArray(list) || list.length === 0) return undefined;
  const out = list
    .map((entry) => formatAddress(entry))
    .filter((entry): entry is string => entry !== undefined);
  return out.length === 0 ? undefined : out;
}

/** Copies `select` fields this projection does not already cover. */
function copyExtras(
  message: GraphMessage,
  extras: readonly string[],
  into: Record<string, unknown>,
): void {
  for (const field of extras) {
    const value = message[field];
    if (value !== undefined && value !== null) into[field] = value;
  }
}

/**
 * Graph property names are kept verbatim (`receivedDateTime`, not `received`)
 * so the model can reuse them in a follow-up `filter` or `orderby` argument.
 */
function summarizeMessage(
  message: GraphMessage,
  extras: readonly string[],
): Record<string, unknown> {
  const out: Record<string, unknown> = {
    id: message.id ?? undefined,
    subject: message.subject ?? undefined,
    from: formatAddress(message.from ?? message.sender),
    toRecipients: formatAddresses(message.toRecipients),
    receivedDateTime: message.receivedDateTime ?? undefined,
    isRead: message.isRead ?? undefined,
    hasAttachments: message.hasAttachments ?? undefined,
    // Graph caps bodyPreview at 255 characters, so it needs no truncation.
    bodyPreview: message.bodyPreview ?? undefined,
    webLink: message.webLink ?? undefined,
  };
  copyExtras(message, extras, out);
  return out;
}

function summarizeFolder(folder: GraphMailFolder, depth: number): Record<string, unknown> {
  const children = depth > 0 && Array.isArray(folder.childFolders) ? folder.childFolders : undefined;
  return {
    id: folder.id ?? undefined,
    displayName: folder.displayName ?? undefined,
    parentFolderId: folder.parentFolderId ?? undefined,
    unreadItemCount: folder.unreadItemCount ?? undefined,
    totalItemCount: folder.totalItemCount ?? undefined,
    childFolderCount: folder.childFolderCount ?? undefined,
    childFolders:
      children === undefined || children.length === 0
        ? undefined
        : children.map((child) => summarizeFolder(child, depth - 1)),
  };
}

function toRecipientList(addresses: readonly string[] | undefined): Recipient[] | undefined {
  if (addresses === undefined || addresses.length === 0) return undefined;
  return addresses.map((address) => ({ emailAddress: { address } }));
}

const DATETIME_PARTS =
  /^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?)(Z|[+-]\d{2}:\d{2})?)?$/;

/**
 * Graph rejects a bare date in a `$filter` on `receivedDateTime` — the literal
 * has to be a full datetime with a zone. `isoDate` happily accepts `2026-01-31`,
 * so widen a date-only bound to cover the whole day rather than silently
 * comparing against midnight at one end of the range.
 */
function graphDateTime(value: string, endOfDay: boolean): string {
  const iso = isoDate(value);
  const parts = DATETIME_PARTS.exec(iso);
  if (parts === null) return iso;

  const date = parts[1] ?? iso;
  const time = parts[2];
  const offset = parts[3];

  if (time === undefined) return `${date}T${endOfDay ? '23:59:59' : '00:00:00'}Z`;
  // Graph wants seconds; `isoDate` allows HH:mm.
  return `${date}T${time.length === 5 ? `${time}:00` : time}${offset ?? 'Z'}`;
}

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const listMessagesInput = z.object({
  folderId: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Restrict to one mail folder. Accepts a folder id from mail_list_folders or a ' +
        'well-known name: inbox, drafts, sentitems, deleteditems, archive, junkemail, ' +
        'outbox, clutter, conversationhistory, scheduled. Omit to list the whole mailbox.',
    ),
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .default(10)
    .describe('How many messages to return. Defaults to 10; 50 is the ceiling.'),
  skip: z
    .number()
    .int()
    .min(0)
    .max(10000)
    .optional()
    .describe('Messages to skip before returning results. Use it to page: 0, then 10, then 20.'),
  filter: z
    .string()
    .min(1)
    .max(1000)
    .optional()
    .describe(
      'Raw OData $filter, e.g. "importance eq \'high\'" or ' +
        '"receivedDateTime ge 2026-01-01T00:00:00Z". Combined with unreadOnly using and.',
    ),
  orderby: z
    .string()
    .regex(ORDERBY_CLAUSE, 'must be OData property names with optional asc/desc')
    .default('receivedDateTime desc')
    .describe('Sort order. Defaults to "receivedDateTime desc" (newest first).'),
  select: z
    .array(z.string().min(1).regex(SELECT_FIELD, 'must be a bare Microsoft Graph property name'))
    .min(1)
    .max(30)
    .optional()
    .describe(
      'Graph message property names to return instead of the default projection, e.g. ' +
        '["subject","importance","conversationId"]. id is always included. Do not ask ' +
        'for "body" here — use mail_get_message.',
    ),
  unreadOnly: z
    .boolean()
    .default(false)
    .describe('Return only unread messages (adds "isRead eq false" to the filter).'),
});

const searchMessagesInput = z.object({
  query: z
    .string()
    .trim()
    .min(1)
    .max(400)
    .describe(
      'Words to look for. Without a date range this is a KQL search across subject, ' +
        'body, and participants, so "subject:budget" or "quarterly review" both work. ' +
        'With after/before it degrades to a subject substring match — see the tool ' +
        'description.',
    ),
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .default(10)
    .describe('How many messages to return. Defaults to 10; 50 is the ceiling.'),
  from: emailArg.optional().describe('Only messages sent by this exact address.'),
  to: emailArg.optional().describe('Only messages addressed to this exact address.'),
  after: z
    .string()
    .min(4)
    .optional()
    .describe(
      'Only messages received at or after this ISO-8601 date or datetime, e.g. ' +
        '"2026-01-01" or "2026-01-01T09:00:00Z". A bare date starts at 00:00:00Z. ' +
        'Supplying this switches the tool to $filter mode.',
    ),
  before: z
    .string()
    .min(4)
    .optional()
    .describe(
      'Only messages received at or before this ISO-8601 date or datetime. A bare date ' +
        'ends at 23:59:59Z that day. Supplying this switches the tool to $filter mode.',
    ),
  hasAttachments: z.boolean().optional().describe('Only messages with (or without) attachments.'),
});

const getMessageInput = z.object({
  id: z.string().min(1).describe('The message id, as returned by mail_list_messages.'),
  format: z
    .enum(['text', 'html', 'preview'])
    .default('text')
    .describe(
      'How to return the body. "text" (default) asks Graph for plain text and strips any ' +
        'HTML that comes back anyway. "html" returns the raw markup, which costs several ' +
        'times more tokens. "preview" skips the body entirely and returns only the ' +
        "255-character bodyPreview — use it when you just need to confirm you have the " +
        'right message.',
    ),
  maxBodyChars: z
    .number()
    .int()
    .min(200)
    .max(100000)
    .default(8000)
    .describe(
      'Character budget for the body. Defaults to 8000, which covers most messages. The ' +
        "server's own output cap still applies on top of this.",
    ),
});

const listFoldersInput = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('How many top-level folders to return. Defaults to 50.'),
  includeChildren: z
    .boolean()
    .default(false)
    .describe(
      'Also return one level of subfolders under each top-level folder. Nesting deeper ' +
        'than one level needs a second call with that folder as the parent.',
    ),
});

const listAttachmentsInput = z.object({
  id: z.string().min(1).describe('The message id, as returned by mail_list_messages.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .default(20)
    .describe('How many attachments to return. Defaults to 20.'),
});

const draftFields = {
  subject: z.string().max(255).describe('Subject line.'),
  body: z.string().min(1).max(150000).describe('Message body, in the format given by contentType.'),
  cc: z.array(emailArg).max(50).optional().describe('Carbon-copy recipients.'),
  bcc: z.array(emailArg).max(50).optional().describe('Blind carbon-copy recipients.'),
  contentType: z
    .enum(['Text', 'HTML'])
    .default('Text')
    .describe('How to interpret body. "Text" (default) is safest; "HTML" renders markup.'),
};

const sendInput = z.object({
  to: z.array(emailArg).min(1).max(50).describe('Primary recipients. At least one is required.'),
  ...draftFields,
  saveToSentItems: z
    .boolean()
    .default(true)
    .describe('Keep a copy in Sent Items. Defaults to true.'),
});

const createDraftInput = z.object({
  to: z
    .array(emailArg)
    .max(50)
    .optional()
    .describe('Primary recipients. Optional — a draft may be saved without any.'),
  ...draftFields,
});

const replyInput = z.object({
  id: z.string().min(1).describe('Id of the message being replied to.'),
  comment: z
    .string()
    .min(1)
    .max(150000)
    .describe('Your reply text. It is placed above the quoted original as plain text.'),
  replyAll: z
    .boolean()
    .default(false)
    .describe('Reply to every recipient rather than only the sender. Defaults to false.'),
});

const forwardInput = z.object({
  id: z.string().min(1).describe('Id of the message being forwarded.'),
  to: z.array(emailArg).min(1).max(50).describe('Who to forward it to. At least one is required.'),
  comment: z.string().max(150000).optional().describe('Optional note above the forwarded message.'),
});

const moveInput = z.object({
  id: z.string().min(1).describe('Id of the message to move.'),
  destinationId: z
    .string()
    .min(1)
    .describe(
      'Target folder id from mail_list_folders, or a well-known name: inbox, archive, ' +
        'deleteditems, junkemail, drafts, sentitems.',
    ),
});

const markReadInput = z.object({
  id: z.string().min(1).describe('Id of the message to flag.'),
  isRead: z.boolean().default(true).describe('True marks it read, false marks it unread.'),
});

const deleteInput = z.object({
  id: z.string().min(1).describe('Id of the message to delete.'),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const mailModule: ToolModule = {
  group: GROUP,
  build({ graph, config }: ToolDeps): ToolDefinition[] {
    return [
      {
        name: 'mail_list_messages',
        title: 'List mail messages',
        description:
          'Lists Outlook messages newest-first, 10 per call by default (50 max), as a compact ' +
          'projection: id, subject, from, toRecipients, receivedDateTime, isRead, ' +
          'hasAttachments, bodyPreview and webLink. Bodies are never included — bodyPreview is ' +
          "Graph's first 255 characters, and mail_get_message returns the rest. Page with " +
          '`skip`. Note that Exchange rejects some filter/orderby combinations with ' +
          '"The restriction or sort order is too complex"; when that happens, drop `orderby` ' +
          'or sort on the same property you filtered.',
        inputSchema: listMessagesInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { folderId, top, skip, filter, orderby, select, unreadOnly } =
            listMessagesInput.parse(args);

          // Message ids and well-known folder names are safe characters, but a
          // folder id is caller-supplied and lands in a path segment.
          const path =
            folderId === undefined
              ? '/me/messages'
              : `/me/mailFolders/${encodeURIComponent(folderId)}/messages`;

          const clauses: string[] = [];
          if (unreadOnly) clauses.push('isRead eq false');
          // Parenthesised so a caller's `a or b` cannot swallow the unread clause.
          if (filter !== undefined) clauses.push(`(${filter})`);

          const fields = [...new Set(['id', ...(select ?? DEFAULT_MESSAGE_SELECT)])];

          const res = await graph.request({
            path,
            method: 'GET',
            query: buildQuery({
              select: fields,
              filter: clauses.length > 0 ? clauses.join(' and ') : undefined,
              orderby,
              top,
              skip,
            }),
            scopes: READ_SCOPES,
          });

          const messages = extractCollection<GraphMessage>(res.data);
          const extras = fields.filter((field) => !SUMMARY_FIELDS.has(field));

          return {
            count: messages.length,
            messages: messages.map((message) => summarizeMessage(message, extras)),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'mail_search_messages',
        title: 'Search mail messages',
        description:
          'Finds messages by text, returning the same compact projection as ' +
          'mail_list_messages, 10 per call by default (50 max). Two modes, because Graph ' +
          'will not combine them: with no date range it uses $search, a relevance-ranked ' +
          'KQL search over subject, body and participants (results are NOT in date order, ' +
          'and $search supports no date syntax). Passing `after` or `before` switches to ' +
          '$filter on receivedDateTime, which sorts newest-first but can only match `query` ' +
          'as a substring of the subject — body text is not searched in that mode. The ' +
          'response reports which mode ran. To search body text within a period, search ' +
          'first and filter the dates yourself.',
        inputSchema: searchMessagesInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { query, top, from, to, after, before, hasAttachments } =
            searchMessagesInput.parse(args);

          const fields = [...DEFAULT_MESSAGE_SELECT];
          const extras: string[] = [];
          const dateBounded = after !== undefined || before !== undefined;

          if (dateBounded) {
            const clauses: string[] = [];
            if (after !== undefined) {
              clauses.push(`receivedDateTime ge ${graphDateTime(after, false)}`);
            }
            if (before !== undefined) {
              clauses.push(`receivedDateTime le ${graphDateTime(before, true)}`);
            }
            if (from !== undefined) {
              clauses.push(`from/emailAddress/address eq '${escapeODataString(from)}'`);
            }
            if (to !== undefined) {
              clauses.push(
                `toRecipients/any(r:r/emailAddress/address eq '${escapeODataString(to)}')`,
              );
            }
            if (hasAttachments !== undefined) {
              clauses.push(`hasAttachments eq ${hasAttachments ? 'true' : 'false'}`);
            }
            clauses.push(`contains(subject,'${escapeODataString(query)}')`);

            const res = await graph.request({
              path: '/me/messages',
              method: 'GET',
              query: buildQuery({
                select: fields,
                filter: clauses.join(' and '),
                orderby: 'receivedDateTime desc',
                top,
              }),
              scopes: READ_SCOPES,
            });

            const messages = extractCollection<GraphMessage>(res.data);
            return {
              mode: 'filter',
              matchedOn: 'subject substring, because a date range was given',
              count: messages.length,
              messages: messages.map((message) => summarizeMessage(message, extras)),
              nextLink: res.nextLink,
            };
          }

          // KQL, not OData: the whole expression is one double-quoted $search
          // value, and Graph rejects $orderby and $skip alongside it.
          const terms = [query];
          if (from !== undefined) terms.push(`from:${from}`);
          if (to !== undefined) terms.push(`to:${to}`);
          if (hasAttachments !== undefined) terms.push(`hasattachment:${hasAttachments}`);

          const res = await graph.request({
            path: '/me/messages',
            method: 'GET',
            query: buildQuery({
              select: fields,
              search: quoteSearch(terms.join(' AND ')),
              top,
            }),
            scopes: READ_SCOPES,
          });

          const messages = extractCollection<GraphMessage>(res.data);
          return {
            mode: 'search',
            matchedOn: 'full-text KQL, ranked by relevance rather than date',
            count: messages.length,
            messages: messages.map((message) => summarizeMessage(message, extras)),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'mail_get_message',
        title: 'Get one mail message',
        description:
          'Returns one message with its body plus sender, recipients, timestamps, ' +
          'importance, conversationId and webLink. The body defaults to plain text and is ' +
          'capped at 8000 characters; raise maxBodyChars for a long thread, or pass ' +
          'format:"preview" to skip the body when you only need to identify the message. ' +
          'Attachment contents are never included — use mail_list_attachments. Reading a ' +
          'message here does not mark it read; use mail_mark_read for that.',
        inputSchema: getMessageInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, format, maxBodyChars } = getMessageInput.parse(args);
          const wantsBody = format !== 'preview';

          // Message ids are base64-ish and contain '/', '+' and '=' — every one
          // of which changes the URL's meaning if left unescaped.
          const res = await graph.request<GraphMessage>({
            path: `/me/messages/${encodeURIComponent(id)}`,
            method: 'GET',
            query: buildQuery({ select: wantsBody ? [...DETAIL_SELECT, 'body'] : DETAIL_SELECT }),
            // Asking Exchange to convert the body server-side beats stripping
            // tags locally: it resolves entities and inline styles properly.
            headers: wantsBody
              ? { Prefer: `outlook.body-content-type="${format}"` }
              : undefined,
            scopes: READ_SCOPES,
          });

          const message = res.data ?? {};
          const budget = Math.max(200, Math.min(maxBodyChars, config.maxOutputChars));

          let body: string | undefined;
          let bodyChars: number | undefined;
          if (wantsBody) {
            const raw = message.body?.content ?? '';
            // The Prefer header is advisory; some mailboxes return HTML anyway.
            const stillHtml = (message.body?.contentType ?? '').toLowerCase() === 'html';
            const text = format === 'html' || !stillHtml ? raw : stripHtml(raw);
            bodyChars = text.length;
            body = truncateText(text, budget);
          }

          return {
            id: message.id ?? undefined,
            subject: message.subject ?? undefined,
            from: formatAddress(message.from ?? message.sender),
            toRecipients: formatAddresses(message.toRecipients),
            ccRecipients: formatAddresses(message.ccRecipients),
            receivedDateTime: message.receivedDateTime ?? undefined,
            sentDateTime: message.sentDateTime ?? undefined,
            isRead: message.isRead ?? undefined,
            isDraft: message.isDraft ?? undefined,
            hasAttachments: message.hasAttachments ?? undefined,
            importance: message.importance ?? undefined,
            conversationId: message.conversationId ?? undefined,
            parentFolderId: message.parentFolderId ?? undefined,
            webLink: message.webLink ?? undefined,
            bodyFormat: format,
            bodyPreview: wantsBody ? undefined : (message.bodyPreview ?? undefined),
            body,
            bodyTruncated: bodyChars !== undefined && bodyChars > budget ? true : undefined,
          };
        },
      },
      {
        name: 'mail_list_folders',
        title: 'List mail folders',
        description:
          'Lists the top-level Outlook mail folders with their ids, unread counts and total ' +
          'counts, 50 per call by default. Pass includeChildren to also get one level of ' +
          'subfolders. Only top-level folders are returned otherwise, so a nested folder ' +
          'needs a second call. Use an id from here as folderId in mail_list_messages, ' +
          'though the well-known names (inbox, archive, sentitems) work there without ' +
          'looking anything up.',
        inputSchema: listFoldersInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { top, includeChildren } = listFoldersInput.parse(args);

          const res = await graph.request({
            path: '/me/mailFolders',
            method: 'GET',
            query: buildQuery({
              select: [
                'id',
                'displayName',
                'parentFolderId',
                'childFolderCount',
                'unreadItemCount',
                'totalItemCount',
              ],
              expand: includeChildren ? 'childFolders' : undefined,
              top,
            }),
            scopes: READ_SCOPES,
          });

          const folders = extractCollection<GraphMailFolder>(res.data);
          return {
            count: folders.length,
            folders: folders.map((folder) => summarizeFolder(folder, includeChildren ? 1 : 0)),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'mail_list_attachments',
        title: 'List message attachments',
        description:
          'Lists attachment metadata for one message — id, name, contentType, size in bytes, ' +
          'whether it is inline, and its kind (file, item, or reference) — 20 per call by ' +
          'default. File contents are deliberately never returned: a single attachment ' +
          'would exceed the output budget many times over. Check hasAttachments on the ' +
          'message first; inline images count as attachments, so a message with no visible ' +
          'attachment can still list several.',
        inputSchema: listAttachmentsInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, top } = listAttachmentsInput.parse(args);

          const res = await graph.request({
            path: `/me/messages/${encodeURIComponent(id)}/attachments`,
            method: 'GET',
            // This $select is load-bearing: without it Graph returns
            // `contentBytes`, the whole base64 file, for every fileAttachment.
            query: buildQuery({
              select: ['id', 'name', 'contentType', 'size', 'isInline', 'lastModifiedDateTime'],
              top,
            }),
            scopes: READ_SCOPES,
          });

          const attachments = extractCollection<GraphAttachment>(res.data);
          return {
            messageId: id,
            count: attachments.length,
            attachments: attachments.map((attachment) => ({
              id: attachment.id ?? undefined,
              name: attachment.name ?? undefined,
              contentType: attachment.contentType ?? undefined,
              size: attachment.size ?? undefined,
              isInline: attachment.isInline ?? undefined,
              kind: (attachment['@odata.type'] ?? '').replace('#microsoft.graph.', '') || undefined,
              lastModifiedDateTime: attachment.lastModifiedDateTime ?? undefined,
            })),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'mail_send',
        title: 'Send mail',
        description:
          'Sends a message immediately from the signed-in mailbox — there is no undo and no ' +
          'confirmation step, so confirm the recipients and text with the user first. ' +
          'Graph returns no message id for a send: if you need one, use mail_create_draft ' +
          'instead. A copy lands in Sent Items unless saveToSentItems is false. Sending ' +
          'needs the Mail.Send permission, which is separate from read/write access.',
        inputSchema: sendInput,
        write: true,
        scopes: SEND_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { to, subject, body, cc, bcc, contentType, saveToSentItems } = sendInput.parse(args);

          const res = await graph.request({
            path: '/me/sendMail',
            method: 'POST',
            body: {
              message: {
                subject,
                body: { contentType, content: body },
                toRecipients: toRecipientList(to),
                ccRecipients: toRecipientList(cc),
                bccRecipients: toRecipientList(bcc),
              },
              saveToSentItems,
            },
            scopes: SEND_SCOPES,
          });

          // sendMail answers 202 Accepted with an empty body.
          return { sent: true, status: res.status, to, cc, bcc, subject, saveToSentItems };
        },
      },
      {
        name: 'mail_create_draft',
        title: 'Create a mail draft',
        description:
          'Creates an unsent draft in the Drafts folder and returns its id and webLink. ' +
          'Nothing is transmitted — this tool cannot send, and there is no companion tool ' +
          'that sends an existing draft, so the user must send it from Outlook (or you can ' +
          'use mail_send instead). Recipients are optional on a draft.',
        inputSchema: createDraftInput,
        write: true,
        scopes: MUTATE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { to, subject, body, cc, bcc, contentType } = createDraftInput.parse(args);

          const res = await graph.request<GraphMessage>({
            path: '/me/messages',
            method: 'POST',
            body: {
              subject,
              body: { contentType, content: body },
              toRecipients: toRecipientList(to),
              ccRecipients: toRecipientList(cc),
              bccRecipients: toRecipientList(bcc),
            },
            scopes: MUTATE_SCOPES,
          });

          const draft = res.data ?? {};
          return {
            created: true,
            id: draft.id ?? undefined,
            subject: draft.subject ?? undefined,
            toRecipients: formatAddresses(draft.toRecipients),
            isDraft: draft.isDraft ?? undefined,
            createdDateTime: draft.createdDateTime ?? undefined,
            webLink: draft.webLink ?? undefined,
          };
        },
      },
      {
        name: 'mail_reply',
        title: 'Reply to a message',
        description:
          'Sends a reply to an existing message immediately — it does not create a draft, ' +
          'and there is no undo, so confirm the text with the user first. Your comment is ' +
          'placed above the quoted original as plain text; Graph builds the subject and ' +
          'recipients itself. replyAll includes every original recipient. Sending needs the ' +
          'Mail.Send permission, which is separate from read/write access.',
        inputSchema: replyInput,
        write: true,
        scopes: SEND_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, comment, replyAll } = replyInput.parse(args);

          const res = await graph.request({
            path: `/me/messages/${encodeURIComponent(id)}/${replyAll ? 'replyAll' : 'reply'}`,
            method: 'POST',
            body: { comment },
            scopes: SEND_SCOPES,
          });

          return { replied: true, status: res.status, id, replyAll };
        },
      },
      {
        name: 'mail_forward',
        title: 'Forward a message',
        description:
          'Forwards an existing message immediately, attachments included — it does not ' +
          'create a draft and there is no undo, so confirm the recipients with the user ' +
          'first. Your comment, if any, is placed above the forwarded content. Sending ' +
          'needs the Mail.Send permission, which is separate from read/write access.',
        inputSchema: forwardInput,
        write: true,
        scopes: SEND_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, to, comment } = forwardInput.parse(args);

          const res = await graph.request({
            path: `/me/messages/${encodeURIComponent(id)}/forward`,
            method: 'POST',
            body: { comment: comment ?? '', toRecipients: toRecipientList(to) },
            scopes: SEND_SCOPES,
          });

          return { forwarded: true, status: res.status, id, to };
        },
      },
      {
        name: 'mail_move_message',
        title: 'Move a message to a folder',
        description:
          'Moves a message to another mail folder. The move assigns the message a NEW id — ' +
          'the old one stops resolving, so use the id returned here for any follow-up call. ' +
          'destinationId accepts a folder id from mail_list_folders or a well-known name ' +
          'such as archive or deleteditems; moving to deleteditems is how you soft-delete ' +
          'without mail_delete_message.',
        inputSchema: moveInput,
        write: true,
        scopes: MUTATE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, destinationId } = moveInput.parse(args);

          const res = await graph.request<GraphMessage>({
            path: `/me/messages/${encodeURIComponent(id)}/move`,
            method: 'POST',
            body: { destinationId },
            scopes: MUTATE_SCOPES,
          });

          const moved = res.data ?? {};
          return {
            moved: true,
            previousId: id,
            id: moved.id ?? undefined,
            subject: moved.subject ?? undefined,
            parentFolderId: moved.parentFolderId ?? undefined,
            webLink: moved.webLink ?? undefined,
          };
        },
      },
      {
        name: 'mail_mark_read',
        title: 'Mark a message read or unread',
        description:
          'Sets the read flag on one message. Defaults to marking it read; pass isRead:false ' +
          'to mark it unread again. Reading a message with mail_get_message does not change ' +
          'this flag, so unread counts stay accurate until you call this.',
        inputSchema: markReadInput,
        write: true,
        scopes: MUTATE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, isRead } = markReadInput.parse(args);

          const res = await graph.request<GraphMessage>({
            path: `/me/messages/${encodeURIComponent(id)}`,
            method: 'PATCH',
            // PATCH echoes the whole message back; trim it to the three fields
            // worth reporting rather than paying for a second body.
            query: buildQuery({ select: ['id', 'subject', 'isRead'] }),
            body: { isRead },
            scopes: MUTATE_SCOPES,
          });

          const updated = res.data ?? {};
          return {
            updated: true,
            id: updated.id ?? id,
            subject: updated.subject ?? undefined,
            isRead: updated.isRead ?? isRead,
          };
        },
      },
      {
        name: 'mail_delete_message',
        title: 'Delete a message',
        description:
          'Deletes a message. This is a soft delete: the message moves to Deleted Items and ' +
          'the user can restore it from Outlook, but this server has no tool to undo it, so ' +
          'confirm with the user first. Deleting a message already in Deleted Items removes ' +
          'it permanently.',
        inputSchema: deleteInput,
        write: true,
        scopes: MUTATE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id } = deleteInput.parse(args);

          const res = await graph.request({
            path: `/me/messages/${encodeURIComponent(id)}`,
            method: 'DELETE',
            scopes: MUTATE_SCOPES,
          });

          // DELETE answers 204 No Content.
          return { deleted: true, status: res.status, id };
        },
      },
    ];
  },
};
