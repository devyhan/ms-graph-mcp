/**
 * The `chat` group: Microsoft Teams one-to-one, group, and meeting chats.
 *
 * Channel conversations are deliberately not here. A channel message is a
 * different Graph resource behind ChannelMessage.Read.All, which needs tenant
 * admin consent; the `teams` group owns it so this group stays user-consentable.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { createPaginator } from '../graph/paginate.js';
import { isoDate } from '../util/odata.js';
import { stripHtml, truncateText } from '../util/truncate.js';

const GROUP = GROUPS['chat']!;

/**
 * Copied, not aliased: a tool's `scopes` array feeds the consent set the login
 * path builds, and handing out the live `GROUPS` array would let a caller that
 * sorts or dedupes in place mutate the group catalogue.
 */
const READ_SCOPES = [...GROUP.readScopes];

/**
 * Chat.ReadWrite is what actually authorises the POST; Chat.Read rides along
 * because every write tool in this server requests the group's read scopes too.
 */
const WRITE_SCOPES = [...new Set([...GROUP.readScopes, ...GROUP.writeScopes])];

/** Graph caps `$top` at 50 on both collections here, and 400s above it rather than clamping. */
const MAX_TOP = 50;

/** Per-message body budget. The worst case for one call is 50 x this. */
const BODY_CHARS = 1000;

/** Teams rejects a chatMessage whose body exceeds 28 KB. */
const MAX_MESSAGE_CHARS = 28000;

/**
 * Teams allows one request per second per app per tenant against a given chat.
 * 1100 ms is that limit plus headroom, so a long walk never earns a 429 that
 * would cost more time in backoff than the pacing costs outright.
 */
const HISTORY_MIN_INTERVAL_MS = 1100;

/**
 * Wall-clock ceiling for one history walk. At the pacing above this is roughly
 * 54 requests, so `maxPages` is what actually ends a large walk; the budget is
 * the backstop for when Graph itself is slow.
 */
const HISTORY_BUDGET_MS = 60000;

/** 25 paced pages is about 28 seconds of request time — the most one call should spend. */
const HISTORY_MAX_PAGES = 25;

const MAX_HISTORY_MESSAGES = 1000;

/** Per-message body budget for a walk. The worst case for one call is maxMessages x this. */
const MAX_HISTORY_BODY_CHARS = 2000;

// ---------------------------------------------------------------------------
// Graph payload shapes (only the parts these tools project)
// ---------------------------------------------------------------------------

interface ConversationMemberPayload {
  id?: string | null;
  displayName?: string | null;
  email?: string | null;
  userId?: string | null;
  roles?: string[] | null;
  visibleHistoryStartDateTime?: string | null;
}

interface ChatPayload {
  id?: string | null;
  topic?: string | null;
  chatType?: string | null;
  createdDateTime?: string | null;
  lastUpdatedDateTime?: string | null;
  webUrl?: string | null;
  viewpoint?: {
    isHidden?: boolean | null;
    lastMessageReadDateTime?: string | null;
  } | null;
  onlineMeetingInfo?: { joinWebUrl?: string | null; calendarEventId?: string | null } | null;
  members?: ConversationMemberPayload[] | null;
}

interface ChatMessagePayload {
  id?: string | null;
  createdDateTime?: string | null;
  lastEditedDateTime?: string | null;
  deletedDateTime?: string | null;
  messageType?: string | null;
  importance?: string | null;
  from?: {
    user?: { id?: string | null; displayName?: string | null } | null;
    application?: { displayName?: string | null } | null;
  } | null;
  body?: { content?: string | null; contentType?: string | null } | null;
  attachments?: unknown[] | null;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Turns a chat id into one safe path segment.
 *
 * Teams ids look like `19:...@thread.v2`. Both `:` and `@` are legal inside a
 * path segment, but the id arrives from a model, so it is percent-encoded to
 * guarantee it can never contribute a segment of its own. `/`, `\`, `%` and
 * whitespace are rejected outright instead of encoded: no real chat id contains
 * them, and encoding them would emit `%2F` / `%25`, which the Graph path
 * validator refuses anyway — better a clear error here than a confusing one there.
 */
function chatSegment(raw: string): string {
  const id = raw.trim();
  if (id.length === 0) {
    throw new Error('chatId must be a non-empty Teams chat id, e.g. "19:...@thread.v2".');
  }
  if (/[/\\%\s]/.test(id)) {
    throw new Error(
      `Invalid chatId ${JSON.stringify(raw)}: pass the id exactly as chat_list_chats returned it, ` +
        'e.g. "19:0d1f...@thread.v2". It must not contain a slash, a percent sign, or whitespace.',
    );
  }
  return encodeURIComponent(id);
}

function projectMember(member: ConversationMemberPayload): Record<string, unknown> {
  return {
    // The membership id addresses the member *inside this chat*; userId is the
    // directory object id. They are not interchangeable.
    membershipId: member.id ?? undefined,
    displayName: member.displayName ?? undefined,
    email: member.email ?? undefined,
    userId: member.userId ?? undefined,
    roles: member.roles !== null && member.roles !== undefined && member.roles.length > 0
      ? member.roles
      : undefined,
  };
}

function projectChat(chat: ChatPayload, includeMembers: boolean): Record<string, unknown> {
  const members = chat.members;
  return {
    id: chat.id ?? undefined,
    // Always null for one-to-one chats — the member list is the only name they have.
    topic: chat.topic ?? undefined,
    chatType: chat.chatType ?? undefined,
    createdDateTime: chat.createdDateTime ?? undefined,
    lastUpdatedDateTime: chat.lastUpdatedDateTime ?? undefined,
    members:
      includeMembers && Array.isArray(members) ? members.map(projectMember) : undefined,
  };
}

function projectMessage(msg: ChatMessagePayload): Record<string, unknown> {
  const body = msg.body;
  const raw = body?.content ?? '';
  // Teams stores nearly every message as HTML even when the user typed plain text.
  const text = body?.contentType === 'html' ? stripHtml(raw) : raw.trim();
  const attachments = Array.isArray(msg.attachments) ? msg.attachments.length : 0;
  const messageType = msg.messageType ?? undefined;
  const importance = msg.importance ?? undefined;

  return {
    id: msg.id ?? undefined,
    createdDateTime: msg.createdDateTime ?? undefined,
    // `from` is null on system event messages (joins, leaves, meeting start).
    from: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? undefined,
    body: text.length > 0 ? truncateText(text, BODY_CHARS) : undefined,
    attachments: attachments > 0 ? attachments : undefined,
    // Only the non-default values are worth a key, so rows stay short.
    messageType: messageType !== undefined && messageType !== 'message' ? messageType : undefined,
    importance: importance !== undefined && importance !== 'normal' ? importance : undefined,
    deleted: msg.deletedDateTime ? true : undefined,
    edited: msg.lastEditedDateTime ? true : undefined,
  };
}

const DATETIME_PARTS =
  /^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?)(Z|[+-]\d{2}:\d{2})?)?$/;

/**
 * Graph rejects a bare date as a `$filter` literal on createdDateTime — the
 * literal has to be a full datetime with a zone. `isoDate` accepts `2026-01-31`,
 * so widen a date-only bound to the edge of the day it names rather than
 * silently comparing against midnight at one end of the range.
 */
function graphInstant(value: string, endOfDay: boolean): string {
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

/** Joins, leaves and meeting-start rows carry a messageType other than 'message'. */
function isOrdinaryMessage(msg: ChatMessagePayload): boolean {
  return (msg.messageType ?? 'message') === 'message';
}

/**
 * Builds the walk's stop predicate. `since` cannot be pushed to Graph: the
 * documented `$filter` on createdDateTime supports `lt` only, so the lower bound
 * of the window is enforced here, on a collection that is already ordered newest
 * first.
 */
function olderThan(thresholdMs: number): (msg: ChatMessagePayload) => boolean {
  return (msg) => {
    const created = msg.createdDateTime;
    if (typeof created !== 'string') return false;
    const parsed = Date.parse(created);
    // An unreadable timestamp must not end the walk; skipping it is recoverable,
    // truncating the caller's window at it is not.
    return !Number.isNaN(parsed) && parsed < thresholdMs;
  };
}

/**
 * Like `projectMessage`, but with a caller-chosen body budget, and without the
 * `importance` and `edited` flags: across hundreds of rows those cost more of the
 * output budget than they inform a reading of the conversation.
 */
function projectHistoryMessage(msg: ChatMessagePayload, bodyChars: number): Record<string, unknown> {
  const body = msg.body;
  const raw = body?.content ?? '';
  const text = body?.contentType === 'html' ? stripHtml(raw) : raw.trim();
  const attachments = Array.isArray(msg.attachments) ? msg.attachments.length : 0;
  const messageType = msg.messageType ?? undefined;

  return {
    id: msg.id ?? undefined,
    createdDateTime: msg.createdDateTime ?? undefined,
    // Null rather than absent: a system event genuinely has no sender, and on a
    // bodyChars=0 timeline that distinction carries most of the row's meaning.
    from: msg.from?.user?.displayName ?? msg.from?.application?.displayName ?? null,
    body: bodyChars > 0 && text.length > 0 ? truncateText(text, bodyChars) : undefined,
    // The count only. Attachment contents are a separate fetch and would swamp a walk.
    attachments: attachments > 0 ? attachments : undefined,
    // Elided when ordinary, which is nearly every row here, so a present
    // messageType always means an event row.
    messageType: messageType !== undefined && messageType !== 'message' ? messageType : undefined,
    deleted: msg.deletedDateTime ? true : undefined,
  };
}

/** The createdDateTime of the oldest and newest message actually returned. */
function historyRange(messages: readonly ChatMessagePayload[]): {
  from: string | undefined;
  to: string | undefined;
} {
  let oldestMs = Number.POSITIVE_INFINITY;
  let newestMs = Number.NEGATIVE_INFINITY;
  let from: string | undefined;
  let to: string | undefined;

  for (const msg of messages) {
    const created = msg.createdDateTime;
    if (typeof created !== 'string') continue;
    const parsed = Date.parse(created);
    if (Number.isNaN(parsed)) continue;
    if (parsed < oldestMs) {
      oldestMs = parsed;
      from = created;
    }
    if (parsed > newestMs) {
      newestMs = parsed;
      to = created;
    }
  }

  return { from, to };
}

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const chatIdField = z
  .string()
  .min(1)
  .describe(
    'The Teams chat id, exactly as chat_list_chats returned it — a thread id such as ' +
      '"19:0d1f9c...@thread.v2". This is not a user id, an email address, or a team id.',
  );

const listChatsInput = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(MAX_TOP)
    .default(20)
    .describe('How many chats to return, 1-50 (default 20). Graph rejects values above 50.'),
  filter: z
    .string()
    .min(1)
    .max(300)
    .optional()
    .describe(
      "OData $filter over the chat resource. Only a narrow set works: chatType eq 'oneOnOne', " +
        "chatType eq 'group', chatType eq 'meeting', and lastUpdatedDateTime gt " +
        '2026-01-01T00:00:00Z. Topic text and member names are NOT filterable and there is no ' +
        '$search on chats. Double any apostrophe inside a quoted literal.',
    ),
  expandMembers: z
    .boolean()
    .default(false)
    .describe(
      "Also return each chat's members ($expand=members). Turn this on to identify one-to-one " +
        'chats, whose topic is always null. It multiplies the output size, so leave it off when ' +
        'listing many chats. Graph rejects $expand=members alongside some $filter expressions; ' +
        'drop the filter if the call fails.',
    ),
});

const getChatInput = z.object({ chatId: chatIdField });

const listMessagesInput = z.object({
  chatId: chatIdField,
  top: z
    .number()
    .int()
    .min(1)
    .max(MAX_TOP)
    .default(20)
    .describe('How many messages to return, 1-50 (default 20). Graph rejects values above 50.'),
  orderby: z
    .enum(['createdDateTime desc', 'lastModifiedDateTime desc'])
    .default('createdDateTime desc')
    .describe(
      'Sort order. Graph supports descending order only on this collection, so the newest ' +
        'message is always first; ascending forms are rejected with 400.',
    ),
});

const listMembersInput = z.object({ chatId: chatIdField });

const sendMessageInput = z.object({
  chatId: chatIdField,
  content: z
    .string()
    .min(1)
    .max(MAX_MESSAGE_CHARS)
    .describe(
      'The message text. Posted immediately as the signed-in user and visible to everyone in ' +
        'the chat; there is no tool here to edit or delete it afterwards. Teams rejects bodies ' +
        'over 28 KB.',
    ),
  contentType: z
    .enum(['text', 'html'])
    .default('text')
    .describe(
      'How to interpret `content`. Use "text" unless formatting is needed; "html" accepts only ' +
        "Teams' limited subset (b, i, u, a, ul, ol, li, br, p, pre, code, blockquote). An <at> " +
        'tag typed into the HTML does NOT notify anyone, because a real @mention needs a ' +
        'separate mentions array this tool does not send.',
    ),
});

const fetchHistoryInput = z.object({
  chatId: chatIdField,
  since: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Lower bound of the window: stop as soon as a message older than this ISO-8601 date or ' +
        'datetime is reached, e.g. "2026-08-01" or "2026-08-01T09:00:00Z". A bare date means the ' +
        'start of that day, UTC. Enforced while reading rather than by Graph, so a distant `since` ' +
        'still costs a request per 50 messages. Omit to walk back until another limit ends the call.',
    ),
  until: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Upper bound of the window: start from messages strictly older than this ISO-8601 date or ' +
        'datetime. A bare date means the end of that day, UTC. Omit to start from the newest ' +
        'message. Prefer `cursor` over `until` when continuing a previous call — the cursor resumes ' +
        'exactly where the walk stopped, while `until` re-runs the query from a timestamp.',
    ),
  maxMessages: z
    .number()
    .int()
    .min(1)
    .max(MAX_HISTORY_MESSAGES)
    .default(200)
    .describe(
      `How many messages to walk back, 1-${MAX_HISTORY_MESSAGES} (default 200), rounded up to a ` +
        'whole page of 50 — asking for 175 can return up to 200. Graph returns 50 per request and ' +
        'this tool waits about a second between requests, so 200 takes roughly 5 seconds and 1000 ' +
        'roughly half a minute. Ask for what you will actually read: a large value spends real ' +
        'time and can still be cut short by the output limit.',
    ),
  bodyChars: z
    .number()
    .int()
    .min(0)
    .max(MAX_HISTORY_BODY_CHARS)
    .default(400)
    .describe(
      `How much of each message body to keep, 0-${MAX_HISTORY_BODY_CHARS} characters (default 400). ` +
        'HTML is stripped first and longer bodies are cut with a marker. Use 0 to omit bodies ' +
        'entirely and get a who-and-when timeline, which is what makes a 1000-message walk fit.',
    ),
  includeSystem: z
    .boolean()
    .default(false)
    .describe(
      'Include system event messages — someone joined or left, a meeting started or ended. They ' +
        'have no body and no sender and are identified by messageType. Off by default because they ' +
        'dominate meeting chats and say nothing about the conversation.',
    ),
  cursor: z
    .string()
    .min(1)
    .optional()
    .describe(
      'The `cursor` returned by a previous chat_fetch_history call, to continue further back from ' +
        'exactly where that call stopped. Pass the same chatId, `since`, `bodyChars` and ' +
        '`includeSystem` alongside it: the cursor carries the server-side query, but the window ' +
        'bound and the projection are applied here and are not baked into it. Omit to start from ' +
        'the newest message.',
    ),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const chatModule: ToolModule = {
  group: GROUP,
  build({ graph }: ToolDeps): ToolDefinition[] {
    // Stateless factory: each walk carries its own pacing, budget and cursor.
    const paginator = createPaginator(graph);

    return [
      {
        name: 'chat_list_chats',
        title: 'List my chats',
        description:
          'Lists the Teams chats the signed-in user belongs to (one-to-one, group, and meeting ' +
          'chats — never channel conversations, which the teams group covers). Returns id, topic, ' +
          'chatType, createdDateTime and lastUpdatedDateTime for up to `top` chats (default 20), ' +
          'plus a count and a nextLink when more exist. Two traps: the order is unspecified, so ' +
          'this is NOT a "most recent chats" list — filter or sort on lastUpdatedDateTime ' +
          'yourself; and one-to-one chats always have a null topic, so pass expandMembers to see ' +
          'who they are with.',
        inputSchema: listChatsInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { top, filter, expandMembers } = listChatsInput.parse(args);
          const res = await graph.request({
            path: '/me/chats',
            method: 'GET',
            query: {
              $top: top,
              $filter: filter,
              $expand: expandMembers ? 'members' : undefined,
            },
            scopes: READ_SCOPES,
          });
          const chats = extractCollection<ChatPayload>(res.data);
          return {
            count: chats.length,
            nextLink: res.nextLink,
            chats: chats.map((chat) => projectChat(chat, expandMembers)),
          };
        },
      },
      {
        name: 'chat_get_chat',
        title: 'Get a chat',
        description:
          'Returns one Teams chat by id: topic, chatType, created and last-updated timestamps, ' +
          'the Teams deep link, whether the signed-in user has hidden it, how far they have read, ' +
          'and the meeting join URL for a meeting chat. Members are not included — call ' +
          'chat_list_members for those, which is the only way to name a one-to-one chat since its ' +
          'topic is always null. A chat the signed-in user is not a member of returns 403.',
        inputSchema: getChatInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { chatId } = getChatInput.parse(args);
          const res = await graph.request<ChatPayload>({
            path: `/chats/${chatSegment(chatId)}`,
            method: 'GET',
            scopes: READ_SCOPES,
          });
          const chat: ChatPayload = res.data ?? {};
          return {
            id: chat.id ?? undefined,
            topic: chat.topic ?? undefined,
            chatType: chat.chatType ?? undefined,
            createdDateTime: chat.createdDateTime ?? undefined,
            lastUpdatedDateTime: chat.lastUpdatedDateTime ?? undefined,
            webUrl: chat.webUrl ?? undefined,
            // viewpoint is the signed-in user's own view of the chat, not shared state.
            isHidden: chat.viewpoint?.isHidden ?? undefined,
            lastMessageReadDateTime: chat.viewpoint?.lastMessageReadDateTime ?? undefined,
            onlineMeetingJoinUrl: chat.onlineMeetingInfo?.joinWebUrl ?? undefined,
          };
        },
      },
      {
        name: 'chat_list_messages',
        title: 'List chat messages',
        description:
          'Reads recent messages from one Teams chat, newest first. Each row is a compact ' +
          'projection: id, createdDateTime, the sender display name, the body stripped of HTML ' +
          `and truncated to ${BODY_CHARS} characters, and the number of attachments. Returns up ` +
          'to `top` messages (default 20) plus a count and a nextLink when older messages exist. ' +
          'System event messages (someone joined or left, a meeting started) carry no body and no ' +
          'sender — they are marked with messageType. Deleted messages still appear, flagged ' +
          'deleted and with the body gone. Attachment contents, reactions, and @mention targets ' +
          'are not returned. There is no $search and no date filter on this collection: to reach ' +
          'older messages, page with nextLink.',
        inputSchema: listMessagesInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { chatId, top, orderby } = listMessagesInput.parse(args);
          const res = await graph.request({
            path: `/chats/${chatSegment(chatId)}/messages`,
            method: 'GET',
            query: { $top: top, $orderby: orderby },
            scopes: READ_SCOPES,
          });
          const messages = extractCollection<ChatMessagePayload>(res.data);
          return {
            chatId,
            count: messages.length,
            nextLink: res.nextLink,
            messages: messages.map(projectMessage),
          };
        },
      },
      {
        name: 'chat_list_members',
        title: 'List chat members',
        description:
          'Lists everyone in one Teams chat: the per-chat membership id, display name, email, ' +
          'directory user id, and roles (an owner of a group chat has ["owner"]; ordinary ' +
          'members have none). This is how a one-to-one chat gets a name, since its topic is ' +
          'always null. Guests and anonymous meeting participants come back with a null email ' +
          'and sometimes a null userId, so match on displayName for those.',
        inputSchema: listMembersInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { chatId } = listMembersInput.parse(args);
          const res = await graph.request({
            path: `/chats/${chatSegment(chatId)}/members`,
            method: 'GET',
            scopes: READ_SCOPES,
          });
          const members = extractCollection<ConversationMemberPayload>(res.data);
          return {
            chatId,
            count: members.length,
            nextLink: res.nextLink,
            members: members.map(projectMember),
          };
        },
      },
      {
        name: 'chat_send_message',
        title: 'Send a chat message',
        description:
          'Posts a message to a Teams chat as the signed-in user. This is immediate and public ' +
          'to the chat, and nothing here can edit or delete it afterwards, so confirm the ' +
          'recipient chat id and the wording before calling. Teams throttles sending to roughly ' +
          'one message per second per user and will return 429 on a burst: send one message, not ' +
          'a loop over many chats. Returns the new message id and its createdDateTime. The ' +
          'signed-in user must already be a member of the chat; otherwise Graph returns 403.',
        inputSchema: sendMessageInput,
        write: true,
        scopes: WRITE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { chatId, content, contentType } = sendMessageInput.parse(args);
          const res = await graph.request<ChatMessagePayload>({
            path: `/chats/${chatSegment(chatId)}/messages`,
            method: 'POST',
            body: { body: { content, contentType } },
            scopes: WRITE_SCOPES,
          });
          const created: ChatMessagePayload = res.data ?? {};
          return {
            id: created.id ?? undefined,
            chatId,
            createdDateTime: created.createdDateTime ?? undefined,
            contentType,
          };
        },
      },
      {
        name: 'chat_fetch_history',
        title: 'Fetch chat history',
        description:
          'Reads a long run of one Teams chat, newest first, paging automatically so you do not ' +
          'have to follow nextLink yourself. Use this instead of repeated chat_list_messages calls ' +
          'whenever you need more than one page — summarising a conversation, finding when ' +
          'something was decided, reconstructing a thread. It paces itself at about one request ' +
          'per second because that is the rate Teams allows per chat, and each request returns 50 ' +
          'messages, so maxMessages 200 takes roughly 5 seconds and 1000 roughly half a minute; ' +
          'ask only for what you will read. Bound the window with `since` (stop once messages ' +
          'older than this are reached) and `until` (start below this timestamp). Each row is id, ' +
          'createdDateTime, the sender display name — null on system events, the app name on bot ' +
          'posts — the body stripped of HTML and truncated to bodyChars, and an attachment count; ' +
          'set bodyChars to 0 for a bodyless timeline. Deleted messages still appear, flagged ' +
          'deleted with the body gone. Raw HTML, reactions, @mention targets and attachment ' +
          'contents are never returned, and there is no text search on this collection. `range` ' +
          'reports the createdDateTime of the oldest and newest message actually returned — read ' +
          'it to confirm you got the window you asked for. `reason` says why the walk ended: ' +
          '"complete" means the chat has no older messages, "stopped" means `since` was reached ' +
          'and the window is fully covered, while "maxItems", "maxPages" and "budget" mean history ' +
          'remains — in those three cases call this tool again with the returned `cursor` to ' +
          'continue further back.',
        inputSchema: fetchHistoryInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { chatId, since, until, maxMessages, bodyChars, includeSystem, cursor } =
            fetchHistoryInput.parse(args);

          const sinceMs = since === undefined ? undefined : Date.parse(graphInstant(since, false));
          if (sinceMs !== undefined && Number.isNaN(sinceMs)) {
            throw new Error(
              `Invalid since value ${JSON.stringify(since)}: expected an ISO-8601 date or datetime, ` +
                'e.g. "2026-08-01" or "2026-08-01T09:00:00Z".',
            );
          }

          const walk = await paginator.walk<ChatMessagePayload>({
            path: `/chats/${chatSegment(chatId)}/messages`,
            query: {
              $top: MAX_TOP,
              // Fixed, not a parameter, for two reasons. It is the stable ordering:
              // lastModifiedDateTime changes when a reaction is added or removed, so a
              // backward walk ordered by it can duplicate or skip rows as old messages
              // jump to the head (an inference from the documented behaviour of
              // lastModifiedDateTime, not a documented statement). And it is a hard
              // precondition for the $filter below — Graph silently IGNORES a $filter
              // whose property does not match $orderby, so a mismatch here would return
              // the newest messages while looking like it honoured `until`.
              $orderby: 'createdDateTime desc',
              $filter:
                until === undefined
                  ? undefined
                  : `createdDateTime lt ${graphInstant(until, true)}`,
            },
            scopes: READ_SCOPES,
            // Rounded up to whole pages on purpose. The walk cuts a page short when
            // this cap is reached mid-page, but the cursor it hands back points past
            // that whole page — so a cap that can land mid-page would silently skip
            // the rest of it on resume. Keeping the cap on a page boundary makes the
            // cursor an exact continuation; the cost is at most 49 rows more than
            // asked for.
            maxItems: Math.ceil(maxMessages / MAX_TOP) * MAX_TOP,
            // Deliberately the page count maxMessages implies, not more: a chat thick
            // with system events therefore ends on 'maxPages' with fewer rows than
            // asked for, which the cursor resumes, rather than silently spending the
            // whole minute filtering.
            maxPages: Math.min(Math.ceil(maxMessages / MAX_TOP), HISTORY_MAX_PAGES),
            minIntervalMs: HISTORY_MIN_INTERVAL_MS,
            budgetMs: HISTORY_BUDGET_MS,
            cursor,
            keep: includeSystem ? undefined : isOrdinaryMessage,
            stop: sinceMs === undefined ? undefined : olderThan(sinceMs),
          });

          const range = historyRange(walk.items);
          return {
            chatId,
            count: walk.items.length,
            pages: walk.pages,
            reason: walk.reason,
            range,
            // Only meaningful while history remains; a completed walk has nothing to resume.
            cursor: walk.reason === 'complete' ? undefined : walk.nextLink,
            messages: walk.items.map((msg) => projectHistoryMessage(msg, bodyChars)),
          };
        },
      },
    ];
  },
};
