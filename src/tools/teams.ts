/**
 * The `teams` group: Microsoft Teams teams, channels, channel messages, and
 * team rosters.
 *
 * Every scope this group uses needs tenant admin consent, so the whole group is
 * gated behind `--org-mode` by the catalogue. The tool descriptions repeat that
 * fact because a 403 here is almost always a missing consent grant rather than
 * a bad argument, and the model needs to be able to say so.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { createPaginator } from '../graph/paginate.js';
import { isoDate } from '../util/odata.js';
import { stripHtml, truncateText } from '../util/truncate.js';

const GROUP = GROUPS['teams']!;

/**
 * Least privilege is per tool, not per group: listing channels does not need
 * ChannelMessage.Read.All, and reading messages does not need the channel
 * scope. The scope strings live only in the group catalogue, so they are
 * selected out of it by prefix rather than retyped here.
 */
function groupScope(prefix: string): string {
  const match = [...GROUP.readScopes, ...GROUP.writeScopes].find((scope) =>
    scope.startsWith(prefix),
  );
  if (match === undefined) {
    throw new Error(`The "teams" tool group meta is missing a ${prefix}* scope.`);
  }
  return match;
}

const TEAM_SCOPES = [groupScope('Team.ReadBasic')];
const CHANNEL_SCOPES = [groupScope('Channel.ReadBasic')];
const MESSAGE_READ_SCOPES = [groupScope('ChannelMessage.Read')];

/**
 * Sending is the one write here and it is genuinely write-only: the POST echoes
 * the created message back, so ChannelMessage.Send alone is enough. Adding the
 * group's read scopes would widen the token request for no gain, and in a
 * tenant that consented to Send but not ChannelMessage.Read.All it would turn a
 * working send into a consent failure.
 */
const MESSAGE_SEND_SCOPES = [groupScope('ChannelMessage.Send')];

/**
 * Not in the group meta on purpose. `GET /teams/{id}/members` is the only tool
 * here that reads a roster, and Graph gates that behind TeamMember.Read.All —
 * Team.ReadBasic.All does not cover it. Keeping it on the one tool that needs
 * it avoids asking every Teams user to consent to roster access.
 */
const TEAM_MEMBER_SCOPES = ['TeamMember.Read.All'];

// ---------------------------------------------------------------------------
// Graph payload shapes (only the parts these tools project)
// ---------------------------------------------------------------------------

interface IdentityRef {
  id?: string | null;
  displayName?: string | null;
}

interface IdentitySet {
  user?: (IdentityRef & { userIdentityType?: string | null }) | null;
  application?: IdentityRef | null;
  device?: IdentityRef | null;
}

interface TeamPayload {
  id?: string | null;
  displayName?: string | null;
  description?: string | null;
  visibility?: string | null;
  isArchived?: boolean | null;
  webUrl?: string | null;
}

interface ChannelPayload {
  id?: string | null;
  displayName?: string | null;
  description?: string | null;
  email?: string | null;
  webUrl?: string | null;
  membershipType?: string | null;
  createdDateTime?: string | null;
  isFavoriteByDefault?: boolean | null;
  isArchived?: boolean | null;
  tenantId?: string | null;
}

interface ChatMessagePayload {
  id?: string | null;
  replyToId?: string | null;
  etag?: string | null;
  messageType?: string | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  lastEditedDateTime?: string | null;
  deletedDateTime?: string | null;
  subject?: string | null;
  summary?: string | null;
  importance?: string | null;
  locale?: string | null;
  webUrl?: string | null;
  from?: IdentitySet | null;
  body?: { contentType?: string | null; content?: string | null } | null;
  attachments?:
    | Array<{
        id?: string | null;
        name?: string | null;
        contentType?: string | null;
        contentUrl?: string | null;
      }>
    | null;
  mentions?: Array<{ mentionText?: string | null; mentioned?: IdentitySet | null }> | null;
  reactions?: Array<{ reactionType?: string | null; user?: IdentitySet | null }> | null;
  channelIdentity?: { teamId?: string | null; channelId?: string | null } | null;
}

/**
 * `$expand=replies` hangs a page of the reply collection off each root message,
 * carrying its own paging annotation when the thread has more replies than Graph
 * is willing to inline.
 */
interface ChannelMessageWithReplies extends ChatMessagePayload {
  replies?: ChatMessagePayload[] | null;
  'replies@odata.nextLink'?: string | null;
}

interface ConversationMemberPayload {
  id?: string | null;
  displayName?: string | null;
  roles?: string[] | null;
  userId?: string | null;
  email?: string | null;
  tenantId?: string | null;
  visibleHistoryStartDateTime?: string | null;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Teams ids are pchar-safe in practice — a channel id looks like
 * `19:a1b2…@thread.tacv2` — and Graph is happier receiving `:` and `@`
 * literally, so those two are restored after encoding. Everything that could
 * change the shape of the URL (a slash, `?`, `#`, a stray `%`, whitespace) stays
 * percent-encoded, which is what keeps a hallucinated id from becoming an extra
 * path segment or a query string.
 */
function segment(id: string): string {
  return encodeURIComponent(id).replace(/%3A/gi, ':').replace(/%40/g, '@');
}

/** Graph message bodies are HTML far more often than not. */
function bodyText(
  body: { contentType?: string | null; content?: string | null } | null | undefined,
  maxChars: number,
): string | undefined {
  const content = body?.content;
  if (maxChars === 0 || typeof content !== 'string' || content.length === 0) {
    return undefined;
  }
  const text = body?.contentType === 'html' ? stripHtml(content) : content.trim();
  return text.length === 0 ? undefined : truncateText(text, maxChars);
}

function senderName(from: IdentitySet | null | undefined): string | undefined {
  return (
    from?.user?.displayName ??
    from?.application?.displayName ??
    from?.device?.displayName ??
    undefined
  );
}

/** Undefined members are dropped by JSON.stringify, which is the projection. */
function compactMessage(
  message: ChatMessagePayload,
  bodyChars: number,
): Record<string, unknown> {
  const attachments = message.attachments ?? [];
  const mentions = message.mentions ?? [];
  const reactions = message.reactions ?? [];

  return {
    id: message.id ?? undefined,
    createdDateTime: message.createdDateTime ?? undefined,
    // `from` is null on system event messages (joins, channel renames).
    from: senderName(message.from),
    fromUserId: message.from?.user?.id ?? undefined,
    // 'message' is the overwhelming majority; only the odd ones are worth naming.
    messageType:
      message.messageType && message.messageType !== 'message'
        ? message.messageType
        : undefined,
    importance:
      message.importance && message.importance !== 'normal'
        ? message.importance
        : undefined,
    subject: message.subject ?? undefined,
    replyToId: message.replyToId ?? undefined,
    body: bodyText(message.body, bodyChars),
    attachmentCount: attachments.length > 0 ? attachments.length : undefined,
    mentionCount: mentions.length > 0 ? mentions.length : undefined,
    reactionCount: reactions.length > 0 ? reactions.length : undefined,
    // Flags rather than timestamps: in a list, only the fact matters, and a
    // tombstone is the one case where an empty body is not a bug.
    deleted: message.deletedDateTime ? true : undefined,
    edited: message.lastEditedDateTime ? true : undefined,
    webUrl: message.webUrl ?? undefined,
  };
}

function compactChannel(channel: ChannelPayload): Record<string, unknown> {
  return {
    id: channel.id ?? undefined,
    displayName: channel.displayName ?? undefined,
    description: channel.description ?? undefined,
    membershipType: channel.membershipType ?? undefined,
    createdDateTime: channel.createdDateTime ?? undefined,
    webUrl: channel.webUrl ?? undefined,
  };
}

// ---------------------------------------------------------------------------
// Channel history walk
// ---------------------------------------------------------------------------

/** Documented `$top` maximum on the channel messages collection; its own default is 20. */
const HISTORY_PAGE_SIZE = 50;

/**
 * Teams allows one request per second per app per tenant against a given
 * channel, and that is the binding limit for a history walk. The extra 100ms is
 * headroom for clock skew, since sitting exactly on the limit collects 429s.
 */
const HISTORY_MIN_INTERVAL_MS = 1100;

/** Wall-clock ceiling for one walk, so a long history cannot hang the caller. */
const HISTORY_BUDGET_MS = 60_000;

/** At 1.1s per request 25 pages is already ~28s; the budget usually bites first. */
const HISTORY_MAX_PAGES = 25;

/**
 * Replies projected inline per root message. `$expand=replies` can return 200 of
 * them, and 200 replies on each of 200 messages would blow both this server's
 * output cap and the client's. The rest are reachable through
 * teams_list_message_replies.
 */
const INLINE_REPLY_LIMIT = 10;

/**
 * A history row is narrower than `compactMessage`: fields that are cheap on a
 * page of 20 are expensive on a walk of 1000, so webUrl and the mention and
 * reaction counts are dropped. `lastModifiedDateTime` is added instead, because
 * it is the property Graph orders this collection by and the only way for the
 * caller to see why a row arrived where it did.
 */
function historyMessage(message: ChatMessagePayload, bodyChars: number): Record<string, unknown> {
  const attachments = message.attachments ?? [];
  return {
    id: message.id ?? undefined,
    createdDateTime: message.createdDateTime ?? undefined,
    lastModifiedDateTime: message.lastModifiedDateTime ?? undefined,
    from: senderName(message.from),
    messageType:
      message.messageType && message.messageType !== 'message' ? message.messageType : undefined,
    body: bodyText(message.body, bodyChars),
    attachmentCount: attachments.length > 0 ? attachments.length : undefined,
    deleted: message.deletedDateTime ? true : undefined,
  };
}

function historyRoot(
  message: ChannelMessageWithReplies,
  bodyChars: number,
  includeReplies: boolean,
): Record<string, unknown> {
  const projected = historyMessage(message, bodyChars);
  if (!includeReplies) return projected;

  const replies = message.replies ?? [];
  // `replyCount` is what this response carried, not the thread's true length:
  // Graph inlines a bounded page and hands back its own nextLink for the rest.
  const more = replies.length > INLINE_REPLY_LIMIT || typeof message['replies@odata.nextLink'] === 'string';
  return {
    ...projected,
    replyCount: replies.length > 0 ? replies.length : undefined,
    replies:
      replies.length > 0
        ? replies.slice(0, INLINE_REPLY_LIMIT).map((reply) => historyMessage(reply, bodyChars))
        : undefined,
    moreReplies: more ? true : undefined,
  };
}

const ISO_OFFSET = /(?:Z|[+-]\d{2}:\d{2})$/;

/**
 * A window bound as epoch milliseconds. The comparison happens in this process,
 * not in a `$filter`, so the bound has to be an instant rather than the OData
 * literal the other tools build.
 */
function windowBound(value: string | undefined, endOfDay: boolean): number | undefined {
  if (value === undefined) return undefined;
  const iso = isoDate(value);

  // A date-only bound covers the whole day. Reading "until: 2026-01-31" as
  // midnight would silently drop everything the caller asked for that day.
  const instant = !iso.includes('T')
    ? `${iso}T${endOfDay ? '23:59:59.999' : '00:00:00.000'}Z`
    : // A bare wall time is read as UTC, the zone Graph stamps createdDateTime in.
      ISO_OFFSET.test(iso)
      ? iso
      : `${iso}Z`;

  const ms = Date.parse(instant);
  if (Number.isNaN(ms)) {
    throw new Error(`Invalid ISO-8601 datetime: ${JSON.stringify(value)}.`);
  }
  return ms;
}

/** Oldest and newest creation time actually returned. */
interface CreatedRange {
  from: string;
  to: string;
}

/**
 * Scans for the extremes rather than reading the first and last rows: the walk
 * arrives in reply-chain modification order, so the ends of the array are not
 * the ends of the date range.
 */
function createdRange(messages: ChatMessagePayload[]): CreatedRange | undefined {
  let from: string | undefined;
  let to: string | undefined;
  let fromMs = Number.POSITIVE_INFINITY;
  let toMs = Number.NEGATIVE_INFINITY;

  for (const message of messages) {
    const created = message.createdDateTime;
    if (typeof created !== 'string') continue;
    const ms = Date.parse(created);
    if (Number.isNaN(ms)) continue;
    if (ms < fromMs) {
      fromMs = ms;
      from = created;
    }
    if (ms > toMs) {
      toMs = ms;
      to = created;
    }
  }

  return from === undefined || to === undefined ? undefined : { from, to };
}

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const teamId = z
  .string()
  .min(1)
  .describe(
    'Team id (the group id GUID), as returned by teams_list_joined. Not the team display name.',
  );

const channelId = z
  .string()
  .min(1)
  .describe(
    'Channel id from teams_list_channels, e.g. "19:a1b2c3@thread.tacv2". Not the channel display name.',
  );

const messageId = z
  .string()
  .min(1)
  .describe(
    'Channel message id from teams_list_channel_messages. Graph uses the message ' +
      'creation timestamp in milliseconds as the id, e.g. "1737052800000".',
  );

const listJoinedInput = z.object({});

const listChannelsInput = z.object({
  teamId,
  membershipType: z
    .enum(['standard', 'private', 'shared'])
    .optional()
    .describe(
      'Return only channels of this kind. Omit for every channel the signed-in user can see. ' +
        'Private and shared channels are only listed when the user is a member of them.',
    ),
});

const getChannelInput = z.object({ teamId, channelId });

const listChannelMessagesInput = z.object({
  teamId,
  channelId,
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .default(20)
    .describe(
      'How many top-level messages to return, newest first. Graph caps this at 50 and ' +
        'throttles this endpoint near one request per second per channel, so keep it small ' +
        'and page with nextLink rather than retrying a large value.',
    ),
  bodyChars: z
    .number()
    .int()
    .min(0)
    .max(4000)
    .default(400)
    .describe(
      'Maximum characters of each message body to return after HTML is stripped to plain ' +
        'text. Use 0 to omit bodies and return only the message metadata.',
    ),
});

const getChannelMessageInput = z.object({
  teamId,
  channelId,
  messageId,
  bodyChars: z
    .number()
    .int()
    .min(0)
    .max(20000)
    .default(4000)
    .describe(
      'Maximum characters of the message body to return after HTML is stripped to plain text.',
    ),
});

const listRepliesInput = z.object({
  teamId,
  channelId,
  messageId: messageId.describe(
    'Id of the top-level message whose replies to list. A reply id is not accepted here: ' +
      'replies are one level deep, so pass the parent message id from teams_list_channel_messages.',
  ),
  top: z
    .number()
    .int()
    .min(1)
    .max(50)
    .default(20)
    .describe('How many replies to return, newest first. Graph caps this at 50.'),
  bodyChars: z
    .number()
    .int()
    .min(0)
    .max(4000)
    .default(400)
    .describe(
      'Maximum characters of each reply body to return after HTML is stripped to plain text. ' +
        'Use 0 to omit bodies.',
    ),
});

const fetchChannelHistoryInput = z.object({
  teamId,
  channelId,
  since: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Inclusive lower bound on createdDateTime, ISO-8601, e.g. "2026-01-01" or ' +
        '"2026-01-01T09:00:00Z". A bare date covers the whole day; a datetime with no offset ' +
        'is read as UTC. Filtered in this server, not by Graph, so it narrows the result but ' +
        'never makes the walk cheaper.',
    ),
  until: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Inclusive upper bound on createdDateTime, same format and same client-side handling ' +
        'as `since`.',
    ),
  maxMessages: z
    .number()
    .int()
    .min(1)
    .max(1000)
    .default(200)
    .describe(
      'How many root messages to return, rounded up to a whole page of 50 — asking for 30 can ' +
        'return up to 50 — and also how far the walk reads: ceil(maxMessages/50) pages, capped ' +
        'at 25. Messages dropped by since/until still consume that page budget, ' +
        'so a window over an older period needs a large maxMessages and probably several calls ' +
        'through `cursor`.',
    ),
  bodyChars: z
    .number()
    .int()
    .min(0)
    .max(2000)
    .default(400)
    .describe(
      'Maximum characters of each body to return after HTML is stripped to plain text; 0 omits ' +
        'bodies entirely. This is the main lever on output size — 200 messages at 400 characters ' +
        'is already most of one response.',
    ),
  includeReplies: z
    .boolean()
    .default(false)
    .describe(
      'Expand each root message\'s replies inline. At most 10 replies per message are projected ' +
        'and the reply collection\'s own nextLink is not followed, so this gives thread shape, ' +
        'not a full thread. It multiplies output, so lower maxMessages and bodyChars with it.',
    ),
  cursor: z
    .string()
    .min(1)
    .optional()
    .describe(
      'Continue a previous call: pass its `cursor` verbatim. The cursor already carries the ' +
        'server-side query, so pass the same teamId, channelId, since, until and includeReplies ' +
        'as the original call — those are re-applied here, and changing them mid-walk produces ' +
        'inconsistent pages.',
    ),
});

const listMembersInput = z.object({
  teamId,
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(50)
    .describe('How many members to return per page. Use nextLink for large teams.'),
});

const sendChannelMessageInput = z.object({
  teamId,
  channelId,
  content: z
    .string()
    .min(1)
    .max(28000)
    .describe(
      'The message body. Plain text unless contentType is "html". Posted as the signed-in ' +
        'user; there is no way to post on someone else\'s behalf.',
    ),
  contentType: z
    .enum(['text', 'html'])
    .default('text')
    .describe(
      'How to interpret `content`. Use "html" only for real markup — with "text", Teams ' +
        'renders angle brackets literally instead of as tags.',
    ),
  subject: z
    .string()
    .min(1)
    .max(255)
    .optional()
    .describe(
      'Optional bold subject line shown above the message. Channel posts show it; omit it for ' +
        'a plain conversational message.',
    ),
  importance: z
    .enum(['normal', 'high', 'urgent'])
    .optional()
    .describe(
      'Message importance. "urgent" repeatedly notifies the channel for 20 minutes, so only ' +
        'use it when the user explicitly asks for it. Defaults to normal.',
    ),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const teamsModule: ToolModule = {
  group: GROUP,
  build({ graph }: ToolDeps): ToolDefinition[] {
    return [
      {
        name: 'teams_list_joined',
        title: 'List my teams',
        description:
          'Lists the Microsoft Teams teams the signed-in user is a member of, returning id, ' +
          'displayName, description, visibility and archived state. Start here: every other ' +
          'teams tool needs a teamId from this list. This endpoint ignores $top, $filter and ' +
          '$orderby and returns a trimmed team object, so properties other than the ones above ' +
          'come back empty even though the full team resource has them. Requires tenant admin ' +
          'consent (Team.ReadBasic.All); a 403 here means the tenant has not granted it.',
        inputSchema: listJoinedInput,
        scopes: TEAM_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          listJoinedInput.parse(args);
          const res = await graph.request({
            path: '/me/joinedTeams',
            method: 'GET',
            scopes: TEAM_SCOPES,
          });
          const teams = extractCollection<TeamPayload>(res.data);
          return {
            count: teams.length,
            teams: teams.map((team) => ({
              id: team.id ?? undefined,
              displayName: team.displayName ?? undefined,
              description: team.description ?? undefined,
              visibility: team.visibility ?? undefined,
              isArchived: team.isArchived === true ? true : undefined,
              webUrl: team.webUrl ?? undefined,
            })),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'teams_list_channels',
        title: 'List team channels',
        description:
          'Lists the channels of one team, returning id, displayName, description, ' +
          'membershipType and webUrl for each. The channelId returned here is what every ' +
          'channel message tool needs. Only channels the signed-in user can see are listed: ' +
          'private and shared channels appear only where the user is a member. Requires ' +
          'tenant admin consent (Channel.ReadBasic.All).',
        inputSchema: listChannelsInput,
        scopes: CHANNEL_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = listChannelsInput.parse(args);
          const res = await graph.request({
            path: `/teams/${segment(input.teamId)}/channels`,
            method: 'GET',
            // Built from a closed enum, never from raw model text.
            query:
              input.membershipType === undefined
                ? undefined
                : { $filter: `membershipType eq '${input.membershipType}'` },
            scopes: CHANNEL_SCOPES,
          });
          const channels = extractCollection<ChannelPayload>(res.data);
          return {
            teamId: input.teamId,
            count: channels.length,
            channels: channels.map(compactChannel),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'teams_get_channel',
        title: 'Get a channel',
        description:
          'Returns one channel: displayName, description, membershipType, creation time, the ' +
          'channel email address (empty unless the channel has one provisioned) and the deep ' +
          'link webUrl. Requires tenant admin consent (Channel.ReadBasic.All).',
        inputSchema: getChannelInput,
        scopes: CHANNEL_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = getChannelInput.parse(args);
          const res = await graph.request<ChannelPayload>({
            path: `/teams/${segment(input.teamId)}/channels/${segment(input.channelId)}`,
            method: 'GET',
            scopes: CHANNEL_SCOPES,
          });
          const channel = res.data ?? {};
          return {
            teamId: input.teamId,
            ...compactChannel(channel),
            email: channel.email ?? undefined,
            isFavoriteByDefault: channel.isFavoriteByDefault ?? undefined,
            isArchived: channel.isArchived === true ? true : undefined,
          };
        },
      },
      {
        name: 'teams_list_channel_messages',
        title: 'List channel messages',
        description:
          'Lists the top-level messages of a channel, newest first, with sender, timestamp, ' +
          'plain-text body (HTML stripped and truncated to `bodyChars`) and attachment, mention ' +
          'and reaction counts. Default page size is 20 and Graph caps it at 50. Replies are ' +
          'NOT included — each message here is a conversation root; use ' +
          'teams_list_message_replies for the thread. Deleted messages still appear as ' +
          'tombstones with deletedDateTime set and no body, and system events (member added, ' +
          'channel renamed) appear with a messageType other than "message". This endpoint is ' +
          'throttled near one request per second per channel, so keep `top` small instead of ' +
          'retrying. Requires tenant admin consent (ChannelMessage.Read.All).',
        inputSchema: listChannelMessagesInput,
        scopes: MESSAGE_READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = listChannelMessagesInput.parse(args);
          const res = await graph.request({
            path: `/teams/${segment(input.teamId)}/channels/${segment(input.channelId)}/messages`,
            method: 'GET',
            query: { $top: input.top },
            scopes: MESSAGE_READ_SCOPES,
          });
          const messages = extractCollection<ChatMessagePayload>(res.data);
          return {
            teamId: input.teamId,
            channelId: input.channelId,
            count: messages.length,
            messages: messages.map((message) => compactMessage(message, input.bodyChars)),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'teams_get_channel_message',
        title: 'Get a channel message',
        description:
          'Returns one channel message in full: sender, timestamps, plain-text body truncated ' +
          'to `bodyChars`, attachment names, mention text and reaction types. Use this after ' +
          'teams_list_channel_messages when a truncated body is not enough. A reply id is not ' +
          'valid here — fetch replies with teams_list_message_replies. Requires tenant admin ' +
          'consent (ChannelMessage.Read.All).',
        inputSchema: getChannelMessageInput,
        scopes: MESSAGE_READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = getChannelMessageInput.parse(args);
          const res = await graph.request<ChatMessagePayload>({
            path:
              `/teams/${segment(input.teamId)}/channels/${segment(input.channelId)}` +
              `/messages/${segment(input.messageId)}`,
            method: 'GET',
            scopes: MESSAGE_READ_SCOPES,
          });
          const message = res.data ?? {};
          const attachments = message.attachments ?? [];
          const mentions = message.mentions ?? [];
          const reactions = message.reactions ?? [];

          return {
            teamId: input.teamId,
            channelId: input.channelId,
            ...compactMessage(message, input.bodyChars),
            lastModifiedDateTime: message.lastModifiedDateTime ?? undefined,
            lastEditedDateTime: message.lastEditedDateTime ?? undefined,
            deletedDateTime: message.deletedDateTime ?? undefined,
            locale: message.locale ?? undefined,
            // Contents are not inlined: contentUrl usually points at a SharePoint
            // file that needs a separate, differently scoped read.
            attachments:
              attachments.length > 0
                ? attachments.map((attachment) => ({
                    id: attachment.id ?? undefined,
                    name: attachment.name ?? undefined,
                    contentType: attachment.contentType ?? undefined,
                    contentUrl: attachment.contentUrl ?? undefined,
                  }))
                : undefined,
            mentions:
              mentions.length > 0
                ? mentions.map((mention) => ({
                    text: mention.mentionText ?? undefined,
                    userId: mention.mentioned?.user?.id ?? undefined,
                    displayName: mention.mentioned?.user?.displayName ?? undefined,
                  }))
                : undefined,
            reactions:
              reactions.length > 0
                ? reactions.map((reaction) => ({
                    type: reaction.reactionType ?? undefined,
                    from: senderName(reaction.user),
                  }))
                : undefined,
          };
        },
      },
      {
        name: 'teams_list_message_replies',
        title: 'List message replies',
        description:
          'Lists the replies to one top-level channel message, newest first, with the same ' +
          'compact projection as teams_list_channel_messages. Default page size is 20, capped ' +
          'at 50. Teams threads are exactly one level deep, so replies never have replies of ' +
          'their own. Like the message list, this endpoint is throttled near one request per ' +
          'second per channel. Requires tenant admin consent (ChannelMessage.Read.All).',
        inputSchema: listRepliesInput,
        scopes: MESSAGE_READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = listRepliesInput.parse(args);
          const res = await graph.request({
            path:
              `/teams/${segment(input.teamId)}/channels/${segment(input.channelId)}` +
              `/messages/${segment(input.messageId)}/replies`,
            method: 'GET',
            query: { $top: input.top },
            scopes: MESSAGE_READ_SCOPES,
          });
          const replies = extractCollection<ChatMessagePayload>(res.data);
          return {
            teamId: input.teamId,
            channelId: input.channelId,
            messageId: input.messageId,
            count: replies.length,
            replies: replies.map((reply) => compactMessage(reply, input.bodyChars)),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'teams_fetch_channel_history',
        title: 'Fetch channel history',
        description:
          'Walks a channel\'s message history and returns a bounded page of root messages with ' +
          'sender, creation and modification times, plain-text body (HTML stripped, truncated to ' +
          '`bodyChars`) and attachment count. Prefer this over paging ' +
          'teams_list_channel_messages by hand: it paces itself at one request per second per ' +
          'channel, ends on maxMessages, 25 pages or a 60-second budget, and hands back a ' +
          '`cursor` to resume from. IMPORTANT — this endpoint supports only $top and $expand: ' +
          'there is no $filter, no $orderby and no $select, so a since/until window cannot be ' +
          'pushed to the server. It is applied here after each page is fetched, which means a ' +
          'window costs a full walk up to the page budget and can still come back incomplete. ' +
          'Graph returns this collection in a fixed order — the last modified date of the whole ' +
          'reply chain, NOT creation order — so an old message whose thread was touched recently ' +
          'arrives early while a quiet old thread arrives late, and a single new reaction can ' +
          'reshuffle the order between calls. Read `range` as a description of what actually came ' +
          'back, never as proof the window is complete. `reason` says why the walk ended: ' +
          '"complete" means the channel ran out, while "maxItems", "maxPages" and "budget" all ' +
          'mean there is more — call again with `cursor`. With includeReplies, Graph inlines up ' +
          'to 200 replies per message under their own replies@odata.nextLink; this tool projects ' +
          'at most 10 per message and does NOT follow that inner link, so use ' +
          'teams_list_message_replies for a complete thread. Requires tenant admin consent ' +
          '(ChannelMessage.Read.All).',
        inputSchema: fetchChannelHistoryInput,
        scopes: MESSAGE_READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = fetchChannelHistoryInput.parse(args);
          const since = windowBound(input.since, false);
          const until = windowBound(input.until, true);
          const windowed = since !== undefined || until !== undefined;

          const inWindow = (message: ChatMessagePayload): boolean => {
            if (!windowed) return true;
            const created = message.createdDateTime;
            const ms = typeof created === 'string' ? Date.parse(created) : Number.NaN;
            // An undated row cannot be placed in the window, and keeping it would
            // make `range` describe something other than what was returned.
            if (Number.isNaN(ms)) return false;
            if (since !== undefined && ms < since) return false;
            if (until !== undefined && ms > until) return false;
            return true;
          };

          const paginator = createPaginator(graph);
          const walk = await paginator.walk<ChannelMessageWithReplies>({
            path: `/teams/${segment(input.teamId)}/channels/${segment(input.channelId)}/messages`,
            // $top is the only knob this collection has besides $expand. No
            // $orderby, no $filter, no $select — the projection below is the
            // only way to make a row smaller.
            query: input.includeReplies
              ? { $top: HISTORY_PAGE_SIZE, $expand: 'replies' }
              : { $top: HISTORY_PAGE_SIZE },
            scopes: MESSAGE_READ_SCOPES,
            // Rounded up to a whole page, exactly as the chat history walk does.
            // The walk cuts a page short when this cap is reached mid-page, but the
            // cursor it hands back is a $skiptoken that points past that whole page
            // — so a cap landing mid-page silently drops the unread remainder on
            // resume. Keeping the cap on a page boundary makes the cursor an exact
            // continuation; the cost is at most 49 rows more than asked for.
            maxItems: Math.ceil(input.maxMessages / HISTORY_PAGE_SIZE) * HISTORY_PAGE_SIZE,
            maxPages: Math.min(
              Math.ceil(input.maxMessages / HISTORY_PAGE_SIZE),
              HISTORY_MAX_PAGES,
            ),
            minIntervalMs: HISTORY_MIN_INTERVAL_MS,
            budgetMs: HISTORY_BUDGET_MS,
            cursor: input.cursor,
            // The window is a `keep`, deliberately not a `stop`. Graph orders
            // this collection by the modification date of the whole reply chain,
            // not by creation date, so a message created years ago surfaces late
            // in the walk as soon as someone replies to or reacts to its thread.
            // A `stop` at the first out-of-range row would end the walk on that
            // ordering and silently truncate the window; `keep` drops the row and
            // keeps reading. (The chat history tool can afford `stop` because
            // /chats/{id}/messages accepts $orderby=createdDateTime desc.)
            keep: inWindow,
          });

          const range = createdRange(walk.items);
          return {
            teamId: input.teamId,
            channelId: input.channelId,
            count: walk.items.length,
            pages: walk.pages,
            reason: walk.reason,
            elapsedMs: walk.elapsedMs,
            throttled: walk.throttled ? true : undefined,
            range,
            messages: walk.items.map((message) =>
              historyRoot(message, input.bodyChars, input.includeReplies),
            ),
            cursor: walk.nextLink,
          };
        },
      },
      {
        name: 'teams_list_members',
        title: 'List team members',
        description:
          'Lists the members of a team with displayName, email, the directory userId and their ' +
          'roles ("owner" for owners, empty for ordinary members, "guest" for guests). The `id` ' +
          'on each entry is a membership id scoped to the team, not the user id — pass userId, ' +
          'not id, to directory tools. This is the team roster; private and shared channels ' +
          'have their own smaller rosters. Requires tenant admin consent, and a scope ' +
          '(TeamMember.Read.All) beyond the rest of this group, so it can 403 even when the ' +
          'other teams tools work.',
        inputSchema: listMembersInput,
        scopes: TEAM_MEMBER_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = listMembersInput.parse(args);
          const res = await graph.request({
            path: `/teams/${segment(input.teamId)}/members`,
            method: 'GET',
            query: { $top: input.top },
            scopes: TEAM_MEMBER_SCOPES,
          });
          const members = extractCollection<ConversationMemberPayload>(res.data);
          return {
            teamId: input.teamId,
            count: members.length,
            members: members.map((member) => ({
              id: member.id ?? undefined,
              userId: member.userId ?? undefined,
              displayName: member.displayName ?? undefined,
              email: member.email ?? undefined,
              roles: member.roles ?? [],
            })),
            nextLink: res.nextLink,
          };
        },
      },
      {
        name: 'teams_send_channel_message',
        title: 'Send a channel message',
        description:
          'Posts a new top-level message to a channel as the signed-in user and returns the ' +
          'created message id, timestamp and deep link. This is visible to everyone in the ' +
          'channel and cannot be recalled by this server, so confirm the exact wording with the ' +
          'user first. It always starts a new conversation — there is no reply support here, so ' +
          'do not use it to answer an existing thread. Requires tenant admin consent ' +
          '(ChannelMessage.Send).',
        inputSchema: sendChannelMessageInput,
        write: true,
        scopes: MESSAGE_SEND_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = sendChannelMessageInput.parse(args);
          const body: Record<string, unknown> = {
            body: { contentType: input.contentType, content: input.content },
          };
          if (input.subject !== undefined) body['subject'] = input.subject;
          if (input.importance !== undefined) body['importance'] = input.importance;

          const res = await graph.request<ChatMessagePayload>({
            path: `/teams/${segment(input.teamId)}/channels/${segment(input.channelId)}/messages`,
            method: 'POST',
            body,
            scopes: MESSAGE_SEND_SCOPES,
          });
          const message = res.data ?? {};
          return {
            sent: true,
            id: message.id ?? undefined,
            teamId: message.channelIdentity?.teamId ?? input.teamId,
            channelId: message.channelIdentity?.channelId ?? input.channelId,
            createdDateTime: message.createdDateTime ?? undefined,
            from: senderName(message.from),
            webUrl: message.webUrl ?? undefined,
          };
        },
      },
    ];
  },
};
