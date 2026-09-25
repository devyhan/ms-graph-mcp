/**
 * The `me` group: read-only access to the signed-in user's own profile and
 * Outlook mailbox settings.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { stripHtml, truncateText } from '../util/truncate.js';

const GROUP = GROUPS['me']!;

/**
 * Copied, not aliased: a tool's `scopes` array is an input to the consent set
 * the login path builds, and handing out the live `GROUPS` array would let a
 * caller that sorts or dedupes in place mutate the group catalogue.
 */
const PROFILE_SCOPES = [...GROUP.readScopes];

/**
 * Without a `$select`, `/me` returns roughly two dozen properties, most of
 * which are null in a typical tenant. This is the set worth spending tokens on.
 */
const DEFAULT_PROFILE_SELECT = [
  'id',
  'displayName',
  'mail',
  'userPrincipalName',
  'jobTitle',
  'officeLocation',
  'preferredLanguage',
];

/**
 * `$select` values reach Graph as a raw query value. Every selectable property
 * on the `user` resource is a bare identifier, so anything else is either a
 * caller mistake or an attempt to smuggle extra OData into the query string.
 */
const SELECT_FIELD = /^[A-Za-z][A-Za-z0-9]*$/;

/**
 * `MailboxSettings.Read` is deliberately absent from the group meta. Group
 * scopes stay the coarse grouping and admin-consent flag; the per-tool `scopes`
 * arrays are the precise source the consent set is built from, so listing it
 * here is what actually gets it requested at login. Reading Outlook
 * configuration is a separate consent decision from reading the directory
 * profile, and only this one tool needs it.
 */
const MAILBOX_SETTINGS_SCOPES = ['MailboxSettings.Read'];

/** The parts of `mailboxSettings` this tool projects. */
interface DateTimeTimeZone {
  dateTime?: string | null;
  timeZone?: string | null;
}

interface MailboxSettingsPayload {
  timeZone?: string | null;
  dateFormat?: string | null;
  timeFormat?: string | null;
  userPurpose?: string | null;
  delegateMeetingMessageDeliveryOptions?: string | null;
  language?: { locale?: string | null; displayName?: string | null } | null;
  workingHours?: {
    daysOfWeek?: string[] | null;
    startTime?: string | null;
    endTime?: string | null;
    timeZone?: { name?: string | null } | null;
  } | null;
  automaticRepliesSetting?: {
    status?: string | null;
    externalAudience?: string | null;
    internalReplyMessage?: string | null;
    externalReplyMessage?: string | null;
    scheduledStartDateTime?: DateTimeTimeZone | null;
    scheduledEndDateTime?: DateTimeTimeZone | null;
  } | null;
}

const profileInput = z.object({
  select: z
    .array(
      z
        .string()
        .min(1)
        .regex(SELECT_FIELD, 'must be a bare Microsoft Graph property name'),
    )
    .min(1)
    .max(30)
    .optional()
    .describe(
      'Graph `user` property names to return instead of the default set, e.g. ' +
        '["displayName","department","mobilePhone"]. Bare property names only: ' +
        'navigation paths such as "manager/displayName" and any other OData ' +
        'expression are rejected. Omit for the compact default projection.',
    ),
});

const mailboxSettingsInput = z.object({
  replyBodyChars: z
    .number()
    .int()
    .min(0)
    .max(4000)
    .default(600)
    .describe(
      'Maximum characters of each automatic-reply message to return after HTML is ' +
        'stripped. Use 0 to omit the reply bodies entirely and keep only the status ' +
        'and schedule.',
    ),
});

/** Keeps only the requested keys, dropping the nulls Graph pads a user with. */
function projectProfile(
  data: unknown,
  fields: readonly string[],
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (typeof data !== 'object' || data === null) return out;
  const record = data as Record<string, unknown>;
  for (const field of fields) {
    const value = record[field];
    if (value !== undefined && value !== null && value !== '') out[field] = value;
  }
  return out;
}

/** Graph returns working hours as `08:00:00.0000000`; the fraction is noise. */
function clockTime(value: string | null | undefined): string | undefined {
  if (typeof value !== 'string' || value.length === 0) return undefined;
  const match = /^(\d{2}:\d{2}(?::\d{2})?)/.exec(value);
  return match?.[1] ?? value;
}

/** Automatic-reply messages are HTML; the model wants short plain text. */
function replyBody(
  html: string | null | undefined,
  maxChars: number,
): string | undefined {
  if (maxChars === 0 || typeof html !== 'string' || html.length === 0) return undefined;
  const text = stripHtml(html);
  return text.length === 0 ? undefined : truncateText(text, maxChars);
}

export const meModule: ToolModule = {
  group: GROUP,
  build({ graph }: ToolDeps): ToolDefinition[] {
    return [
      {
        name: 'me_get_profile',
        title: 'Get my profile',
        description:
          "Returns the signed-in user's own Entra ID profile from /me: by default id, " +
          'displayName, mail, userPrincipalName, jobTitle, officeLocation and ' +
          'preferredLanguage. Pass `select` to ask for different properties instead. ' +
          'Properties Graph returns as null or empty are omitted, so a missing key means ' +
          '"not set" rather than "not requested". `mail` is null for accounts without an ' +
          'Exchange mailbox and personal Microsoft accounts leave the work fields empty — ' +
          'userPrincipalName is the identifier that is always present.',
        inputSchema: profileInput,
        scopes: PROFILE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { select } = profileInput.parse(args);
          // `id` is always worth the handful of characters: every follow-up call
          // that targets this user needs it.
          const fields = [...new Set(['id', ...(select ?? DEFAULT_PROFILE_SELECT)])];
          const res = await graph.request({
            path: '/me',
            method: 'GET',
            query: { $select: fields.join(',') },
            scopes: PROFILE_SCOPES,
          });
          return projectProfile(res.data, fields);
        },
      },
      {
        name: 'me_get_mailbox_settings',
        title: 'Get my mailbox settings',
        description:
          "Returns the signed-in user's Outlook mailbox configuration: time zone, date " +
          'and time format, locale, working hours, and the automatic-reply ' +
          '(out-of-office) setting. Automatic-reply bodies are HTML in Graph and come ' +
          'back stripped to plain text and truncated to `replyBodyChars`. The scheduled ' +
          'start and end times are only meaningful when status is "scheduled" — when ' +
          'status is "alwaysEnabled" or "disabled" they are stale leftovers from the last ' +
          'schedule. Accounts with no Exchange Online mailbox fail with 404 ' +
          'MailboxNotEnabledForRESTAPI.',
        inputSchema: mailboxSettingsInput,
        scopes: MAILBOX_SETTINGS_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { replyBodyChars } = mailboxSettingsInput.parse(args);
          const res = await graph.request<MailboxSettingsPayload>({
            path: '/me/mailboxSettings',
            method: 'GET',
            scopes: MAILBOX_SETTINGS_SCOPES,
          });
          const settings = res.data ?? {};
          const hours = settings.workingHours;
          const replies = settings.automaticRepliesSetting;

          return {
            timeZone: settings.timeZone ?? undefined,
            locale: settings.language?.locale ?? undefined,
            dateFormat: settings.dateFormat ?? undefined,
            timeFormat: settings.timeFormat ?? undefined,
            userPurpose: settings.userPurpose ?? undefined,
            delegateMeetingMessageDeliveryOptions:
              settings.delegateMeetingMessageDeliveryOptions ?? undefined,
            workingHours: hours
              ? {
                  daysOfWeek: hours.daysOfWeek ?? [],
                  startTime: clockTime(hours.startTime),
                  endTime: clockTime(hours.endTime),
                  // Working hours carry their own zone, which can differ from the
                  // mailbox zone above.
                  timeZone: hours.timeZone?.name ?? undefined,
                }
              : undefined,
            automaticReplies: replies
              ? {
                  status: replies.status ?? undefined,
                  externalAudience: replies.externalAudience ?? undefined,
                  scheduledStart: replies.scheduledStartDateTime?.dateTime ?? undefined,
                  scheduledEnd: replies.scheduledEndDateTime?.dateTime ?? undefined,
                  scheduleTimeZone:
                    replies.scheduledStartDateTime?.timeZone ?? undefined,
                  internalReplyMessage: replyBody(
                    replies.internalReplyMessage,
                    replyBodyChars,
                  ),
                  externalReplyMessage: replyBody(
                    replies.externalReplyMessage,
                    replyBodyChars,
                  ),
                }
              : undefined,
          };
        },
      },
    ];
  },
};
