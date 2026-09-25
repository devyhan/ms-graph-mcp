/**
 * Outlook calendar tools: reading events, checking availability, and managing
 * invitations.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { buildQuery, isoDate } from '../util/odata.js';
import { stripHtml, truncateText } from '../util/truncate.js';

const GROUP = GROUPS['calendar']!;

const READ_SCOPES = GROUP.readScopes;
const WRITE_SCOPES = [...GROUP.readScopes, ...GROUP.writeScopes];

/**
 * getSchedule and findMeetingTimes read *other people's* free/busy blocks.
 * Delegated Calendars.Read only covers the signed-in user's own calendars, so
 * Graph rejects both endpoints with ErrorAccessDenied unless the .Shared
 * variant was consented. It is not in the group meta because no other calendar
 * tool needs it.
 */
const SHARED_READ_SCOPES = [...GROUP.readScopes, 'Calendars.Read.Shared'];

// ---------------------------------------------------------------------------
// Graph shapes (only the fields these tools project)
// ---------------------------------------------------------------------------

interface EmailAddress {
  name?: string | null;
  address?: string | null;
}

interface DateTimeTimeZone {
  dateTime?: string | null;
  timeZone?: string | null;
}

interface GraphAttendee {
  type?: string | null;
  status?: { response?: string | null; time?: string | null } | null;
  emailAddress?: EmailAddress | null;
}

interface GraphEvent {
  id?: string;
  subject?: string | null;
  bodyPreview?: string | null;
  body?: { contentType?: string | null; content?: string | null } | null;
  start?: DateTimeTimeZone | null;
  end?: DateTimeTimeZone | null;
  isAllDay?: boolean | null;
  isCancelled?: boolean | null;
  isOnlineMeeting?: boolean | null;
  onlineMeeting?: { joinUrl?: string | null } | null;
  location?: { displayName?: string | null } | null;
  organizer?: { emailAddress?: EmailAddress | null } | null;
  attendees?: GraphAttendee[] | null;
  responseStatus?: { response?: string | null; time?: string | null } | null;
  showAs?: string | null;
  importance?: string | null;
  sensitivity?: string | null;
  type?: string | null;
  seriesMasterId?: string | null;
  recurrence?: {
    pattern?: { type?: string | null; interval?: number | null } | null;
    range?: { type?: string | null; startDate?: string | null; endDate?: string | null } | null;
  } | null;
  webLink?: string | null;
  categories?: string[] | null;
  hasAttachments?: boolean | null;
  reminderMinutesBeforeStart?: number | null;
  lastModifiedDateTime?: string | null;
}

interface GraphCalendar {
  id?: string;
  name?: string | null;
  color?: string | null;
  hexColor?: string | null;
  isDefaultCalendar?: boolean | null;
  canEdit?: boolean | null;
  canShare?: boolean | null;
  canViewPrivateItems?: boolean | null;
  owner?: EmailAddress | null;
}

interface ScheduleInformation {
  scheduleId?: string | null;
  availabilityView?: string | null;
  error?: { message?: string | null; responseCode?: string | null } | null;
  workingHours?: {
    daysOfWeek?: string[] | null;
    startTime?: string | null;
    endTime?: string | null;
    timeZone?: { name?: string | null } | null;
  } | null;
  scheduleItems?: Array<{
    status?: string | null;
    subject?: string | null;
    location?: string | null;
    isPrivate?: boolean | null;
    start?: DateTimeTimeZone | null;
    end?: DateTimeTimeZone | null;
  }> | null;
}

interface MeetingTimeSuggestion {
  confidence?: number | null;
  organizerAvailability?: string | null;
  suggestionReason?: string | null;
  meetingTimeSlot?: { start?: DateTimeTimeZone | null; end?: DateTimeTimeZone | null } | null;
  attendeeAvailability?: Array<{
    availability?: string | null;
    attendee?: { emailAddress?: EmailAddress | null } | null;
  }> | null;
  locations?: Array<{ displayName?: string | null }> | null;
}

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/**
 * A Prefer header value is spliced into the raw request headers, so a model
 * supplying `UTC"\r\nX-Evil: 1` could forge a header. Windows and IANA zone
 * names never need anything outside this set.
 */
const TIME_ZONE_NAME = /^[A-Za-z0-9 ()+_./-]{1,64}$/;

function preferHeader(timeZone: string | undefined, plainTextBody: boolean): Record<string, string> | undefined {
  const prefer: string[] = [];

  if (timeZone !== undefined && timeZone.length > 0) {
    if (!TIME_ZONE_NAME.test(timeZone)) {
      throw new Error(
        `Invalid timeZone ${JSON.stringify(timeZone)}. Use an IANA name such as "Europe/Berlin" or a Windows name such as "Pacific Standard Time".`,
      );
    }
    prefer.push(`outlook.timezone="${timeZone}"`);
  }

  // Ask Outlook to convert HTML bodies for us; stripHtml stays as a fallback
  // because older mailboxes still answer with HTML.
  if (plainTextBody) prefer.push('outlook.body-content-type="text"');

  return prefer.length > 0 ? { Prefer: prefer.join(', ') } : undefined;
}

/**
 * Graph's DateTimeTimeZone wants a *naive* wall-clock string plus a separate
 * zone name; handing it "2026-01-31T09:00:00Z" together with a timeZone is
 * ambiguous and Graph reads it inconsistently. So: an argument carrying an
 * explicit offset is an absolute instant and is normalised to UTC, while one
 * without an offset is a local wall time interpreted in `timeZone`.
 */
function toDateTimeTimeZone(value: string, timeZone: string | undefined): { dateTime: string; timeZone: string } {
  const iso = isoDate(value);
  const hasOffset = /(?:Z|[+-]\d{2}:\d{2})$/.test(iso);

  if (hasOffset) {
    const instant = new Date(iso);
    return { dateTime: instant.toISOString().slice(0, 19), timeZone: 'UTC' };
  }

  const dateTime = iso.includes('T') ? iso : `${iso}T00:00:00`;
  return { dateTime, timeZone: timeZone ?? 'UTC' };
}

function formatEmail(e: EmailAddress | null | undefined): string | undefined {
  if (e === null || e === undefined) return undefined;
  const address = typeof e.address === 'string' && e.address.length > 0 ? e.address : undefined;
  const name = typeof e.name === 'string' && e.name.length > 0 ? e.name : undefined;
  if (address === undefined) return name;
  return name === undefined || name === address ? address : `${name} <${address}>`;
}

function formatWhen(v: DateTimeTimeZone | null | undefined): string | undefined {
  if (v === null || v === undefined || typeof v.dateTime !== 'string') return undefined;
  return typeof v.timeZone === 'string' && v.timeZone.length > 0 ? `${v.dateTime} ${v.timeZone}` : v.dateTime;
}

/** Graph echoes `@odata.etag` / `@odata.context` on every entity; the model never needs them. */
function withoutODataKeys(item: unknown): unknown {
  if (item === null || typeof item !== 'object' || Array.isArray(item)) return item;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(item as Record<string, unknown>)) {
    if (key.startsWith('@odata.')) continue;
    out[key] = value;
  }
  return out;
}

const MAX_ATTENDEES_SHOWN = 50;
const MAX_SCHEDULE_ITEMS_SHOWN = 25;

function summarizeEvent(ev: GraphEvent): Record<string, unknown> {
  const out: Record<string, unknown> = {
    id: ev.id,
    subject: ev.subject ?? '(no subject)',
    start: formatWhen(ev.start),
    end: formatWhen(ev.end),
  };

  if (ev.isAllDay === true) out['isAllDay'] = true;
  if (ev.isCancelled === true) out['isCancelled'] = true;

  const location = ev.location?.displayName;
  if (typeof location === 'string' && location.length > 0) out['location'] = location;

  const organizer = formatEmail(ev.organizer?.emailAddress);
  if (organizer !== undefined) out['organizer'] = organizer;

  const attendees = ev.attendees ?? [];
  if (attendees.length > 0) out['attendeeCount'] = attendees.length;

  if (typeof ev.showAs === 'string' && ev.showAs !== 'busy') out['showAs'] = ev.showAs;

  const myResponse = ev.responseStatus?.response;
  if (typeof myResponse === 'string' && myResponse !== 'none') out['myResponse'] = myResponse;

  if (ev.isOnlineMeeting === true) {
    const joinUrl = ev.onlineMeeting?.joinUrl;
    out['onlineMeetingJoinUrl'] = typeof joinUrl === 'string' ? joinUrl : '(join URL not returned)';
  }

  if (typeof ev.type === 'string' && ev.type !== 'singleInstance') out['type'] = ev.type;
  if (typeof ev.seriesMasterId === 'string') out['seriesMasterId'] = ev.seriesMasterId;

  return out;
}

function detailEvent(ev: GraphEvent, bodyChars: number): Record<string, unknown> {
  const out = summarizeEvent(ev);

  const attendees = ev.attendees ?? [];
  if (attendees.length > 0) {
    out['attendees'] = attendees.slice(0, MAX_ATTENDEES_SHOWN).map((a) => ({
      who: formatEmail(a.emailAddress),
      type: a.type ?? undefined,
      response: a.status?.response ?? undefined,
    }));
    if (attendees.length > MAX_ATTENDEES_SHOWN) {
      out['attendeesOmitted'] = attendees.length - MAX_ATTENDEES_SHOWN;
    }
  }

  const raw = ev.body?.content;
  const isHtml = ev.body?.contentType === 'html';
  const text = typeof raw === 'string' && raw.length > 0 ? (isHtml ? stripHtml(raw) : raw) : (ev.bodyPreview ?? '');
  if (text.length > 0) out['body'] = truncateText(text.trim(), bodyChars);

  const pattern = ev.recurrence?.pattern;
  if (pattern !== null && pattern !== undefined) {
    out['recurrence'] = {
      pattern: pattern.type ?? undefined,
      interval: pattern.interval ?? undefined,
      range: ev.recurrence?.range?.type ?? undefined,
      until: ev.recurrence?.range?.endDate ?? undefined,
    };
  }

  if (ev.hasAttachments === true) out['hasAttachments'] = true;
  if (typeof ev.importance === 'string' && ev.importance !== 'normal') out['importance'] = ev.importance;
  if (typeof ev.sensitivity === 'string' && ev.sensitivity !== 'normal') out['sensitivity'] = ev.sensitivity;
  if (Array.isArray(ev.categories) && ev.categories.length > 0) out['categories'] = ev.categories;
  if (typeof ev.reminderMinutesBeforeStart === 'number') {
    out['reminderMinutesBeforeStart'] = ev.reminderMinutesBeforeStart;
  }
  if (typeof ev.lastModifiedDateTime === 'string') out['lastModifiedDateTime'] = ev.lastModifiedDateTime;
  if (typeof ev.webLink === 'string') out['webLink'] = ev.webLink;

  return out;
}

/** Builds the attendee entries `POST /me/events` and `findMeetingTimes` both take. */
function graphAttendees(
  list: ReadonlyArray<{ address: string; name?: string | undefined; type?: string | undefined }>,
): Array<Record<string, unknown>> {
  return list.map((a) => ({
    emailAddress: { address: a.address, name: a.name },
    type: a.type ?? 'required',
  }));
}

// ---------------------------------------------------------------------------
// Schemas
// ---------------------------------------------------------------------------

const attendeeSchema = z.object({
  address: z.string().min(1).describe('SMTP address of the attendee, e.g. jane@contoso.com.'),
  name: z.string().optional().describe('Display name to show on the invitation. Optional.'),
  type: z
    .enum(['required', 'optional', 'resource'])
    .optional()
    .describe("Attendance type. 'resource' is for rooms and equipment. Defaults to 'required'."),
});

const timeZoneField = z
  .string()
  .optional()
  .describe(
    'IANA or Windows time-zone name (e.g. "Europe/Berlin", "Pacific Standard Time") used to interpret naive start/end values and to render returned times.',
  );

const listEventsSchema = z.object({
  start: z
    .string()
    .optional()
    .describe(
      'ISO-8601 start of the window, e.g. 2026-01-31T00:00:00Z. Must be given together with `end`; supplying both switches to calendarView, which expands recurring series into individual occurrences.',
    ),
  end: z
    .string()
    .optional()
    .describe('ISO-8601 end of the window. Must be given together with `start`.'),
  calendarId: z
    .string()
    .min(1)
    .optional()
    .describe('Calendar to read from, from calendar_list_calendars. Defaults to the primary calendar.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(25)
    .describe('Maximum events to return in this page. Defaults to 25.'),
  select: z
    .array(z.string())
    .optional()
    .describe(
      'Graph field names to return, e.g. ["id","subject","start"]. When given, the raw selected fields are returned instead of the compact projection.',
    ),
  orderby: z
    .string()
    .optional()
    .describe('OData sort expression, e.g. "start/dateTime desc". Defaults to "start/dateTime" ascending.'),
  timeZone: timeZoneField,
});

const getEventSchema = z.object({
  id: z.string().min(1).describe('Event id, from calendar_list_events.'),
  timeZone: timeZoneField,
});

const listCalendarsSchema = z.object({
  top: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(50)
    .describe('Maximum calendars to return. Defaults to 50.'),
});

const getScheduleSchema = z.object({
  schedules: z
    .array(z.string().min(1))
    .min(1)
    .max(20)
    .describe('SMTP addresses of the people or rooms to check, at most 20 per call.'),
  start: z.string().describe('ISO-8601 start of the window to check, e.g. 2026-01-31T08:00:00Z.'),
  end: z
    .string()
    .describe('ISO-8601 end of the window. Keep the window under a few days: the availability string grows with it.'),
  availabilityViewInterval: z
    .number()
    .int()
    .min(5)
    .max(1440)
    .default(30)
    .describe('Minutes each character of availabilityView represents. Defaults to 30.'),
  timeZone: timeZoneField,
});

const findMeetingTimesSchema = z.object({
  attendees: z
    .array(z.string().min(1))
    .min(1)
    .max(20)
    .describe('SMTP addresses of the required attendees. The signed-in user is the organizer and must not be listed.'),
  durationMinutes: z
    .number()
    .int()
    .min(5)
    .max(1440)
    .describe('Meeting length in minutes.'),
  start: z
    .string()
    .optional()
    .describe('ISO-8601 earliest acceptable start. Must be given together with `end`; omit both to let Graph pick the window.'),
  end: z.string().optional().describe('ISO-8601 latest acceptable end. Must be given together with `start`.'),
  maxCandidates: z
    .number()
    .int()
    .min(1)
    .max(100)
    .default(10)
    .describe('Maximum suggestions to return. Defaults to 10.'),
  timeZone: timeZoneField,
});

const createEventSchema = z.object({
  subject: z.string().min(1).describe('Event subject line.'),
  start: z
    .string()
    .describe('ISO-8601 start. Without an offset it is read as local time in `timeZone`; with a Z or ±HH:MM offset it is an absolute instant.'),
  end: z.string().describe('ISO-8601 end, interpreted the same way as `start`.'),
  timeZone: timeZoneField,
  attendees: z
    .array(attendeeSchema)
    .optional()
    .describe('People and rooms to invite. Creating the event sends them an invitation immediately.'),
  body: z.string().optional().describe('Event description shown in the invitation.'),
  bodyType: z
    .enum(['text', 'html'])
    .default('text')
    .describe("How to interpret `body`. Defaults to 'text'."),
  location: z.string().optional().describe('Free-text location, e.g. "Room 4B" or "Zurich office".'),
  isAllDay: z
    .boolean()
    .optional()
    .describe('Make this an all-day event. Graph then requires start and end to be midnight and at least 24h apart.'),
  isOnlineMeeting: z
    .boolean()
    .optional()
    .describe('Attach a Microsoft Teams meeting link.'),
});

const updateEventSchema = z.object({
  id: z.string().min(1).describe('Event id to update.'),
  subject: z.string().min(1).optional().describe('New subject line.'),
  start: z
    .string()
    .optional()
    .describe('New ISO-8601 start. Send `end` too whenever the new start would land after the existing end.'),
  end: z.string().optional().describe('New ISO-8601 end.'),
  timeZone: timeZoneField,
  attendees: z
    .array(attendeeSchema)
    .optional()
    .describe('Replaces the entire attendee list — anyone omitted is uninvited, so send the full list, not just additions.'),
  body: z.string().optional().describe('Replacement event description.'),
  bodyType: z.enum(['text', 'html']).default('text').describe("How to interpret `body`. Defaults to 'text'."),
  location: z.string().optional().describe('Replacement location.'),
  isAllDay: z.boolean().optional().describe('Switch the event to or from all-day.'),
  isOnlineMeeting: z.boolean().optional().describe('Add a Teams meeting link. Removing one after the fact is not supported.'),
  showAs: z
    .enum(['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'])
    .optional()
    .describe('How the event shows on the free/busy view.'),
  categories: z.array(z.string()).optional().describe('Replaces the category list.'),
  reminderMinutesBeforeStart: z
    .number()
    .int()
    .min(0)
    .max(20160)
    .optional()
    .describe('Reminder lead time in minutes.'),
});

const deleteEventSchema = z.object({
  id: z.string().min(1).describe('Event id to delete.'),
});

const respondEventSchema = z.object({
  id: z.string().min(1).describe('Event id of the invitation to respond to.'),
  response: z
    .enum(['accept', 'decline', 'tentativelyAccept'])
    .describe('The reply to send.'),
  comment: z.string().optional().describe('Optional message included with the reply.'),
  sendResponse: z
    .boolean()
    .default(true)
    .describe('Whether the organizer is notified. Defaults to true.'),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const calendarModule: ToolModule = {
  group: GROUP,

  build({ graph, config }: ToolDeps): ToolDefinition[] {
    // A single event body must not be able to consume the whole output budget.
    const bodyChars = Math.min(4000, Math.max(300, Math.floor(config.maxOutputChars / 10)));

    return [
      {
        name: 'calendar_list_events',
        title: 'List calendar events',
        description:
          'Lists events with subject, start, end, location, organizer and attendee count, plus a nextLink when more pages exist. Returns 25 events by default. Passing both start and end switches to calendarView, which expands recurring meetings into their individual occurrences; without them the raw event list is returned, where a recurring series appears once as its series master. Times without an explicit offset in start/end are read as UTC — pass a timeZone to control how returned times are rendered.',
        group: 'calendar',
        scopes: READ_SCOPES,
        inputSchema: listEventsSchema,
        handler: async (args) => {
          const a = listEventsSchema.parse(args);
          const { start, end } = a;

          if ((start === undefined) !== (end === undefined)) {
            throw new Error(
              'calendar_list_events needs both start and end, or neither. Graph rejects a calendarView with only one bound.',
            );
          }

          const window =
            start !== undefined && end !== undefined
              ? { startDateTime: isoDate(start), endDateTime: isoDate(end) }
              : undefined;
          const base =
            a.calendarId === undefined
              ? '/me'
              : `/me/calendars/${encodeURIComponent(a.calendarId)}`;
          const path = `${base}/${window === undefined ? 'events' : 'calendarView'}`;

          const query = buildQuery({
            select: a.select,
            top: a.top,
            orderby: a.orderby ?? 'start/dateTime',
            ...window,
          });

          const res = await graph.request<unknown>({
            path,
            method: 'GET',
            query,
            headers: preferHeader(a.timeZone, false),
            scopes: READ_SCOPES,
          });

          const events = extractCollection<GraphEvent>(res.data);
          const items =
            a.select === undefined ? events.map(summarizeEvent) : events.map(withoutODataKeys);

          return {
            source:
              window === undefined ? 'events (series masters)' : 'calendarView (recurrences expanded)',
            count: items.length,
            items,
            nextLink: res.nextLink,
          };
        },
      },

      {
        name: 'calendar_get_event',
        title: 'Get a calendar event',
        description:
          'Returns one event in full: subject, start, end, location, organizer, the attendee list with each response, the description body (HTML converted to text and capped), recurrence summary, and the Outlook web link. Use calendar_list_events first to find the id.',
        group: 'calendar',
        scopes: READ_SCOPES,
        inputSchema: getEventSchema,
        handler: async (args) => {
          const a = getEventSchema.parse(args);

          // Event ids are base64-ish and routinely contain '/' and '+'.
          const res = await graph.request<GraphEvent>({
            path: `/me/events/${encodeURIComponent(a.id)}`,
            method: 'GET',
            headers: preferHeader(a.timeZone, true),
            scopes: READ_SCOPES,
          });

          return detailEvent(res.data, bodyChars);
        },
      },

      {
        name: 'calendar_list_calendars',
        title: 'List calendars',
        description:
          "Lists the signed-in user's calendars with id, name, colour, owner and whether they can be edited or shared. Returns 50 by default. Use the returned id as calendarId in calendar_list_events.",
        group: 'calendar',
        scopes: READ_SCOPES,
        inputSchema: listCalendarsSchema,
        handler: async (args) => {
          const a = listCalendarsSchema.parse(args);

          const res = await graph.request<unknown>({
            path: '/me/calendars',
            method: 'GET',
            query: buildQuery({ top: a.top }),
            scopes: READ_SCOPES,
          });

          const calendars = extractCollection<GraphCalendar>(res.data);

          return {
            count: calendars.length,
            items: calendars.map((c) => ({
              id: c.id,
              name: c.name ?? undefined,
              isDefault: c.isDefaultCalendar === true ? true : undefined,
              canEdit: c.canEdit ?? undefined,
              canShare: c.canShare ?? undefined,
              canViewPrivateItems: c.canViewPrivateItems === true ? true : undefined,
              owner: formatEmail(c.owner),
              color: c.hexColor !== null && c.hexColor !== undefined && c.hexColor.length > 0 ? c.hexColor : (c.color ?? undefined),
            })),
            nextLink: res.nextLink,
          };
        },
      },

      {
        name: 'calendar_get_schedule',
        title: 'Get free/busy schedule',
        description:
          'Returns free/busy availability for up to 20 people or rooms over a time window, as an availabilityView string where each character covers one interval (0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere), plus working hours and any visible busy blocks. Read-only despite being an HTTP POST. Subject and location are only returned for people who share those details with you; everything else comes back as a bare busy block.',
        group: 'calendar',
        scopes: SHARED_READ_SCOPES,
        inputSchema: getScheduleSchema,
        handler: async (args) => {
          const a = getScheduleSchema.parse(args);

          const res = await graph.request<unknown>({
            path: '/me/calendar/getSchedule',
            method: 'POST',
            body: {
              schedules: a.schedules,
              startTime: toDateTimeTimeZone(a.start, a.timeZone),
              endTime: toDateTimeTimeZone(a.end, a.timeZone),
              availabilityViewInterval: a.availabilityViewInterval,
            },
            headers: preferHeader(a.timeZone, false),
            scopes: SHARED_READ_SCOPES,
          });

          const schedules = extractCollection<ScheduleInformation>(res.data);

          return {
            legend: '0=free, 1=tentative, 2=busy, 3=out of office, 4=working elsewhere',
            intervalMinutes: a.availabilityViewInterval,
            count: schedules.length,
            schedules: schedules.map((s) => {
              const items = s.scheduleItems ?? [];
              const hours = s.workingHours;
              return {
                scheduleId: s.scheduleId ?? undefined,
                availabilityView: s.availabilityView ?? undefined,
                error: s.error?.message ?? undefined,
                workingHours:
                  hours === null || hours === undefined
                    ? undefined
                    : {
                        days: hours.daysOfWeek ?? undefined,
                        from: hours.startTime ?? undefined,
                        to: hours.endTime ?? undefined,
                        timeZone: hours.timeZone?.name ?? undefined,
                      },
                items: items.slice(0, MAX_SCHEDULE_ITEMS_SHOWN).map((it) => ({
                  status: it.status ?? undefined,
                  subject: it.subject ?? undefined,
                  location: it.location ?? undefined,
                  start: formatWhen(it.start),
                  end: formatWhen(it.end),
                })),
                itemsOmitted:
                  items.length > MAX_SCHEDULE_ITEMS_SHOWN ? items.length - MAX_SCHEDULE_ITEMS_SHOWN : undefined,
              };
            }),
          };
        },
      },

      {
        name: 'calendar_find_meeting_times',
        title: 'Find meeting times',
        description:
          'Suggests meeting slots that work for the signed-in user and the given attendees, ranked by confidence, with the attendees who are unavailable in each slot. Returns 10 suggestions by default. Read-only despite being an HTTP POST. All attendees are treated as required, and an empty result comes back with emptySuggestionsReason explaining why (commonly "AttendeesUnavailable" or a window that is too narrow).',
        group: 'calendar',
        scopes: SHARED_READ_SCOPES,
        inputSchema: findMeetingTimesSchema,
        handler: async (args) => {
          const a = findMeetingTimesSchema.parse(args);

          if ((a.start === undefined) !== (a.end === undefined)) {
            throw new Error('calendar_find_meeting_times needs both start and end, or neither.');
          }

          const body: Record<string, unknown> = {
            attendees: graphAttendees(a.attendees.map((address) => ({ address }))),
            meetingDuration: `PT${a.durationMinutes}M`,
            maxCandidates: a.maxCandidates,
            isOrganizerOptional: false,
            returnSuggestionReasons: true,
          };

          if (a.start !== undefined && a.end !== undefined) {
            body['timeConstraint'] = {
              activityDomain: 'work',
              timeSlots: [
                {
                  start: toDateTimeTimeZone(a.start, a.timeZone),
                  end: toDateTimeTimeZone(a.end, a.timeZone),
                },
              ],
            };
          }

          const res = await graph.request<{
            emptySuggestionsReason?: string | null;
            meetingTimeSuggestions?: MeetingTimeSuggestion[] | null;
          }>({
            path: '/me/findMeetingTimes',
            method: 'POST',
            body,
            headers: preferHeader(a.timeZone, false),
            scopes: SHARED_READ_SCOPES,
          });

          const suggestions = res.data.meetingTimeSuggestions ?? [];

          return {
            emptySuggestionsReason: res.data.emptySuggestionsReason ?? undefined,
            count: suggestions.length,
            suggestions: suggestions.map((s) => {
              const unavailable = (s.attendeeAvailability ?? [])
                .filter((x) => x.availability !== 'free')
                .map((x) => `${formatEmail(x.attendee?.emailAddress) ?? 'unknown'} (${x.availability ?? 'unknown'})`);
              return {
                start: formatWhen(s.meetingTimeSlot?.start),
                end: formatWhen(s.meetingTimeSlot?.end),
                confidence: s.confidence ?? undefined,
                organizerAvailability: s.organizerAvailability ?? undefined,
                unavailable: unavailable.length > 0 ? unavailable : undefined,
                reason: s.suggestionReason ?? undefined,
              };
            }),
          };
        },
      },

      {
        name: 'calendar_create_event',
        title: 'Create a calendar event',
        description:
          'Creates an event on the primary calendar and returns the created event. Listing attendees sends them a meeting invitation immediately, so confirm the details with the user first. A start or end without an explicit UTC offset is treated as local time in `timeZone` (UTC when that is omitted).',
        group: 'calendar',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: createEventSchema,
        handler: async (args) => {
          const a = createEventSchema.parse(args);

          const body: Record<string, unknown> = {
            subject: a.subject,
            start: toDateTimeTimeZone(a.start, a.timeZone),
            end: toDateTimeTimeZone(a.end, a.timeZone),
          };

          if (a.body !== undefined) {
            body['body'] = { contentType: a.bodyType === 'html' ? 'HTML' : 'Text', content: a.body };
          }
          if (a.location !== undefined) body['location'] = { displayName: a.location };
          if (a.attendees !== undefined) body['attendees'] = graphAttendees(a.attendees);
          if (a.isAllDay !== undefined) body['isAllDay'] = a.isAllDay;
          if (a.isOnlineMeeting === true) {
            body['isOnlineMeeting'] = true;
            // Graph defaults the provider to the tenant's setting and errors on
            // some tenants when it is left unset alongside isOnlineMeeting.
            body['onlineMeetingProvider'] = 'teamsForBusiness';
          }

          const res = await graph.request<GraphEvent>({
            path: '/me/events',
            method: 'POST',
            body,
            headers: preferHeader(a.timeZone, true),
            scopes: WRITE_SCOPES,
          });

          return { created: true, event: detailEvent(res.data, bodyChars) };
        },
      },

      {
        name: 'calendar_update_event',
        title: 'Update a calendar event',
        description:
          'Patches the given fields on an event and returns the updated event; omitted fields are left alone. Two traps: sending `attendees` replaces the whole list, so anyone left out is uninvited, and changing the time or attendees of a meeting sends an update to everyone invited. Updating a recurring series master changes every occurrence.',
        group: 'calendar',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: updateEventSchema,
        handler: async (args) => {
          const a = updateEventSchema.parse(args);

          const body: Record<string, unknown> = {};
          if (a.subject !== undefined) body['subject'] = a.subject;
          if (a.start !== undefined) body['start'] = toDateTimeTimeZone(a.start, a.timeZone);
          if (a.end !== undefined) body['end'] = toDateTimeTimeZone(a.end, a.timeZone);
          if (a.body !== undefined) {
            body['body'] = { contentType: a.bodyType === 'html' ? 'HTML' : 'Text', content: a.body };
          }
          if (a.location !== undefined) body['location'] = { displayName: a.location };
          if (a.attendees !== undefined) body['attendees'] = graphAttendees(a.attendees);
          if (a.isAllDay !== undefined) body['isAllDay'] = a.isAllDay;
          if (a.showAs !== undefined) body['showAs'] = a.showAs;
          if (a.categories !== undefined) body['categories'] = a.categories;
          if (a.reminderMinutesBeforeStart !== undefined) {
            body['reminderMinutesBeforeStart'] = a.reminderMinutesBeforeStart;
          }
          if (a.isOnlineMeeting === true) {
            body['isOnlineMeeting'] = true;
            body['onlineMeetingProvider'] = 'teamsForBusiness';
          }

          if (Object.keys(body).length === 0) {
            throw new Error('calendar_update_event needs at least one field to change besides id.');
          }

          const res = await graph.request<GraphEvent>({
            path: `/me/events/${encodeURIComponent(a.id)}`,
            method: 'PATCH',
            body,
            headers: preferHeader(a.timeZone, true),
            scopes: WRITE_SCOPES,
          });

          return { updated: true, event: detailEvent(res.data, bodyChars) };
        },
      },

      {
        name: 'calendar_delete_event',
        title: 'Delete a calendar event',
        description:
          'Deletes an event, moving it to Deleted Items. If the signed-in user organized a meeting, every attendee is sent a cancellation; if they were only invited, this removes it from their calendar without telling the organizer — use calendar_respond_event with decline for that. Deleting a series master deletes every occurrence.',
        group: 'calendar',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: deleteEventSchema,
        handler: async (args) => {
          const a = deleteEventSchema.parse(args);

          const res = await graph.request<unknown>({
            path: `/me/events/${encodeURIComponent(a.id)}`,
            method: 'DELETE',
            scopes: WRITE_SCOPES,
          });

          return { deleted: true, id: a.id, status: res.status };
        },
      },

      {
        name: 'calendar_respond_event',
        title: 'Respond to a meeting invitation',
        description:
          'Accepts, declines, or tentatively accepts a meeting invitation, optionally with a comment to the organizer. Graph returns no content on success. Only works on events where the signed-in user is an attendee — responding to an event they organized fails. Proposing a new time is not supported.',
        group: 'calendar',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: respondEventSchema,
        handler: async (args) => {
          const a = respondEventSchema.parse(args);

          // `response` is a closed enum, so it is safe as a path segment.
          const res = await graph.request<unknown>({
            path: `/me/events/${encodeURIComponent(a.id)}/${a.response}`,
            method: 'POST',
            body: { comment: a.comment ?? '', sendResponse: a.sendResponse },
            scopes: WRITE_SCOPES,
          });

          return { responded: a.response, id: a.id, organizerNotified: a.sendResponse, status: res.status };
        },
      },
    ];
  },
};
