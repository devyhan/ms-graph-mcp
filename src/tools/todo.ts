/**
 * Microsoft To Do tools.
 *
 * Tasks.Read and Tasks.ReadWrite exist only as delegated permissions —
 * Microsoft publishes no application permission for To Do — so every path here
 * is rooted at `/me` and an app-only deployment cannot use this group at all.
 */

import { z } from 'zod';
import type { ToolDefinition, ToolDeps, ToolGroupMeta, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { stripHtml, truncateText } from '../util/truncate.js';

// The group catalogue is a static literal; `todo` is always present.
const GROUP: ToolGroupMeta = GROUPS['todo']!;
const READ_SCOPES: string[] = [...GROUP.readScopes];
const WRITE_SCOPES: string[] = [...GROUP.readScopes, ...GROUP.writeScopes];

const LISTS_PATH = '/me/todo/lists';

/** `YYYY-MM-DD`, optionally followed by a time and an offset. */
const ISO_DATE_RE =
  /^\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2})?(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})?)?$/;

const ISO_DATE_MESSAGE = "Use 'YYYY-MM-DD' or a full ISO 8601 date-time such as '2026-09-10T14:30:00'.";

// ---------------------------------------------------------------------------
// Graph payload shapes (only the fields these tools project)
// ---------------------------------------------------------------------------

interface DateTimeTimeZone {
  dateTime?: string | null;
  timeZone?: string | null;
}

interface ItemBody {
  content?: string | null;
  contentType?: string | null;
}

interface TodoTaskList {
  id?: string;
  displayName?: string;
  isOwner?: boolean;
  isShared?: boolean;
  wellknownListName?: string;
}

interface TodoTask {
  id?: string;
  title?: string;
  body?: ItemBody | null;
  status?: string;
  importance?: string;
  isReminderOn?: boolean;
  categories?: string[] | null;
  hasAttachments?: boolean;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  dueDateTime?: DateTimeTimeZone | null;
  startDateTime?: DateTimeTimeZone | null;
  reminderDateTime?: DateTimeTimeZone | null;
  completedDateTime?: DateTimeTimeZone | null;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * To Do ids are base64-derived and can carry `+`, `/` and `=`, any of which
 * would otherwise change the shape of the URL.
 */
function seg(id: string): string {
  return encodeURIComponent(id);
}

function compact<T extends Record<string, unknown>>(value: T): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, v] of Object.entries(value)) {
    if (v !== undefined && v !== null) out[key] = v;
  }
  return out;
}

/**
 * Graph models `dueDateTime` and `reminderDateTime` as `dateTimeTimeZone`
 * complex types, so a bare ISO string is rejected with a 400. A date-only value
 * becomes midnight, which is how the To Do clients store all-day due dates.
 */
function toDateTimeTimeZone(value: string, timeZone: string): { dateTime: string; timeZone: string } {
  const raw = value.trim().replace(' ', 'T');
  if (raw.length <= 10) return { dateTime: `${raw}T00:00:00`, timeZone };

  if (!/(?:Z|[+-]\d{2}:\d{2})$/.test(raw)) return { dateTime: raw, timeZone };

  // An explicit offset wins over the timeZone argument: normalise to UTC wall
  // time so Graph does not apply a second zone shift on top of the offset.
  const ms = Date.parse(raw);
  if (Number.isNaN(ms)) throw new Error(`Could not parse date-time "${value}". ${ISO_DATE_MESSAGE}`);
  return { dateTime: new Date(ms).toISOString().slice(0, 19), timeZone: 'UTC' };
}

function readDateTime(value: DateTimeTimeZone | null | undefined): string | undefined {
  if (value === null || value === undefined) return undefined;
  const { dateTime, timeZone } = value;
  if (typeof dateTime !== 'string' || dateTime === '') return undefined;
  return typeof timeZone === 'string' && timeZone !== '' ? `${dateTime} ${timeZone}` : dateTime;
}

function readBody(body: ItemBody | null | undefined, limit: number): string | undefined {
  if (body === null || body === undefined || typeof body.content !== 'string') return undefined;
  const text = body.contentType === 'html' ? stripHtml(body.content) : body.content;
  const trimmed = text.trim();
  return trimmed === '' ? undefined : truncateText(trimmed, limit);
}

function projectList(list: TodoTaskList): Record<string, unknown> {
  return compact({
    id: list.id,
    displayName: list.displayName,
    wellknownListName:
      list.wellknownListName === 'none' ? undefined : list.wellknownListName,
    isOwner: list.isOwner === false ? false : undefined,
    isShared: list.isShared === true ? true : undefined,
  });
}

function projectTaskSummary(task: TodoTask, bodyLimit: number): Record<string, unknown> {
  const categories = task.categories;
  return compact({
    id: task.id,
    title: task.title,
    status: task.status,
    importance: task.importance === 'normal' ? undefined : task.importance,
    dueDateTime: readDateTime(task.dueDateTime),
    reminderDateTime: task.isReminderOn === true ? readDateTime(task.reminderDateTime) : undefined,
    completedDateTime: readDateTime(task.completedDateTime),
    categories: categories !== null && categories !== undefined && categories.length > 0 ? categories : undefined,
    hasAttachments: task.hasAttachments === true ? true : undefined,
    lastModifiedDateTime: task.lastModifiedDateTime,
    bodyPreview: readBody(task.body, bodyLimit),
  });
}

function projectTaskDetail(task: TodoTask, bodyLimit: number): Record<string, unknown> {
  const categories = task.categories;
  return compact({
    id: task.id,
    title: task.title,
    status: task.status,
    importance: task.importance,
    isReminderOn: task.isReminderOn,
    dueDateTime: readDateTime(task.dueDateTime),
    startDateTime: readDateTime(task.startDateTime),
    reminderDateTime: readDateTime(task.reminderDateTime),
    completedDateTime: readDateTime(task.completedDateTime),
    categories: categories !== null && categories !== undefined && categories.length > 0 ? categories : undefined,
    hasAttachments: task.hasAttachments,
    createdDateTime: task.createdDateTime,
    lastModifiedDateTime: task.lastModifiedDateTime,
    body: readBody(task.body, bodyLimit),
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
// Input schemas
// ---------------------------------------------------------------------------

const listListsSchema = z.object({});

const listTasksSchema = z.object({
  listId: z.string().min(1).describe('Id of the To Do list, as returned by todo_list_lists.'),
  top: z
    .number()
    .int()
    .min(1)
    .max(200)
    .default(50)
    .describe('Maximum number of tasks to return in this page. Defaults to 50.'),
  filter: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Raw OData $filter, ANDed with the completed-task filter. Example: \"importance eq 'high'\" or \"dueDateTime/dateTime ge '2026-01-01T00:00:00'\".",
    ),
  includeCompleted: z
    .boolean()
    .default(false)
    .describe('Include completed tasks. Defaults to false, which adds status ne \'completed\' to the filter.'),
});

const getTaskSchema = z.object({
  listId: z.string().min(1).describe('Id of the To Do list that holds the task.'),
  taskId: z.string().min(1).describe('Id of the task, as returned by todo_list_tasks.'),
});

const createTaskSchema = z.object({
  listId: z.string().min(1).describe('Id of the To Do list to create the task in.'),
  title: z.string().min(1).describe('Task title. This is the only required task field.'),
  body: z.string().optional().describe('Plain-text notes stored on the task body.'),
  dueDateTime: z
    .string()
    .regex(ISO_DATE_RE, ISO_DATE_MESSAGE)
    .optional()
    .describe(`Due date, '2026-09-10' or a full ISO 8601 date-time. ${ISO_DATE_MESSAGE}`),
  reminderDateTime: z
    .string()
    .regex(ISO_DATE_RE, ISO_DATE_MESSAGE)
    .optional()
    .describe('Reminder date-time. Supplying this also turns the reminder on.'),
  importance: z
    .enum(['low', 'normal', 'high'])
    .optional()
    .describe('Task importance. Graph defaults to normal.'),
  timeZone: z
    .string()
    .min(1)
    .default('UTC')
    .describe(
      "Time zone the due and reminder wall-clock times are expressed in, e.g. 'UTC' or 'Pacific Standard Time'. Ignored for values that already carry a UTC offset.",
    ),
});

const updateTaskSchema = z.object({
  listId: z.string().min(1).describe('Id of the To Do list that holds the task.'),
  taskId: z.string().min(1).describe('Id of the task to update.'),
  title: z.string().min(1).optional().describe('New task title.'),
  body: z.string().optional().describe('Replacement plain-text notes. Pass an empty string to clear them.'),
  status: z
    .enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'])
    .optional()
    .describe('New task status. Use todo_complete_task for the common completed case.'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('New task importance.'),
  dueDateTime: z
    .string()
    .regex(ISO_DATE_RE, ISO_DATE_MESSAGE)
    .optional()
    .describe(`New due date. ${ISO_DATE_MESSAGE}`),
  reminderDateTime: z
    .string()
    .regex(ISO_DATE_RE, ISO_DATE_MESSAGE)
    .optional()
    .describe('New reminder date-time. Supplying this also turns the reminder on.'),
  clearDueDateTime: z
    .boolean()
    .optional()
    .describe('Remove the due date. Omitting dueDateTime alone never clears it.'),
  clearReminderDateTime: z
    .boolean()
    .optional()
    .describe('Remove the reminder and turn it off.'),
  timeZone: z
    .string()
    .min(1)
    .default('UTC')
    .describe(
      "Time zone for the due and reminder wall-clock times. Ignored for values that already carry a UTC offset.",
    ),
});

const completeTaskSchema = z.object({
  listId: z.string().min(1).describe('Id of the To Do list that holds the task.'),
  taskId: z.string().min(1).describe('Id of the task to mark completed.'),
});

const deleteTaskSchema = z.object({
  listId: z.string().min(1).describe('Id of the To Do list that holds the task.'),
  taskId: z.string().min(1).describe('Id of the task to delete permanently.'),
});

const createListSchema = z.object({
  displayName: z.string().min(1).describe('Name of the new To Do list, as it appears in the To Do apps.'),
});

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const todoModule: ToolModule = {
  group: GROUP,

  build({ graph, config }: ToolDeps): ToolDefinition[] {
    // The task body is the only unbounded field, so give it a slice of the
    // output budget rather than a hard-coded guess.
    const detailBodyChars = Math.min(4000, Math.max(500, Math.floor(config.maxOutputChars / 8)));
    const previewBodyChars = 200;

    return [
      {
        name: 'todo_list_lists',
        title: 'List To Do lists',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          "Lists the signed-in user's Microsoft To Do lists with their id, display name and well-known name (for example defaultList for the built-in Tasks list). Follows up to three pages, which covers any realistic number of lists. Use the returned id as listId for every other todo tool; To Do is delegated-only, so these are always the caller's own lists.",
        inputSchema: listListsSchema,
        handler: async () => {
          const res = await graph.request({
            path: LISTS_PATH,
            query: { $top: 100 },
            scopes: READ_SCOPES,
            maxPages: 3,
          });
          const lists = extractCollection<TodoTaskList>(res.data).map(projectList);
          return collectionResult(lists, 'lists', res.nextLink);
        },
      },

      {
        name: 'todo_list_tasks',
        title: 'List To Do tasks',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          "Lists tasks in one To Do list, returning id, title, status, importance, due and reminder times and a short body preview per task, plus count and nextLink when more pages exist. Page size defaults to 50. Completed tasks are excluded unless includeCompleted is true. To Do tasks do not support $search: narrow with filter instead, and note that dueDateTime is a complex type, so a date filter reads dueDateTime/dateTime ge '2026-01-01T00:00:00'.",
        inputSchema: listTasksSchema,
        handler: async (args) => {
          const { listId, top, filter, includeCompleted } = listTasksSchema.parse(args);

          const clauses: string[] = [];
          if (!includeCompleted) clauses.push("status ne 'completed'");
          if (filter !== undefined) clauses.push(`(${filter})`);

          const res = await graph.request({
            path: `${LISTS_PATH}/${seg(listId)}/tasks`,
            query: {
              $top: top,
              $filter: clauses.length > 0 ? clauses.join(' and ') : undefined,
            },
            scopes: READ_SCOPES,
          });

          const tasks = extractCollection<TodoTask>(res.data).map((task) =>
            projectTaskSummary(task, previewBodyChars),
          );
          return collectionResult(tasks, 'tasks', res.nextLink);
        },
      },

      {
        name: 'todo_get_task',
        title: 'Get a To Do task',
        group: GROUP.name,
        scopes: READ_SCOPES,
        description:
          'Reads one To Do task in full: title, status, importance, reminder state, start/due/completed times, categories and the notes body (truncated). Needs both the list id and the task id, which come from todo_list_lists and todo_list_tasks.',
        inputSchema: getTaskSchema,
        handler: async (args) => {
          const { listId, taskId } = getTaskSchema.parse(args);
          const res = await graph.request<TodoTask>({
            path: `${LISTS_PATH}/${seg(listId)}/tasks/${seg(taskId)}`,
            scopes: READ_SCOPES,
          });
          return projectTaskDetail(res.data, detailBodyChars);
        },
      },

      {
        name: 'todo_create_task',
        title: 'Create a To Do task',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          'Creates a task in a To Do list and returns the created task. Only title is required. Due and reminder values are sent as Graph dateTimeTimeZone objects using the timeZone argument (default UTC), and a date-only dueDateTime becomes midnight, matching how the To Do apps store all-day due dates. Supplying reminderDateTime also switches the reminder on, since Graph leaves isReminderOn false otherwise.',
        inputSchema: createTaskSchema,
        handler: async (args) => {
          const parsed = createTaskSchema.parse(args);
          const { listId, title, body, dueDateTime, reminderDateTime, importance, timeZone } = parsed;

          const payload: Record<string, unknown> = { title };
          if (body !== undefined) payload['body'] = { content: body, contentType: 'text' };
          if (importance !== undefined) payload['importance'] = importance;
          if (dueDateTime !== undefined) payload['dueDateTime'] = toDateTimeTimeZone(dueDateTime, timeZone);
          if (reminderDateTime !== undefined) {
            payload['reminderDateTime'] = toDateTimeTimeZone(reminderDateTime, timeZone);
            payload['isReminderOn'] = true;
          }

          const res = await graph.request<TodoTask>({
            path: `${LISTS_PATH}/${seg(listId)}/tasks`,
            method: 'POST',
            body: payload,
            scopes: WRITE_SCOPES,
          });
          return projectTaskDetail(res.data, detailBodyChars);
        },
      },

      {
        name: 'todo_update_task',
        title: 'Update a To Do task',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          'Patches an existing To Do task and returns the updated task. Only the supplied fields change; omitting dueDateTime or reminderDateTime leaves the existing value alone, so use clearDueDateTime or clearReminderDateTime to remove one. Due and reminder values are sent as Graph dateTimeTimeZone objects using the timeZone argument (default UTC).',
        inputSchema: updateTaskSchema,
        handler: async (args) => {
          const parsed = updateTaskSchema.parse(args);
          const {
            listId,
            taskId,
            title,
            body,
            status,
            importance,
            dueDateTime,
            reminderDateTime,
            clearDueDateTime,
            clearReminderDateTime,
            timeZone,
          } = parsed;

          const payload: Record<string, unknown> = {};
          if (title !== undefined) payload['title'] = title;
          if (body !== undefined) payload['body'] = { content: body, contentType: 'text' };
          if (status !== undefined) payload['status'] = status;
          if (importance !== undefined) payload['importance'] = importance;

          if (clearDueDateTime === true) payload['dueDateTime'] = null;
          else if (dueDateTime !== undefined) payload['dueDateTime'] = toDateTimeTimeZone(dueDateTime, timeZone);

          if (clearReminderDateTime === true) {
            payload['reminderDateTime'] = null;
            payload['isReminderOn'] = false;
          } else if (reminderDateTime !== undefined) {
            payload['reminderDateTime'] = toDateTimeTimeZone(reminderDateTime, timeZone);
            payload['isReminderOn'] = true;
          }

          if (Object.keys(payload).length === 0) {
            throw new Error(
              'todo_update_task needs at least one field to change: title, body, status, importance, dueDateTime, reminderDateTime, clearDueDateTime or clearReminderDateTime.',
            );
          }

          const res = await graph.request<TodoTask>({
            path: `${LISTS_PATH}/${seg(listId)}/tasks/${seg(taskId)}`,
            method: 'PATCH',
            body: payload,
            scopes: WRITE_SCOPES,
          });
          return projectTaskDetail(res.data, detailBodyChars);
        },
      },

      {
        name: 'todo_complete_task',
        title: 'Complete a To Do task',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          "Marks a To Do task completed by patching its status to 'completed', and returns the updated task including the completedDateTime Graph stamps on it. Recurring tasks are rolled forward by the service rather than closed outright.",
        inputSchema: completeTaskSchema,
        handler: async (args) => {
          const { listId, taskId } = completeTaskSchema.parse(args);
          const res = await graph.request<TodoTask>({
            path: `${LISTS_PATH}/${seg(listId)}/tasks/${seg(taskId)}`,
            method: 'PATCH',
            body: { status: 'completed' },
            scopes: WRITE_SCOPES,
          });
          return projectTaskDetail(res.data, detailBodyChars);
        },
      },

      {
        name: 'todo_delete_task',
        title: 'Delete a To Do task',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          'Deletes a To Do task permanently and returns a confirmation. There is no recycle bin for To Do tasks, so prefer todo_complete_task when the user only wants it out of the way.',
        inputSchema: deleteTaskSchema,
        handler: async (args) => {
          const { listId, taskId } = deleteTaskSchema.parse(args);
          const res = await graph.request({
            path: `${LISTS_PATH}/${seg(listId)}/tasks/${seg(taskId)}`,
            method: 'DELETE',
            scopes: WRITE_SCOPES,
          });
          return { deleted: true, listId, taskId, status: res.status };
        },
      },

      {
        name: 'todo_create_list',
        title: 'Create a To Do list',
        group: GROUP.name,
        write: true,
        scopes: WRITE_SCOPES,
        description:
          'Creates a new To Do list and returns its id and display name. Display names are not required to be unique, so check todo_list_lists first if the user expects one list per name.',
        inputSchema: createListSchema,
        handler: async (args) => {
          const { displayName } = createListSchema.parse(args);
          const res = await graph.request<TodoTaskList>({
            path: LISTS_PATH,
            method: 'POST',
            body: { displayName },
            scopes: WRITE_SCOPES,
          });
          return projectList(res.data);
        },
      },
    ];
  },
};
