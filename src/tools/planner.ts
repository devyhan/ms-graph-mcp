/**
 * Microsoft Planner tools.
 *
 * Two Planner-specific behaviours shape everything below:
 *
 * 1. Every PATCH and DELETE needs an `If-Match` header carrying the item's
 *    current `@odata.etag`. The handlers here GET the item first and send that
 *    etag automatically, so callers never have to see one.
 * 2. Planner collection endpoints do not reliably honour `$top`/`$skip`, so
 *    sending one risks a hard 400. Result caps are therefore applied to the
 *    parsed response instead, and `maxPages: 1` keeps a single server page.
 */

import { z } from 'zod';

import type { GraphClient, ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { extractCollection } from '../graph/client.js';
import { isoDate } from '../util/odata.js';
import { truncateText } from '../util/truncate.js';
import { GROUPS } from './groups.js';

const GROUP = GROUPS['planner'];
if (GROUP === undefined) throw new Error('Tool group "planner" is missing from GROUPS.');

const READ_SCOPES: string[] = GROUP.readScopes;
const WRITE_SCOPES: string[] = [...new Set([...GROUP.readScopes, ...GROUP.writeScopes])];

/** Planner payloads are small; a low cap keeps a plan-wide listing readable. */
const DEFAULT_LIMIT = 25;
const MAX_LIMIT = 100;
/** Task descriptions are free text and occasionally enormous. */
const DESCRIPTION_MAX_CHARS = 4000;

const ASSIGNMENT_TYPE = '#microsoft.graph.plannerAssignment';
/** Planner order hints are lexicographic; ' !' sorts an item to the top. */
const TOP_ORDER_HINT = ' !';

// ---------------------------------------------------------------------------
// Graph shapes (only the fields these tools project)
// ---------------------------------------------------------------------------

interface PlannerPlan {
  id?: string;
  title?: string;
  owner?: string;
  createdDateTime?: string;
  container?: { type?: string; containerId?: string; url?: string };
  '@odata.etag'?: string;
}

interface PlannerBucket {
  id?: string;
  name?: string;
  planId?: string;
  orderHint?: string;
}

interface PlannerTask {
  id?: string;
  planId?: string;
  bucketId?: string;
  title?: string;
  percentComplete?: number;
  priority?: number;
  startDateTime?: string | null;
  dueDateTime?: string | null;
  completedDateTime?: string | null;
  createdDateTime?: string;
  hasDescription?: boolean;
  checklistItemCount?: number;
  activeChecklistItemCount?: number;
  referenceCount?: number;
  assignments?: Record<string, unknown>;
  '@odata.etag'?: string;
}

interface PlannerChecklistItem {
  title?: string;
  isChecked?: boolean;
  orderHint?: string;
}

interface PlannerReference {
  alias?: string;
  type?: string;
  previewPriority?: string;
}

interface PlannerTaskDetails {
  id?: string;
  description?: string | null;
  previewType?: string;
  checklist?: Record<string, unknown>;
  references?: Record<string, unknown>;
}

interface Assignment {
  '@odata.type': string;
  orderHint: string;
}

// ---------------------------------------------------------------------------
// Projections
// ---------------------------------------------------------------------------

/** Graph mixes `@odata.*` annotations into open maps; those are never user ids. */
function mapKeys(map: Record<string, unknown> | undefined): string[] {
  if (map === null || typeof map !== 'object') return [];
  return Object.keys(map).filter((key) => !key.startsWith('@'));
}

function projectPlan(plan: PlannerPlan): Record<string, unknown> {
  return {
    id: plan.id,
    title: plan.title,
    // Planner plans hang off a Microsoft 365 group; `owner` is that group's id.
    groupId: plan.owner,
    containerType: plan.container?.type,
    containerId: plan.container?.containerId,
    createdDateTime: plan.createdDateTime,
  };
}

function projectBucket(bucket: PlannerBucket): Record<string, unknown> {
  return {
    id: bucket.id,
    name: bucket.name,
    planId: bucket.planId,
    orderHint: bucket.orderHint,
  };
}

function projectTask(task: PlannerTask): Record<string, unknown> {
  return {
    id: task.id,
    title: task.title,
    planId: task.planId,
    bucketId: task.bucketId,
    percentComplete: task.percentComplete,
    priority: task.priority,
    startDateTime: task.startDateTime ?? undefined,
    dueDateTime: task.dueDateTime ?? undefined,
    completedDateTime: task.completedDateTime ?? undefined,
    createdDateTime: task.createdDateTime,
    assigneeIds: mapKeys(task.assignments),
    hasDescription: task.hasDescription,
    checklistItemCount: task.checklistItemCount,
    activeChecklistItemCount: task.activeChecklistItemCount,
    referenceCount: task.referenceCount,
  };
}

function projectDetails(details: PlannerTaskDetails): Record<string, unknown> {
  const checklist = mapKeys(details.checklist).map((id) => {
    const item = (details.checklist?.[id] ?? {}) as PlannerChecklistItem;
    return { id, title: item.title, isChecked: item.isChecked === true };
  });

  const references = mapKeys(details.references).map((encoded) => {
    const ref = (details.references?.[encoded] ?? {}) as PlannerReference;
    return { url: decodeUrlKey(encoded), alias: ref.alias, type: ref.type };
  });

  const description = typeof details.description === 'string' ? details.description : undefined;
  return {
    description: description === undefined ? undefined : truncateText(description, DESCRIPTION_MAX_CHARS),
    previewType: details.previewType,
    checklist,
    references,
  };
}

const HAS_OFFSET = /(?:Z|[+-]\d{2}:\d{2})$/;

/**
 * Planner date properties are `DateTimeOffset`. `isoDate` validates but happily
 * passes through a bare `2026-01-31` or an offset-less `2026-01-31T09:00`, both
 * of which Graph rejects, so fill in the missing parts as UTC here.
 */
function plannerDateTime(value: string): string {
  const normalized = isoDate(value);
  if (!normalized.includes('T')) return `${normalized}T00:00:00Z`;
  return HAS_OFFSET.test(normalized) ? normalized : `${normalized}Z`;
}

/** Reference keys are percent-encoded URLs; a malformed one must not throw. */
function decodeUrlKey(key: string): string {
  try {
    return decodeURIComponent(key);
  } catch {
    return key;
  }
}

interface Page<T> {
  count: number;
  items: T[];
  truncated?: boolean;
  fetched?: number;
  nextLink?: string;
}

function page<T>(all: T[], limit: number, nextLink: string | undefined): Page<T> {
  const items = all.slice(0, limit);
  const result: Page<T> = { count: items.length, items };
  if (all.length > items.length) {
    result.truncated = true;
    result.fetched = all.length;
  }
  if (nextLink !== undefined) result.nextLink = nextLink;
  return result;
}

// ---------------------------------------------------------------------------
// Write helpers
// ---------------------------------------------------------------------------

function taskPath(taskId: string): string {
  // Planner ids are base64url-ish, but encoding costs nothing and stops a
  // hand-typed id containing `/` or `?` from rewriting the URL.
  return `/planner/tasks/${encodeURIComponent(taskId)}`;
}

function assignment(): Assignment {
  return { '@odata.type': ASSIGNMENT_TYPE, orderHint: TOP_ORDER_HINT };
}

function newAssignments(assigneeIds: readonly string[]): Record<string, Assignment> {
  const out: Record<string, Assignment> = {};
  for (const raw of assigneeIds) {
    const id = raw.trim();
    if (id.length > 0) out[id] = assignment();
  }
  return out;
}

/**
 * Planner treats `assignments` as an open map patch, not a replacement: an
 * omitted user stays assigned. Turning the caller's "these are the assignees"
 * intent into that shape means explicitly nulling everyone who dropped off.
 */
function assignmentPatch(
  assigneeIds: readonly string[],
  current: Record<string, unknown> | undefined,
): Record<string, Assignment | null> {
  const wanted = new Set(assigneeIds.map((id) => id.trim()).filter((id) => id.length > 0));
  const patch: Record<string, Assignment | null> = {};
  for (const id of wanted) patch[id] = assignment();
  for (const id of mapKeys(current)) {
    if (!wanted.has(id)) patch[id] = null;
  }
  return patch;
}

/** Reads a task purely to obtain the etag every Planner write must echo back. */
async function readTaskForWrite(
  graph: GraphClient,
  taskId: string,
): Promise<{ etag: string; task: PlannerTask }> {
  const res = await graph.request<PlannerTask>({
    path: taskPath(taskId),
    scopes: WRITE_SCOPES,
  });
  const task = res.data ?? {};
  const etag = task['@odata.etag'];
  if (typeof etag !== 'string' || etag.length === 0) {
    throw new Error(
      `Planner task ${taskId} was returned without an @odata.etag, so the required If-Match header cannot be built. Retry, or confirm the id names a task rather than a plan or bucket.`,
    );
  }
  return { etag, task };
}

// ---------------------------------------------------------------------------
// Schemas
// ---------------------------------------------------------------------------

const limitField = z
  .number()
  .int()
  .min(1)
  .max(MAX_LIMIT)
  .default(DEFAULT_LIMIT)
  .describe(`Maximum items to return, 1-${MAX_LIMIT}. Applied after the response arrives.`);

const planIdField = z
  .string()
  .min(1)
  .describe('Planner plan id, as returned by planner_list_my_plans.');

const taskIdField = z
  .string()
  .min(1)
  .describe('Planner task id, as returned by planner_list_my_tasks or planner_list_plan_tasks.');

const assigneeIdsField = z
  .array(z.string().min(1))
  .describe('Microsoft Entra user object ids (GUIDs, not email addresses) to assign the task to.');

const dueDateTimeField = z
  .string()
  .min(1)
  .describe('Due date as ISO-8601, e.g. 2026-01-31T17:00:00Z. A bare YYYY-MM-DD is accepted and becomes midnight UTC.');

const percentCompleteField = z
  .number()
  .int()
  .min(0)
  .max(100)
  .describe('Progress: 0 = not started, 1-99 = in progress, 100 = completed. Planner has no separate status field.');

const priorityField = z
  .number()
  .int()
  .min(0)
  .max(10)
  .describe('Priority 0-10. Planner shows 1 as urgent, 3 important, 5 medium, 9 low.');

const ListMyPlansSchema = z.object({ limit: limitField });
const ListMyTasksSchema = z.object({ limit: limitField });
const GetPlanSchema = z.object({ planId: planIdField });
const ListBucketsSchema = z.object({ planId: planIdField, limit: limitField });
const ListPlanTasksSchema = z.object({ planId: planIdField, limit: limitField });

const GetTaskSchema = z.object({
  taskId: taskIdField,
  includeDetails: z
    .boolean()
    .default(false)
    .describe('Also fetch the task details (description, checklist, attachment references) and merge them in. Costs one extra Graph call.'),
});

const CreateTaskSchema = z.object({
  planId: planIdField,
  title: z.string().min(1).describe('Task title.'),
  bucketId: z
    .string()
    .min(1)
    .optional()
    .describe('Bucket to file the task under. Omit to leave it unbucketed; use planner_list_buckets to find one.'),
  assigneeIds: assigneeIdsField.optional(),
  dueDateTime: dueDateTimeField.optional(),
  percentComplete: percentCompleteField.optional(),
  priority: priorityField.optional(),
});

const UpdateTaskSchema = z.object({
  taskId: taskIdField,
  title: z.string().min(1).optional().describe('New task title.'),
  bucketId: z.string().min(1).optional().describe('Move the task into this bucket.'),
  assigneeIds: assigneeIdsField
    .optional()
    .describe('Replaces the assignee set: anyone currently assigned but absent from this list is unassigned. Pass [] to clear all assignees.'),
  dueDateTime: dueDateTimeField.optional(),
  clearDueDateTime: z
    .boolean()
    .optional()
    .describe('Remove the due date. Overrides dueDateTime when both are given.'),
  percentComplete: percentCompleteField.optional(),
  priority: priorityField.optional(),
});

const DeleteTaskSchema = z.object({ taskId: taskIdField });

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

type Handler = ToolDefinition['handler'];

function bind<T>(schema: z.ZodType<T>, run: (args: T) => Promise<unknown>): Handler {
  return async (args: Record<string, unknown>): Promise<unknown> => run(schema.parse(args));
}

export const plannerModule: ToolModule = {
  group: GROUP,
  build({ graph }: ToolDeps): ToolDefinition[] {
    return [
      {
        name: 'planner_list_my_plans',
        title: 'List my Planner plans',
        description:
          `Lists the Planner plans the signed-in user can see, with id, title, owning Microsoft 365 group id and container. Returns up to ${DEFAULT_LIMIT} by default. Only plans in groups the user belongs to appear; a plan shared by link but not joined will be missing.`,
        group: 'planner',
        scopes: READ_SCOPES,
        inputSchema: ListMyPlansSchema,
        handler: bind(ListMyPlansSchema, async (a) => {
          const res = await graph.request({ path: '/me/planner/plans', scopes: READ_SCOPES, maxPages: 1 });
          const plans = extractCollection<PlannerPlan>(res.data).map(projectPlan);
          return page(plans, a.limit, res.nextLink);
        }),
      },
      {
        name: 'planner_list_my_tasks',
        title: 'List my Planner tasks',
        description:
          `Lists Planner tasks assigned to the signed-in user across every plan, with title, planId, bucketId, due date, percentComplete and assignee ids. Returns up to ${DEFAULT_LIMIT} by default. Completed tasks are included; filter on percentComplete === 100 to separate them.`,
        group: 'planner',
        scopes: READ_SCOPES,
        inputSchema: ListMyTasksSchema,
        handler: bind(ListMyTasksSchema, async (a) => {
          const res = await graph.request({ path: '/me/planner/tasks', scopes: READ_SCOPES, maxPages: 1 });
          const tasks = extractCollection<PlannerTask>(res.data).map(projectTask);
          return page(tasks, a.limit, res.nextLink);
        }),
      },
      {
        name: 'planner_get_plan',
        title: 'Get a Planner plan',
        description:
          'Returns one Planner plan: id, title, owning Microsoft 365 group id, container and creation time. Use planner_list_buckets and planner_list_plan_tasks for its contents.',
        group: 'planner',
        scopes: READ_SCOPES,
        inputSchema: GetPlanSchema,
        handler: bind(GetPlanSchema, async (a) => {
          const res = await graph.request<PlannerPlan>({
            path: `/planner/plans/${encodeURIComponent(a.planId)}`,
            scopes: READ_SCOPES,
          });
          return projectPlan(res.data ?? {});
        }),
      },
      {
        name: 'planner_list_buckets',
        title: 'List Planner buckets',
        description:
          `Lists the buckets (columns) of a plan with id, name and orderHint. Returns up to ${DEFAULT_LIMIT} by default. Bucket order follows orderHint lexicographically, not the response order.`,
        group: 'planner',
        scopes: READ_SCOPES,
        inputSchema: ListBucketsSchema,
        handler: bind(ListBucketsSchema, async (a) => {
          const res = await graph.request({
            path: `/planner/plans/${encodeURIComponent(a.planId)}/buckets`,
            scopes: READ_SCOPES,
            maxPages: 1,
          });
          const buckets = extractCollection<PlannerBucket>(res.data).map(projectBucket);
          return page(buckets, a.limit, res.nextLink);
        }),
      },
      {
        name: 'planner_list_plan_tasks',
        title: 'List tasks in a Planner plan',
        description:
          `Lists every task in a plan regardless of assignee, with title, bucketId, due date, percentComplete and assignee ids. Returns up to ${DEFAULT_LIMIT} by default. Descriptions and checklists are not included; call planner_get_task with includeDetails for those.`,
        group: 'planner',
        scopes: READ_SCOPES,
        inputSchema: ListPlanTasksSchema,
        handler: bind(ListPlanTasksSchema, async (a) => {
          const res = await graph.request({
            path: `/planner/plans/${encodeURIComponent(a.planId)}/tasks`,
            scopes: READ_SCOPES,
            maxPages: 1,
          });
          const tasks = extractCollection<PlannerTask>(res.data).map(projectTask);
          return page(tasks, a.limit, res.nextLink);
        }),
      },
      {
        name: 'planner_get_task',
        title: 'Get a Planner task',
        description:
          'Returns one Planner task. With includeDetails the task details resource is fetched too and merged in, adding the description, checklist items and attachment references.',
        group: 'planner',
        scopes: READ_SCOPES,
        inputSchema: GetTaskSchema,
        handler: bind(GetTaskSchema, async (a) => {
          const res = await graph.request<PlannerTask>({ path: taskPath(a.taskId), scopes: READ_SCOPES });
          const task = projectTask(res.data ?? {});
          if (!a.includeDetails) return task;

          const detailsRes = await graph.request<PlannerTaskDetails>({
            path: `${taskPath(a.taskId)}/details`,
            scopes: READ_SCOPES,
          });
          return { ...task, ...projectDetails(detailsRes.data ?? {}) };
        }),
      },
      {
        name: 'planner_create_task',
        title: 'Create a Planner task',
        description:
          'Creates a task in a plan and returns it. Assignees are Entra user object ids, not email addresses. The description and checklist cannot be set here: create the task, then set them through the task details resource.',
        group: 'planner',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: CreateTaskSchema,
        handler: bind(CreateTaskSchema, async (a) => {
          const body: Record<string, unknown> = { planId: a.planId, title: a.title };
          if (a.bucketId !== undefined) body['bucketId'] = a.bucketId;
          if (a.assigneeIds !== undefined) body['assignments'] = newAssignments(a.assigneeIds);
          if (a.dueDateTime !== undefined) body['dueDateTime'] = plannerDateTime(a.dueDateTime);
          if (a.percentComplete !== undefined) body['percentComplete'] = a.percentComplete;
          if (a.priority !== undefined) body['priority'] = a.priority;

          const res = await graph.request<PlannerTask>({
            path: '/planner/tasks',
            method: 'POST',
            body,
            scopes: WRITE_SCOPES,
          });
          return { created: true, ...projectTask(res.data ?? {}) };
        }),
      },
      {
        name: 'planner_update_task',
        title: 'Update a Planner task',
        description:
          'Updates a Planner task and returns it. The task is re-read first so the required If-Match etag is supplied automatically; a 412 means someone else edited the task in between, so re-read and retry. assigneeIds replaces the whole assignee set.',
        group: 'planner',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: UpdateTaskSchema,
        handler: bind(UpdateTaskSchema, async (a) => {
          const { etag, task } = await readTaskForWrite(graph, a.taskId);

          const body: Record<string, unknown> = {};
          if (a.title !== undefined) body['title'] = a.title;
          if (a.bucketId !== undefined) body['bucketId'] = a.bucketId;
          if (a.assigneeIds !== undefined) body['assignments'] = assignmentPatch(a.assigneeIds, task.assignments);
          if (a.clearDueDateTime === true) body['dueDateTime'] = null;
          else if (a.dueDateTime !== undefined) body['dueDateTime'] = plannerDateTime(a.dueDateTime);
          if (a.percentComplete !== undefined) body['percentComplete'] = a.percentComplete;
          if (a.priority !== undefined) body['priority'] = a.priority;

          if (Object.keys(body).length === 0) {
            throw new Error('planner_update_task needs at least one field to change besides taskId.');
          }

          const res = await graph.request<PlannerTask | null>({
            path: taskPath(a.taskId),
            method: 'PATCH',
            body,
            // Planner answers PATCH with 204 unless representation is requested.
            headers: { 'If-Match': etag, Prefer: 'return=representation' },
            scopes: WRITE_SCOPES,
          });

          const updated = res.data;
          if (updated !== null && typeof updated === 'object' && typeof updated.id === 'string') {
            return { updated: true, ...projectTask(updated) };
          }
          return { updated: true, id: a.taskId, changed: Object.keys(body) };
        }),
      },
      {
        name: 'planner_delete_task',
        title: 'Delete a Planner task',
        description:
          'Permanently deletes a Planner task; there is no recycle bin. The task is re-read first so the required If-Match etag is supplied automatically; a 412 means the task changed in between.',
        group: 'planner',
        scopes: WRITE_SCOPES,
        write: true,
        inputSchema: DeleteTaskSchema,
        handler: bind(DeleteTaskSchema, async (a) => {
          const { etag, task } = await readTaskForWrite(graph, a.taskId);
          await graph.request({
            path: taskPath(a.taskId),
            method: 'DELETE',
            headers: { 'If-Match': etag },
            scopes: WRITE_SCOPES,
          });
          return { deleted: true, id: a.taskId, title: task.title };
        }),
      },
    ];
  },
};
