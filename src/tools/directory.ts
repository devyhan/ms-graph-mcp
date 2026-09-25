/**
 * The `directory` group: read-only lookups of Microsoft Entra ID users and
 * groups.
 *
 * Every scope here (`User.Read.All`, `Group.Read.All`) needs tenant admin
 * consent, which is why the group meta carries `requiresAdminConsent` and the
 * whole group stays off unless `--org-mode` is passed.
 *
 * Directory collections are not Outlook collections: `$skip` is not supported,
 * paging runs on an opaque `$skiptoken`, and anything beyond a trivial
 * `$filter` counts as an "advanced query" that Graph rejects with a 400 unless
 * the request carries `ConsistencyLevel: eventual` *and* `$count=true`. The
 * list tools therefore always send `$count=true`; the client derives the header
 * from it.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { quoteSearch } from '../util/odata.js';
import { truncateText } from '../util/truncate.js';

const GROUP = GROUPS['directory']!;

/**
 * Per-tool `scopes` are what the login path turns into the consent set, so a
 * user lookup that asked for `Group.Read.All` would widen the token for
 * nothing. The subsets are derived from the group meta rather than retyped, and
 * fall back to the whole read set if the catalogue ever stops carrying a
 * matching scope — a too-wide token beats a token with no directory scope at all.
 */
function readScopesFor(prefix: string): string[] {
  const subset = GROUP.readScopes.filter((scope) => scope.startsWith(prefix));
  return subset.length > 0 ? subset : [...GROUP.readScopes];
}

const USER_SCOPES = readScopesFor('User.');
const GROUP_SCOPES = readScopesFor('Group.');

/** `getMemberGroups` reads a user *and* the groups that user belongs to. */
const USER_AND_GROUP_SCOPES = [...new Set([...USER_SCOPES, ...GROUP_SCOPES])];

/** Group descriptions are free text and occasionally hold a pasted policy document. */
const DESCRIPTION_MAX_CHARS = 500;

/** Graph's own ceiling for `$top` on directory collections. */
const MAX_TOP = 999;

/** The properties worth spending tokens on for a user row in a list. */
const USER_LIST_SELECT = [
  'id',
  'displayName',
  'userPrincipalName',
  'mail',
  'jobTitle',
  'department',
  'accountEnabled',
];

/**
 * A single user view. `/users/{id}` without a `$select` returns a fixed subset
 * that omits most of this, so the detail read has to name every property it
 * wants. Deliberately excluded: `proxyAddresses`, `assignedLicenses` and
 * `assignedPlans`, which are long, noisy, and never what the question was.
 */
const USER_DETAIL_SELECT = [
  ...USER_LIST_SELECT,
  'givenName',
  'surname',
  'mailNickname',
  'companyName',
  'officeLocation',
  'employeeId',
  'employeeType',
  'userType',
  'businessPhones',
  'mobilePhone',
  'streetAddress',
  'city',
  'state',
  'postalCode',
  'country',
  'usageLocation',
  'preferredLanguage',
  'createdDateTime',
  'onPremisesSyncEnabled',
  'onPremisesSamAccountName',
];

const GROUP_LIST_SELECT = [
  'id',
  'displayName',
  'description',
  'mail',
  'groupTypes',
  'securityEnabled',
  'mailEnabled',
  'visibility',
];

const GROUP_DETAIL_SELECT = [
  ...GROUP_LIST_SELECT,
  'mailNickname',
  'classification',
  'createdDateTime',
  'renewedDateTime',
  'expirationDateTime',
  'membershipRule',
  'membershipRuleProcessingState',
  'isAssignableToRole',
  'resourceProvisioningOptions',
  'onPremisesSyncEnabled',
  'onPremisesSamAccountName',
];

/**
 * The `$search` fields Graph indexes for the `user` resource. A term is matched
 * per field, so the clauses are OR-ed together to get "name or address".
 */
const USER_SEARCH_FIELDS = ['displayName', 'mail'] as const;

/**
 * A caller-supplied `$select` reaches Graph as a raw query value. Every
 * selectable property on `user` and `group` is a bare identifier, so anything
 * else is a mistake or an attempt to smuggle extra OData into the query string.
 */
const SELECT_FIELD = /^[A-Za-z][A-Za-z0-9]*$/;

/** One `$orderby` clause: a bare property name, optionally `asc` or `desc`. */
const ORDERBY_CLAUSE = /^[A-Za-z][A-Za-z0-9]*(?:\s+(?:asc|desc))?$/;

// ---------------------------------------------------------------------------
// Graph payload shapes (only the parts that are projected)
// ---------------------------------------------------------------------------

interface UserPayload {
  '@odata.type'?: string | null;
  id?: string | null;
  displayName?: string | null;
  givenName?: string | null;
  surname?: string | null;
  userPrincipalName?: string | null;
  mail?: string | null;
  mailNickname?: string | null;
  jobTitle?: string | null;
  department?: string | null;
  companyName?: string | null;
  officeLocation?: string | null;
  employeeId?: string | null;
  employeeType?: string | null;
  userType?: string | null;
  accountEnabled?: boolean | null;
  businessPhones?: string[] | null;
  mobilePhone?: string | null;
  streetAddress?: string | null;
  city?: string | null;
  state?: string | null;
  postalCode?: string | null;
  country?: string | null;
  usageLocation?: string | null;
  preferredLanguage?: string | null;
  createdDateTime?: string | null;
  onPremisesSyncEnabled?: boolean | null;
  onPremisesSamAccountName?: string | null;
}

interface GroupPayload {
  '@odata.type'?: string | null;
  id?: string | null;
  displayName?: string | null;
  description?: string | null;
  mail?: string | null;
  mailNickname?: string | null;
  mailEnabled?: boolean | null;
  securityEnabled?: boolean | null;
  groupTypes?: string[] | null;
  visibility?: string | null;
  classification?: string | null;
  createdDateTime?: string | null;
  renewedDateTime?: string | null;
  expirationDateTime?: string | null;
  membershipRule?: string | null;
  membershipRuleProcessingState?: string | null;
  isAssignableToRole?: boolean | null;
  resourceProvisioningOptions?: string[] | null;
  onPremisesSyncEnabled?: boolean | null;
  onPremisesSamAccountName?: string | null;
}

/**
 * `members`, `directReports` and `manager` are `directoryObject` collections:
 * a single response can mix users, groups, devices and service principals, and
 * only `@odata.type` says which is which.
 */
type DirectoryObjectPayload = UserPayload & GroupPayload;

// ---------------------------------------------------------------------------
// Projection helpers
// ---------------------------------------------------------------------------

function asRecord(value: unknown): Record<string, unknown> | undefined {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return undefined;
  return value as Record<string, unknown>;
}

/**
 * Graph pads a directory object with a null for every property that was asked
 * for and never set, which in a typical tenant is most of the payload. Dropping
 * them means a missing key reads as "not set", not "not requested".
 */
function dropEmpty(obj: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(obj)) {
    if (value === undefined || value === null || value === '') continue;
    if (Array.isArray(value) && value.length === 0) continue;
    out[key] = value;
  }
  return out;
}

/** Keeps only the named properties of a raw Graph object, dropping nulls. */
function pick(data: unknown, fields: readonly string[]): Record<string, unknown> {
  const record = asRecord(data);
  if (record === undefined) return {};
  const out: Record<string, unknown> = {};
  for (const field of fields) {
    const value = record[field];
    if (value !== undefined && value !== null && value !== '') out[field] = value;
  }
  return out;
}

function trimmed(value: string | null | undefined): string | undefined {
  if (typeof value !== 'string') return undefined;
  const text = value.trim();
  return text.length === 0 ? undefined : text;
}

/** `#microsoft.graph.user` -> `user`, so a mixed collection is readable at a glance. */
function kindOf(entry: { '@odata.type'?: string | null }): string | undefined {
  const raw = trimmed(entry['@odata.type']);
  if (raw === undefined) return undefined;
  return trimmed(raw.replace(/^#?microsoft\.graph\./i, ''));
}

function projectUserSummary(user: UserPayload): Record<string, unknown> {
  return dropEmpty({
    id: user.id,
    displayName: user.displayName,
    userPrincipalName: user.userPrincipalName,
    mail: user.mail,
    jobTitle: user.jobTitle,
    department: user.department,
    accountEnabled: user.accountEnabled,
  });
}

function projectUserDetail(user: UserPayload): Record<string, unknown> {
  return dropEmpty({
    ...projectUserSummary(user),
    givenName: user.givenName,
    surname: user.surname,
    mailNickname: user.mailNickname,
    companyName: user.companyName,
    officeLocation: user.officeLocation,
    employeeId: user.employeeId,
    employeeType: user.employeeType,
    userType: user.userType,
    mobilePhone: user.mobilePhone,
    businessPhones: user.businessPhones,
    streetAddress: user.streetAddress,
    city: user.city,
    state: user.state,
    postalCode: user.postalCode,
    country: user.country,
    usageLocation: user.usageLocation,
    preferredLanguage: user.preferredLanguage,
    createdDateTime: user.createdDateTime,
    onPremisesSyncEnabled: user.onPremisesSyncEnabled,
    onPremisesSamAccountName: user.onPremisesSamAccountName,
  });
}

function projectGroupSummary(group: GroupPayload): Record<string, unknown> {
  const description = trimmed(group.description);
  return dropEmpty({
    id: group.id,
    displayName: group.displayName,
    description: description === undefined ? undefined : truncateText(description, DESCRIPTION_MAX_CHARS),
    mail: group.mail,
    groupTypes: group.groupTypes,
    securityEnabled: group.securityEnabled,
    mailEnabled: group.mailEnabled,
    visibility: group.visibility,
  });
}

function projectGroupDetail(group: GroupPayload): Record<string, unknown> {
  return dropEmpty({
    ...projectGroupSummary(group),
    mailNickname: group.mailNickname,
    classification: group.classification,
    createdDateTime: group.createdDateTime,
    renewedDateTime: group.renewedDateTime,
    expirationDateTime: group.expirationDateTime,
    membershipRule: group.membershipRule,
    membershipRuleProcessingState: group.membershipRuleProcessingState,
    isAssignableToRole: group.isAssignableToRole,
    resourceProvisioningOptions: group.resourceProvisioningOptions,
    onPremisesSyncEnabled: group.onPremisesSyncEnabled,
    onPremisesSamAccountName: group.onPremisesSamAccountName,
  });
}

/** One entry of a heterogeneous `directoryObject` collection. */
function projectDirectoryObject(entry: DirectoryObjectPayload): Record<string, unknown> {
  const kind = kindOf(entry);
  if (kind === 'group') {
    return dropEmpty({ type: kind, ...projectGroupSummary(entry) });
  }
  return dropEmpty({
    type: kind,
    ...projectUserSummary(entry),
    // Devices and service principals have neither a UPN nor a job title, so the
    // user projection above yields just id and displayName for them.
  });
}

/**
 * Directory collections page on an opaque `$skiptoken`, never on `$skip`.
 * Handing the token back separately saves the model from having to parse a URL
 * to continue a listing.
 */
function skipTokenFrom(nextLink: string | undefined): string | undefined {
  if (nextLink === undefined) return undefined;
  let url: URL;
  try {
    url = new URL(nextLink);
  } catch {
    return undefined;
  }
  const params = url.searchParams;
  return trimmed(params.get('$skiptoken') ?? params.get('$skipToken'));
}

function collectionResult<T>(
  itemsKey: string,
  data: unknown,
  nextLink: string | undefined,
  project: (item: T) => Record<string, unknown>,
): Record<string, unknown> {
  const items = extractCollection<T>(data);
  const total = asRecord(data)?.['@odata.count'];
  const out: Record<string, unknown> = {
    count: items.length,
    [itemsKey]: items.map(project),
  };
  if (typeof total === 'number') out['totalCount'] = total;
  if (nextLink !== undefined) out['nextLink'] = nextLink;
  const token = skipTokenFrom(nextLink);
  if (token !== undefined) out['skipToken'] = token;
  return out;
}

/**
 * Ids reach the URL encoded because a UPN is a legal id here, and a guest UPN
 * carries `#` (`alice_contoso.com#EXT#@fabrikam.onmicrosoft.com`) which would
 * otherwise truncate the path at the fragment.
 */
function userPath(id: string, suffix = ''): string {
  return `/users/${encodeURIComponent(id)}${suffix}`;
}

function groupPath(id: string, suffix = ''): string {
  return `/groups/${encodeURIComponent(id)}${suffix}`;
}

/**
 * Builds the `$search` expression for a term. The term is quoted rather than
 * concatenated so one containing a space, a quote or the literal `OR` cannot
 * break out of its clause.
 */
function userSearchExpression(term: string): string {
  return USER_SEARCH_FIELDS.map((field) => quoteSearch(`${field}:${term}`)).join(' OR ');
}

// ---------------------------------------------------------------------------
// Shared argument pieces
// ---------------------------------------------------------------------------

const userId = z
  .string()
  .min(1)
  .max(256)
  .describe(
    'The user\'s object id (a GUID) or userPrincipalName, e.g. "irene@contoso.com". ' +
      'Guest UPNs containing "#EXT#" are accepted as-is; they are URL-encoded here.',
  );

const groupId = z
  .string()
  .min(1)
  .max(256)
  .describe('The group\'s object id (a GUID), as returned by directory_list_groups.');

const topArg = z
  .number()
  .int()
  .min(1)
  .max(MAX_TOP)
  .default(25)
  .describe(`How many rows to return in this page (max ${MAX_TOP}). Graph itself defaults to 100.`);

const skipTokenArg = z
  .string()
  .min(1)
  .max(4000)
  .optional()
  .describe(
    'Cursor for the next page: pass the `skipToken` from the previous result, keeping ' +
      'every other argument identical. Directory collections do not support numeric ' +
      'skipping, so this is the only way to page.',
  );

const selectArg = z
  .array(z.string().min(1).regex(SELECT_FIELD, 'must be a bare Microsoft Graph property name'))
  .min(1)
  .max(30)
  .optional();

const orderbyArg = z
  .string()
  .min(1)
  .max(200)
  .refine(
    (value) => value.split(',').every((clause) => ORDERBY_CLAUSE.test(clause.trim())),
    'must be comma-separated bare property names, each optionally followed by asc or desc',
  )
  .optional();

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const listUsersInput = z.object({
  top: topArg,
  filter: z
    .string()
    .min(1)
    .max(512)
    .optional()
    .describe(
      'Raw OData $filter over `user` properties, e.g. "department eq \'Sales\'", ' +
        '"startswith(displayName,\'Ir\')" or "endsWith(mail,\'@contoso.com\')". Single ' +
        'quotes inside a literal must be doubled. `endsWith`, `ne` and `not` work here ' +
        'because the request always enables advanced queries.',
    ),
  select: selectArg.describe(
    'Graph `user` property names to return instead of the default compact projection, ' +
      'e.g. ["displayName","officeLocation","employeeId"]. Bare property names only; ' +
      'OData expressions and navigation paths are rejected. `id` is always included.',
  ),
  orderby: orderbyArg.describe(
    'Sort order, e.g. "displayName" or "userPrincipalName desc". Bare property names ' +
      'only. Sorting a filtered list is allowed here, but Graph will not sort on a ' +
      'property it cannot index, such as `assignedLicenses`.',
  ),
  skipToken: skipTokenArg,
});

const searchUsersInput = z.object({
  query: z
    .string()
    .min(1)
    .max(200)
    .describe(
      'What to look for, e.g. "Irene" or "contoso.com". Matched against displayName and ' +
        'mail as whole words or word prefixes.',
    ),
  top: topArg,
  select: selectArg.describe(
    'Graph `user` property names to return instead of the default compact projection. ' +
      'Bare property names only; `id` is always included.',
  ),
  skipToken: skipTokenArg,
});

const getUserInput = z.object({ id: userId });

const userRelationInput = z.object({ id: userId });

const directReportsInput = z.object({
  id: userId,
  top: topArg,
  skipToken: skipTokenArg,
});

const listGroupsInput = z.object({
  top: topArg,
  filter: z
    .string()
    .min(1)
    .max(512)
    .optional()
    .describe(
      'Raw OData $filter over `group` properties, e.g. "securityEnabled eq true", ' +
        '"startswith(displayName,\'Eng\')" or "groupTypes/any(c:c eq \'Unified\')". ' +
        'Single quotes inside a literal must be doubled.',
    ),
  select: selectArg.describe(
    'Graph `group` property names to return instead of the default compact projection, ' +
      'e.g. ["displayName","membershipRule"]. Bare property names only; `id` is always ' +
      'included.',
  ),
  skipToken: skipTokenArg,
});

const getGroupInput = z.object({ id: groupId });

const groupMembersInput = z.object({
  id: groupId,
  top: topArg,
  skipToken: skipTokenArg,
});

const userGroupsInput = z.object({ id: userId });

// ---------------------------------------------------------------------------
// Module
// ---------------------------------------------------------------------------

export const directoryModule: ToolModule = {
  group: GROUP,
  build({ graph }: ToolDeps): ToolDefinition[] {
    return [
      {
        name: 'directory_list_users',
        title: 'List directory users',
        description:
          'Lists users in the Microsoft Entra ID tenant, 25 per page by default (max ' +
          `${MAX_TOP}). Each row carries id, displayName, userPrincipalName, mail, ` +
          'jobTitle, department and accountEnabled; pass `select` for other properties. ' +
          'A `skipToken` in the result means more pages exist — call again with it set, ' +
          'since directory collections cannot be skipped numerically. Requires tenant ' +
          'admin consent for User.Read.All, and results are eventually consistent, so a ' +
          'user created seconds ago may not appear yet.',
        inputSchema: listUsersInput,
        scopes: USER_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { top, filter, select, orderby, skipToken } = listUsersInput.parse(args);
          const fields = select === undefined ? USER_LIST_SELECT : [...new Set(['id', ...select])];
          const res = await graph.request({
            path: '/users',
            method: 'GET',
            query: {
              // `$count=true` is what makes the client send `ConsistencyLevel: eventual`,
              // which in turn is what lets `endsWith`, `ne`, `not` and a filtered
              // `$orderby` work at all on a directory resource.
              $count: true,
              $top: top,
              $filter: filter,
              $orderby: orderby,
              $select: fields.join(','),
              $skiptoken: skipToken,
            },
            scopes: USER_SCOPES,
          });
          return collectionResult<UserPayload>('users', res.data, res.nextLink, (user) =>
            select === undefined ? projectUserSummary(user) : pick(user, fields),
          );
        },
      },
      {
        name: 'directory_search_users',
        title: 'Search directory users',
        description:
          'Finds users by name or email address, returning the same compact rows as ' +
          'directory_list_users, 25 per page by default. The term is matched against ' +
          'displayName and mail only, as whole words or word prefixes — there are no ' +
          'wildcards, no substring matching inside a word, and no date or numeric ' +
          'ranges. To match a userPrincipalName or any other property, use ' +
          'directory_list_users with a `startswith` filter instead. Requires tenant ' +
          'admin consent for User.Read.All.',
        inputSchema: searchUsersInput,
        scopes: USER_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { query, top, select, skipToken } = searchUsersInput.parse(args);
          const fields = select === undefined ? USER_LIST_SELECT : [...new Set(['id', ...select])];
          const res = await graph.request({
            path: '/users',
            method: 'GET',
            query: {
              $search: userSearchExpression(query),
              // Graph rejects `$search` on a directory resource unless the request is an
              // advanced query, which means both this and the eventual-consistency header.
              $count: true,
              $top: top,
              $select: fields.join(','),
              $skiptoken: skipToken,
            },
            scopes: USER_SCOPES,
          });
          return collectionResult<UserPayload>('users', res.data, res.nextLink, (user) =>
            select === undefined ? projectUserSummary(user) : pick(user, fields),
          );
        },
      },
      {
        name: 'directory_get_user',
        title: 'Get a directory user',
        description:
          'Returns one user by object id or userPrincipalName, with names, contact ' +
          'details, job title, department, office, employee id, account state, usage ' +
          'location and on-premises sync fields. Properties Graph reports as empty are ' +
          'omitted, so a missing key means the field is not set in the directory. An ' +
          'unknown id fails with 404 Request_ResourceNotFound. Requires tenant admin ' +
          'consent for User.Read.All.',
        inputSchema: getUserInput,
        scopes: USER_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id } = getUserInput.parse(args);
          const res = await graph.request<UserPayload>({
            path: userPath(id),
            method: 'GET',
            query: { $select: USER_DETAIL_SELECT.join(',') },
            scopes: USER_SCOPES,
          });
          return projectUserDetail(res.data ?? {});
        },
      },
      {
        name: 'directory_get_user_manager',
        title: 'Get a user\'s manager',
        description:
          "Returns the user's manager as a compact row (id, displayName, " +
          'userPrincipalName, mail, jobTitle, department). A user with no manager ' +
          'assigned fails with 404 Request_ResourceNotFound — that is the normal answer ' +
          'for executives and service accounts, not a broken id. Only the direct ' +
          'manager is returned; walk the chain by calling this again with the manager\'s ' +
          'id. Requires tenant admin consent for User.Read.All.',
        inputSchema: userRelationInput,
        scopes: USER_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id } = userRelationInput.parse(args);
          // No `$select`: `manager` is typed as `directoryObject`, and Graph rejects a
          // select naming user-only properties on it without an OData type cast.
          const res = await graph.request<DirectoryObjectPayload>({
            path: userPath(id, '/manager'),
            method: 'GET',
            scopes: USER_SCOPES,
          });
          return projectDirectoryObject(res.data ?? {});
        },
      },
      {
        name: 'directory_list_direct_reports',
        title: "List a user's direct reports",
        description:
          'Lists the users who report directly to the given user, 25 per page by default ' +
          `(max ${MAX_TOP}), as compact rows. Only one level down: reports of reports ` +
          'are not included. An empty list is the normal answer for an individual ' +
          'contributor. Requires tenant admin consent for User.Read.All.',
        inputSchema: directReportsInput,
        scopes: USER_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, top, skipToken } = directReportsInput.parse(args);
          // As with `manager`, this navigation is a `directoryObject` collection, so it
          // is fetched whole and projected here rather than trimmed with `$select`.
          const res = await graph.request({
            path: userPath(id, '/directReports'),
            method: 'GET',
            query: { $top: top, $skiptoken: skipToken },
            scopes: USER_SCOPES,
          });
          return collectionResult<DirectoryObjectPayload>(
            'directReports',
            res.data,
            res.nextLink,
            projectDirectoryObject,
          );
        },
      },
      {
        name: 'directory_list_groups',
        title: 'List directory groups',
        description:
          'Lists groups in the Microsoft Entra ID tenant, 25 per page by default (max ' +
          `${MAX_TOP}). Each row carries id, displayName, a truncated description, mail, ` +
          'groupTypes, securityEnabled, mailEnabled and visibility; a `groupTypes` ' +
          'containing "Unified" marks a Microsoft 365 group, and one containing ' +
          '"DynamicMembership" marks a rule-based group. Page with `skipToken`. ' +
          'Requires tenant admin consent for Group.Read.All.',
        inputSchema: listGroupsInput,
        scopes: GROUP_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { top, filter, select, skipToken } = listGroupsInput.parse(args);
          const fields = select === undefined ? GROUP_LIST_SELECT : [...new Set(['id', ...select])];
          const res = await graph.request({
            path: '/groups',
            method: 'GET',
            query: {
              $count: true,
              $top: top,
              $filter: filter,
              $select: fields.join(','),
              $skiptoken: skipToken,
            },
            scopes: GROUP_SCOPES,
          });
          return collectionResult<GroupPayload>('groups', res.data, res.nextLink, (group) =>
            select === undefined ? projectGroupSummary(group) : pick(group, fields),
          );
        },
      },
      {
        name: 'directory_get_group',
        title: 'Get a directory group',
        description:
          'Returns one group by object id, with its description, mail nickname, type ' +
          'flags, visibility, classification, lifecycle dates, dynamic membership rule ' +
          'and on-premises sync fields. A `resourceProvisioningOptions` containing ' +
          '"Team" means the group is backed by a Microsoft Teams team. Member counts are ' +
          'not included — use directory_list_group_members. Requires tenant admin ' +
          'consent for Group.Read.All.',
        inputSchema: getGroupInput,
        scopes: GROUP_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id } = getGroupInput.parse(args);
          const res = await graph.request<GroupPayload>({
            path: groupPath(id),
            method: 'GET',
            query: { $select: GROUP_DETAIL_SELECT.join(',') },
            scopes: GROUP_SCOPES,
          });
          return projectGroupDetail(res.data ?? {});
        },
      },
      {
        name: 'directory_list_group_members',
        title: 'List group members',
        description:
          `Lists the direct members of a group, 25 per page by default (max ${MAX_TOP}). ` +
          'Membership is mixed, so every row carries a `type` ("user", "group", ' +
          '"device", "servicePrincipal"); nested groups are returned as single rows ' +
          'rather than expanded, so a member of a nested group does not appear here. ' +
          'Page with `skipToken`. Requires tenant admin consent for Group.Read.All.',
        inputSchema: groupMembersInput,
        scopes: GROUP_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, top, skipToken } = groupMembersInput.parse(args);
          // `members` is a `directoryObject` collection: a `$select` naming user-only
          // properties needs a `/microsoft.graph.user` cast, which would then drop every
          // non-user member. Fetching whole objects and projecting here keeps both.
          const res = await graph.request({
            path: groupPath(id, '/members'),
            method: 'GET',
            query: { $top: top, $skiptoken: skipToken },
            scopes: GROUP_SCOPES,
          });
          return collectionResult<DirectoryObjectPayload>(
            'members',
            res.data,
            res.nextLink,
            projectDirectoryObject,
          );
        },
      },
      {
        name: 'directory_list_user_groups',
        title: "List a user's group memberships",
        description:
          'Returns the ids of every group the user belongs to, including memberships ' +
          'inherited through nested groups. Ids only: resolve the ones that matter with ' +
          'directory_get_group, since names are not part of the response. Security, ' +
          'distribution and Microsoft 365 groups are all included, directory roles are ' +
          'not, and a user in thousands of groups returns thousands of ids. Requires ' +
          'tenant admin consent for User.Read.All and Group.Read.All.',
        inputSchema: userGroupsInput,
        // Not a write: getMemberGroups is a POST only because its options travel in a
        // body. It changes nothing, so it stays available under --read-only.
        scopes: USER_AND_GROUP_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id } = userGroupsInput.parse(args);
          const res = await graph.request({
            path: userPath(id, '/getMemberGroups'),
            method: 'POST',
            // `true` would silently drop distribution and Microsoft 365 groups, which is
            // rarely what a "what groups is this person in" question means.
            body: { securityEnabledOnly: false },
            scopes: USER_AND_GROUP_SCOPES,
          });
          const groupIds = extractCollection<string>(res.data).filter(
            (value): value is string => typeof value === 'string',
          );
          const out: Record<string, unknown> = { count: groupIds.length, groupIds };
          if (res.nextLink !== undefined) out['nextLink'] = res.nextLink;
          return out;
        },
      },
    ];
  },
};
