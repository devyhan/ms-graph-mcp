/**
 * The `contacts` group: the signed-in user's Outlook personal contacts.
 *
 * Everything here targets `/me/contacts`, which is the default Contacts folder.
 * Contacts filed in a custom folder live under `/me/contactFolders/{id}/contacts`
 * and do not appear in these results.
 */

import { z } from 'zod';

import type { ToolDefinition, ToolDeps, ToolModule } from '../contracts.js';
import { GROUPS } from './groups.js';
import { extractCollection } from '../graph/client.js';
import { quoteSearch } from '../util/odata.js';
import { truncateText } from '../util/truncate.js';

const GROUP = GROUPS['contacts']!;

/**
 * Copied, not aliased: a tool's `scopes` array is an input to the consent set
 * the login path builds, and handing out the live `GROUPS` array would let a
 * caller that sorts or dedupes in place mutate the group catalogue.
 */
const READ_SCOPES = [...GROUP.readScopes];

/** Write tools project the resulting contact back to the caller, so they need read access too. */
const WRITE_SCOPES = [...new Set([...GROUP.readScopes, ...GROUP.writeScopes])];

/** `personalNotes` is free text and occasionally holds an entire pasted email. */
const NOTES_MAX_CHARS = 1000;

/**
 * Outlook contacts have exactly three email slots. Graph rejects a fourth with
 * `ErrorInvalidPropertySet`, so the cap is enforced before the round trip.
 */
const MAX_EMAIL_ADDRESSES = 3;

/** Properties worth spending tokens on for a row in a list. */
const LIST_SELECT = [
  'id',
  'displayName',
  'givenName',
  'surname',
  'companyName',
  'jobTitle',
  'emailAddresses',
  'mobilePhone',
  'businessPhones',
];

/** Everything a single-contact view is expected to answer. */
const DETAIL_SELECT = [
  ...LIST_SELECT,
  'middleName',
  'nickName',
  'department',
  'officeLocation',
  'homePhones',
  'imAddresses',
  'personalNotes',
  'birthday',
  'categories',
  'parentFolderId',
  'businessAddress',
  'homeAddress',
  'createdDateTime',
  'lastModifiedDateTime',
];

/**
 * A caller-supplied `$select` reaches Graph as a raw query value. Every
 * selectable property on the `contact` resource is a bare identifier, so
 * anything else is a mistake or an attempt to smuggle OData into the query.
 */
const SELECT_FIELD = /^[A-Za-z][A-Za-z0-9]*$/;

/** One `$orderby` clause: a bare property name, optionally `asc` or `desc`. */
const ORDERBY_CLAUSE = /^[A-Za-z][A-Za-z0-9]*(?:\s+(?:asc|desc))?$/;

// ---------------------------------------------------------------------------
// Graph payload shapes (only the parts that are projected)
// ---------------------------------------------------------------------------

interface EmailAddressPayload {
  address?: string | null;
  name?: string | null;
}

interface PhysicalAddressPayload {
  street?: string | null;
  city?: string | null;
  state?: string | null;
  postalCode?: string | null;
  countryOrRegion?: string | null;
}

interface ContactPayload {
  id?: string | null;
  displayName?: string | null;
  givenName?: string | null;
  middleName?: string | null;
  surname?: string | null;
  nickName?: string | null;
  companyName?: string | null;
  jobTitle?: string | null;
  department?: string | null;
  officeLocation?: string | null;
  emailAddresses?: EmailAddressPayload[] | null;
  mobilePhone?: string | null;
  businessPhones?: string[] | null;
  homePhones?: string[] | null;
  imAddresses?: string[] | null;
  personalNotes?: string | null;
  birthday?: string | null;
  categories?: string[] | null;
  parentFolderId?: string | null;
  businessAddress?: PhysicalAddressPayload | null;
  homeAddress?: PhysicalAddressPayload | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
}

// ---------------------------------------------------------------------------
// Projection helpers
// ---------------------------------------------------------------------------

/**
 * Graph pads a contact with a null for every property it was asked for, which
 * is most of the payload for a typical contact. A missing key in the output
 * therefore means "not set", not "not requested".
 */
function dropEmpty(obj: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(obj)) {
    if (value === undefined || value === null || value === '') continue;
    if (Array.isArray(value) && value.length === 0) continue;
    if (
      typeof value === 'object' &&
      !Array.isArray(value) &&
      Object.keys(value as Record<string, unknown>).length === 0
    ) {
      continue;
    }
    out[key] = value;
  }
  return out;
}

/** Keeps only the named properties of a raw Graph object, dropping nulls. */
function pick(data: unknown, fields: readonly string[]): Record<string, unknown> {
  if (typeof data !== 'object' || data === null) return {};
  const record = data as Record<string, unknown>;
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

function projectEmails(
  list: EmailAddressPayload[] | null | undefined,
): Array<{ address: string; name?: string }> {
  if (!Array.isArray(list)) return [];
  const out: Array<{ address: string; name?: string }> = [];
  for (const entry of list) {
    const address = trimmed(entry?.address);
    if (address === undefined) continue;
    const name = trimmed(entry?.name);
    out.push(name === undefined ? { address } : { address, name });
  }
  return out;
}

function projectAddress(
  address: PhysicalAddressPayload | null | undefined,
): Record<string, unknown> | undefined {
  if (address === null || address === undefined) return undefined;
  const parts = dropEmpty({
    street: address.street,
    city: address.city,
    state: address.state,
    postalCode: address.postalCode,
    countryOrRegion: address.countryOrRegion,
  });
  return Object.keys(parts).length === 0 ? undefined : parts;
}

/** One row of a collection. */
function projectSummary(contact: ContactPayload): Record<string, unknown> {
  return dropEmpty({
    id: contact.id,
    displayName: contact.displayName,
    givenName: contact.givenName,
    surname: contact.surname,
    companyName: contact.companyName,
    jobTitle: contact.jobTitle,
    emails: projectEmails(contact.emailAddresses),
    mobilePhone: contact.mobilePhone,
    businessPhones: contact.businessPhones,
  });
}

/** A single contact, with the fields a follow-up action actually needs. */
function projectDetail(contact: ContactPayload): Record<string, unknown> {
  const notes = trimmed(contact.personalNotes);
  return dropEmpty({
    ...projectSummary(contact),
    middleName: contact.middleName,
    nickName: contact.nickName,
    department: contact.department,
    officeLocation: contact.officeLocation,
    homePhones: contact.homePhones,
    imAddresses: contact.imAddresses,
    personalNotes: notes === undefined ? undefined : truncateText(notes, NOTES_MAX_CHARS),
    birthday: contact.birthday,
    categories: contact.categories,
    parentFolderId: contact.parentFolderId,
    businessAddress: projectAddress(contact.businessAddress),
    homeAddress: projectAddress(contact.homeAddress),
    createdDateTime: contact.createdDateTime,
    lastModifiedDateTime: contact.lastModifiedDateTime,
  });
}

function collectionResult(
  data: unknown,
  nextLink: string | undefined,
  project: (contact: ContactPayload) => Record<string, unknown>,
): Record<string, unknown> {
  const items = extractCollection<ContactPayload>(data);
  // `contacts` stays present even when empty: an absent key reads as "not
  // returned", and "no matches" is the answer worth stating plainly.
  return {
    count: items.length,
    ...(nextLink === undefined ? {} : { nextLink }),
    contacts: items.map(project),
  };
}

/**
 * Outlook item ids are long and opaque, and ids carried over from EWS can
 * contain characters that are not URL-safe, so they never reach the path raw.
 */
function contactPath(id: string): string {
  return `/me/contacts/${encodeURIComponent(id)}`;
}

// ---------------------------------------------------------------------------
// Shared argument pieces
// ---------------------------------------------------------------------------

const contactId = z
  .string()
  .min(1)
  .describe(
    'The contact id, as returned by contacts_list or contacts_search. This is the ' +
      'long opaque Outlook item id, not an email address or display name.',
  );

const emailAddressesArg = z
  .array(
    z.object({
      address: z
        .string()
        .min(1)
        .max(256)
        .describe('The SMTP address, e.g. "irene@contoso.com".'),
      name: z
        .string()
        .min(1)
        .max(256)
        .optional()
        .describe('Display label shown beside the address in Outlook. Defaults to blank.'),
    }),
  )
  .min(1)
  .max(MAX_EMAIL_ADDRESSES);

const topArg = z
  .number()
  .int()
  .min(1)
  .max(100)
  .default(25)
  .describe('How many contacts to return in this page. Graph itself defaults to 10.');

const skipArg = z
  .number()
  .int()
  .min(0)
  .max(10000)
  .optional()
  .describe(
    'How many contacts to skip before this page, for paging: pass the running total ' +
      'already seen. Only meaningful while `nextLink` is present in the previous result.',
  );

const selectArg = z
  .array(z.string().min(1).regex(SELECT_FIELD, 'must be a bare Microsoft Graph property name'))
  .min(1)
  .max(30)
  .optional()
  .describe(
    'Graph `contact` property names to return instead of the default compact ' +
      'projection, e.g. ["displayName","department","birthday"]. Bare property names ' +
      'only; OData expressions and navigation paths are rejected. `id` is always included.',
  );

function toEmailBody(
  entries: Array<{ address: string; name?: string }> | undefined,
): Array<{ address: string; name?: string }> | undefined {
  if (entries === undefined) return undefined;
  return entries.map((entry) =>
    entry.name === undefined ? { address: entry.address } : { address: entry.address, name: entry.name },
  );
}

// ---------------------------------------------------------------------------
// Input schemas
// ---------------------------------------------------------------------------

const listInput = z.object({
  top: topArg,
  skip: skipArg,
  filter: z
    .string()
    .min(1)
    .max(512)
    .optional()
    .describe(
      'Raw OData $filter over `contact` properties, e.g. ' +
        "\"companyName eq 'Contoso'\" or \"startswith(surname,'Mc')\". Single quotes " +
        'inside a literal must be doubled. Filtering on emailAddresses is not supported ' +
        'by Outlook — use contacts_search for that.',
    ),
  orderby: z
    .string()
    .min(1)
    .max(200)
    .refine(
      (value) => value.split(',').every((clause) => ORDERBY_CLAUSE.test(clause.trim())),
      'must be comma-separated bare property names, each optionally followed by asc or desc',
    )
    .optional()
    .describe(
      'Sort order, e.g. "displayName" or "lastModifiedDateTime desc". Bare property ' +
        'names only. Omit to get Outlook\'s own ordering.',
    ),
  select: selectArg,
});

const getInput = z.object({
  id: contactId,
});

const searchInput = z.object({
  query: z
    .string()
    .min(1)
    .max(200)
    .describe(
      'What to look for. A plain term such as "Irene" or "contoso.com" matches across ' +
        'display name, given/surname, company and every email address. A field-scoped ' +
        'term such as "displayName:Irene" or "emailAddresses:contoso.com" narrows it.',
    ),
  top: topArg,
  skip: skipArg,
});

const createInput = z.object({
  givenName: z.string().min(1).max(256).describe("The contact's first name."),
  surname: z.string().min(1).max(256).optional().describe("The contact's last name."),
  emailAddresses: emailAddressesArg
    .optional()
    .describe(
      `Up to ${MAX_EMAIL_ADDRESSES} email addresses, in Outlook slot order (Email 1, 2, 3). ` +
        'Outlook has no fourth slot.',
    ),
  mobilePhone: z
    .string()
    .min(1)
    .max(64)
    .optional()
    .describe('Mobile phone number, as free text, e.g. "+1 425 555 0109".'),
  companyName: z.string().min(1).max(256).optional().describe('Employer or organisation.'),
  jobTitle: z.string().min(1).max(256).optional().describe('Job title at that organisation.'),
});

const updateInput = z.object({
  id: contactId,
  givenName: z.string().min(1).max(256).optional().describe("The contact's first name."),
  surname: z.string().min(1).max(256).optional().describe("The contact's last name."),
  displayName: z
    .string()
    .min(1)
    .max(256)
    .optional()
    .describe(
      'The name Outlook files the contact under. Outlook derives it from the given and ' +
        'surname when it has never been set explicitly; setting it here stops that.',
    ),
  emailAddresses: emailAddressesArg
    .optional()
    .describe(
      `Replacement list of up to ${MAX_EMAIL_ADDRESSES} email addresses. This REPLACES ` +
        'every existing address rather than adding to them, so include the ones to keep. ' +
        'Read the contact first if unsure.',
    ),
  mobilePhone: z.string().min(1).max(64).optional().describe('Mobile phone number, as free text.'),
  companyName: z.string().min(1).max(256).optional().describe('Employer or organisation.'),
  jobTitle: z.string().min(1).max(256).optional().describe('Job title at that organisation.'),
});

const deleteInput = z.object({
  id: contactId,
});

export const contactsModule: ToolModule = {
  group: GROUP,
  build({ graph }: ToolDeps): ToolDefinition[] {
    return [
      {
        name: 'contacts_list',
        title: 'List contacts',
        description:
          "Lists personal contacts from the signed-in user's default Outlook Contacts " +
          'folder, 25 per page by default (max 100). Each row carries id, displayName, ' +
          'given/surname, company, job title, email addresses and phone numbers; pass ' +
          '`select` for different properties. A `nextLink` in the result means more pages ' +
          'exist — call again with `skip` set to the number already seen. Contacts stored ' +
          'in a custom contact folder are not included.',
        inputSchema: listInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { top, skip, filter, orderby, select } = listInput.parse(args);
          const fields = select === undefined ? LIST_SELECT : [...new Set(['id', ...select])];
          const res = await graph.request({
            path: '/me/contacts',
            method: 'GET',
            query: {
              $top: top,
              $skip: skip,
              $filter: filter,
              $orderby: orderby,
              $select: fields.join(','),
            },
            scopes: READ_SCOPES,
          });
          return collectionResult(res.data, res.nextLink, (contact) =>
            select === undefined ? projectSummary(contact) : pick(contact, fields),
          );
        },
      },
      {
        name: 'contacts_get',
        title: 'Get a contact',
        description:
          'Returns one personal contact by id, with names, company, job title, department, ' +
          'every email address, phone numbers, postal addresses, birthday, categories and ' +
          'timestamps. Personal notes are truncated to ' +
          `${NOTES_MAX_CHARS} characters. Properties Graph reports as empty are omitted, ` +
          'so a missing key means the field is not set. An id from a different mailbox or ' +
          'a deleted contact fails with 404 ErrorItemNotFound.',
        inputSchema: getInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id } = getInput.parse(args);
          const res = await graph.request<ContactPayload>({
            path: contactPath(id),
            method: 'GET',
            query: { $select: DETAIL_SELECT.join(',') },
            scopes: READ_SCOPES,
          });
          return projectDetail(res.data ?? {});
        },
      },
      {
        name: 'contacts_search',
        title: 'Search contacts',
        description:
          'Full-text search over the default Contacts folder using Outlook `$search`, ' +
          'returning the same compact rows as contacts_list, 25 per page by default. ' +
          'Results come back ranked by relevance, which is why `$filter` and `$orderby` ' +
          'cannot be combined with a search and are not offered here — use contacts_list ' +
          'for filtered or sorted results. Search matches whole words and prefixes only: ' +
          'it will not do date ranges, wildcards in the middle of a term, or "not equal" ' +
          'style conditions.',
        inputSchema: searchInput,
        scopes: READ_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { query, top, skip } = searchInput.parse(args);
          const res = await graph.request({
            path: '/me/contacts',
            method: 'GET',
            query: {
              $search: quoteSearch(query),
              $top: top,
              $skip: skip,
              $select: LIST_SELECT.join(','),
            },
            scopes: READ_SCOPES,
          });
          return collectionResult(res.data, res.nextLink, projectSummary);
        },
      },
      {
        name: 'contacts_create',
        title: 'Create a contact',
        description:
          "Creates a personal contact in the signed-in user's default Contacts folder and " +
          'returns the stored contact, including the new id. Outlook builds displayName ' +
          'from givenName and surname. Outlook holds at most ' +
          `${MAX_EMAIL_ADDRESSES} email addresses per contact. Graph does not de-duplicate: ` +
          'calling this twice creates two contacts, so search first when the contact may ' +
          'already exist.',
        inputSchema: createInput,
        write: true,
        scopes: WRITE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const input = createInput.parse(args);
          const res = await graph.request<ContactPayload>({
            path: '/me/contacts',
            method: 'POST',
            body: dropEmpty({
              givenName: input.givenName,
              surname: input.surname,
              companyName: input.companyName,
              jobTitle: input.jobTitle,
              mobilePhone: input.mobilePhone,
              emailAddresses: toEmailBody(input.emailAddresses),
            }),
            scopes: WRITE_SCOPES,
          });
          return projectDetail(res.data ?? {});
        },
      },
      {
        name: 'contacts_update',
        title: 'Update a contact',
        description:
          'Updates the given fields of one personal contact and returns the stored result. ' +
          'Omitted fields are left alone, but `emailAddresses` is a collection: supplying ' +
          'it REPLACES every address on the contact, so read the contact first and resend ' +
          'the ones to keep. There is no optimistic-concurrency check here — a concurrent ' +
          'edit from Outlook is silently overwritten. At least one field besides `id` is ' +
          'required.',
        inputSchema: updateInput,
        write: true,
        scopes: WRITE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id, emailAddresses, ...fields } = updateInput.parse(args);
          const body = dropEmpty({
            givenName: fields.givenName,
            surname: fields.surname,
            displayName: fields.displayName,
            companyName: fields.companyName,
            jobTitle: fields.jobTitle,
            mobilePhone: fields.mobilePhone,
            emailAddresses: toEmailBody(emailAddresses),
          });
          if (Object.keys(body).length === 0) {
            throw new Error(
              'contacts_update needs at least one field to change besides `id`, e.g. jobTitle or mobilePhone.',
            );
          }
          const res = await graph.request<ContactPayload>({
            path: contactPath(id),
            method: 'PATCH',
            body,
            scopes: WRITE_SCOPES,
          });
          return projectDetail(res.data ?? {});
        },
      },
      {
        name: 'contacts_delete',
        title: 'Delete a contact',
        description:
          'Deletes one personal contact. The contact moves to the Deleted Items folder, ' +
          'so it is recoverable from Outlook but not through these tools, and Graph ' +
          'returns no body — only a confirmation of the id. Deleting an id that is already ' +
          'gone fails with 404 ErrorItemNotFound.',
        inputSchema: deleteInput,
        write: true,
        scopes: WRITE_SCOPES,
        group: GROUP.name,
        handler: async (args) => {
          const { id } = deleteInput.parse(args);
          const res = await graph.request({
            path: contactPath(id),
            method: 'DELETE',
            scopes: WRITE_SCOPES,
          });
          return { id, deleted: true, status: res.status };
        },
      },
    ];
  },
};
