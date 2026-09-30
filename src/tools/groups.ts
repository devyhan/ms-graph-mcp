/**
 * The tool-group catalogue: the single source of truth for which delegated
 * scopes each area of Microsoft 365 needs, and which areas a plain user can
 * consent to on their own.
 */

import type { ToolGroupMeta } from '../contracts.js';

/**
 * Admin-consent flags come from Microsoft's permissions reference and are NOT
 * a stylistic choice — do not "simplify" them away.
 *
 * The delegated Mail.*, Calendars.*, Files.*, Tasks.*, Contacts.*, Chat.* and
 * Sites.* scopes below are user-consentable: a signed-in user can grant them
 * without involving a tenant administrator. By contrast User.Read.All,
 * Group.Read.All, ChannelMessage.Read.All (plus Team.ReadBasic.All /
 * Channel.ReadBasic.All in most tenants) and every
 * DeviceManagementManagedDevices.* scope require tenant admin consent even in
 * delegated mode. Groups that touch those are flagged `requiresAdminConsent`
 * and stay disabled unless `--org-mode` is passed, so a personal-account user
 * never sees a consent prompt they cannot satisfy.
 */
export const GROUPS: Record<string, ToolGroupMeta> = {
  me: {
    name: 'me',
    title: 'My profile',
    description: 'The signed-in user profile stored in Microsoft Entra ID.',
    readScopes: ['User.Read'],
    writeScopes: [],
    requiresAdminConsent: false,
  },
  mail: {
    name: 'mail',
    title: 'Mail',
    description: 'Outlook mail: messages, folders, attachments, and sending.',
    readScopes: ['Mail.Read'],
    writeScopes: ['Mail.ReadWrite', 'Mail.Send'],
    requiresAdminConsent: false,
  },
  calendar: {
    name: 'calendar',
    title: 'Calendar',
    description: 'Outlook calendar events, availability, and meeting invitations.',
    readScopes: ['Calendars.Read'],
    writeScopes: ['Calendars.ReadWrite'],
    requiresAdminConsent: false,
  },
  files: {
    name: 'files',
    title: 'Files',
    description: 'OneDrive files and folders, including content upload and download.',
    readScopes: ['Files.Read'],
    writeScopes: ['Files.ReadWrite'],
    requiresAdminConsent: false,
  },
  todo: {
    name: 'todo',
    title: 'To Do',
    description: 'Microsoft To Do task lists and their tasks.',
    readScopes: ['Tasks.Read'],
    writeScopes: ['Tasks.ReadWrite'],
    requiresAdminConsent: false,
  },
  planner: {
    name: 'planner',
    title: 'Planner',
    description: 'Microsoft Planner plans, buckets, and assigned tasks.',
    readScopes: ['Tasks.Read'],
    writeScopes: ['Tasks.ReadWrite'],
    requiresAdminConsent: false,
  },
  contacts: {
    name: 'contacts',
    title: 'Contacts',
    description: 'Outlook personal contacts and contact folders.',
    readScopes: ['Contacts.Read'],
    writeScopes: ['Contacts.ReadWrite'],
    requiresAdminConsent: false,
  },
  chat: {
    name: 'chat',
    title: 'Chat',
    description: 'Microsoft Teams one-to-one and group chats the user belongs to.',
    readScopes: ['Chat.Read'],
    writeScopes: ['Chat.ReadWrite'],
    requiresAdminConsent: false,
  },
  sharepoint: {
    name: 'sharepoint',
    title: 'SharePoint',
    description: 'SharePoint sites, document libraries, and lists.',
    readScopes: ['Sites.Read.All'],
    writeScopes: ['Sites.ReadWrite.All'],
    requiresAdminConsent: false,
  },
  search: {
    name: 'search',
    title: 'Search',
    description: 'Microsoft Search queries spanning Outlook mail, OneDrive, and SharePoint.',
    readScopes: ['Mail.Read', 'Files.Read', 'Sites.Read.All'],
    writeScopes: [],
    requiresAdminConsent: false,
  },
  teams: {
    name: 'teams',
    title: 'Teams',
    description: 'Microsoft Teams teams, channels, and channel messages.',
    readScopes: ['Team.ReadBasic.All', 'Channel.ReadBasic.All', 'ChannelMessage.Read.All'],
    writeScopes: ['ChannelMessage.Send'],
    requiresAdminConsent: true,
  },
  directory: {
    name: 'directory',
    title: 'Directory',
    description: 'Microsoft Entra ID directory lookups for users and groups.',
    readScopes: ['User.Read.All', 'Group.Read.All'],
    writeScopes: [],
    requiresAdminConsent: true,
  },
  intune: {
    name: 'intune',
    title: 'Intune',
    description: 'Microsoft Intune managed device inventory.',
    readScopes: ['DeviceManagementManagedDevices.Read.All'],
    writeScopes: [],
    requiresAdminConsent: true,
  },
};

/** Catalogue order, used everywhere a stable group ordering is needed. */
export const GROUP_NAMES: string[] = Object.keys(GROUPS);

const PERSONAL_GROUPS = ['me', 'mail', 'calendar', 'files', 'todo', 'contacts', 'search'];
const WORK_GROUPS = [...PERSONAL_GROUPS, 'chat', 'planner', 'sharepoint'];
const ADMIN_GROUPS = ['directory', 'intune', 'teams'];

/** Named bundles selectable with `--preset`. */
export const PRESETS: Record<string, string[]> = {
  personal: PERSONAL_GROUPS,
  work: WORK_GROUPS,
  admin: ADMIN_GROUPS,
  all: [...GROUP_NAMES],
};

/** Groups enabled when neither `--groups` nor `--preset` is given. */
export const DEFAULT_GROUPS: string[] = PERSONAL_GROUPS;

/** Preset names, in a stable order for help text and error messages. */
export const PRESET_NAMES: string[] = Object.keys(PRESETS);

function unknownGroupError(name: string): Error {
  return new Error(
    `Unknown tool group "${name}". Valid groups: ${GROUP_NAMES.join(', ')}.`,
  );
}

/**
 * Turns `--groups` / `--preset` into the final enabled group list.
 *
 * Both selectors may be combined; their union is taken. Admin-consent groups
 * are dropped unless `orgMode` is set, because requesting those scopes on a
 * personal or unprivileged account produces a consent prompt the user cannot
 * approve.
 */
export function resolveGroups(opts: {
  groups?: string[];
  preset?: string;
  orgMode: boolean;
}): string[] {
  const selected = new Set<string>();
  let explicit = false;

  if (opts.preset !== undefined) {
    const preset = PRESETS[opts.preset];
    if (preset === undefined) {
      throw new Error(
        `Unknown preset "${opts.preset}". Valid presets: ${PRESET_NAMES.join(', ')}.`,
      );
    }
    for (const name of preset) selected.add(name);
    explicit = true;
  }

  if (opts.groups !== undefined && opts.groups.length > 0) {
    for (const raw of opts.groups) {
      const name = raw.trim();
      if (GROUPS[name] === undefined) throw unknownGroupError(raw);
      selected.add(name);
    }
    explicit = true;
  }

  if (!explicit) {
    for (const name of DEFAULT_GROUPS) selected.add(name);
  }

  const ordered = GROUP_NAMES.filter((name) => selected.has(name));
  if (opts.orgMode) return ordered;

  const kept: string[] = [];
  const dropped: string[] = [];
  for (const name of ordered) {
    const meta = GROUPS[name];
    if (meta !== undefined && meta.requiresAdminConsent) dropped.push(name);
    else kept.push(name);
  }

  if (dropped.length > 0) {
    // stdio transport owns stdout; diagnostics must go to stderr.
    process.stderr.write(
      `[microsoft-graph-mcp] Skipping tool group(s) ${dropped.join(', ')}: ` +
        'their Microsoft Graph scopes require tenant admin consent. ' +
        'Pass --org-mode (or MS365_MCP_ORG_MODE=1) to enable them.\n',
    );
  }

  return kept;
}

/**
 * The delegated scopes to request for a set of groups. `offline_access` keeps
 * refresh tokens working and `User.Read` is needed to identify the account,
 * so both are always included.
 */
export function scopesForGroups(names: string[], readOnly: boolean): string[] {
  const scopes = new Set<string>(['offline_access', 'User.Read']);
  for (const name of names) {
    const meta = GROUPS[name];
    if (meta === undefined) throw unknownGroupError(name);
    for (const scope of meta.readScopes) scopes.add(scope);
    if (!readOnly) {
      for (const scope of meta.writeScopes) scopes.add(scope);
    }
  }
  return [...scopes].sort();
}

/**
 * Scopes only a tenant administrator can grant, derived from the catalogue
 * rather than kept by hand.
 *
 * A scope counts as admin-only when every group that carries it is flagged
 * `requiresAdminConsent`. Subtracting the user-consentable groups matters:
 * several scopes appear in more than one group, and a scope a plain user can
 * already grant somewhere else is not something to send anyone to an
 * administrator for.
 *
 * This exists so the server can tell a blocked user the truth. "Run login
 * again" is wrong advice for a scope no amount of signing in will produce, and
 * it was the advice this server used to give.
 */
/**
 * Admin-consent scopes that deliberately live on one tool instead of in a
 * group's meta, where the derivation below cannot see them.
 *
 * `TeamMember.Read.All` is declared only by the roster tool in
 * `src/tools/teams.ts`, so that enabling the Teams group does not ask every
 * user to consent to reading membership. That is a good reason and the list
 * stays; what keeps it honest is the test that fails if a tool in an
 * admin-flagged group ever declares a scope missing from here.
 */
const TOOL_ONLY_ADMIN_SCOPES = ['TeamMember.Read.All'];

export const ADMIN_CONSENT_SCOPES: ReadonlySet<string> = (() => {
  const admin = new Set<string>(TOOL_ONLY_ADMIN_SCOPES);
  const user = new Set<string>();
  for (const meta of Object.values(GROUPS)) {
    const target = meta.requiresAdminConsent ? admin : user;
    for (const scope of [...meta.readScopes, ...meta.writeScopes]) target.add(scope);
  }
  for (const scope of user) admin.delete(scope);
  return admin;
})();

/**
 * The group a caller would enable to get a scope requested at login, preferring
 * one that needs no administrator when a scope appears in several.
 */
export function groupOwningScope(scope: string): string | undefined {
  let fallback: string | undefined;
  for (const meta of Object.values(GROUPS)) {
    if (![...meta.readScopes, ...meta.writeScopes].includes(scope)) continue;
    if (!meta.requiresAdminConsent) return meta.name;
    fallback ??= meta.name;
  }
  return fallback;
}
