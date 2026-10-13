/**
 * Best-effort mapping from a Graph path to the delegated scopes that path needs.
 *
 * Two consumers share this table so they can never disagree about what a path
 * needs. The 403 hint in `graph/errors.ts` names the alternatives as prose for
 * a person to read; `graph_request` asks MSAL for the minimal set instead of
 * the union of every enabled group, so a path whose scope is granted succeeds
 * even when an unrelated scope in the union is missing consent. A guess that
 * turns out too narrow degrades to a 403 whose hint names the rest — the same
 * narrowing `search_query` already does per entity type.
 */

export interface ScopeGuess {
  /** Matches the Graph-relative paths this entry covers. First match wins. */
  readonly test: RegExp;
  /**
   * The smallest scope set a token request for these paths should ask for:
   * the weakest scope that can satisfy some call there, not every
   * alternative. MSAL mints a token valid for every scope asked for, so
   * asking for the alternatives too would demand consent for all of them.
   */
  readonly scopes: readonly string[];
  /** The alternatives as prose, for the 403 hint. */
  readonly hint: string;
}

/**
 * Ordered most specific first: `/me/messages` must match the mail entry, not
 * the `/me` fallback at the end.
 */
export const SCOPE_GUESSES: ReadonlyArray<ScopeGuess> = [
  { test: /\/(messages|mailFolders|sendMail|mailboxSettings|inferenceClassification)\b/i, scopes: ['Mail.Read'], hint: 'Mail.Read, Mail.ReadWrite or Mail.Send' },
  { test: /\/(events|calendar|calendars|calendarView|calendarGroups)\b/i, scopes: ['Calendars.Read'], hint: 'Calendars.Read or Calendars.ReadWrite' },
  { test: /^\/me\/contacts\b/i, scopes: ['Contacts.Read'], hint: 'Contacts.Read or Contacts.ReadWrite' },
  { test: /\/(drive|drives|shares)\b/i, scopes: ['Files.Read'], hint: 'Files.Read, Files.Read.All or Files.ReadWrite.All' },
  { test: /^\/sites\b/i, scopes: ['Sites.Read.All'], hint: 'Sites.Read.All or Sites.ReadWrite.All' },
  { test: /\/(chats|chatMessage)\b/i, scopes: ['Chat.Read'], hint: 'Chat.Read, Chat.ReadWrite or ChatMessage.Send' },
  { test: /(^\/teams|joinedTeams|\/channels)\b/i, scopes: ['Team.ReadBasic.All'], hint: 'Team.ReadBasic.All, Channel.ReadBasic.All or ChannelMessage.Read.All' },
  { test: /\/todo\b/i, scopes: ['Tasks.Read'], hint: 'Tasks.Read or Tasks.ReadWrite' },
  { test: /^\/planner\b/i, scopes: ['Tasks.Read'], hint: 'Tasks.Read or Tasks.ReadWrite, plus Group.Read.All for plans in a group' },
  { test: /^\/groups\b/i, scopes: ['Group.Read.All'], hint: 'Group.Read.All or GroupMember.Read.All (tenant admin consent)' },
  { test: /^\/(applications|servicePrincipals)\b/i, scopes: ['Application.Read.All'], hint: 'Application.Read.All (tenant admin consent)' },
  { test: /^\/deviceManagement\b/i, scopes: ['DeviceManagementManagedDevices.Read.All'], hint: 'DeviceManagementManagedDevices.Read.All (tenant admin consent)' },
  { test: /^\/(users|directoryObjects|directoryRoles|administrativeUnits|contacts|organization|devices)\b/i, scopes: ['User.Read.All'], hint: 'User.Read.All or Directory.Read.All (tenant admin consent)' },
  { test: /^\/me\b/i, scopes: ['User.Read'], hint: 'User.Read' },
];

/**
 * The minimal scopes a token request for `path` should ask for, or undefined
 * when the path matches nothing known and the caller should fall back to the
 * union of the enabled groups. Returns a copy; the table is shared.
 */
export function guessScopesForPath(path: string): string[] | undefined {
  for (const guess of SCOPE_GUESSES) {
    if (guess.test.test(path)) return [...guess.scopes];
  }
  return undefined;
}
