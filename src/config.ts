/**
 * Command-line and environment parsing for the Microsoft Graph MCP server.
 *
 * Every option has a long flag and an `MS365_MCP_*` environment fallback;
 * flags win. Anything unrecognised is a hard error, so a typo can never
 * silently downgrade the server's capabilities.
 */

import { readFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import type { AuthFlow, OutputFormat, ServerConfig } from './contracts.js';
import { GROUPS, GROUP_NAMES, PRESET_NAMES, PRESETS, resolveGroupsDetailed } from './tools/groups.js';

/** What the process should do once configuration is resolved. */
export type Command =
  | { kind: 'serve' }
  | { kind: 'login'; scopes?: string[]; full?: boolean }
  | { kind: 'logout' }
  | { kind: 'status' }
  | { kind: 'permissions' }
  | { kind: 'help'; text: string }
  | { kind: 'version' };

/** Cap on serialised tool output when `--max-output-chars` is not given. */
export const DEFAULT_MAX_OUTPUT_CHARS = 60_000;

/** Every accepted `--auth-flow` value, in help-text order. */
export const AUTH_FLOWS: AuthFlow[] = ['auto', 'browser', 'device'];

/** Sign-in flow used when `--auth-flow` is not given. */
export const DEFAULT_AUTH_FLOW: AuthFlow = 'auto';

/** Loopback port for the browser redirect when `--auth-port` is not given. */
export const DEFAULT_AUTH_PORT = 0;

/** Every accepted `--output` value, in help-text order. */
export const OUTPUT_FORMATS: OutputFormat[] = ['json', 'toon'];

/** Tool result format used when `--output` is not given. */
export const DEFAULT_OUTPUT: OutputFormat = 'json';

/**
 * One line per flow for `--help`. Both flows are OAuth 2.0; they differ only in
 * how the user proves consent, so the wording must not imply otherwise.
 */
const AUTH_FLOW_HELP: Record<AuthFlow, string> = {
  auto: 'Try browser, and on failure log the reason to stderr and fall back to device.',
  browser: 'Authorization code + PKCE (RFC 7636) on a loopback redirect; opens a browser.',
  device: 'Device authorization grant (RFC 8628); prints a code to enter elsewhere.',
};

/**
 * Lowest port the loopback redirect listener may bind.
 *
 * Ports below this are privileged on Unix and the listener runs as the ordinary
 * user, so binding one would fail at sign-in time rather than at parse time.
 */
const LOWEST_UNPRIVILEGED_PORT = 1024;
const HIGHEST_PORT = 65_535;

/** Sovereign-cloud endpoints. `graphHost` is a bare hostname, never a URL. */
export const CLOUDS: Record<string, { graphHost: string; authorityHost: string }> = {
  global: { graphHost: 'graph.microsoft.com', authorityHost: 'login.microsoftonline.com' },
  usgov: { graphHost: 'graph.microsoft.us', authorityHost: 'login.microsoftonline.us' },
  usgovdod: { graphHost: 'dod-graph.microsoft.us', authorityHost: 'login.microsoftonline.us' },
  china: {
    graphHost: 'microsoftgraph.chinacloudapi.cn',
    authorityHost: 'login.partner.microsoftonline.cn',
  },
};

export const CLOUD_NAMES: string[] = Object.keys(CLOUDS);

/** Cloud used when `--cloud` is absent, and the only one the shipped app serves. */
export const DEFAULT_CLOUD = 'global';

/** Raised for any bad flag, value, group, or preset. Callers print and exit. */
export class ConfigError extends Error {
  override readonly name = 'ConfigError';
}

// ---------------------------------------------------------------------------
// The application this package signs in with
// ---------------------------------------------------------------------------

/**
 * The shared multi-tenant Entra app this project ships so that a new user
 * does not have to register their own. Swap this single constant when the
 * registration changes. Empty string means "not yet registered" and every
 * code path must degrade to asking the user for --client-id.
 *
 * The replacement MUST be multi-tenant: a single-tenant registration is
 * rejected with AADSTS700016 in every other tenant, and a GUID that was never
 * an application is rejected with AADSTS700038 - both only after the user has
 * already typed their password, which is exactly why the default here is empty
 * rather than a syntactically valid placeholder. It must also be a public
 * client ("Allow public client flows" on, an http://localhost redirect under
 * "Mobile and desktop applications"), and, if its audience includes personal
 * Microsoft accounts, carry api.requestedAccessTokenVersion = 2.
 *
 * It will be registered in the global cloud only, so `resolveClientId` refuses
 * to hand it to a sovereign cloud.
 */
export const SHIPPED_CLIENT_ID = '';

/** True when this build ships a default application to fall back on. */
export function hasShippedClientId(): boolean {
  return SHIPPED_CLIENT_ID.trim() !== '';
}

/** True while no application is configured, so sign-in cannot even be attempted. */
export function isClientIdUnset(clientId: string): boolean {
  return clientId.trim() === '';
}

/**
 * A `${...}` placeholder that nothing substituted.
 *
 * Both the Claude Code plugin manifest and ordinary shell config interpolate
 * `${NAME}` before the value reaches us. When that does not happen — a typo in
 * the key, a config option the user skipped, a syntax the host does not in fact
 * support — the literal text arrives instead, and it is not inert: an
 * unsubstituted tenant id is concatenated into the authority URL and the run
 * fails much later with a confusing Entra error rather than here with a clear
 * one. Cheap to detect, because no real GUID or tenant name contains braces.
 */
export function unsubstitutedPlaceholder(value: string): string | null {
  const match = value.trim().match(/^\$\{([^}]*)\}$/);
  return match ? (match[1] ?? '') : null;
}

/** What to tell someone whose placeholder never got a value. */
export function placeholderGuidance(field: string, raw: string): string {
  const key = unsubstitutedPlaceholder(raw) ?? raw;
  return [
    `${field} arrived as the literal text "${raw}", which means nothing replaced it.`,
    '',
    `If this came from a Claude Code plugin, "${key}" is a plugin setting that was`,
    'never filled in. Reconfigure the plugin and supply a value.',
    '',
    'If it came from a shell or an MCP config, the variable it refers to is unset in',
    'the environment the server was started from. Note that a GUI-launched client',
    'does not read your shell profile — export it where that client can see it, or',
    'pass the value on the command line instead.',
  ].join('\n');
}

/**
 * Which candidate supplied `config.clientId`.
 *
 * Reported rather than inferred: the user is about to grant an application
 * access to their mail, files and calendar, and "whose application is this"
 * is the question they need answered before consenting.
 */
export type ClientIdSource = 'flag' | 'env' | 'shipped' | 'unset';

/** The candidates `resolveClientId` chooses between, highest precedence first. */
export interface ClientIdCandidates {
  /** `--client-id` on the command line. */
  flag?: string | undefined;
  /** `MS365_MCP_CLIENT_ID` from the environment. */
  env?: string | undefined;
  /** The compiled-in default; production callers pass `SHIPPED_CLIENT_ID`. */
  shipped: string;
  /** Selected cloud name; the shipped app exists in `DEFAULT_CLOUD` alone. */
  cloud: string;
}

/** The chosen application, where it came from, and what was refused on the way. */
export interface ClientIdResolution {
  /** Empty string when no candidate was usable. */
  clientId: string;
  source: ClientIdSource;
  /**
   * Set only when a candidate existed but was deliberately not used, so the
   * caller can explain the gap instead of reporting a bare "unset".
   */
  reason?: string;
}

/**
 * Resolves the application to sign in with: `--client-id`, then
 * `MS365_MCP_CLIENT_ID`, then the shipped app, then nothing.
 *
 * Takes the shipped value as a parameter instead of reading the module
 * constant so that every branch stays testable, including the ones that only
 * come alive once a real registration ships.
 */
export function resolveClientId(candidates: ClientIdCandidates): ClientIdResolution {
  const flag = candidates.flag?.trim() ?? '';
  // Checked before the value is accepted, not after: an unsubstituted
  // placeholder is a configuration mistake wearing the costume of a client ID,
  // and every later error it causes points somewhere else.
  if (unsubstitutedPlaceholder(flag) !== null) {
    return { clientId: '', source: 'unset', reason: placeholderGuidance('--client-id', flag) };
  }
  if (flag !== '') return { clientId: flag, source: 'flag' };

  const env = candidates.env?.trim() ?? '';
  if (unsubstitutedPlaceholder(env) !== null) {
    return {
      clientId: '',
      source: 'unset',
      reason: placeholderGuidance('MS365_MCP_CLIENT_ID', env),
    };
  }
  if (env !== '') return { clientId: env, source: 'env' };

  const shipped = candidates.shipped.trim();
  if (shipped === '') return { clientId: '', source: 'unset' };

  // A global-cloud application ID is meaningless to a sovereign authority: it
  // would buy an AADSTS700038 after the password prompt. Refuse it here, and
  // say why, rather than letting the user discover it at sign-in.
  if (candidates.cloud !== DEFAULT_CLOUD) {
    return {
      clientId: '',
      source: 'unset',
      reason:
        `The application shipped with this package is registered in the ${DEFAULT_CLOUD} ` +
        `cloud only, so it cannot be used with --cloud ${candidates.cloud}. Register an ` +
        `application in that cloud and name it with --client-id.`,
    };
  }

  return { clientId: shipped, source: 'shipped' };
}

/** One phrase naming the application, for `status` and for consent prompts. */
export function describeClientIdSource(source: ClientIdSource): string {
  switch (source) {
    case 'flag':
      return '--client-id';
    case 'env':
      return 'MS365_MCP_CLIENT_ID';
    case 'shipped':
      return 'shipped with this package (shared multi-tenant application)';
    case 'unset':
      return 'not configured';
  }
}

/**
 * Shown before consenting with the shipped application.
 *
 * Shipping an application removes the REGISTRATION barrier, never the CONSENT
 * barrier, and the user still deserves to know whose app is doing the asking.
 */
export const SHIPPED_APP_DISCLOSURE = [
  'This sign-in uses the application shipped with ms-graph-mcp, not one',
  'you registered. Consenting grants that application the delegated scopes',
  'listed above against your own mailbox, files and calendar.',
  '',
  'In a work or school tenant the first sign-in also creates an enterprise',
  'application there, visible to that tenant\'s administrators and in the Entra',
  'sign-in logs. The default tenant consent policy withholds Mail.Read,',
  'Calendars.*, Chat.Read, Tasks.*, Files.Read.All, Sites.Read.All and',
  'MailboxSettings.Read from ordinary user consent, so an administrator may',
  'still have to approve the request.',
  '',
  'Pass --client-id <id> to sign in with your own registration instead.',
].join('\n');

/**
 * The registration walkthrough.
 *
 * This is the FALLBACK path once an application ships, so the lead depends on
 * whether one has: promising "you normally do not need this" would be a lie in
 * a build whose `SHIPPED_CLIENT_ID` is still empty.
 */
export function buildClientIdSetupMessage(shipped: string): string {
  const lead =
    shipped.trim() === ''
      ? [
          'No Entra application (client) ID configured, and this build ships no',
          'default application to fall back on, so sign-in cannot start.',
          '',
          'Pass an application your organisation already has with --client-id <id>',
          '(or set MS365_MCP_CLIENT_ID), or register one of your own below.',
        ]
      : [
          'No Entra application (client) ID configured.',
          '',
          'You normally do not need this: the package ships a shared multi-tenant',
          'application and uses it whenever no client ID is set. Register your own',
          'only when that default will not do - an employer that requires its own',
          'registration, a sovereign cloud (--cloud usgov, usgovdod, china), or a',
          'tenant whose administrators block the shared application.',
        ];

  return [
    ...lead,
    '',
    'To register your own application:',
    '  1. Open https://entra.microsoft.com > Applications > App registrations > New registration.',
    '  2. Under "Supported account types" pick the audience you need',
    '     (personal Microsoft accounts require the multitenant + personal option,',
    '     which in turn requires api.requestedAccessTokenVersion = 2 in the app',
    '     manifest; a version 1 token is rejected for those accounts).',
    '  3. Under "Authentication", add a "Mobile and desktop applications" platform',
    '     with the redirect URI http://localhost - the browser sign-in catches its',
    '     redirect there - and set "Allow public client flows" to Yes, which the',
    '     device code sign-in needs.',
    '  4. Copy the Application (client) ID from the Overview page.',
    '',
    'If the tenant refuses the bare http://localhost wildcard, register an exact',
    'http://localhost:PORT redirect URI instead and start the server with',
    '--auth-port PORT so the loopback listener binds that same port.',
    '',
    'Then start the server with --client-id <id> (or set MS365_MCP_CLIENT_ID).',
    'A personal Microsoft account also needs --tenant-id consumers: a refresh token',
    'issued through the common authority is rejected at the first refresh.',
  ].join('\n');
}

/** Guidance shown by `status` and `login` while the client ID is unset. */
export const CLIENT_ID_SETUP_MESSAGE = buildClientIdSetupMessage(SHIPPED_CLIENT_ID);

/**
 * Everything worth printing for an unset client ID: why a candidate was
 * refused, when one was, followed by the walkthrough.
 */
export function clientIdGuidance(reason?: string): string {
  if (reason === undefined || reason.trim() === '') return CLIENT_ID_SETUP_MESSAGE;
  return `${reason}\n\n${CLIENT_ID_SETUP_MESSAGE}`;
}

/** Absolute Graph base URL for a config, e.g. `https://graph.microsoft.com`. */
export function graphBaseUrl(config: Pick<ServerConfig, 'graphHost'>): string {
  return `https://${config.graphHost}`;
}

interface FlagSpec {
  flag: string;
  alias?: string;
  env?: string;
  value?: string;
  summary: string;
  fallback?: string;
  note?: string;
}

/**
 * The `--client-id` row. Its wording follows the build: with an application
 * shipped the flag is an override, without one it is a prerequisite.
 */
function clientIdFlagSpec(): FlagSpec {
  const base: FlagSpec = {
    flag: '--client-id',
    env: 'MS365_MCP_CLIENT_ID',
    value: '<id>',
    summary: 'Entra application (client) ID.',
  };
  return hasShippedClientId()
    ? { ...base, fallback: 'the application shipped with this package (global cloud only)' }
    : { ...base, note: 'required before sign-in; this build ships no default application' };
}

/** Every accepted option, in help-text order. Drives parsing AND `--help`. */
const FLAGS: FlagSpec[] = [
  clientIdFlagSpec(),
  {
    flag: '--tenant-id',
    env: 'MS365_MCP_TENANT_ID',
    value: '<id>',
    summary: 'Tenant ID, or common / organizations / consumers.',
    fallback: 'common',
  },
  {
    flag: '--cloud',
    env: 'MS365_MCP_CLOUD',
    value: '<name>',
    summary: `Sovereign cloud: ${CLOUD_NAMES.join(' | ')}.`,
    fallback: 'global',
  },
  {
    flag: '--auth-flow',
    env: 'MS365_MCP_AUTH_FLOW',
    value: '<name>',
    summary: `Interactive sign-in flow: ${AUTH_FLOWS.join(' | ')}.`,
    fallback: DEFAULT_AUTH_FLOW,
  },
  {
    flag: '--auth-port',
    env: 'MS365_MCP_AUTH_PORT',
    value: '<n>',
    summary: 'Loopback port the browser flow redirects to.',
    fallback: `${DEFAULT_AUTH_PORT} (the OS picks a free port)`,
  },
  {
    flag: '--scopes',
    value: '<a,b,c>',
    summary: 'Delegated scopes for `login` to request on top of the minimal set.',
  },
  {
    flag: '--full',
    summary: 'Ask `login` for every scope the configuration may need.',
  },
  {
    flag: '--groups',
    env: 'MS365_MCP_GROUPS',
    value: '<a,b,c>',
    summary: 'Comma-separated tool groups to enable; combines with --preset (union is taken).',
  },
  {
    flag: '--preset',
    env: 'MS365_MCP_PRESET',
    value: '<name>',
    summary: `Group bundle: ${PRESET_NAMES.join(' | ')}; combines with --groups (union is taken).`,
  },
  { flag: '--read-only', env: 'MS365_MCP_READ_ONLY', summary: 'Hide every tool that writes.' },
  {
    flag: '--org-mode',
    env: 'MS365_MCP_ORG_MODE',
    summary: 'Enable groups whose scopes need tenant admin consent.',
  },
  {
    flag: '--discovery',
    env: 'MS365_MCP_DISCOVERY',
    summary: 'Expose only discover_tools / call_tool instead of every tool.',
  },
  {
    flag: '--beta',
    env: 'MS365_MCP_ALLOW_BETA',
    summary: 'Allow requests against the Graph beta endpoint.',
  },
  {
    flag: '--allow-generic-write',
    env: 'MS365_MCP_ALLOW_GENERIC_WRITE',
    summary: 'Let the generic graph_request tool send write methods.',
  },
  {
    flag: '--max-output-chars',
    env: 'MS365_MCP_MAX_OUTPUT_CHARS',
    value: '<n>',
    summary: 'Truncate serialised tool output beyond this many characters.',
    fallback: String(DEFAULT_MAX_OUTPUT_CHARS),
  },
  {
    flag: '--output',
    env: 'MS365_MCP_OUTPUT',
    value: '<name>',
    summary: `Tool result format: ${OUTPUT_FORMATS.join(' | ')}.`,
    fallback: DEFAULT_OUTPUT,
  },
  {
    flag: '--cache-dir',
    env: 'MS365_MCP_CACHE_DIR',
    value: '<path>',
    summary: 'Directory holding the token cache.',
    fallback: 'OS config directory',
  },
  {
    flag: '--verbose',
    alias: '-v',
    env: 'MS365_MCP_VERBOSE',
    summary: 'Log Graph requests to stderr.',
  },
  { flag: '--help', alias: '-h', summary: 'Show this help and exit.' },
  { flag: '--version', summary: 'Print the package version and exit.' },
];

const VALUE_FLAGS = new Set(FLAGS.filter((f) => f.value !== undefined).map((f) => f.flag));

/** Flags that mean true when bare and also accept an explicit `=true`/`=false`. */
const BOOLEAN_FLAGS = new Set([
  '--full',
  '--read-only',
  '--org-mode',
  '--discovery',
  '--beta',
  '--allow-generic-write',
  '--verbose',
  '-v',
]);

const SUBCOMMANDS = ['serve', 'login', 'logout', 'status', 'permissions'] as const;
type Subcommand = (typeof SUBCOMMANDS)[number];

function isSubcommand(token: string): token is Subcommand {
  return (SUBCOMMANDS as readonly string[]).includes(token);
}

function allFlagNames(): string[] {
  const names: string[] = [];
  for (const spec of FLAGS) {
    names.push(spec.flag);
    if (spec.alias !== undefined) names.push(spec.alias);
  }
  return names;
}

function envString(env: NodeJS.ProcessEnv, key: string): string | undefined {
  const raw = env[key];
  if (raw === undefined) return undefined;
  const trimmed = raw.trim();
  return trimmed === '' ? undefined : trimmed;
}

const TRUTHY = new Set(['1', 'true', 'yes', 'on']);
const FALSY = new Set(['0', 'false', 'no', 'off']);

function envBool(env: NodeJS.ProcessEnv, key: string): boolean | undefined {
  const raw = env[key];
  if (raw === undefined) return undefined;
  const value = raw.trim().toLowerCase();
  if (value === '') return undefined;
  if (TRUTHY.has(value)) return true;
  if (FALSY.has(value)) return false;
  throw new ConfigError(
    `${key} must be one of 1/0, true/false, yes/no, on/off; got "${raw}".`,
  );
}

/**
 * Parses the `=value` on a boolean flag. Accepts the same vocabulary as the
 * boolean environment variables, so `--read-only=false` switches off
 * `MS365_MCP_READ_ONLY=true`. Only the `=value` form is read; a bare flag
 * still means true and never consumes the next argument.
 */
function parseFlagBool(raw: string, name: string): boolean {
  const value = raw.trim().toLowerCase();
  if (TRUTHY.has(value)) return true;
  if (FALSY.has(value)) return false;
  throw new ConfigError(
    `Option "${name}" takes an optional boolean value ` +
      `(one of 1/0, true/false, yes/no, on/off); got "${raw}".`,
  );
}

function parsePositiveInt(raw: string, source: string): number {
  const parsed = Number(raw);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new ConfigError(`${source} must be a positive whole number; got "${raw}".`);
  }
  return parsed;
}

/**
 * Resolves an `--auth-flow` value.
 *
 * `interactive` was the undocumented value that turned on the browser flow
 * before it had a flag, so it keeps working as an alias rather than becoming a
 * hard error for anyone already setting it.
 */
function parseAuthFlow(raw: string, source: string): AuthFlow {
  const value = raw.trim().toLowerCase();
  if (value === 'interactive') return 'browser';
  for (const flow of AUTH_FLOWS) {
    if (flow === value) return flow;
  }
  throw new ConfigError(
    `Unknown auth flow "${raw}" for ${source}. Valid flows: ${AUTH_FLOWS.join(', ')}.`,
  );
}

/** Resolves an `--output` value. Case-insensitive, like `--auth-flow`. */
function parseOutput(raw: string, source: string): OutputFormat {
  const value = raw.trim().toLowerCase();
  for (const format of OUTPUT_FORMATS) {
    if (format === value) return format;
  }
  throw new ConfigError(
    `Unknown output format "${raw}" for ${source}. Valid formats: ${OUTPUT_FORMATS.join(', ')}.`,
  );
}

/** Resolves an `--auth-port` value: 0 (let the OS pick) or 1024-65535. */
function parseAuthPort(raw: string, source: string): number {
  const trimmed = raw.trim();
  const parsed = Number(trimmed);
  if (trimmed === '' || !Number.isInteger(parsed)) {
    throw new ConfigError(
      `${source} must be a whole number: 0 to let the OS pick a free port, ` +
        `or a port in ${LOWEST_UNPRIVILEGED_PORT}-${HIGHEST_PORT}; got "${raw}".`,
    );
  }
  if (parsed === 0) return 0;
  if (parsed > 0 && parsed < LOWEST_UNPRIVILEGED_PORT) {
    throw new ConfigError(
      `${source} must not be a privileged port; got ${parsed}. The loopback ` +
        `redirect listener runs unprivileged and cannot bind below ` +
        `${LOWEST_UNPRIVILEGED_PORT}. Use 0 to let the OS pick a free port, ` +
        `or a port in ${LOWEST_UNPRIVILEGED_PORT}-${HIGHEST_PORT}.`,
    );
  }
  if (parsed > HIGHEST_PORT || parsed < 0) {
    throw new ConfigError(
      `${source} must be 0 or a port in ${LOWEST_UNPRIVILEGED_PORT}-${HIGHEST_PORT}; ` +
        `got "${raw}".`,
    );
  }
  return parsed;
}

function splitList(raw: string, source: string, what = 'group'): string[] {
  const items = raw
    .split(',')
    .map((item) => item.trim())
    .filter((item) => item !== '');
  if (items.length === 0) {
    throw new ConfigError(`${source} must name at least one ${what}.`);
  }
  return items;
}

/** Resolved defaults for the token cache, kept outside the package directory. */
function defaultCacheDir(env: NodeJS.ProcessEnv): string {
  // Deliberately NOT renamed alongside the package. This names the directory
  // holding the encrypted token cache, so changing it would silently sign out
  // everyone who upgraded and send them back through an interactive login and,
  // in a locked-down tenant, another admin consent round.
  const appName = 'microsoft-graph-mcp';
  if (process.platform === 'win32') {
    const appData = envString(env, 'APPDATA') ?? path.join(os.homedir(), 'AppData', 'Roaming');
    return path.join(appData, appName);
  }
  const configHome = envString(env, 'XDG_CONFIG_HOME') ?? path.join(os.homedir(), '.config');
  return path.join(configHome, appName);
}

function padRight(text: string, width: number): string {
  return text.length >= width ? text : text + ' '.repeat(width - text.length);
}

/** Help text, generated from the flag table and the group catalogue. */
export function buildHelpText(): string {
  const lines: string[] = [
    'ms-graph-mcp - Model Context Protocol server for Microsoft 365.',
    '',
    'Usage: ms-graph-mcp [command] [options]',
    '',
    'Commands:',
    '  serve         Run the MCP server on stdio (default).',
    '  login         Sign in interactively and store the token cache (minimal scopes;',
    '                widen with --scopes or --full).',
    '  logout        Clear the stored token cache.',
    '  status        Print the signed-in account and resolved configuration.',
    '  permissions   List the delegated Graph scopes the enabled groups may request.',
    '',
    'Options:',
  ];

  const rendered = FLAGS.map((spec) => {
    const alias = spec.alias === undefined ? '' : `, ${spec.alias}`;
    const value = spec.value === undefined ? '' : ` ${spec.value}`;
    return { spec, left: `  ${spec.flag}${alias}${value}` };
  });
  const width = Math.max(...rendered.map((entry) => entry.left.length)) + 2;

  for (const { spec, left } of rendered) {
    const notes: string[] = [];
    if (spec.env !== undefined) notes.push(`env ${spec.env}`);
    if (spec.fallback !== undefined) notes.push(`default ${spec.fallback}`);
    if (spec.note !== undefined) notes.push(spec.note);
    const suffix = notes.length === 0 ? '' : ` [${notes.join('; ')}]`;
    lines.push(`${padRight(left, width)}${spec.summary}${suffix}`);
  }

  lines.push('', 'Sign-in flows (--auth-flow); both are OAuth 2.0:');
  const flowWidth = Math.max(...AUTH_FLOWS.map((name) => name.length)) + 4;
  for (const flow of AUTH_FLOWS) {
    lines.push(`  ${padRight(flow, flowWidth)}${AUTH_FLOW_HELP[flow]}`);
  }

  lines.push('', 'Tool groups:');
  const groupWidth = Math.max(...GROUP_NAMES.map((name) => name.length)) + 4;
  for (const name of GROUP_NAMES) {
    const meta = GROUPS[name];
    if (meta === undefined) continue;
    const admin = meta.requiresAdminConsent ? ' (needs --org-mode)' : '';
    lines.push(`  ${padRight(name, groupWidth)}${meta.description}${admin}`);
  }

  lines.push('', 'Presets:');
  const presetWidth = Math.max(...PRESET_NAMES.map((name) => name.length)) + 4;
  for (const name of PRESET_NAMES) {
    const members = PRESETS[name] ?? [];
    lines.push(`  ${padRight(name, presetWidth)}${members.join(', ')}`);
  }

  // With an application shipped, --client-id is an override and the examples
  // must not imply otherwise. Without one, every example needs it.
  const id = hasShippedClientId() ? '' : ' --client-id <id>';
  lines.push(
    '',
    'Examples:',
    `  ms-graph-mcp login${id}`,
    `  ms-graph-mcp login${id} --scopes Mail.Read,Calendars.Read`,
    `  ms-graph-mcp${id} --preset work --read-only`,
    `  ms-graph-mcp${id} --groups mail,calendar --verbose`,
    `  ms-graph-mcp login${id} --auth-flow device`,
    `  ms-graph-mcp login${id} --auth-flow browser --auth-port 8400`,
    '',
    'Groups flagged above need tenant admin consent and stay disabled without --org-mode.',
    hasShippedClientId()
      ? 'Sign-in uses the application shipped with this package unless --client-id or\n' +
        'MS365_MCP_CLIENT_ID names another one. That removes the registration step, not\n' +
        'the consent step: many scopes still need a tenant administrator, and the shipped\n' +
        'application is registered in the global cloud only.'
      : 'This build ships no default application: pass --client-id <id> (or set\n' +
        'MS365_MCP_CLIENT_ID) before login. Run `ms-graph-mcp status` for the\n' +
        'registration walkthrough.',
  );

  return lines.join('\n');
}

/** Package name, read from package.json next to the compiled output. */
export function packageName(): string {
  try {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const raw = readFileSync(path.join(here, '..', 'package.json'), 'utf8');
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed === 'object' && parsed !== null && 'name' in parsed) {
      const name = (parsed as { name?: unknown }).name;
      if (typeof name === 'string' && name.trim() !== '') return name;
    }
  } catch {
    // A packaged install may not ship package.json beside dist/.
  }
  return '@devyhan/ms-graph-mcp';
}

/** Package version, read from package.json next to the compiled output. */
export function packageVersion(): string {
  try {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const raw = readFileSync(path.join(here, '..', 'package.json'), 'utf8');
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed === 'object' && parsed !== null && 'version' in parsed) {
      const version = (parsed as { version?: unknown }).version;
      if (typeof version === 'string') return version;
    }
  } catch {
    // A packaged install may not ship package.json beside dist/.
  }
  return '0.0.0-unknown';
}

interface RawOptions {
  /** `--client-id`, kept apart from the env value so the source is reportable. */
  clientIdFlag?: string;
  /** `MS365_MCP_CLIENT_ID`. */
  clientIdEnv?: string;
  tenantId?: string;
  cloud?: string;
  groups?: string[];
  preset?: string;
  readOnly?: boolean;
  orgMode?: boolean;
  discovery?: boolean;
  beta?: boolean;
  allowGenericWrite?: boolean;
  maxOutputChars?: number;
  output?: OutputFormat;
  cacheDir?: string;
  verbose?: boolean;
  authFlow?: AuthFlow;
  authPort?: number;
  /** `--scopes`: extra scopes for `login`. Read only by the login command. */
  loginScopes?: string[];
  /** `--full`: `login` asks for the whole configuration instead of the minimal set. */
  loginFull?: boolean;
}

function fromEnv(env: NodeJS.ProcessEnv): RawOptions {
  const raw: RawOptions = {};
  const clientId = envString(env, 'MS365_MCP_CLIENT_ID');
  if (clientId !== undefined) raw.clientIdEnv = clientId;
  const tenantId = envString(env, 'MS365_MCP_TENANT_ID');
  if (tenantId !== undefined) raw.tenantId = tenantId;
  const cloud = envString(env, 'MS365_MCP_CLOUD');
  if (cloud !== undefined) raw.cloud = cloud.toLowerCase();
  const groups = envString(env, 'MS365_MCP_GROUPS');
  if (groups !== undefined) raw.groups = splitList(groups, 'MS365_MCP_GROUPS');
  const preset = envString(env, 'MS365_MCP_PRESET');
  if (preset !== undefined) raw.preset = preset.toLowerCase();
  const maxOutput = envString(env, 'MS365_MCP_MAX_OUTPUT_CHARS');
  if (maxOutput !== undefined) {
    raw.maxOutputChars = parsePositiveInt(maxOutput, 'MS365_MCP_MAX_OUTPUT_CHARS');
  }
  const output = envString(env, 'MS365_MCP_OUTPUT');
  if (output !== undefined) raw.output = parseOutput(output, 'MS365_MCP_OUTPUT');
  const cacheDir = envString(env, 'MS365_MCP_CACHE_DIR');
  if (cacheDir !== undefined) raw.cacheDir = cacheDir;
  const authFlow = envString(env, 'MS365_MCP_AUTH_FLOW');
  if (authFlow !== undefined) raw.authFlow = parseAuthFlow(authFlow, 'MS365_MCP_AUTH_FLOW');
  const authPort = envString(env, 'MS365_MCP_AUTH_PORT');
  if (authPort !== undefined) raw.authPort = parseAuthPort(authPort, 'MS365_MCP_AUTH_PORT');

  const bools: Array<[keyof RawOptions, string]> = [
    ['readOnly', 'MS365_MCP_READ_ONLY'],
    ['orgMode', 'MS365_MCP_ORG_MODE'],
    ['discovery', 'MS365_MCP_DISCOVERY'],
    ['beta', 'MS365_MCP_ALLOW_BETA'],
    ['allowGenericWrite', 'MS365_MCP_ALLOW_GENERIC_WRITE'],
    ['verbose', 'MS365_MCP_VERBOSE'],
  ];
  for (const [key, name] of bools) {
    const value = envBool(env, name);
    if (value !== undefined) {
      // Every entry above names a boolean field.
      (raw as Record<string, unknown>)[key] = value;
    }
  }
  return raw;
}

function unknownFlagError(token: string): ConfigError {
  return new ConfigError(
    `Unknown option "${token}". Valid options: ${allFlagNames().join(', ')}.`,
  );
}

/** Everything `parseConfig` resolves from the command line and environment. */
export interface ParsedConfig {
  config: ServerConfig;
  command: Command;
  /**
   * Which candidate supplied `config.clientId`. It lives here rather than on
   * `ServerConfig`, a shared contract this module does not own.
   */
  clientIdSource: ClientIdSource;
  /**
   * Why a usable-looking candidate was refused, when one was. Feed it to
   * `clientIdGuidance` so the user hears the reason, not just the symptom.
   */
  clientIdReason?: string;
}

/** `buildConfig` returns the resolution too, so `parseConfig` can pass it on. */
interface BuiltConfig {
  config: ServerConfig;
  resolution: ClientIdResolution;
}

function toParsedConfig(built: BuiltConfig, command: Command): ParsedConfig {
  const result: ParsedConfig = {
    config: built.config,
    command,
    clientIdSource: built.resolution.source,
  };
  if (built.resolution.reason !== undefined) result.clientIdReason = built.resolution.reason;
  return result;
}

/**
 * Parses `argv` (the arguments after the script name) plus `env` into a fully
 * resolved config, the command to run, and where the client ID came from.
 */
export function parseConfig(argv: string[], env: NodeJS.ProcessEnv): ParsedConfig {
  // Help and version short-circuit before any validation, so they print even
  // alongside a bad flag or a bad environment value. The config they carry is
  // defaults-only: neither command reads settings beyond the cache directory
  // and verbosity used around printing.
  if (argv.includes('--help') || argv.includes('-h')) {
    return toParsedConfig(buildConfig({}, env), {
      kind: 'help',
      text: buildHelpText(),
    });
  }
  if (argv.includes('--version')) {
    return toParsedConfig(buildConfig({}, env), { kind: 'version' });
  }

  const raw = fromEnv(env);
  let subcommand: Subcommand = 'serve';
  let sawSubcommand = false;

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (token === undefined) continue;

    if (!token.startsWith('-')) {
      if (sawSubcommand) {
        throw new ConfigError(
          `Unexpected argument "${token}". Only one command is allowed: ${SUBCOMMANDS.join(', ')}.`,
        );
      }
      if (!isSubcommand(token)) {
        throw new ConfigError(
          `Unknown command "${token}". Valid commands: ${SUBCOMMANDS.join(', ')}.`,
        );
      }
      subcommand = token;
      sawSubcommand = true;
      continue;
    }

    const eq = token.indexOf('=');
    const name = eq === -1 ? token : token.slice(0, eq);
    const inlineValue = eq === -1 ? undefined : token.slice(eq + 1);

    if (inlineValue !== undefined && !VALUE_FLAGS.has(name) && !BOOLEAN_FLAGS.has(name)) {
      if (allFlagNames().includes(name)) {
        throw new ConfigError(`Option "${name}" does not take a value.`);
      }
      throw unknownFlagError(name);
    }

    const takeValue = (): string => {
      if (inlineValue !== undefined) {
        if (inlineValue === '') throw new ConfigError(`Option "${name}" requires a value.`);
        return inlineValue;
      }
      const next = argv[i + 1];
      if (next === undefined || next.startsWith('-')) {
        throw new ConfigError(
          `Option "${name}" requires a value. ` +
            `Write ${name}=<value> if the value itself starts with "-".`,
        );
      }
      i += 1;
      return next;
    };

    switch (name) {
      case '--client-id':
        raw.clientIdFlag = takeValue();
        break;
      case '--tenant-id':
        raw.tenantId = takeValue();
        break;
      case '--cloud':
        raw.cloud = takeValue().toLowerCase();
        break;
      case '--groups':
        raw.groups = splitList(takeValue(), '--groups');
        break;
      case '--preset':
        raw.preset = takeValue().toLowerCase();
        break;
      case '--max-output-chars':
        raw.maxOutputChars = parsePositiveInt(takeValue(), '--max-output-chars');
        break;
      case '--output':
        raw.output = parseOutput(takeValue(), '--output');
        break;
      case '--cache-dir':
        raw.cacheDir = takeValue();
        break;
      case '--auth-flow':
        raw.authFlow = parseAuthFlow(takeValue(), '--auth-flow');
        break;
      case '--auth-port':
        raw.authPort = parseAuthPort(takeValue(), '--auth-port');
        break;
      case '--scopes':
        raw.loginScopes = splitList(takeValue(), '--scopes', 'scope');
        break;
      case '--full':
        raw.loginFull = inlineValue === undefined ? true : parseFlagBool(inlineValue, name);
        break;
      case '--read-only':
        raw.readOnly = inlineValue === undefined ? true : parseFlagBool(inlineValue, name);
        break;
      case '--org-mode':
        raw.orgMode = inlineValue === undefined ? true : parseFlagBool(inlineValue, name);
        break;
      case '--discovery':
        raw.discovery = inlineValue === undefined ? true : parseFlagBool(inlineValue, name);
        break;
      case '--beta':
        raw.beta = inlineValue === undefined ? true : parseFlagBool(inlineValue, name);
        break;
      case '--allow-generic-write':
        raw.allowGenericWrite =
          inlineValue === undefined ? true : parseFlagBool(inlineValue, name);
        break;
      case '--verbose':
      case '-v':
        raw.verbose = inlineValue === undefined ? true : parseFlagBool(inlineValue, name);
        break;
      default:
        throw unknownFlagError(name);
    }
  }

  // Login-only arguments ride on the command, not the server config: `serve`
  // never reads them, and putting per-invocation asks next to the session
  // settings would suggest it does. On any other command they would silently
  // do nothing, so name them on stderr instead of dropping them quietly. Only
  // `--full` in its enabling form counts: an explicit `--full=false` asks for
  // nothing even on `login`, so there is nothing lost to warn about.
  if (subcommand !== 'login') {
    const ignored: string[] = [];
    if (raw.loginScopes !== undefined) ignored.push('--scopes');
    if (raw.loginFull === true) ignored.push('--full');
    if (ignored.length > 0) {
      const flags = ignored.join(' and ');
      const verb = ignored.length === 1 ? 'is' : 'are';
      try {
        process.stderr.write(
          `[ms-graph-mcp] Warning: ${flags} ${verb} only used by \`login\` and ` +
            `will be ignored for \`${subcommand}\`.\n`,
        );
      } catch {
        // A closed stderr must not turn a clean parse into a crash.
      }
    }
  }
  const command: Command =
    subcommand === 'login'
      ? {
          kind: 'login',
          ...(raw.loginScopes === undefined ? {} : { scopes: raw.loginScopes }),
          ...(raw.loginFull === true ? { full: true as const } : {}),
        }
      : { kind: subcommand };
  return toParsedConfig(buildConfig(raw, env), command);
}

function buildConfig(raw: RawOptions, env: NodeJS.ProcessEnv): BuiltConfig {
  const cloudName = raw.cloud ?? DEFAULT_CLOUD;
  const cloud = CLOUDS[cloudName];
  if (cloud === undefined) {
    throw new ConfigError(
      `Unknown cloud "${cloudName}". Valid clouds: ${CLOUD_NAMES.join(', ')}.`,
    );
  }

  // The tenant is the one placeholder that fails silently. It is concatenated
  // straight into the authority URL, so an unsubstituted value produces a
  // plausible-looking https://login.microsoftonline.com/${...} and the run dies
  // much later against Entra, blaming something else. `groups` and the numeric
  // options are already caught by their own validation.
  const rawTenant = raw.tenantId ?? '';
  if (unsubstitutedPlaceholder(rawTenant) !== null) {
    throw new ConfigError(placeholderGuidance('--tenant-id', rawTenant));
  }
  const tenantId = rawTenant === '' ? 'common' : rawTenant;
  const orgMode = raw.orgMode ?? false;

  const resolution = resolveClientId({
    flag: raw.clientIdFlag,
    env: raw.clientIdEnv,
    shipped: SHIPPED_CLIENT_ID,
    cloud: cloudName,
  });

  let groups: string[];
  let droppedGroups: string[];
  try {
    const selection: { groups?: string[]; preset?: string; orgMode: boolean } = { orgMode };
    if (raw.groups !== undefined) selection.groups = raw.groups;
    if (raw.preset !== undefined) selection.preset = raw.preset;
    const resolved = resolveGroupsDetailed(selection);
    groups = resolved.groups;
    droppedGroups = resolved.dropped;
  } catch (error) {
    throw new ConfigError(error instanceof Error ? error.message : String(error));
  }

  return {
    resolution,
    config: {
      clientId: resolution.clientId,
      tenantId,
      authority: `https://${cloud.authorityHost}/${tenantId}`,
      graphHost: cloud.graphHost,
      groups,
      droppedGroups,
      readOnly: raw.readOnly ?? false,
      // Tools opt into beta explicitly; --beta only lifts the ban on doing so.
      graphVersion: 'v1.0',
      allowBeta: raw.beta ?? false,
      discovery: raw.discovery ?? false,
      orgMode,
      maxOutputChars: raw.maxOutputChars ?? DEFAULT_MAX_OUTPUT_CHARS,
      output: raw.output ?? DEFAULT_OUTPUT,
      verbose: raw.verbose ?? false,
      cacheDir: raw.cacheDir ?? defaultCacheDir(env),
      authFlow: raw.authFlow ?? DEFAULT_AUTH_FLOW,
      authPort: raw.authPort ?? DEFAULT_AUTH_PORT,
      allowGenericWrite: raw.allowGenericWrite ?? false,
    },
  };
}
