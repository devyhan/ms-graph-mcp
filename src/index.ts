#!/usr/bin/env node
/**
 * Executable entry point.
 *
 * Two very different jobs live here. `serve` speaks MCP over stdio, where
 * stdout carries JSON-RPC framing and a stray `console.log` corrupts the
 * session. Every other subcommand is an ordinary CLI run by a human, where
 * stdout is the answer. `out()` is only ever called on the CLI paths.
 */

import type { AuthProvider, GraphClient, ServerConfig, ToolDefinition } from './contracts.js';
import { serveStdio } from '@modelcontextprotocol/server/stdio';

import { cachePaths, clearCache, createAuthProvider, describeAccount } from './auth/index.js';
import type { ClientIdSource, ParsedConfig } from './config.js';
import {
  ConfigError,
  SHIPPED_APP_DISCLOSURE,
  clientIdGuidance,
  describeClientIdSource,
  graphBaseUrl,
  isClientIdUnset,
  packageVersion,
  parseConfig,
} from './config.js';
import { createGraphClient } from './graph/client.js';
import { collectTools, createServerFactory } from './server.js';
import { GROUPS, scopesForGroups } from './tools/groups.js';
import { createLogger, logger, setLogger } from './util/logger.js';

/** Scopes every session asks for regardless of which groups are enabled. */
const BASE_SCOPES = ['offline_access', 'User.Read'];

/** CLI-only. Never reachable from `serve`. */
function out(text: string): void {
  process.stdout.write(text);
}

function err(text: string): void {
  try {
    process.stderr.write(text);
  } catch {
    // A closed stderr must not turn a clean shutdown into a crash.
  }
}

/**
 * A client that refuses to talk.
 *
 * `collectTools` only closes over the client — no module calls it while
 * building definitions — so the commands that merely enumerate tools can skip
 * constructing MSAL entirely, and therefore work before a client ID is set.
 */
function inertGraphClient(): GraphClient {
  const refuse = (): never => {
    throw new Error('This command inspects configuration only and makes no Graph calls.');
  };
  return { request: refuse, batch: refuse, follow: refuse };
}

interface Session {
  auth: AuthProvider;
  graph: GraphClient;
  tools: ToolDefinition[];
}

function buildSession(config: ServerConfig, clientIdSource: ClientIdSource): Session {
  // The source travels with the config so the provider can name the app it is
  // about to send the user to: "whose application is this" is the question a
  // person needs answered before they consent, and it stops being self-evident
  // the moment a shared application ships as the default.
  const auth = createAuthProvider({ config, clientIdSource });
  const graph = createGraphClient({ auth, config });
  return { auth, graph, tools: collectTools({ graph, config }) };
}

/**
 * The consent set for a sign-in.
 *
 * The group catalogue is the coarse grouping; individual tools declare extra
 * scopes it does not list (`me_get_mailbox_settings` needs
 * `MailboxSettings.Read`, for instance), and a scope nobody asks for at login
 * is a scope that 403s at call time. Under `--read-only` the write tools are
 * suppressed, so their scopes are left out rather than consented to for nothing.
 */
/**
 * One line describing how `login` will sign the user in. Both flows are OAuth
 * 2.0, so the wording names the grant rather than implying one of them is "the
 * OAuth one".
 */
function describeAuthFlow(config: ServerConfig): string {
  const port =
    config.authPort > 0
      ? `http://localhost:${config.authPort}`
      : 'http://localhost (OS-assigned port)';

  switch (config.authFlow) {
    case 'browser':
      return `browser only, authorization code + PKCE, redirect ${port}`;
    case 'device':
      return 'device code only, no browser opened';
    default:
      return `auto: browser first (authorization code + PKCE, redirect ${port}), device code as fallback`;
  }
}

function consentScopes(config: ServerConfig, tools: ToolDefinition[]): string[] {
  const scopes = new Set(scopesForGroups(config.groups, config.readOnly));
  for (const tool of tools) {
    if (config.readOnly && tool.write === true) continue;
    for (const scope of tool.scopes) scopes.add(scope);
  }
  return [...scopes].sort();
}

/** The same scope set, attributed to the group that asks for it. */
function scopesByGroup(
  config: ServerConfig,
  tools: ToolDefinition[],
): Array<{ name: string; title: string; adminConsent: boolean; scopes: string[] }> {
  const buckets = new Map<string, Set<string>>();

  for (const name of config.groups) {
    const meta = GROUPS[name];
    if (meta === undefined) continue;
    const set = new Set(meta.readScopes);
    if (!config.readOnly) for (const scope of meta.writeScopes) set.add(scope);
    buckets.set(name, set);
  }

  for (const tool of tools) {
    if (config.readOnly && tool.write === true) continue;
    const set = buckets.get(tool.group) ?? new Set<string>();
    for (const scope of tool.scopes) set.add(scope);
    buckets.set(tool.group, set);
  }

  return [...buckets.entries()].map(([name, set]) => {
    const meta = GROUPS[name];
    return {
      name,
      title: meta?.title ?? name,
      adminConsent: meta?.requiresAdminConsent ?? false,
      // The base scopes are reported once on their own; repeating them per
      // group turns a consent checklist into noise.
      scopes: [...set].filter((scope) => !BASE_SCOPES.includes(scope)).sort(),
    };
  });
}

function pad(label: string): string {
  return label.padEnd(16);
}

// ---------------------------------------------------------------------------
// Commands
// ---------------------------------------------------------------------------

async function runLogin(parsed: ParsedConfig): Promise<number> {
  const { config, clientIdSource } = parsed;

  // Checked before MSAL is constructed and before any socket is opened: an
  // empty client ID can only come back as AADSTS700038 seconds later, after
  // the user has already picked an account, and it says less than this does.
  if (isClientIdUnset(config.clientId)) {
    err(`${clientIdGuidance(parsed.clientIdReason)}\n`);
    return 1;
  }

  const { auth, tools } = buildSession(config, clientIdSource);
  const scopes = consentScopes(config, tools);

  // Progress goes to stderr so a caller can pipe the account line on stdout.
  err(`Signing in to ${config.authority} for ${scopes.length} delegated scopes.\n`);
  // Shipping an application removes the REGISTRATION barrier, never the
  // CONSENT one, and the user is about to grant OUR app access to their
  // mailbox. Say so while they can still stop.
  if (clientIdSource === 'shipped') err(`\n${SHIPPED_APP_DISCLOSURE}\n\n`);

  let account;
  try {
    account = await auth.login(scopes);
  } catch (error) {
    // The auth layer writes its user guidance INTO the message — which app was
    // refused, which tenant, which flag fixes it — and logs the MSAL stack to
    // stderr separately to keep it out of the way. Letting this reach the
    // top-level handler would undo that: it prefixes "microsoft-graph-mcp
    // failed:" and reprints the whole guidance as the first line of a stack
    // trace. Print the guidance, keep the stack behind --verbose.
    err(`${describe(error)}\n`);
    if (config.verbose) err(`${describeWithStack(error)}\n`);
    return 1;
  }

  out(`Signed in as ${describeAccount(account)}\n`);
  out(`Token cache: ${cachePaths(config.cacheDir).cacheFile}\n`);
  return 0;
}

async function runLogout(parsed: ParsedConfig): Promise<number> {
  const { config } = parsed;
  const paths = cachePaths(config.cacheDir);
  try {
    await createAuthProvider({ config, clientIdSource: parsed.clientIdSource }).logout();
  } catch (error) {
    // A cache written by a different client ID, or a corrupt one, still has to
    // be removable — that is the whole point of `logout`.
    err(`MSAL sign-out failed (${describe(error)}); removing the cache files directly.\n`);
    await clearCache(config.cacheDir);
  }
  out(`Signed out. Removed ${paths.cacheFile} and ${paths.keyFile}.\n`);
  return 0;
}

async function runStatus(parsed: ParsedConfig): Promise<number> {
  const { config, clientIdSource } = parsed;
  const unset = isClientIdUnset(config.clientId);
  // No client ID means no MSAL and no account to look up, but the rest of the
  // report is still worth printing — that is how a user diagnoses the gap.
  const tools = collectTools({ graph: inertGraphClient(), config });

  out(`microsoft-graph-mcp ${packageVersion()}\n\n`);
  out(`${pad('Account:')}${unset ? 'no client ID configured' : await currentAccount(config)}\n`);
  // Never a stand-in GUID: a syntactically valid one looks configured, gets
  // copied into bug reports, and reaches Entra only to earn an AADSTS700038.
  out(`${pad('Client ID:')}${unset ? 'not configured' : config.clientId}\n`);
  // Only when there IS one: "not configured / not configured" says it twice
  // and buries the guidance printed at the end of the report.
  if (!unset) out(`${pad('Client ID from:')}${describeClientIdSource(clientIdSource)}\n`);
  out(`${pad('Tenant:')}${config.tenantId}\n`);
  out(`${pad('Authority:')}${config.authority}\n`);
  out(`${pad('Graph:')}${graphBaseUrl(config)}/${config.graphVersion}\n`);
  out(
    `${pad('Mode:')}${config.readOnly ? 'read-only' : 'read-write'}` +
      `, discovery ${config.discovery ? 'on' : 'off'}` +
      `, org-mode ${config.orgMode ? 'on' : 'off'}` +
      `, beta ${config.allowBeta ? 'allowed' : 'blocked'}` +
      `, generic writes ${config.allowGenericWrite ? 'allowed' : 'blocked'}\n`,
  );
  out(`${pad('Sign-in:')}${describeAuthFlow(config)}\n`);
  out(`${pad('Groups:')}${config.groups.join(', ')} (+ generic)\n`);
  out(`${pad('Tools:')}${tools.length}\n`);
  out(`${pad('Cache:')}${cachePaths(config.cacheDir).cacheFile}\n`);

  const scopes = consentScopes(config, tools);
  out(`\nDelegated scopes requested (${scopes.length}):\n`);
  for (const scope of scopes) out(`  ${scope}\n`);

  const adminGroups = config.groups.filter((name) => GROUPS[name]?.requiresAdminConsent === true);
  if (adminGroups.length > 0) {
    out(
      `\nGroups needing tenant admin consent: ${adminGroups.join(', ')}.\n` +
        'Sign-in fails for these until an administrator grants consent for the application.\n',
    );
  }
  if (unset) {
    out(`\n${clientIdGuidance(parsed.clientIdReason)}\n`);
  } else if (clientIdSource === 'shipped') {
    out(`\n${SHIPPED_APP_DISCLOSURE}\n`);
  }
  return 0;
}

/** The signed-in account, or why it could not be read. Never throws. */
async function currentAccount(config: ServerConfig): Promise<string> {
  try {
    return describeAccount(await createAuthProvider({ config }).getAccount());
  } catch (error) {
    return `could not read the token cache: ${describe(error)}`;
  }
}

function runPermissions(config: ServerConfig): number {
  const tools = collectTools({ graph: inertGraphClient(), config });

  out('Delegated Microsoft Graph permissions for this configuration.\n');
  out(`Mode: ${config.readOnly ? 'read-only (write scopes omitted)' : 'read-write'}\n\n`);

  out('Always requested\n');
  for (const scope of BASE_SCOPES) out(`  ${scope}\n`);

  for (const group of scopesByGroup(config, tools)) {
    // The generic group is the union of every other group by construction, so
    // repeating it here would double the length of a consent checklist.
    if (group.name === 'generic' || group.scopes.length === 0) continue;
    const flag = group.adminConsent ? '  [TENANT ADMIN CONSENT REQUIRED]' : '';
    out(`\n${group.name} — ${group.title}${flag}\n`);
    for (const scope of group.scopes) out(`  ${scope}\n`);
  }

  out('\ngeneric — Generic Graph access\n');
  out('  (no scopes of its own: graph_request reuses the union of the scopes above)\n');

  const all = consentScopes(config, tools);
  out(`\nAll ${all.length} scopes, for an admin consent request:\n`);
  out(`${all.join(' ')}\n`);
  return 0;
}

function runServe(parsed: ParsedConfig): void {
  const { config, clientIdSource } = parsed;
  if (isClientIdUnset(config.clientId)) {
    // stderr, never stdout: stdout is the JSON-RPC frame and a banner on it
    // corrupts the session before the client has finished initializing.
    err(
      '\n**********************************************************************\n' +
        '  microsoft-graph-mcp: NO ENTRA CLIENT ID CONFIGURED.\n' +
        '  The server will start and list its tools, but every Graph call will\n' +
        '  fail. Pass --client-id <id> or set MS365_MCP_CLIENT_ID, then run\n' +
        '  `npx ms-graph-mcp login`.\n' +
        '**********************************************************************\n\n',
    );
    if (parsed.clientIdReason !== undefined) err(`${parsed.clientIdReason}\n\n`);
  }

  const { auth, graph, tools } = buildSession(config, clientIdSource);
  const handle = serveStdio(createServerFactory({ config, graph, auth, tools }), {
    onerror: (error) => logger.error('Transport error.', { message: error.message }),
  });

  let closing = false;
  const shutdown = (signal: string): void => {
    if (closing) return;
    closing = true;
    err(`[graph-mcp] Received ${signal}; closing the stdio transport.\n`);
    void handle
      .close()
      .catch((error: unknown) => err(`[graph-mcp] Close failed: ${describe(error)}\n`))
      .finally(() => process.exit(0));
  };
  process.on('SIGINT', () => shutdown('SIGINT'));
  process.on('SIGTERM', () => shutdown('SIGTERM'));

  // Without these the process can die on a background rejection while the
  // client sits waiting on a response that will never come. Loud and non-zero
  // is the only useful failure for a supervised stdio server.
  process.on('unhandledRejection', (reason: unknown) => {
    err(`[graph-mcp] FATAL unhandled rejection: ${describeWithStack(reason)}\n`);
    process.exit(1);
  });
  process.on('uncaughtException', (error: unknown) => {
    err(`[graph-mcp] FATAL uncaught exception: ${describeWithStack(error)}\n`);
    process.exit(1);
  });

  logger.info('Serving Microsoft Graph MCP over stdio.', {
    version: packageVersion(),
    groups: config.groups,
    tools: tools.length,
    discovery: config.discovery,
    readOnly: config.readOnly,
  });
}

// ---------------------------------------------------------------------------
// Entry
// ---------------------------------------------------------------------------

function describe(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function describeWithStack(error: unknown): string {
  if (error instanceof Error) return error.stack ?? error.message;
  return String(error);
}

/** Resolves to an exit code, or `undefined` when the process should keep running. */
async function main(): Promise<number | undefined> {
  let parsed;
  try {
    parsed = parseConfig(process.argv.slice(2), process.env);
  } catch (error) {
    if (error instanceof ConfigError) {
      err(`${error.message}\n\nRun with --help for the full list of options.\n`);
      return 2;
    }
    throw error;
  }

  const { config, command } = parsed;
  setLogger(createLogger(config.verbose));

  switch (command.kind) {
    case 'help':
      out(`${command.text}\n`);
      return 0;
    case 'version':
      out(`${packageVersion()}\n`);
      return 0;
    case 'login':
      return runLogin(parsed);
    case 'logout':
      return runLogout(parsed);
    case 'status':
      return runStatus(parsed);
    case 'permissions':
      return runPermissions(config);
    case 'serve':
      runServe(parsed);
      return undefined;
  }
}

main()
  .then((code) => {
    if (code !== undefined) process.exit(code);
  })
  .catch((error: unknown) => {
    err(`microsoft-graph-mcp failed: ${describeWithStack(error)}\n`);
    process.exit(1);
  });
