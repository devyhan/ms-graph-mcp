/**
 * Assembly: turns the tool modules into a registered MCP server.
 *
 * `serveStdio` calls the factory once per connection and serves both the
 * 2026-07-28 stateless revision and the legacy 2025-11-25 handshake from it, so
 * the factory registers everything every time and closes over nothing mutable.
 */

import { McpServer } from '@modelcontextprotocol/server';

import type {
  AuthProvider,
  GraphClient,
  ServerConfig,
  ToolDefinition,
  ToolDeps,
  ToolModule,
} from './contracts.js';
import { GraphError, InteractionRequiredError } from './contracts.js';
import { packageVersion } from './config.js';
import { GROUP_NAMES } from './tools/groups.js';
import { buildDiscoveryTools } from './tools/discovery.js';
import { genericModule } from './tools/generic.js';
import { calendarModule } from './tools/calendar.js';
import { chatModule } from './tools/chat.js';
import { contactsModule } from './tools/contacts.js';
import { directoryModule } from './tools/directory.js';
import { filesModule } from './tools/files.js';
import { intuneModule } from './tools/intune.js';
import { mailModule } from './tools/mail.js';
import { meModule } from './tools/me.js';
import { plannerModule } from './tools/planner.js';
import { searchModule } from './tools/search.js';
import { sharepointModule } from './tools/sharepoint.js';
import { teamsModule } from './tools/teams.js';
import { todoModule } from './tools/todo.js';
import { logger } from './util/logger.js';
import { serializeResult } from './util/truncate.js';

/** The server name reported in `initialize`. Clients key their config off it. */
export const SERVER_NAME = 'microsoft-graph-mcp';

/** Every domain module, keyed by the group name it owns. */
const MODULES: Record<string, ToolModule> = {
  me: meModule,
  mail: mailModule,
  calendar: calendarModule,
  files: filesModule,
  todo: todoModule,
  planner: plannerModule,
  contacts: contactsModule,
  chat: chatModule,
  sharepoint: sharepointModule,
  search: searchModule,
  teams: teamsModule,
  directory: directoryModule,
  intune: intuneModule,
};

/** The tool always exposed in discovery mode alongside the discovery pair. */
const PERMISSIONS_TOOL = 'graph_list_permissions';

/**
 * Tools whose effect a user cannot undo from inside this server: mail leaves
 * the tenant, a deleted item is gone. `destructiveHint` is what a client uses
 * to decide whether to ask before running one.
 */
const DESTRUCTIVE = /(?:^|_)(?:delete|send|forward|reply|respond)(?:_|$)/;

/**
 * `graph_request` carries no verb in its name, and with --allow-generic-write it
 * can issue DELETE against anything the token reaches. `call_tool` carries no
 * verb either and dispatches to any tool in the catalogue, sending mail and
 * deleting files among them. Both are the most dangerous tools here, so they are
 * named rather than pattern-matched: a client that gates destructive calls on
 * the annotation must see the truth about the dispatcher, not about the
 * dispatch.
 */
const ALWAYS_DESTRUCTIVE = new Set(['graph_request', 'call_tool']);

/** Writes that land on the same state however many times they run. */
const IDEMPOTENT_WRITE = /(?:^|_)(?:update|mark|move|complete|set)(?:_|$)/;

/**
 * Assembles the enabled modules' definitions in catalogue order.
 *
 * The generic module is appended unconditionally: it is not selectable with
 * `--groups`, and the escape hatch plus `graph_list_permissions` are what a
 * caller needs when a domain group turns out to be switched off.
 */
export function collectTools(deps: ToolDeps): ToolDefinition[] {
  const enabled = new Set(deps.config.groups);
  const tools: ToolDefinition[] = [];

  for (const name of GROUP_NAMES) {
    if (!enabled.has(name)) continue;
    const module = MODULES[name];
    if (module === undefined) {
      logger.warn('No tool module implements this group; skipping.', { group: name });
      continue;
    }
    tools.push(...module.build(deps));
  }

  tools.push(...genericModule.build(deps));
  return tools;
}

/** Renders a Graph failure as text the model can act on without a stack trace. */
function describeGraphError(error: GraphError): string {
  const lines = [
    `Microsoft Graph returned ${error.status} (${error.code}) for ${error.path}.`,
    error.message,
  ];
  if (error.hint !== undefined) lines.push(`Hint: ${error.hint}`);
  if (error.requestId !== undefined) lines.push(`Graph request id: ${error.requestId}`);
  return lines.join('\n');
}

function signInMessage(error: InteractionRequiredError): string {
  return [
    'Not signed in to Microsoft 365, or the stored session can no longer be refreshed.',
    'The user must run `npx @devyhan/ms-graph-mcp login` in a terminal, complete the sign-in, ' +
      'and then retry this call. This server cannot prompt for credentials itself, because ' +
      'the MCP transport owns stdio.',
    `Detail: ${error.message}`,
  ].join('\n');
}

interface FactoryDeps {
  config: ServerConfig;
  graph: GraphClient;
  auth: AuthProvider;
  tools: ToolDefinition[];
}

/**
 * Builds the factory handed to `serveStdio`.
 *
 * `auth` is held rather than used: it is what makes the closure a complete
 * description of the session, and callers already have it to hand.
 */
export function createServerFactory(deps: FactoryDeps): () => McpServer {
  const { config, tools } = deps;

  const available = config.readOnly ? tools.filter((tool) => tool.write !== true) : tools;

  // Discovery gets the unfiltered catalogue on purpose: it applies --read-only
  // itself, and knowing a suppressed tool exists is how `call_tool` can explain
  // why it will not run rather than claiming the name is unknown.
  const exposed = config.discovery
    ? [
        ...buildDiscoveryTools(tools, config),
        ...available.filter((tool) => tool.name === PERMISSIONS_TOOL),
      ]
    : available;

  return function factory(): McpServer {
    const server = new McpServer(
      { name: SERVER_NAME, version: packageVersion() },
      { capabilities: { tools: {} } },
    );

    for (const def of exposed) {
      const write = def.write === true;
      server.registerTool(
        def.name,
        {
          title: def.title,
          description: def.description,
          inputSchema: def.inputSchema,
          annotations: {
            readOnlyHint: !write,
            destructiveHint: write && (DESTRUCTIVE.test(def.name) || ALWAYS_DESTRUCTIVE.has(def.name)),
            idempotentHint: !write || IDEMPOTENT_WRITE.test(def.name),
          },
        },
        async (args: unknown) => {
          try {
            const value = await def.handler((args ?? {}) as Record<string, unknown>);
            const { text } = serializeResult(value, config.maxOutputChars);
            return { content: [{ type: 'text' as const, text }] };
          } catch (error) {
            return {
              content: [{ type: 'text' as const, text: toolErrorText(def.name, error) }],
              isError: true,
            };
          }
        },
      );
    }

    logger.debug('Registered tools for a new connection.', {
      count: exposed.length,
      discovery: config.discovery,
      readOnly: config.readOnly,
    });

    return server;
  };
}

/**
 * A stack trace tells the model nothing it can use and can carry a path or a
 * token fragment, so it goes to stderr and only the message goes on the wire.
 */
function toolErrorText(name: string, error: unknown): string {
  if (error instanceof GraphError) return describeGraphError(error);
  if (error instanceof InteractionRequiredError) return signInMessage(error);

  if (error instanceof Error) {
    logger.error('Tool failed.', { tool: name, stack: error.stack ?? error.message });
    return `${name} failed: ${error.message}`;
  }

  logger.error('Tool failed with a non-Error value.', { tool: name, value: String(error) });
  return `${name} failed: ${String(error)}`;
}
