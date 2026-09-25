/**
 * Discovery mode: two tools that stand in for the whole catalogue.
 *
 * A hundred-odd tool schemas is a large, permanent tax on the first turn of
 * every conversation. Under `--discovery` the client sees `discover_tools` and
 * `call_tool` instead, and pays for a tool's schema only when it wants one.
 */

import { z } from 'zod';

import type { ServerConfig, ToolDefinition } from '../contracts.js';

/** Group name reported for the discovery tools themselves. */
const GROUP = 'discovery';

/** How many name suggestions an unknown `call_tool` name is worth. */
const SUGGESTION_COUNT = 5;

const discoverInput = z.object({
  query: z
    .string()
    .optional()
    .describe(
      'Case-insensitive search term matched against tool names, titles and descriptions, ' +
        'e.g. "attachment", "calendar" or "upload". Omit it to browse by group.',
    ),
  group: z
    .string()
    .optional()
    .describe(
      'Restrict the result to one tool group, e.g. "mail" or "files". Call this tool with no ' +
        'arguments to see the available group names and how many tools each holds.',
    ),
});

const callInput = z.object({
  name: z
    .string()
    .describe('Exact tool name from discover_tools, e.g. "mail_list_messages".'),
  arguments: z
    .record(z.string(), z.unknown())
    .optional()
    .describe(
      'Arguments for the target tool, as a JSON object. Pass {} for a tool that takes none. ' +
        'They are validated against the target\'s own schema, and a mismatch comes back as a ' +
        'list of what was wrong rather than a silent failure.',
    ),
});

/** The summary shape `discover_tools` returns — deliberately without schemas. */
interface ToolSummary {
  name: string;
  group: string;
  title: string;
  description: string;
  write: boolean;
}

function summarize(tool: ToolDefinition): ToolSummary {
  return {
    name: tool.name,
    group: tool.group,
    title: tool.title,
    description: tool.description,
    write: tool.write === true,
  };
}

/** Levenshtein distance, capped implicitly by the short strings involved. */
function editDistance(a: string, b: string): number {
  if (a === b) return 0;
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  // One rolling row instead of the full matrix: names are short, but this runs
  // over the entire catalogue on every miss.
  let previous: number[] = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i += 1) {
    const current: number[] = [i];
    for (let j = 1; j <= b.length; j += 1) {
      const substitution = (previous[j - 1] ?? 0) + (a[i - 1] === b[j - 1] ? 0 : 1);
      const deletion = (previous[j] ?? 0) + 1;
      const insertion = (current[j - 1] ?? 0) + 1;
      current[j] = Math.min(substitution, deletion, insertion);
    }
    previous = current;
  }
  return previous[b.length] ?? Math.max(a.length, b.length);
}

/**
 * Names closest to `wanted`. Substring hits come first because a model that
 * asks for `list_messages` wants `mail_list_messages`, which is four edits away
 * but an obvious match.
 */
function suggestNames(wanted: string, tools: ToolDefinition[]): string[] {
  const needle = wanted.toLowerCase();
  const scored = tools.map((tool) => {
    const name = tool.name.toLowerCase();
    const contains = name.includes(needle) || needle.includes(name);
    return { name: tool.name, rank: contains ? -1 : editDistance(needle, name) };
  });
  scored.sort((a, b) => a.rank - b.rank || a.name.localeCompare(b.name));
  return scored.slice(0, SUGGESTION_COUNT).map((entry) => entry.name);
}

function formatIssues(error: z.ZodError): string {
  return error.issues
    .map((issue) => {
      const where = issue.path.length === 0 ? '(root)' : issue.path.join('.');
      return `  ${where}: ${issue.message}`;
    })
    .join('\n');
}

/**
 * Builds the discovery pair.
 *
 * `config` is a parameter rather than a closure over module state because
 * `call_tool` is the one place a write tool can be reached without passing
 * through the registration layer's `--read-only` filter, so it re-checks here.
 */
export function buildDiscoveryTools(
  all: ToolDefinition[],
  config: ServerConfig,
): ToolDefinition[] {
  // Lookup spans the whole catalogue so a suppressed write tool can be answered
  // with "this server is read-only" rather than the misleading "no such tool".
  const byName = new Map(all.map((tool) => [tool.name, tool]));
  const visible = config.readOnly ? all.filter((tool) => tool.write !== true) : all;

  return [
    {
      name: 'discover_tools',
      title: 'Find Microsoft 365 tools',
      description:
        'Searches the Microsoft 365 tool catalogue and returns matching tools as name, group, ' +
        'title, description and whether they write. It deliberately omits argument schemas — ' +
        'call the tool through call_tool and it will tell you exactly what it expected if the ' +
        'arguments are wrong. Called with no arguments it lists the groups (mail, calendar, ' +
        'files, and so on) with a tool count for each, which is the cheapest way to orient. ' +
        'Every tool it names is callable through call_tool.',
      inputSchema: discoverInput,
      scopes: [],
      group: GROUP,
      handler: async (args) => {
        const { query, group } = discoverInput.parse(args);

        if (query === undefined && group === undefined) {
          const counts = new Map<string, number>();
          for (const tool of visible) {
            counts.set(tool.group, (counts.get(tool.group) ?? 0) + 1);
          }
          return {
            totalTools: visible.length,
            groups: [...counts.entries()].map(([name, toolCount]) => ({ name, toolCount })),
            note:
              'Call discover_tools again with `group` set to one of these names, or with a ' +
              '`query` term, to see the tools themselves.',
          };
        }

        const wantedGroup = group?.trim().toLowerCase();
        const needle = query?.trim().toLowerCase();

        const matches = visible.filter((tool) => {
          if (wantedGroup !== undefined && tool.group.toLowerCase() !== wantedGroup) return false;
          if (needle === undefined || needle === '') return true;
          const haystack = `${tool.name} ${tool.title} ${tool.description}`.toLowerCase();
          return haystack.includes(needle);
        });

        if (matches.length === 0) {
          const groups = [...new Set(visible.map((tool) => tool.group))];
          return {
            count: 0,
            tools: [],
            note:
              wantedGroup !== undefined && !groups.includes(wantedGroup)
                ? `No tool group named "${group}". Available groups: ${groups.join(', ')}.`
                : 'Nothing matched. Try a broader term, or call discover_tools with no ' +
                  'arguments to browse the groups.',
          };
        }

        return { count: matches.length, tools: matches.map(summarize) };
      },
    },
    {
      name: 'call_tool',
      title: 'Call a Microsoft 365 tool',
      description:
        'Invokes a tool found with discover_tools by name, passing `arguments` straight through ' +
        'to it. The arguments are validated against that tool\'s own schema first, so a wrong ' +
        'or missing field comes back as a list of exactly which fields were wrong. An unknown ' +
        'name comes back with the closest matching names. The result is whatever the target ' +
        'tool returns.',
      inputSchema: callInput,
      /**
       * A dispatcher is as dangerous as the most dangerous thing it dispatches
       * to. Left unflagged this registered with `readOnlyHint: true`, so a
       * client that prompts before destructive calls would wave through a
       * `call_tool` that sends mail or deletes a file. Under `--read-only` the
       * catalogue it can reach holds no writes, so there it is honestly a read.
       *
       * Marking it a write also means the registration layer suppresses it
       * under `--read-only` — which is why the flag has to be conditional
       * rather than simply `true`, or discovery mode would lose its dispatcher
       * exactly when it is safe.
       */
      write: !config.readOnly,
      scopes: [],
      group: GROUP,
      handler: async (args) => {
        const { name, arguments: rawArgs } = callInput.parse(args);
        const tool = byName.get(name);

        if (tool === undefined) {
          throw new Error(
            `No tool named "${name}". Closest matches: ${suggestNames(name, visible).join(', ')}. ` +
              'Use discover_tools to search the catalogue.',
          );
        }

        // This is the one entry point that bypasses tools/list, so --read-only is
        // enforced here rather than assumed from what discover_tools advertised.
        if (config.readOnly && tool.write === true) {
          throw new Error(
            `"${name}" modifies data and this server is running in --read-only mode. ` +
              'Restart it without --read-only to allow writes.',
          );
        }

        const parsed = tool.inputSchema.safeParse(rawArgs ?? {});
        if (!parsed.success) {
          throw new Error(
            `Arguments for "${name}" did not validate:\n${formatIssues(parsed.error)}`,
          );
        }

        return tool.handler(parsed.data as Record<string, unknown>);
      },
    },
  ];
}
