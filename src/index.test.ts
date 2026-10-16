import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { scopeReadingOnlyAuth } from './auth/index.js';
import { parseConfig } from './config.js';
import type { GraphClient, ToolDefinition } from './contracts.js';
import { collectTools, exposedTools } from './server.js';
import { consentScopesFor } from './tools/groups.js';

/** The compiled CLI next to this compiled test file. */
const CLI = join(fileURLToPath(new URL('.', import.meta.url)), 'index.js');

/**
 * Flags that change which tools exist are driven by argv in these tests, so any
 * ambient environment fallback for them is removed — otherwise a developer's
 * exported `MS365_MCP_READ_ONLY` would silently rewrite the expectation.
 */
function testEnv(cacheDir: string): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = {
    ...process.env,
    MS365_MCP_CACHE_DIR: cacheDir,
    // Suppresses the update check: no registry fetch, no waiting, no network.
    CI: '1',
  };
  delete env['MS365_MCP_GROUPS'];
  delete env['MS365_MCP_PRESET'];
  delete env['MS365_MCP_READ_ONLY'];
  delete env['MS365_MCP_DISCOVERY'];
  delete env['MS365_MCP_ORG_MODE'];
  return env;
}

/** Runs the CLI with the given args and returns the exit code plus captured output. */
async function runCli(
  args: string[],
  env: NodeJS.ProcessEnv,
): Promise<{ code: number; stdout: string; stderr: string }> {
  const child = spawn(process.execPath, [CLI, ...args], { env });
  let stdout = '';
  let stderr = '';
  child.stdout.on('data', (chunk: Buffer) => {
    stdout += chunk.toString('utf8');
  });
  child.stderr.on('data', (chunk: Buffer) => {
    stderr += chunk.toString('utf8');
  });
  const code = await new Promise<number>((resolve, reject) => {
    child.on('error', reject);
    child.on('close', (c) => resolve(c ?? 1));
  });
  return { code, stdout, stderr };
}

/** Runs `status` with the given flags and returns its stdout. */
async function statusOutput(flags: string[], env: NodeJS.ProcessEnv): Promise<string> {
  const child = spawn(process.execPath, [CLI, 'status', ...flags], { env });
  let stdout = '';
  let stderr = '';
  child.stdout.on('data', (chunk: Buffer) => {
    stdout += chunk.toString('utf8');
  });
  child.stderr.on('data', (chunk: Buffer) => {
    stderr += chunk.toString('utf8');
  });
  const code = await new Promise<number>((resolve, reject) => {
    child.on('error', reject);
    child.on('close', (c) => resolve(c ?? 1));
  });
  assert.equal(code, 0, `status ${flags.join(' ')} exited ${code}: ${stderr}`);
  return stdout;
}

function toolsLine(stdout: string): number {
  const match = stdout.match(/^Tools:\s+(\d+)$/m);
  assert.ok(match?.[1] !== undefined, `no Tools: line in:\n${stdout}`);
  return Number(match[1]);
}

/**
 * The count `status` should print: the same catalogue `runStatus` builds,
 * passed through the same exposure filter the server applies.
 */
function expectedExposedCount(flags: string[], env: NodeJS.ProcessEnv): number {
  const { config } = parseConfig(['status', ...flags], env);
  const refuse = (): never => {
    throw new Error('status builds definitions without calling Graph');
  };
  const graph: GraphClient = { request: refuse, batch: refuse, follow: refuse };
  let built: ToolDefinition[] = [];
  built = collectTools({
    graph,
    config,
    auth: scopeReadingOnlyAuth(),
    consentScopes: () => consentScopesFor(config, built),
  });
  return exposedTools(built, config).length;
}

async function statusToolsCount(flags: string[]): Promise<{ reported: number; expected: number }> {
  const cacheDir = await mkdtemp(join(tmpdir(), 'ms-graph-mcp-status-'));
  const env = testEnv(cacheDir);
  const reported = toolsLine(await statusOutput(flags, env));
  return { reported, expected: expectedExposedCount(flags, env) };
}

test('status reports the exposed tool count, not the raw catalogue', async () => {
  const { reported, expected } = await statusToolsCount([]);
  assert.equal(reported, expected);
});

test('status --read-only reports the count with write tools suppressed', async () => {
  const { reported, expected } = await statusToolsCount(['--read-only']);
  assert.equal(reported, expected);
  const full = await statusToolsCount([]);
  assert.ok(
    reported < full.reported,
    `expected the read-only count (${reported}) below the full count (${full.reported})`,
  );
});

test('status --discovery reports the handful a client actually sees', async () => {
  const { reported, expected } = await statusToolsCount(['--discovery']);
  assert.equal(reported, expected);
  const full = await statusToolsCount([]);
  assert.ok(
    reported < full.reported,
    `expected the discovery count (${reported}) below the full count (${full.reported})`,
  );
});

function groupsLine(stdout: string): string {
  const match = stdout.match(/^Groups:\s+(.*)$/m);
  assert.ok(match?.[1] !== undefined, `no Groups: line in:\n${stdout}`);
  return match[1];
}

test('logout with no session says so instead of claiming a removal', async () => {
  const cacheDir = join(await mkdtemp(join(tmpdir(), 'ms-graph-mcp-logout-')), 'fresh-cache');
  const { code, stdout } = await runCli(['logout'], testEnv(cacheDir));
  assert.equal(code, 0);
  assert.match(stdout, /No session found in .*fresh-cache: nothing to remove\./);
  assert.doesNotMatch(stdout, /Signed out\. Removed/);
});

test('logout with a session reports what it removed', async () => {
  const cacheDir = await mkdtemp(join(tmpdir(), 'ms-graph-mcp-logout-'));
  await writeFile(join(cacheDir, 'token-cache.enc'), 'not a real cache', 'utf8');
  await writeFile(join(cacheDir, 'cache.key'), 'not a real key', 'utf8');
  const { code, stdout } = await runCli(['logout'], testEnv(cacheDir));
  assert.equal(code, 0);
  assert.match(stdout, /Signed out\. Removed .*token-cache\.enc and .*cache\.key\./);
  await assert.rejects(stat(join(cacheDir, 'token-cache.enc')));
  await assert.rejects(stat(join(cacheDir, 'cache.key')));
});

test('status names dropped groups and the --org-mode remedy on the Groups line', async () => {
  const cacheDir = await mkdtemp(join(tmpdir(), 'ms-graph-mcp-status-'));
  const { code, stdout } = await runCli(['status', '--groups', 'teams'], testEnv(cacheDir));
  assert.equal(code, 0);
  const line = groupsLine(stdout);
  assert.match(line, /teams skipped/);
  assert.match(line, /--org-mode/);
});

test('status shows kept groups alongside the dropped note', async () => {
  const cacheDir = await mkdtemp(join(tmpdir(), 'ms-graph-mcp-status-'));
  const { code, stdout } = await runCli(['status', '--groups', 'mail,teams'], testEnv(cacheDir));
  assert.equal(code, 0);
  const line = groupsLine(stdout);
  assert.match(line, /mail \(\+ generic\)/);
  assert.match(line, /teams skipped/);
});

test('status with --org-mode shows no dropped note', async () => {
  const cacheDir = await mkdtemp(join(tmpdir(), 'ms-graph-mcp-status-'));
  const { code, stdout } = await runCli(
    ['status', '--groups', 'teams', '--org-mode'],
    testEnv(cacheDir),
  );
  assert.equal(code, 0);
  assert.doesNotMatch(groupsLine(stdout), /skipped/);
});

/** Runs `permissions` with the given flags and returns its stdout. */
async function permissionsOutput(flags: string[]): Promise<string> {
  const cacheDir = await mkdtemp(join(tmpdir(), 'ms-graph-mcp-permissions-'));
  const { code, stdout, stderr } = await runCli(['permissions', ...flags], testEnv(cacheDir));
  assert.equal(code, 0, `permissions ${flags.join(' ')} exited ${code}: ${stderr}`);
  return stdout;
}

test('permissions qualifies the no-admin closing line with the tenant-policy gate', async () => {
  // The unqualified line contradicted the consent disclosure: Microsoft marks
  // these scopes user-consentable, but the default tenant policy still withholds
  // most of them from ordinary users. Both statements must agree.
  const stdout = await permissionsOutput([]);
  assert.match(stdout, /under Microsoft's classification/);
  assert.match(stdout, /tenant consent policy may still withhold/);
});

test('permissions annotates cross-group scopes with a short reason', async () => {
  const stdout = await permissionsOutput([]);
  assert.match(stdout, /MailboxSettings\.Read — me_get_mailbox_settings only/);
  assert.match(stdout, /Calendars\.Read\.Shared — other people's free\/busy/);
  assert.match(stdout, /Files\.Read\.All — items in others' drives/);
  assert.match(stdout, /Mail\.Read — mail attachments \(files_parse_document\)/);
  assert.match(stdout, /Chat\.Read — chatMessage index/);
  assert.match(stdout, /Calendars\.Read — event index/);
  // Scopes a group obviously owns carry no annotation.
  assert.match(stdout, /^  Mail\.Read$/m);
});

test('permissions annotates the teams roster scope under --org-mode', async () => {
  const stdout = await permissionsOutput(['--groups', 'teams', '--org-mode']);
  assert.match(stdout, /TeamMember\.Read\.All.* — team roster only/);
});
