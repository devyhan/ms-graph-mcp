#!/usr/bin/env node
/**
 * Runs the compiled test files, by name.
 *
 * Neither obvious spelling works across the Node versions in CI:
 *
 *   node --test dist/
 *     Node 20 selects only files matching the test-file patterns — 180 tests.
 *     Node 22 and 24 treat every file under an explicitly named directory as a
 *     test, so they also execute `dist/index.js`, which is the CLI. It starts
 *     serving MCP over stdio and never exits, and the job hangs until the
 *     six-hour limit cancels it. Every CI run on this repository did exactly
 *     that before this script existed.
 *
 *   node --test "dist/**‌/*.test.js"
 *     Node 21+ expands the glob itself and this is correct. Node 20 has no
 *     native glob, so the shell expands it, `**` does not recurse in sh, and
 *     only the top level matches — 67 of 180 tests, silently.
 *
 * Enumerating the files here is boring and behaves the same on 20, 22 and 24,
 * and on Windows, where neither shell glob would have worked either. Extra
 * arguments pass through, so `npm test -- --test-name-pattern=fold` works.
 */
import { spawn } from 'node:child_process';
import { readdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..', 'dist');

/** Compiled test files, sorted so failures are reported in a stable order. */
async function findTests(dir) {
  const found = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) found.push(...(await findTests(full)));
    else if (entry.name.endsWith('.test.js')) found.push(full);
  }
  return found.sort();
}

let files;
try {
  files = await findTests(ROOT);
} catch (error) {
  console.error(`Cannot read ${ROOT}: ${error.message}`);
  console.error('Run `npm run build` first — `npm test` does it for you via pretest.');
  process.exit(1);
}

if (files.length === 0) {
  // Silence here would look like a green run with nothing in it, which is how
  // the 67-of-180 version of this went unnoticed.
  console.error(`No *.test.js under ${ROOT}. Did the build emit anything?`);
  process.exit(1);
}

// Passthrough args go BEFORE the file list: these are node's own options, and
// node stops treating them as such once file arguments have started. That is
// how `--test-reporter=tap` reaches node rather than the test files.
const child = spawn(process.execPath, ['--test', ...process.argv.slice(2), ...files], {
  stdio: 'inherit',
});
child.on('exit', (code, signal) => {
  if (signal !== null) process.kill(process.pid, signal);
  else process.exit(code ?? 1);
});
