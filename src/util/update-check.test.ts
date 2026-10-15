import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import test from 'node:test';

import { cachePaths } from '../auth/token-cache.js';
import {
  UPDATE_CHECK_INTERVAL_MS,
  checkForUpdate,
  formatUpdateNotice,
  isNewerVersion,
  isUpdateCheckSuppressed,
  updateCachePath,
} from './update-check.js';
import type { RegistryFetch } from './update-check.js';

const PACKAGE = '@devyhan/ms-graph-mcp';

/** Every test injects `env: {}` so the real CI environment cannot suppress it. */
async function tempDir(): Promise<string> {
  return mkdtemp(join(tmpdir(), 'update-check-'));
}

function fetchReturning(version: string, calls: { count: number }): RegistryFetch {
  return (async () => {
    calls.count += 1;
    return { ok: true, status: 200, json: async () => ({ version }) };
  }) as RegistryFetch;
}

function fetchFailing(calls: { count: number }): RegistryFetch {
  return (async () => {
    calls.count += 1;
    throw new Error('offline');
  }) as RegistryFetch;
}

test('numeric compare treats 0.10.0 as newer than 0.9.0', () => {
  assert.equal(isNewerVersion('0.10.0', '0.9.0'), true);
  assert.equal(isNewerVersion('0.5.0', '0.4.0'), true);
  assert.equal(isNewerVersion('0.4.0', '0.4.0'), false);
  assert.equal(isNewerVersion('0.3.9', '0.4.0'), false);
  assert.equal(isNewerVersion('1.0', '0.9.9'), true);
});

test('numeric compare stays silent on unparseable versions', () => {
  assert.equal(isNewerVersion('soon', '0.4.0'), false);
  assert.equal(isNewerVersion('0.5.0', 'not-a-version'), false);
  assert.equal(isNewerVersion('', '0.4.0'), false);
});

test('suppression honors NO_UPDATE_NOTIFIER and CI', () => {
  assert.equal(isUpdateCheckSuppressed({ NO_UPDATE_NOTIFIER: '1' }), true);
  assert.equal(isUpdateCheckSuppressed({ CI: 'true' }), true);
  assert.equal(isUpdateCheckSuppressed({ GITHUB_ACTIONS: 'true' }), true);
  assert.equal(isUpdateCheckSuppressed({ NO_UPDATE_NOTIFIER: '0' }), false);
  assert.equal(isUpdateCheckSuppressed({ CI: '' }), false);
  assert.equal(isUpdateCheckSuppressed({}), false);
});

test('a suppressed check never touches the network', async () => {
  const calls = { count: 0 };
  const notice = await checkForUpdate({
    cacheDir: await tempDir(),
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: { NO_UPDATE_NOTIFIER: '1' },
    fetchImpl: fetchFailing(calls),
  });
  assert.equal(notice, undefined);
  assert.equal(calls.count, 0);
});

test('an outdated install yields one line naming installed, latest and CHANGELOG', async () => {
  const calls = { count: 0 };
  const notice = await checkForUpdate({
    cacheDir: await tempDir(),
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    fetchImpl: fetchReturning('0.5.0', calls),
  });
  assert.ok(notice !== undefined);
  assert.ok(notice.includes('0.4.0'));
  assert.ok(notice.includes('0.5.0'));
  assert.ok(notice.includes('CHANGELOG'));
  assert.equal(notice.includes('\n'), false);
  assert.equal(calls.count, 1);
});

test('formatUpdateNotice strips the scope and stays on one line', () => {
  const notice = formatUpdateNotice(PACKAGE, '0.4.0', '0.5.0');
  assert.ok(notice.startsWith('ms-graph-mcp update available:'));
  assert.equal(notice.includes('\n'), false);
});

test('an up-to-date install stays silent', async () => {
  const calls = { count: 0 };
  const notice = await checkForUpdate({
    cacheDir: await tempDir(),
    currentVersion: '0.5.0',
    packageName: PACKAGE,
    env: {},
    fetchImpl: fetchReturning('0.5.0', calls),
  });
  assert.equal(notice, undefined);
  assert.equal(calls.count, 1);
});

test('a fetch failure stays silent and is not retried within 24h', async () => {
  const cacheDir = await tempDir();
  const calls = { count: 0 };
  const failing = fetchFailing(calls);
  const base = {
    cacheDir,
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    fetchImpl: failing,
  };
  assert.equal(await checkForUpdate(base), undefined);
  assert.equal(await checkForUpdate(base), undefined);
  assert.equal(calls.count, 1);
});

test('a fresh cache answers without fetching', async () => {
  const cacheDir = await tempDir();
  const warm = { count: 0 };
  const first = await checkForUpdate({
    cacheDir,
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    fetchImpl: fetchReturning('0.5.0', warm),
  });
  assert.ok(first !== undefined);

  const cold = { count: 0 };
  const second = await checkForUpdate({
    cacheDir,
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    fetchImpl: fetchFailing(cold),
  });
  assert.equal(second, first);
  assert.equal(cold.count, 0);
  assert.equal(warm.count, 1);
});

test('a stale cache re-fetches after 24h', async () => {
  const cacheDir = await tempDir();
  const now = { at: 1_000_000 };
  const firstCalls = { count: 0 };
  await checkForUpdate({
    cacheDir,
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    now: () => now.at,
    fetchImpl: fetchReturning('0.4.0', firstCalls),
  });
  assert.equal(firstCalls.count, 1);

  now.at += UPDATE_CHECK_INTERVAL_MS + 1;
  const secondCalls = { count: 0 };
  const notice = await checkForUpdate({
    cacheDir,
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    now: () => now.at,
    fetchImpl: fetchReturning('0.5.0', secondCalls),
  });
  assert.equal(secondCalls.count, 1);
  assert.ok(notice?.includes('0.5.0'));
});

test('the cache is keyed per installed version', async () => {
  const cacheDir = await tempDir();
  const base = {
    cacheDir,
    packageName: PACKAGE,
    env: {},
  };
  const firstCalls = { count: 0 };
  await checkForUpdate({
    ...base,
    currentVersion: '0.4.0',
    fetchImpl: fetchReturning('0.4.0', firstCalls),
  });
  assert.equal(firstCalls.count, 1);

  // Same 24h window, but a new install must not trust the old entry.
  const secondCalls = { count: 0 };
  await checkForUpdate({
    ...base,
    currentVersion: '0.5.0',
    fetchImpl: fetchReturning('0.5.0', secondCalls),
  });
  assert.equal(secondCalls.count, 1);
});

test('a corrupt cache is ignored and the check re-fetches', async () => {
  const cacheDir = await tempDir();
  await writeFile(updateCachePath(cacheDir), 'not json{{{', 'utf8');
  const calls = { count: 0 };
  const notice = await checkForUpdate({
    cacheDir,
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    fetchImpl: fetchReturning('0.5.0', calls),
  });
  assert.equal(calls.count, 1);
  assert.ok(notice?.includes('0.5.0'));
});

test('the check cache sits beside the token cache', async () => {
  const cacheDir = await tempDir();
  assert.equal(dirname(updateCachePath(cacheDir)), dirname(cachePaths(cacheDir).cacheFile));

  const calls = { count: 0 };
  await checkForUpdate({
    cacheDir,
    currentVersion: '0.4.0',
    packageName: PACKAGE,
    env: {},
    fetchImpl: fetchReturning('0.4.0', calls),
  });
  const raw = await readFile(updateCachePath(cacheDir), 'utf8');
  const parsed = JSON.parse(raw) as { installedVersion: string; latestVersion: string };
  assert.equal(parsed.installedVersion, '0.4.0');
  assert.equal(parsed.latestVersion, '0.4.0');
});
