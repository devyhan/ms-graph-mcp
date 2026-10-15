/**
 * Update notice for human CLI runs, in the update-notifier pattern.
 *
 * Once per 24h (per installed version) the latest published version is fetched
 * from the npm registry with a short timeout; when it is newer than the
 * running package, the caller prints a single stderr line. The check never
 * throws, never writes to stdout, and stays silent when suppressed
 * (`NO_UPDATE_NOTIFIER`, CI), offline, or on any error.
 */

import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';

import { cachePaths } from '../auth/token-cache.js';

/** How long a registry answer is trusted before it is re-fetched. */
export const UPDATE_CHECK_INTERVAL_MS = 24 * 60 * 60 * 1000;

/** Upper bound on the registry request; the CLI never waits longer than this. */
export const UPDATE_CHECK_TIMEOUT_MS = 3000;

/** Cache filename, stored beside the token cache. */
export const UPDATE_CHECK_FILE = 'update-check.json';

/** Where the notice points for what changed between releases. */
export const CHANGELOG_URL = 'https://github.com/devyhan/ms-graph-mcp/blob/main/CHANGELOG.md';

/** Environment markers that suppress the check. */
const CI_MARKERS = [
  'CI',
  'CONTINUOUS_INTEGRATION',
  'BUILD_NUMBER',
  'GITHUB_ACTIONS',
  'GITLAB_CI',
  'CIRCLECI',
  'TRAVIS',
  'JENKINS_URL',
  'TEAMCITY_VERSION',
  'TF_BUILD',
  'BITBUCKET_BUILD_NUMBER',
];

/** Minimal fetch surface the check needs; the global fetch satisfies it. */
export type RegistryFetch = (
  url: string,
  init?: { signal?: AbortSignal },
) => Promise<{ ok: boolean; status: number; json(): Promise<unknown> }>;

export interface UpdateCheckOptions {
  /** Token cache directory; the check cache lives beside it. */
  cacheDir: string;
  /** Currently running package version. */
  currentVersion: string;
  /** Registry package name, e.g. `@devyhan/ms-graph-mcp`. */
  packageName: string;
  /** Defaults to `process.env`; injectable so tests stay hermetic. */
  env?: NodeJS.ProcessEnv;
  /** Defaults to `Date.now`; injectable so tests can age the cache. */
  now?: () => number;
  /** Defaults to the global fetch; injectable so tests stay offline. */
  fetchImpl?: RegistryFetch;
  /** Registry timeout in ms; defaults to `UPDATE_CHECK_TIMEOUT_MS`. */
  timeoutMs?: number;
}

interface UpdateCheckCache {
  /** Version that performed the check; a mismatch invalidates the entry. */
  installedVersion: string;
  /** Last fetched `latest`, or null when the fetch failed. */
  latestVersion: string | null;
  /** Epoch millis of the last fetch attempt, successful or not. */
  checkedAt: number;
}

/** Cache path derived from the token-cache location, so both move together. */
export function updateCachePath(cacheDir: string): string {
  return join(dirname(cachePaths(cacheDir).cacheFile), UPDATE_CHECK_FILE);
}

function isTruthyEnv(value: string | undefined): boolean {
  if (value === undefined) return false;
  const normalized = value.trim().toLowerCase();
  return normalized !== '' && normalized !== '0' && normalized !== 'false';
}

/** True when the operator or environment asks for silence. */
export function isUpdateCheckSuppressed(env: NodeJS.ProcessEnv = process.env): boolean {
  if (isTruthyEnv(env['NO_UPDATE_NOTIFIER'])) return true;
  return CI_MARKERS.some((key) => isTruthyEnv(env[key]));
}

/** Numeric prefix of a version (`v1.2.3`, `1.2.3-beta`, `1.2`), or undefined. */
function parseNumericVersion(version: string): number[] | undefined {
  const core = version.trim().replace(/^[v=\s]+/i, '').split('+')[0]?.split('-')[0] ?? '';
  if (core === '') return undefined;
  const parts = core.split('.');
  const numbers: number[] = [];
  for (const part of parts) {
    if (!/^\d+$/.test(part)) return undefined;
    numbers.push(Number(part));
  }
  return numbers.length === 0 ? undefined : numbers;
}

/** Numeric semver comparison; unparseable input compares as "not newer". */
export function isNewerVersion(latest: string, current: string): boolean {
  const a = parseNumericVersion(latest);
  const b = parseNumericVersion(current);
  if (a === undefined || b === undefined) return false;
  const length = Math.max(a.length, b.length);
  for (let i = 0; i < length; i += 1) {
    const x = a[i] ?? 0;
    const y = b[i] ?? 0;
    if (x !== y) return x > y;
  }
  return false;
}

/** The single stderr line printed when an update is available. No newline. */
export function formatUpdateNotice(
  packageName: string,
  currentVersion: string,
  latestVersion: string,
): string {
  const shortName = packageName.includes('/')
    ? packageName.slice(packageName.lastIndexOf('/') + 1)
    : packageName;
  return (
    `${shortName} update available: installed ${currentVersion}, latest ${latestVersion}. ` +
    `See CHANGELOG.md: ${CHANGELOG_URL}`
  );
}

function isCacheShape(value: unknown): value is UpdateCheckCache {
  if (typeof value !== 'object' || value === null) return false;
  const record = value as Record<string, unknown>;
  return (
    typeof record['installedVersion'] === 'string' &&
    (typeof record['latestVersion'] === 'string' || record['latestVersion'] === null) &&
    typeof record['checkedAt'] === 'number' &&
    Number.isFinite(record['checkedAt'])
  );
}

async function readUpdateCache(path: string): Promise<UpdateCheckCache | undefined> {
  try {
    const parsed: unknown = JSON.parse(await readFile(path, 'utf8'));
    return isCacheShape(parsed) ? parsed : undefined;
  } catch {
    return undefined;
  }
}

async function writeUpdateCache(path: string, cache: UpdateCheckCache): Promise<void> {
  try {
    await mkdir(dirname(path), { recursive: true });
    await writeFile(path, JSON.stringify(cache), 'utf8');
  } catch {
    // A cache that cannot persist only costs a re-fetch next run.
  }
}

async function fetchLatestVersion(
  packageName: string,
  fetchImpl: RegistryFetch,
  timeoutMs: number,
): Promise<string | undefined> {
  try {
    const url = `https://registry.npmjs.org/${encodeURIComponent(packageName)}/latest`;
    const response = await fetchImpl(url, { signal: AbortSignal.timeout(timeoutMs) });
    if (!response.ok) return undefined;
    const body: unknown = await response.json();
    if (typeof body !== 'object' || body === null || !('version' in body)) return undefined;
    const version = (body as { version?: unknown }).version;
    return typeof version === 'string' && version.trim() !== '' ? version.trim() : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Returns the one-line update notice when a newer release is known, else
 * undefined. Reads the 24h cache first and re-fetches only when it is missing,
 * stale, or belongs to another installed version. Never throws.
 */
export async function checkForUpdate(options: UpdateCheckOptions): Promise<string | undefined> {
  try {
    const { cacheDir, currentVersion, packageName } = options;
    if (currentVersion.trim() === '' || packageName.trim() === '') return undefined;
    if (isUpdateCheckSuppressed(options.env ?? process.env)) return undefined;

    const path = updateCachePath(cacheDir);
    const cached = await readUpdateCache(path);
    const nowMs = (options.now ?? Date.now)();
    if (
      cached !== undefined &&
      cached.installedVersion === currentVersion &&
      nowMs - cached.checkedAt < UPDATE_CHECK_INTERVAL_MS
    ) {
      if (cached.latestVersion !== null && isNewerVersion(cached.latestVersion, currentVersion)) {
        return formatUpdateNotice(packageName, currentVersion, cached.latestVersion);
      }
      return undefined;
    }

    const latest = await fetchLatestVersion(
      packageName,
      options.fetchImpl ?? fetch,
      options.timeoutMs ?? UPDATE_CHECK_TIMEOUT_MS,
    );
    if (latest === undefined) {
      // Cache the attempt (keeping the previous answer, if any) so an offline
      // machine does not retry on every invocation; this run stays silent.
      await writeUpdateCache(path, {
        installedVersion: currentVersion,
        latestVersion:
          cached !== undefined && cached.installedVersion === currentVersion
            ? cached.latestVersion
            : null,
        checkedAt: nowMs,
      });
      return undefined;
    }

    await writeUpdateCache(path, {
      installedVersion: currentVersion,
      latestVersion: latest,
      checkedAt: nowMs,
    });
    return isNewerVersion(latest, currentVersion)
      ? formatUpdateNotice(packageName, currentVersion, latest)
      : undefined;
  } catch {
    return undefined;
  }
}
