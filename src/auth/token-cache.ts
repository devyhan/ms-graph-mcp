/**
 * Persistent MSAL token cache with no native dependencies.
 *
 * An MSAL cache blob runs to several KB, which overflows the per-item limits of
 * some OS credential stores, so the blob is encrypted to a file and only the
 * 32-byte key is kept beside it. Both live under the user's config directory
 * rather than the package directory: npx installs into content-hashed paths, so
 * anything written next to the code is discarded on the next version bump.
 */

import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { chmod, mkdir, open, readFile, rename, rm, unlink, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';

const CACHE_FILE = 'token-cache.enc';
const KEY_FILE = 'cache.key';

const DIR_MODE = 0o700;
const FILE_MODE = 0o600;

const ALGORITHM = 'aes-256-gcm';
const KEY_BYTES = 32;
const IV_BYTES = 12;
const TAG_BYTES = 16;

function warn(message: string): void {
  // stdout belongs to the MCP transport; diagnostics only ever go to stderr.
  process.stderr.write(`[auth] ${message}\n`);
}

function errorCode(err: unknown): string | undefined {
  if (typeof err === 'object' && err !== null && 'code' in err) {
    const code = (err as { code?: unknown }).code;
    if (typeof code === 'string') return code;
  }
  return undefined;
}

function describeError(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

/**
 * Serialises every cache operation for a given directory. MSAL fires
 * before/afterCacheAccess around each token request, and concurrent tool calls
 * would otherwise interleave a read with a half-finished write.
 */
const locks = new Map<string, Promise<unknown>>();

function runExclusive<T>(key: string, fn: () => Promise<T>): Promise<T> {
  const previous = locks.get(key) ?? Promise.resolve();
  // Run regardless of how the predecessor settled, so one failure cannot wedge
  // the chain for the life of the process.
  const next = previous.then(fn, fn);
  locks.set(
    key,
    next.catch(() => undefined),
  );
  return next;
}

async function ensureDir(cacheDir: string): Promise<void> {
  await mkdir(cacheDir, { recursive: true, mode: DIR_MODE });
}

/** Writes through a temp file so a crash mid-write cannot truncate the cache. */
async function writeSecret(path: string, contents: string): Promise<void> {
  const temp = `${path}.${process.pid.toString(36)}${randomBytes(4).toString('hex')}.tmp`;
  await writeFile(temp, contents, { encoding: 'utf8', mode: FILE_MODE });
  try {
    await rename(temp, path);
  } catch (err) {
    await rm(temp, { force: true });
    throw err;
  }
  // rename preserves the temp file's mode, but an existing target replaced on a
  // filesystem that ignores mode-on-create still gets tightened here.
  await chmod(path, FILE_MODE).catch(() => undefined);
}

/**
 * Loads the cache key, creating it on first use. Uses an exclusive create so two
 * processes racing on first run agree on a single key instead of each clobbering
 * the other's.
 */
async function loadOrCreateKey(cacheDir: string): Promise<Buffer> {
  const keyPath = join(cacheDir, KEY_FILE);

  const existing = await readKey(keyPath);
  if (existing) return existing;

  await ensureDir(cacheDir);
  const key = randomBytes(KEY_BYTES);
  try {
    const handle = await open(keyPath, 'wx', FILE_MODE);
    try {
      await handle.writeFile(key.toString('base64'), 'utf8');
    } finally {
      await handle.close();
    }
    return key;
  } catch (err) {
    if (errorCode(err) === 'EEXIST') {
      const raced = await readKey(keyPath);
      if (raced) return raced;
    }
    throw err;
  }
}

async function readKey(keyPath: string): Promise<Buffer | null> {
  let raw: string;
  try {
    raw = await readFile(keyPath, 'utf8');
  } catch (err) {
    if (errorCode(err) === 'ENOENT') return null;
    throw err;
  }
  const key = Buffer.from(raw.trim(), 'base64');
  if (key.length !== KEY_BYTES) {
    // A truncated or hand-edited key can never decrypt the cache; drop both and
    // fall back to a fresh sign-in rather than failing every request forever.
    warn(`cache key at ${keyPath} is malformed; regenerating (sign-in required)`);
    await rm(keyPath, { force: true });
    return null;
  }
  return key;
}

function encrypt(key: Buffer, plaintext: string): string {
  const iv = randomBytes(IV_BYTES);
  const cipher = createCipheriv(ALGORITHM, key, iv);
  const ciphertext = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  return Buffer.concat([iv, cipher.getAuthTag(), ciphertext]).toString('base64');
}

function decrypt(key: Buffer, encoded: string): string {
  const blob = Buffer.from(encoded.trim(), 'base64');
  if (blob.length <= IV_BYTES + TAG_BYTES) {
    throw new Error('cache blob is too short to contain an IV and auth tag');
  }
  const iv = blob.subarray(0, IV_BYTES);
  const tag = blob.subarray(IV_BYTES, IV_BYTES + TAG_BYTES);
  const ciphertext = blob.subarray(IV_BYTES + TAG_BYTES);
  const decipher = createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]).toString('utf8');
}

/** Removes an unreadable cache so the next sign-in starts from a clean slate. */
async function discardCache(cachePath: string, reason: string): Promise<void> {
  warn(`token cache at ${cachePath} is unreadable (${reason}); starting empty`);
  try {
    await unlink(cachePath);
  } catch (err) {
    if (errorCode(err) !== 'ENOENT') {
      warn(`could not delete the unreadable token cache: ${describeError(err)}`);
    }
  }
}

/**
 * Builds the MSAL cache plugin backed by `${cacheDir}/token-cache.enc`.
 *
 * A cache that cannot be read is never fatal: the plugin warns, deletes the bad
 * file and lets MSAL continue with an empty cache, which downgrades a corrupt
 * file to "please sign in again" instead of a server that will not start.
 */
export function createCachePlugin(cacheDir: string): ICachePlugin {
  const cachePath = join(cacheDir, CACHE_FILE);

  return {
    async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
      await runExclusive(cacheDir, async () => {
        let encoded: string;
        try {
          encoded = await readFile(cachePath, 'utf8');
        } catch (err) {
          if (errorCode(err) !== 'ENOENT') {
            warn(`could not read the token cache: ${describeError(err)}`);
          }
          return;
        }
        if (encoded.trim().length === 0) return;

        let plaintext: string;
        try {
          const key = await loadOrCreateKey(cacheDir);
          plaintext = decrypt(key, encoded);
        } catch (err) {
          await discardCache(cachePath, describeError(err));
          return;
        }

        try {
          ctx.tokenCache.deserialize(plaintext);
        } catch (err) {
          await discardCache(cachePath, `not a valid MSAL cache: ${describeError(err)}`);
        }
      });
    },

    async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
      if (!ctx.cacheHasChanged) return;
      const plaintext = ctx.tokenCache.serialize();
      await runExclusive(cacheDir, async () => {
        try {
          await ensureDir(cacheDir);
          const key = await loadOrCreateKey(cacheDir);
          await writeSecret(cachePath, encrypt(key, plaintext));
        } catch (err) {
          // A cache we cannot persist costs the user a re-login next start; it
          // must not fail the token request that is in flight.
          warn(`could not persist the token cache: ${describeError(err)}`);
        }
      });
    },
  };
}

/** Deletes the encrypted cache and its key. Safe to call when neither exists. */
export async function clearCache(cacheDir: string): Promise<void> {
  await runExclusive(cacheDir, async () => {
    await rm(join(cacheDir, CACHE_FILE), { force: true });
    await rm(join(cacheDir, KEY_FILE), { force: true });
  });
}

/** Absolute paths this module owns, for status output and diagnostics. */
export function cachePaths(cacheDir: string): { cacheFile: string; keyFile: string } {
  return { cacheFile: join(cacheDir, CACHE_FILE), keyFile: join(cacheDir, KEY_FILE) };
}
