import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';

import type { IdentityRecord } from '../contracts.js';
import {
  cachePaths,
  clearCache,
  clearIdentity,
  decrypt,
  encrypt,
  formatIdentityLine,
  loadOrCreateKey,
  readIdentity,
  writeIdentity,
} from './token-cache.js';

async function tempDir(): Promise<string> {
  return mkdtemp(join(tmpdir(), 'token-cache-'));
}

function sampleRecord(): IdentityRecord {
  return {
    homeAccountId: 'uid.tid',
    username: 'irene@contoso.com',
    name: 'Irene',
    tenantId: 'tid',
    clientId: 'client-id',
    authority: 'https://login.microsoftonline.com/tid',
    capturedAt: '2026-10-19T00:00:00.000Z',
    source: 'msal-account',
  };
}

test('identity round-trips through the encrypted sibling file', async () => {
  const dir = await tempDir();
  assert.equal(await readIdentity(dir), null);

  const record = sampleRecord();
  await writeIdentity(dir, record);
  assert.deepEqual(await readIdentity(dir), record);

  // The payload on disk is ciphertext: the username must not appear in clear.
  const raw = await readFile(cachePaths(dir).identityFile, 'utf8');
  assert.ok(!raw.includes(record.username));
});

test('identity and token cache share one key file in the same directory', async () => {
  const dir = await tempDir();
  await writeIdentity(dir, sampleRecord());
  const paths = cachePaths(dir);
  assert.equal(join(dir, 'identity-cache.enc'), paths.identityFile);
  // One key file only: no second key beside the identity file.
  const key = await loadOrCreateKey(dir);
  assert.equal(key.length, 32);
  assert.equal(decrypt(key, await readFile(paths.identityFile, 'utf8')), JSON.stringify(sampleRecord()));
});

test('key regeneration kills the identity snapshot', async () => {
  const dir = await tempDir();
  await writeIdentity(dir, sampleRecord());
  assert.ok(await readIdentity(dir));

  await rm(cachePaths(dir).keyFile);
  // A fresh key cannot open the old blob: reads fall back to absent.
  assert.equal(await readIdentity(dir), null);
});

test('clearCache removes the identity snapshot with the token cache', async () => {
  const dir = await tempDir();
  await writeIdentity(dir, sampleRecord());
  await clearCache(dir);
  assert.equal(await readIdentity(dir), null);
  await assert.rejects(readFile(cachePaths(dir).identityFile, 'utf8'), /ENOENT/);
  await assert.rejects(readFile(cachePaths(dir).keyFile, 'utf8'), /ENOENT/);
});

test('clearIdentity removes only the identity snapshot', async () => {
  const dir = await tempDir();
  await writeIdentity(dir, sampleRecord());
  await clearIdentity(dir);
  assert.equal(await readIdentity(dir), null);
  // The shared key survives, so a later write still works.
  await writeIdentity(dir, sampleRecord());
  assert.deepEqual(await readIdentity(dir), sampleRecord());
});

test('a corrupt identity file reads as absent instead of throwing', async () => {
  const dir = await tempDir();
  const { identityFile } = cachePaths(dir);
  await writeIdentity(dir, sampleRecord());

  await writeFile(identityFile, '!!! not base64 !!!', 'utf8');
  assert.equal(await readIdentity(dir), null);

  await writeFile(identityFile, '', 'utf8');
  assert.equal(await readIdentity(dir), null);
});

test('a well-formed blob with a bad payload reads as absent', async () => {
  const dir = await tempDir();
  const { identityFile } = cachePaths(dir);
  const key = await loadOrCreateKey(dir);

  await writeFile(identityFile, encrypt(key, 'valid ciphertext, not JSON'), 'utf8');
  assert.equal(await readIdentity(dir), null);

  await writeFile(identityFile, encrypt(key, JSON.stringify({ username: 'no-shape' })), 'utf8');
  assert.equal(await readIdentity(dir), null);
});

test('concurrent identity writes serialize without interleaving', async () => {
  const dir = await tempDir();
  const first = { ...sampleRecord(), username: 'first@contoso.com' };
  const second = { ...sampleRecord(), username: 'second@contoso.com', source: 'graph-me' as const };
  await Promise.all([writeIdentity(dir, first), writeIdentity(dir, second)]);
  const read = await readIdentity(dir);
  assert.ok(read !== null);
  assert.ok(read.username === first.username || read.username === second.username);
  assert.deepEqual(read, read.username === first.username ? first : second);
});

test('the identity file is created with owner-only permissions', async (t) => {
  // Windows has no POSIX permission bits: Node synthesizes 0666 and the 0600
  // passed to writeFile is correctly a no-op there, so there is nothing to
  // assert. The product behavior is unchanged on every platform.
  if (process.platform === 'win32') {
    t.skip('POSIX file modes do not exist on Windows');
    return;
  }
  const dir = await tempDir();
  await writeIdentity(dir, sampleRecord());
  const mode = (await stat(cachePaths(dir).identityFile)).mode & 0o777;
  assert.equal(mode, 0o600);
});

test('the identity line names the session without unactionable ids', () => {
  const line = formatIdentityLine(sampleRecord());
  assert.equal(
    line,
    'Signed in as irene@contoso.com (Irene) in tenant tid; ' +
      'identity cached at 2026-10-19T00:00:00.000Z (msal-account).',
  );
  assert.ok(!line.includes('uid.tid'), 'no home account id');
  assert.ok(!line.includes('client-id'), 'no client id');
  assert.ok(!line.includes('login.microsoftonline.com'), 'no authority');
});

test('the identity line degrades when name and tenant are absent', () => {
  const { name: _n, tenantId: _t, ...bare } = sampleRecord();
  assert.equal(
    formatIdentityLine(bare),
    'Signed in as irene@contoso.com; identity cached at 2026-10-19T00:00:00.000Z (msal-account).',
  );
  // A name that repeats the username adds nothing.
  assert.equal(
    formatIdentityLine({ ...bare, name: bare.username }),
    'Signed in as irene@contoso.com; identity cached at 2026-10-19T00:00:00.000Z (msal-account).',
  );
});
