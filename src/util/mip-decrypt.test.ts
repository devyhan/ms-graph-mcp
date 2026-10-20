import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import {
  MIP_DECRYPTOR_ENV,
  MipDecryptError,
  createMipDecryptor,
  decryptWithSidecar,
  resolveMipDecryptorPath,
} from './mip-decrypt.js';

/**
 * Launcher tests behind a fake sidecar: the checked-in fixture speaks the
 * real argv/exit-code protocol with synthetic behaviour per FAKE_MIP_MODE, so
 * no MIP SDK, tenant or network is involved.
 */

const FIXTURE = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
  'src',
  'util',
  'mip-decrypt-fixtures',
  'fake-sidecar.mjs',
);

/** Stages the fixture as an executable under a fresh temp dir. */
async function stageSidecar(name = 'fake-sidecar.mjs'): Promise<{ dir: string; bin: string }> {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'fake-mip-'));
  const bin = path.join(dir, name);
  await fs.copyFile(FIXTURE, bin);
  await fs.chmod(bin, 0o755);
  return { dir, bin };
}

async function withMode<T>(mode: string, fn: () => Promise<T>): Promise<T> {
  const prev = process.env['FAKE_MIP_MODE'];
  process.env['FAKE_MIP_MODE'] = mode;
  try {
    return await fn();
  } finally {
    if (prev === undefined) delete process.env['FAKE_MIP_MODE'];
    else process.env['FAKE_MIP_MODE'] = prev;
  }
}

async function leftovers(): Promise<string[]> {
  return (await fs.readdir(os.tmpdir())).filter((name) => name.startsWith('mip-decrypt-'));
}

async function assertRejectsKind(
  fn: () => Promise<unknown>,
  kind: string,
  message: RegExp,
): Promise<void> {
  await assert.rejects(fn, (error: unknown) => {
    assert.ok(error instanceof MipDecryptError, `want MipDecryptError, got ${String(error)}`);
    assert.equal(error.kind, kind);
    assert.match(error.message, message);
    return true;
  });
}

test('resolution order is explicit > env > PATH, and empty when nothing resolves', async () => {
  const { dir, bin } = await stageSidecar('mip-decrypt');
  try {
    assert.equal(resolveMipDecryptorPath({ explicitPath: '/x', env: { [MIP_DECRYPTOR_ENV]: '/y' } }), '/x');
    assert.equal(resolveMipDecryptorPath({ env: { [MIP_DECRYPTOR_ENV]: '/y' } }), '/y');
    assert.equal(resolveMipDecryptorPath({ env: { PATH: dir } }), bin);
    assert.equal(resolveMipDecryptorPath({ env: {} }), undefined);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('createMipDecryptor returns undefined without a sidecar, a hook with one', async () => {
  assert.equal(createMipDecryptor({ env: {} }), undefined);
  const hook = createMipDecryptor({ explicitPath: '/nonexistent/mip-decrypt' });
  assert.equal(typeof hook, 'function');
});

test('a successful run returns the sidecar output bytes', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    const out = await withMode('success', () =>
      decryptWithSidecar(bin, new Uint8Array([1, 2, 3]), 'r.pdf'),
    );
    assert.deepEqual(Array.from(out.slice(0, 6)), Array.from(Buffer.from('PLAIN:', 'utf8')));
    assert.deepEqual(Array.from(out.slice(6)), [1, 2, 3]);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('the input temp file is created 0600', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    await withMode('check-mode', () => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf'));
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('exit 1 maps to a failure carrying the sidecar diagnosis', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    await withMode('fail', () =>
      assertRejectsKind(() => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf'), 'failed', /synthetic decrypt failure/),
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('exit 2 maps to a usage error', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    await withMode('usage', () =>
      assertRejectsKind(() => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf'), 'usage', /usage error/),
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('a rights refusal maps to no-rights guidance, not a generic failure', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    await withMode('no-rights', () =>
      assertRejectsKind(
        () => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf'),
        'no-rights',
        /no usage rights/,
      ),
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('a missing binary maps to a no-runtime error naming the env var', async () => {
  await assertRejectsKind(
    () => decryptWithSidecar('/nonexistent/mip-decrypt', new Uint8Array([9]), 'r.pdf'),
    'no-runtime',
    new RegExp(MIP_DECRYPTOR_ENV),
  );
});

test('a hung sidecar is killed at the timeout', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    await withMode('hang', () =>
      assertRejectsKind(() => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf', 300), 'timeout', /still decrypting/),
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('empty sidecar output is treated as failed', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    await withMode('empty', () =>
      assertRejectsKind(() => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf'), 'failed', /empty file/),
    );
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('temp dirs are removed on success and on failure', async () => {
  const { dir, bin } = await stageSidecar();
  try {
    assert.deepEqual(await leftovers(), []);
    await withMode('success', () => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf'));
    assert.deepEqual(await leftovers(), []);
    await withMode('fail', () =>
      assert.rejects(() => decryptWithSidecar(bin, new Uint8Array([9]), 'r.pdf')),
    );
    assert.deepEqual(await leftovers(), []);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
});
