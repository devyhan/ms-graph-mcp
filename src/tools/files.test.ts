import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { scopeReadingOnlyAuth } from '../auth/index.js';
import { MIP_DECRYPTOR_ENV } from '../util/mip-decrypt.js';
import type {
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
  ServerConfig,
  ToolDefinition,
} from '../contracts.js';
import { filesModule, filesParseTestHooks } from './files.js';

/** Handler-level tests for the share/drive/parse tools. Graph and bytes are stubbed: no network. */

const config: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['files'],
  readOnly: true,
  graphVersion: 'v1.0',
  allowBeta: false,
  discovery: false,
  orgMode: false,
  maxOutputChars: 60_000,
  output: 'json',
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'auto',
  authPort: 0,
  allowGenericWrite: false,
};

const PDF_BYTES = new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31]);

function setup(
  routes: Record<string, unknown>,
  configOverrides: Partial<ServerConfig> = {},
): {
  requests: GraphRequestOptions[];
  tool: (name: string) => ToolDefinition;
} {
  const requests: GraphRequestOptions[] = [];
  const graph: GraphClient = {
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      requests.push(opts);
      const key = opts.path.startsWith('/shares/') ? '/shares/*' : opts.path;
      if (!(key in routes)) throw new Error(`stub has no route for ${opts.path}`);
      return { status: 200, data: routes[key] as T };
    },
    async batch() {
      throw new Error('unused');
    },
    async follow() {
      throw new Error('unused');
    },
  };
  const defs = filesModule.build({
    graph,
    config: { ...config, ...configOverrides },
    auth: scopeReadingOnlyAuth(),
    consentScopes: () => [],
  });
  return {
    requests,
    tool: (name: string) => {
      const found = defs.find((entry) => entry.name === name);
      assert.ok(found !== undefined, `${name} should exist`);
      return found;
    },
  };
}

test('files_resolve_share turns a link into a drive address', async () => {
  const { requests, tool } = setup({
    '/shares/*': {
      id: 'ITEM',
      name: 'q3.docx',
      size: 41,
      file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
      parentReference: { driveId: 'DRIVE' },
    },
  });
  const result = (await tool('files_resolve_share').handler({
    shareUrl: 'https://tenant.sharepoint.com/:w:/g/x',
  })) as Record<string, unknown>;

  assert.match(requests[0]?.path ?? '', /^\/shares\/u!/);
  assert.equal(result['driveId'], 'DRIVE');
  assert.equal(result['itemId'], 'ITEM');
  assert.equal(result['name'], 'q3.docx');
});

test('files_get_drive_item reads outside the own drive', async () => {
  const { requests, tool } = setup({
    '/drives/DRIVE/items/ITEM': {
      id: 'ITEM',
      name: 'shared.pdf',
      size: 7,
      file: { mimeType: 'application/pdf' },
      parentReference: { driveId: 'DRIVE', path: '/drive/root:/Docs' },
      lastModifiedDateTime: '2026-09-01T00:00:00Z',
    },
  });
  const result = (await tool('files_get_drive_item').handler({
    driveId: 'DRIVE',
    itemId: 'ITEM',
  })) as Record<string, unknown>;

  assert.equal(requests[0]?.path, '/drives/DRIVE/items/ITEM');
  assert.equal(result['name'], 'shared.pdf');
  assert.equal(result['driveId'], 'DRIVE');
  assert.equal(result['folderPath'], '/Docs');
});

test('files_parse_document parses a drive source end to end (stubbed bytes)', async () => {
  filesParseTestHooks.deps = {
    fetchBytes: async () => ({ bytes: PDF_BYTES, nextOffset: PDF_BYTES.length }),
    extract: async (kind) => ({
      units: [{ id: 'page-1', title: 'Page 1', text: `hello ${kind}` }],
      warnings: [],
    }),
  };
  try {
    const { tool } = setup({
      '/drives/DRIVE/items/ITEM': { id: 'ITEM', name: 'r.pdf', size: 6, file: { mimeType: 'application/pdf' } },
    });
    const result = (await tool('files_parse_document').handler({
      driveId: 'DRIVE',
      itemId: 'ITEM',
    })) as {
      fileName: string;
      units: Array<{ text: string }>;
      manifest: { reason: string; unitCount: number };
    };
    assert.equal(result.fileName, 'r.pdf');
    assert.equal(result.units[0]?.text, 'hello pdf');
    assert.equal(result.manifest.reason, 'complete');
    assert.equal(result.manifest.unitCount, 1);
  } finally {
    filesParseTestHooks.deps = undefined;
  }
});

test('files_parse_document decrypts a wrapper when a sidecar resolves (fake binary)', async () => {
  // Bridge: handler -> createMipDecryptor -> fake sidecar -> plaintext parse.
  // Wrapper bytes are synthetic (magic + /Encrypt + marker), never real ones.
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'fake-mip-'));
  const bin = path.join(dir, 'fake-sidecar.mjs');
  await fs.copyFile(
    path.resolve(
      path.dirname(fileURLToPath(import.meta.url)),
      '..',
      '..',
      'src',
      'util',
      'mip-decrypt-fixtures',
      'fake-sidecar.mjs',
    ),
    bin,
  );
  await fs.chmod(bin, 0o755);
  const prevSidecar = process.env[MIP_DECRYPTOR_ENV];
  const prevMode = process.env['FAKE_MIP_MODE'];
  process.env[MIP_DECRYPTOR_ENV] = bin;
  process.env['FAKE_MIP_MODE'] = 'success';
  filesParseTestHooks.deps = {
    fetchBytes: async () => ({
      bytes: new Uint8Array(
        Buffer.from('%PDF-1.7\n1 0 obj\n<< /Filter /Aadrm >>\ntrailer << /Encrypt 1 0 R >>', 'latin1'),
      ),
      nextOffset: 70,
    }),
    extract: async () => ({
      units: [{ id: 'page-1', title: 'Page 1', text: 'real content after decrypt' }],
      warnings: [],
    }),
  };
  try {
    const { tool } = setup({
      '/drives/DRIVE/items/ITEM': { id: 'ITEM', name: 'r.pdf', size: 70, file: { mimeType: 'application/pdf' } },
    });
    const result = (await tool('files_parse_document').handler({
      driveId: 'DRIVE',
      itemId: 'ITEM',
    })) as {
      units: Array<{ text: string }>;
      protection: string;
      decrypted: boolean;
    };
    assert.equal(result.units[0]?.text, 'real content after decrypt');
    assert.equal(result.protection, 'aip-rms');
    assert.equal(result.decrypted, true);
  } finally {
    filesParseTestHooks.deps = undefined;
    if (prevSidecar === undefined) delete process.env[MIP_DECRYPTOR_ENV];
    else process.env[MIP_DECRYPTOR_ENV] = prevSidecar;
    if (prevMode === undefined) delete process.env['FAKE_MIP_MODE'];
    else process.env['FAKE_MIP_MODE'] = prevMode;
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('files_parse_document prefers config.mipDecryptor over the environment', async () => {
  // Wiring proof: the flag value must reach the launcher as explicitPath. The
  // env var points at a missing binary, so only the config path can decrypt —
  // without the wiring this test fails with a no-runtime error instead.
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'cfg-mip-'));
  const bin = path.join(dir, 'fake-sidecar.mjs');
  await fs.copyFile(
    path.resolve(
      path.dirname(fileURLToPath(import.meta.url)),
      '..',
      '..',
      'src',
      'util',
      'mip-decrypt-fixtures',
      'fake-sidecar.mjs',
    ),
    bin,
  );
  await fs.chmod(bin, 0o755);
  const prevSidecar = process.env[MIP_DECRYPTOR_ENV];
  const prevMode = process.env['FAKE_MIP_MODE'];
  process.env[MIP_DECRYPTOR_ENV] = path.join(dir, 'does-not-exist');
  process.env['FAKE_MIP_MODE'] = 'success';
  filesParseTestHooks.deps = {
    fetchBytes: async () => ({
      bytes: new Uint8Array(
        Buffer.from('%PDF-1.7\n1 0 obj\n<< /Filter /Aadrm >>\ntrailer << /Encrypt 1 0 R >>', 'latin1'),
      ),
      nextOffset: 70,
    }),
    extract: async () => ({
      units: [{ id: 'page-1', title: 'Page 1', text: 'decrypted via config path' }],
      warnings: [],
    }),
  };
  try {
    const { tool } = setup(
      {
        '/drives/DRIVE/items/ITEM': { id: 'ITEM', name: 'r.pdf', size: 70, file: { mimeType: 'application/pdf' } },
      },
      { mipDecryptor: bin },
    );
    const result = (await tool('files_parse_document').handler({
      driveId: 'DRIVE',
      itemId: 'ITEM',
    })) as {
      units: Array<{ text: string }>;
      protection: string;
      decrypted: boolean;
    };
    assert.equal(result.units[0]?.text, 'decrypted via config path');
    assert.equal(result.protection, 'aip-rms');
    assert.equal(result.decrypted, true);
  } finally {
    filesParseTestHooks.deps = undefined;
    if (prevSidecar === undefined) delete process.env[MIP_DECRYPTOR_ENV];
    else process.env[MIP_DECRYPTOR_ENV] = prevSidecar;
    if (prevMode === undefined) delete process.env['FAKE_MIP_MODE'];
    else process.env['FAKE_MIP_MODE'] = prevMode;
    await fs.rm(dir, { recursive: true, force: true });
  }
});

test('files_parse_document rejects sourceless and multi-source calls', async () => {
  const { tool } = setup({});
  const parse = tool('files_parse_document');
  await assert.rejects(() => parse.handler({}), /one source/);
  await assert.rejects(
    () => parse.handler({ shareUrl: 'https://example.com/x', driveId: 'D', itemId: 'I' }),
    /exactly one source/,
  );
  await assert.rejects(() => parse.handler({ driveId: 'D' }), /belong together/);
  await assert.rejects(
    () => parse.handler({ cursor: 'abc', shareUrl: 'https://example.com/x' }),
    /either `cursor` or a source/,
  );
  assert.ok(parse.scopes.includes('Mail.Read'), 'the mail-attachment source needs Mail.Read');
  assert.ok(parse.scopes.includes('Files.Read.All'), 'shared drives need Files.Read.All');
});
