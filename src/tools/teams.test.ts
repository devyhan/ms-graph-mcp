import assert from 'node:assert/strict';
import test from 'node:test';

import { scopeReadingOnlyAuth } from '../auth/index.js';
import type {
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
  ServerConfig,
  ToolDefinition,
} from '../contracts.js';
import { teamsModule } from './teams.js';

/** Handler-level tests for the channel files-folder bridge. Suite stays offline: Graph is stubbed. */

const config: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['teams'],
  readOnly: true,
  graphVersion: 'v1.0',
  allowBeta: false,
  discovery: false,
  orgMode: true,
  maxOutputChars: 60_000,
  output: 'json',
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'auto',
  authPort: 0,
  allowGenericWrite: false,
};

function setup(data: unknown): { requests: GraphRequestOptions[]; tool: ToolDefinition } {
  const requests: GraphRequestOptions[] = [];
  const graph: GraphClient = {
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      requests.push(opts);
      return { status: 200, data: data as T };
    },
    async batch() {
      throw new Error('unused');
    },
    async follow() {
      throw new Error('unused');
    },
  };
  const defs = teamsModule.build({ graph, config, auth: scopeReadingOnlyAuth(), consentScopes: () => [] });
  const tool = defs.find((entry) => entry.name === 'teams_get_channel_files_folder');
  assert.ok(tool !== undefined, 'teams_get_channel_files_folder should exist');
  return { requests, tool };
}

test('the files folder comes back as a drive address', async () => {
  const { requests, tool } = setup({
    id: 'FOLDER',
    name: 'General',
    webUrl: 'https://tenant.sharepoint.com/General',
    folder: { childCount: 7 },
    parentReference: { driveId: 'DRIVE' },
  });
  const result = (await tool.handler({ teamId: 'TEAM', channelId: 'CHAN' })) as Record<string, unknown>;

  assert.equal(requests.length, 1);
  assert.equal(requests[0]?.path, '/teams/TEAM/channels/CHAN/filesFolder');
  assert.deepEqual(requests[0]?.scopes, ['Channel.ReadBasic.All']);
  assert.deepEqual(result, {
    teamId: 'TEAM',
    channelId: 'CHAN',
    driveId: 'DRIVE',
    itemId: 'FOLDER',
    name: 'General',
    childCount: 7,
    webUrl: 'https://tenant.sharepoint.com/General',
  });
});

test('a folder without a drive address is a missing-tab error', async () => {
  const { tool } = setup({ id: null, parentReference: null });
  await assert.rejects(() => tool.handler({ teamId: 'TEAM', channelId: 'CHAN' }), /no provisioned Files tab/);
});
