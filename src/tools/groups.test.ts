import assert from 'node:assert/strict';
import test from 'node:test';
import { z } from 'zod';

import type { ToolDefinition } from '../contracts.js';
import {
  MINIMAL_LOGIN_SCOPES,
  consentScopesFor,
  loginScopesFor,
  resolveGroups,
  resolveGroupsDetailed,
  upfrontScopesFor,
} from './groups.js';

test('the minimal set is the account and the refresh token, nothing else', () => {
  assert.deepEqual([...MINIMAL_LOGIN_SCOPES].sort(), ['User.Read', 'offline_access']);
});

test('a bare sign-in asks for the minimal set alone', () => {
  assert.deepEqual(loginScopesFor(), ['User.Read', 'offline_access']);
  assert.deepEqual(loginScopesFor([]), ['User.Read', 'offline_access']);
});

test('named scopes join the minimal set, trimmed and deduped', () => {
  assert.deepEqual(
    loginScopesFor(['Mail.Read', ' Mail.Read ', '  ', 'Calendars.Read']),
    ['Calendars.Read', 'Mail.Read', 'User.Read', 'offline_access'],
  );
});

test('a scope outside the enabled groups consents the same way', () => {
  // No group check here by design: a `graph_request` scopes override can name
  // anything, and Entra — not this function — decides what is grantable.
  assert.ok(loginScopesFor(['ChannelMessage.Read.All']).includes('ChannelMessage.Read.All'));
});

function fakeTool(group: string, scopes: string[], write = false): ToolDefinition {
  return {
    name: `${group}_fake_tool`,
    title: 'Fake tool',
    description: 'A test tool carrying extra scopes.',
    inputSchema: z.object({}),
    write: write ? true : undefined,
    scopes,
    group,
    handler: async () => ({}),
  };
}

test('the upfront set keeps what an ordinary user can grant', () => {
  const tools = [
    fakeTool('mail', ['Mail.Read', 'Mail.Send']),
    fakeTool('files', ['Files.Read', 'Files.Read.All']),
    fakeTool('teams', ['ChannelMessage.Read.All']),
    fakeTool('me', ['MailboxSettings.Read']),
  ];
  const upfront = upfrontScopesFor(
    { groups: ['me', 'mail', 'files', 'teams'], readOnly: false },
    tools,
  );

  for (const kept of ['User.Read', 'offline_access', 'Mail.Send', 'Files.Read', 'Files.ReadWrite']) {
    assert.ok(upfront.includes(kept), `${kept} is grantable and must be kept`);
  }
  for (const dropped of [
    'Mail.Read',
    'Mail.ReadWrite',
    'MailboxSettings.Read',
    'Files.Read.All',
    'ChannelMessage.Read.All',
    'Team.ReadBasic.All',
  ]) {
    assert.ok(!upfront.includes(dropped), `${dropped} would fail the prompt and must be left out`);
  }
  assert.deepEqual([...upfront].sort(), upfront, 'sorted');
});

test('the upfront set is read-only aware', async () => {
  const tools = [fakeTool('mail', ['Mail.Send'], true)];
  const readOnly = upfrontScopesFor({ groups: ['mail'], readOnly: true }, tools);
  assert.ok(!readOnly.includes('Mail.Send'), 'write scopes are left out under --read-only');
  assert.ok(readOnly.includes('User.Read'));

  const readWrite = upfrontScopesFor({ groups: ['mail'], readOnly: false }, tools);
  assert.ok(readWrite.includes('Mail.Send'));
});

test('the upfront set ignores tools outside the enabled groups', () => {
  // A tool from a disabled group must not widen the ask: consenting to a scope
  // whose tools are switched off buys nothing.
  const tools = [fakeTool('chat', ['Chat.ReadWrite']), fakeTool('generic', ['Files.Read'])];
  const upfront = upfrontScopesFor({ groups: ['me'], readOnly: false }, tools);

  assert.ok(!upfront.includes('Chat.ReadWrite'), 'disabled-group extras are left out');
  assert.ok(upfront.includes('Files.Read'), 'generic tools are always in');
});

test('the detailed resolution reports admin-consent groups dropped without orgMode', () => {
  const { groups, dropped } = resolveGroupsDetailed({
    groups: ['mail', 'teams'],
    orgMode: false,
  });
  assert.deepEqual(groups, ['mail']);
  assert.deepEqual(dropped, ['teams']);
});

test('the detailed resolution drops nothing with orgMode', () => {
  const { groups, dropped } = resolveGroupsDetailed({
    groups: ['mail', 'teams'],
    orgMode: true,
  });
  assert.deepEqual(groups, ['mail', 'teams']);
  assert.deepEqual(dropped, []);
});

test('resolveGroups still returns only the kept groups', () => {
  assert.deepEqual(resolveGroups({ groups: ['teams'], orgMode: false }), []);
  assert.deepEqual(resolveGroups({ groups: ['teams'], orgMode: true }), ['teams']);
});

test('the upfront set never exceeds the full consent set', () => {
  const tools = [
    fakeTool('mail', ['Mail.Read', 'Mail.Send']),
    fakeTool('files', ['Files.Read', 'Files.Read.All']),
  ];
  const config = { groups: ['me', 'mail', 'files'], readOnly: false };
  const full = consentScopesFor(config, tools);
  for (const scope of upfrontScopesFor(config, tools)) {
    assert.ok(full.includes(scope), `${scope} must come from the full set`);
  }
});
