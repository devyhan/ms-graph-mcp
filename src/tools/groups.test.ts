import assert from 'node:assert/strict';
import test from 'node:test';

import { MINIMAL_LOGIN_SCOPES, loginScopesFor } from './groups.js';

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
