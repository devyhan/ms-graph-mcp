import test from 'node:test';
import assert from 'node:assert/strict';

import { guessScopesForPath } from './scopes.js';

test('a chat path guesses Chat.Read, not the alternatives', () => {
  assert.deepEqual(guessScopesForPath('/me/chats'), ['Chat.Read']);
});

test('a mail path guesses Mail.Read even under /me', () => {
  // Order matters: /me/messages must match the mail entry, not the /me
  // fallback at the end of the table.
  assert.deepEqual(guessScopesForPath('/me/messages'), ['Mail.Read']);
});

test('a bare /me path guesses User.Read', () => {
  assert.deepEqual(guessScopesForPath('/me'), ['User.Read']);
});

test('an unknown path guesses nothing, so the caller falls back to the union', () => {
  assert.equal(guessScopesForPath('/some/future/resource'), undefined);
});

test('each guess is a copy, so callers cannot mutate the shared table', () => {
  const first = guessScopesForPath('/me/chats');
  assert.ok(first !== undefined);
  first.push('Mail.Read');
  assert.deepEqual(guessScopesForPath('/me/chats'), ['Chat.Read']);
});
