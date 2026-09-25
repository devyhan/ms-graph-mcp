import test from 'node:test';
import assert from 'node:assert/strict';

import { assertSameGraphHost, normalizeGraphPath, resolveGraphUrl } from './paths.js';

const GRAPH = 'https://graph.microsoft.com';

test('normalizeGraphPath rejects absolute and protocol-relative URLs', () => {
  assert.throws(() => normalizeGraphPath('https://evil.com/x'), /relative/i);
  assert.throws(() => normalizeGraphPath('//evil.com/x'), /protocol-relative/i);
});

test('normalizeGraphPath rejects traversal, including percent-encoded separators', () => {
  assert.throws(() => normalizeGraphPath('/users/..%2f..'), /percent-encoded/i);
  assert.throws(() => normalizeGraphPath('/users/../../me'), /".."/);
});

test('normalizeGraphPath rejects backslashes and userinfo tricks', () => {
  assert.throws(() => normalizeGraphPath('/a\\b'), /backslash/i);
  assert.throws(() => normalizeGraphPath('/x@evil.com'), /@/);
});

test('normalizeGraphPath rejects malformed input', () => {
  assert.throws(() => normalizeGraphPath('users'), /start with "\/"/);
  assert.throws(() => normalizeGraphPath(''), /non-empty/i);
  assert.throws(() => normalizeGraphPath('/me/messages?$top=1 OR 2'), /whitespace/i);
});

test('normalizeGraphPath accepts ordinary Graph paths', () => {
  assert.equal(normalizeGraphPath('/me/messages'), '/me/messages');
  assert.equal(normalizeGraphPath('/users/abc-123'), '/users/abc-123');
  // A UPN is legal past the first segment, and OneDrive item paths use colons.
  assert.equal(normalizeGraphPath('/users/user@contoso.com/messages'), '/users/user@contoso.com/messages');
  assert.equal(normalizeGraphPath('/me/drive/root:/report.docx:/content'), '/me/drive/root:/report.docx:/content');
  assert.equal(normalizeGraphPath('/me//messages/'), '/me/messages');
});

test('assertSameGraphHost refuses a nextLink pointing at another host', () => {
  const hostile = new URL('https://evil.example.com/v1.0/me/messages?$skiptoken=abc');
  assert.throws(() => assertSameGraphHost(hostile, GRAPH), /evil\.example\.com/);

  // A lookalike suffix must not pass either.
  assert.throws(() => assertSameGraphHost(new URL('https://graph.microsoft.com.evil.net/v1.0/me'), GRAPH), /expected/i);
  assert.throws(() => assertSameGraphHost(new URL('http://graph.microsoft.com/v1.0/me'), GRAPH), /https/i);
  assert.throws(() => assertSameGraphHost(new URL('https://user:pw@graph.microsoft.com/v1.0/me'), GRAPH), /credentials/i);

  const genuine = new URL('https://GRAPH.microsoft.com/v1.0/me/messages?$skiptoken=abc');
  assert.doesNotThrow(() => assertSameGraphHost(genuine, GRAPH));
  assert.doesNotThrow(() => assertSameGraphHost(genuine, 'graph.microsoft.com'));
});

test('resolveGraphUrl builds a pinned URL and rejects a hostile path', () => {
  const url = resolveGraphUrl(GRAPH, 'v1.0', '/me/messages');
  assert.equal(url.toString(), 'https://graph.microsoft.com/v1.0/me/messages');
  assert.equal(resolveGraphUrl(GRAPH, 'beta', '/users/abc-123').pathname, '/beta/users/abc-123');

  assert.throws(() => resolveGraphUrl(GRAPH, 'v1.0', 'https://evil.com/x'));
  assert.throws(() => resolveGraphUrl('http://graph.microsoft.com', 'v1.0', '/me'), /https/i);
  assert.throws(() => resolveGraphUrl(GRAPH, '../beta', '/me'), /version/i);
});
