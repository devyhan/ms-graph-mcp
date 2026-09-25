import assert from 'node:assert/strict';
import test from 'node:test';

import { serializeResult } from './truncate.js';

/**
 * The failure these guard against: a chat list went over the output limit, the
 * serialiser cut it mid-string, and the caller got a JSON fragment that threw
 * "Bad control character at position 60000". The diagnostic pointed at the data
 * rather than at the size, which is the wrong place to look.
 */

const chat = (i: number) => ({
  id: `19:${'a'.repeat(40)}${i}@thread.v2`,
  topic: `Conversation number ${i}`,
  chatType: 'group',
  lastUpdatedDateTime: '2026-09-23T01:27:06.603Z',
});

test('a result under the limit is returned whole and untouched', () => {
  const value = { count: 2, chats: [chat(1), chat(2)] };
  const out = serializeResult(value, 100_000);
  assert.equal(out.truncated, false);
  assert.deepEqual(JSON.parse(out.text), value);
});

test('an oversized collection stays valid JSON, with fewer rows', () => {
  const value = { count: 400, chats: Array.from({ length: 400 }, (_, i) => chat(i)) };
  const out = serializeResult(value, 6_000);

  assert.equal(out.truncated, true);
  // The whole point: still parseable.
  const parsed = JSON.parse(out.text);
  assert.ok(out.text.length <= 6_000, `expected <= 6000 chars, got ${out.text.length}`);
  assert.ok(parsed.chats.length > 0, 'should keep as many rows as fit');
  assert.ok(parsed.chats.length < 400, 'should have dropped some');
  assert.deepEqual(parsed.chats[0], chat(0), 'kept rows are intact, not clipped');
});

test('the truncation note says what was dropped and how much', () => {
  const value = { chats: Array.from({ length: 400 }, (_, i) => chat(i)) };
  const parsed = JSON.parse(serializeResult(value, 6_000).text);

  assert.equal(parsed.truncated.field, 'chats');
  assert.equal(parsed.truncated.total, 400);
  assert.equal(parsed.truncated.returned, parsed.chats.length);
  assert.match(parsed.truncated.advice, /fewer fields|smaller page/);
});

test('sibling fields survive the trim', () => {
  const value = {
    count: 400,
    nextLink: 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=X',
    chats: Array.from({ length: 400 }, (_, i) => chat(i)),
  };
  const parsed = JSON.parse(serializeResult(value, 6_000).text);
  // The cursor is how the caller recovers the rest; losing it strands them.
  assert.equal(parsed.nextLink, value.nextLink);
  assert.equal(parsed.count, 400);
});

test('the longest array is the one shortened', () => {
  const value = {
    warnings: ['just one'],
    chats: Array.from({ length: 300 }, (_, i) => chat(i)),
  };
  const parsed = JSON.parse(serializeResult(value, 6_000).text);
  assert.equal(parsed.truncated.field, 'chats');
  assert.deepEqual(parsed.warnings, ['just one']);
});

test('one oversized item falls back to a character cut, warning first', () => {
  const value = { blob: 'x'.repeat(50_000) };
  const out = serializeResult(value, 1_000);

  assert.equal(out.truncated, true);
  assert.throws(() => JSON.parse(out.text), 'this case genuinely cannot stay JSON');
  // So the warning has to arrive before the payload, not after it.
  assert.match(out.text.slice(0, 80), /NOT VALID JSON/);
});

test('a bare array with no wrapper object also falls back', () => {
  const out = serializeResult(Array.from({ length: 400 }, (_, i) => chat(i)), 2_000);
  assert.equal(out.truncated, true);
  assert.match(out.text.slice(0, 80), /NOT VALID JSON/);
});

test('originalChars reports the full size, not the returned size', () => {
  const value = { chats: Array.from({ length: 400 }, (_, i) => chat(i)) };
  const out = serializeResult(value, 6_000);
  assert.ok(out.originalChars > 6_000);
  assert.ok(out.text.length <= 6_000);
});
