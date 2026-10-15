import assert from 'node:assert/strict';
import test from 'node:test';

import { decodeToon, encodeToon, ToonError } from './toon.js';

/**
 * The failure these guard against: a compact encoding that saves characters
 * by dropping or blurring data — an unquoted comma splitting one subject
 * into two cells, a `"true"` string coming back as a boolean. Every shape
 * below pins either its exact spelling or a decode round-trip, usually both.
 */

test('uniform object arrays encode as one header plus bare rows', () => {
  const text = encodeToon({
    users: [
      { id: 1, name: 'Ada' },
      { id: 2, name: 'Bob' },
    ],
  });
  assert.equal(text, 'users[2]{id,name}:\n  1,Ada\n  2,Bob');
  assert.deepEqual(decodeToon(text), {
    users: [
      { id: 1, name: 'Ada' },
      { id: 2, name: 'Bob' },
    ],
  });
});

test('primitive arrays encode inline, including at the root', () => {
  assert.equal(encodeToon({ tags: ['a', 'b'] }), 'tags[2]: a,b');
  assert.equal(encodeToon([1, 2, 3]), '[3]: 1,2,3');
  assert.deepEqual(decodeToon('[3]: 1,2,3'), [1, 2, 3]);
  assert.deepEqual(decodeToon('tags[2]: a,b'), { tags: ['a', 'b'] });
});

test('mixed scalar and object arrays use list form', () => {
  const text = encodeToon({ mixed: [1, { two: 2 }] });
  assert.equal(text, 'mixed[2]:\n  - 1\n  -\n    two: 2');
  assert.deepEqual(decodeToon(text), { mixed: [1, { two: 2 }] });
});

test('nested objects indent one level per depth', () => {
  const value = { a: { b: { c: 1 }, d: 2 }, e: 3 };
  const text = encodeToon(value);
  assert.equal(text, 'a:\n  b:\n    c: 1\n  d: 2\ne: 3');
  assert.deepEqual(decodeToon(text), value);
});

test('non-uniform object arrays stay readable as dashed items', () => {
  const value = { rows: [{ a: 1 }, { a: 1, b: 2 }] };
  const text = encodeToon(value);
  assert.equal(text, 'rows[2]:\n  -\n    a: 1\n  -\n    a: 1\n    b: 2');
  assert.deepEqual(decodeToon(text), value);
});

test('nested arrays ride on the dash line', () => {
  const value = { matrix: [[1, 2], [3]] };
  const text = encodeToon(value);
  assert.equal(text, 'matrix[2]:\n  - [2]: 1,2\n  - [1]: 3');
  assert.deepEqual(decodeToon(text), value);
});

test('empty arrays and objects have explicit forms', () => {
  assert.equal(encodeToon({ a: [], o: {} }), 'a: []\no:');
  assert.deepEqual(decodeToon('a: []\no:'), { a: [], o: {} });
  assert.equal(encodeToon([]), '[]');
  assert.deepEqual(decodeToon('[]'), []);
  // The empty document is the empty object, per the TOON spec.
  assert.equal(encodeToon({}), '');
  assert.deepEqual(decodeToon(''), {});
});

test('strings that would misread are quoted', () => {
  const cases: Array<[string, string]> = [
    ['', '""'],
    ['true', '"true"'],
    ['false', '"false"'],
    ['null', '"null"'],
    ['123', '"123"'],
    ['-4.5', '"-4.5"'],
    ['1e3', '"1e3"'],
    [' leading', '" leading"'],
    ['trailing ', '"trailing "'],
    ['a,b', '"a,b"'],
    ['a|b', '"a|b"'],
    ['a: b', '"a: b"'],
    ['a"b', '"a\\"b"'],
    ['a\\b', '"a\\\\b"'],
    ['a\nb', '"a\\nb"'],
    ['a\tb', '"a\\tb"'],
    ['-dash', '"-dash"'],
    ['#hash', '"#hash"'],
    ['a[0]', '"a[0]"'],
    ['a{b}', '"a{b}"'],
    // And the ones that stay bare: `007` is not a JSON number (leading
    // zero), so no decoder can mistake it for one.
    ['007', '007'],
    ['Ada', 'Ada'],
    ['Conversation number 1', 'Conversation number 1'],
    ['a-b_c.d', 'a-b_c.d'],
    ['100%', '100%'],
  ];
  for (const [raw, quoted] of cases) {
    assert.equal(encodeToon({ v: raw }), `v: ${quoted}`, JSON.stringify(raw));
    assert.deepEqual(decodeToon(`v: ${quoted}`), { v: raw }, quoted);
  }
});

test('keys outside the bare pattern are quoted', () => {
  const value = { '@odata.nextLink': 'https://x', 'plain.key_1': 1, 'with space': 2 };
  const text = encodeToon(value);
  assert.equal(text, '"@odata.nextLink": "https://x"\nplain.key_1: 1\n"with space": 2');
  assert.deepEqual(decodeToon(text), value);
});

test('every encoded document decodes back to the same value', () => {
  const values: unknown[] = [
    null,
    true,
    false,
    0,
    -42,
    3.14,
    1e21,
    'plain',
    '  padded  ',
    'snowman ☃ tab\there',
    { a: 1, b: 'x', c: true, d: null, e: [1, 'two', false, null] },
    { nested: { deeper: [{ x: [1, { y: 'z' }] }] } },
    { count: 2, items: [{ id: 'a', n: 1 }, { id: 'b', n: 2 }], cursor: 'abc,def: "quoted"' },
    [
      [1, 2],
      { mixed: [null, 's'] },
    ],
    { '': 'empty key', '0': 'numeric key' },
  ];
  for (const value of values) {
    assert.deepEqual(decodeToon(encodeToon(value)), value, encodeToon(value));
  }
});

test('decode rejects corrupt headers, counts and escapes', () => {
  // Header says two, carries one.
  assert.throws(() => decodeToon('a[2]: 1'), ToonError);
  // Header says two rows, carries one.
  assert.throws(() => decodeToon('a[2]{x}:\n  1'), ToonError);
  // A row with the wrong cell count.
  assert.throws(() => decodeToon('a[1]{x,y}:\n  1'), ToonError);
  // Fields plus inline content.
  assert.throws(() => decodeToon('a[1]{x}: 1'), ToonError);
  // Unterminated strings and bad escapes.
  assert.throws(() => decodeToon('a: "open'), ToonError);
  assert.throws(() => decodeToon('a: "bad\\q"'), ToonError);
  assert.throws(() => decodeToon('a: "lone\\ud800"'), ToonError);
  // A dash without its header, and stray text after a document.
  assert.throws(() => decodeToon('- 1'), ToonError);
  assert.throws(() => decodeToon('a: 1\nb: 2\n  c: 3\norphan'), ToonError);
  // A truncated character-cut starts with the banner, and the banner carries
  // a colon: without the explicit refusal below it would decode as a field
  // and return garbage as data instead of failing.
  assert.throws(
    () => decodeToon('--- OUTPUT TRUNCATED: NOT VALID TOON ---\nblob: xxx'),
    ToonError,
  );
});

test('full-line comments and blank lines are ignored on decode', () => {
  assert.deepEqual(decodeToon('a: 1\n# a comment\n\n  # indented too\nb: 2'), { a: 1, b: 2 });
});

test('hostile values serialise instead of throwing', () => {
  assert.equal(encodeToon({ big: 123n }), 'big: "123"');
  assert.equal(encodeToon({ err: new Error('boom') }), 'err:\n  name: Error\n  message: boom');
  const node: Record<string, unknown> = { id: 'a' };
  node['self'] = node;
  assert.deepEqual(decodeToon(encodeToon({ node })), { node: { id: 'a', self: '[Circular]' } });
  // Shared references are data, not cycles: both paths keep the value.
  const shared = [1, 2];
  assert.deepEqual(decodeToon(encodeToon({ a: shared, b: shared })), { a: [1, 2], b: [1, 2] });
  assert.equal(encodeToon(undefined), 'undefined');
});
