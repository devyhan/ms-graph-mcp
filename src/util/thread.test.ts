import assert from 'node:assert/strict';
import test from 'node:test';

import {
  TRACKING_LINK_MIN_CHARS,
  foldRepeatedBlocks,
  foldTrackingLinks,
  normalizeBodyText,
  shareBudget,
  splitQuotedBlocks,
} from './thread.js';

/**
 * These guard two failures that both reached a real mailbox before anyone
 * noticed.
 *
 * The first was a link rule keyed on length alone. At a 60-character threshold
 * it folded away three references a summary had cited — a design file, a
 * spreadsheet, an API doc — because they were merely long, not opaque. The
 * lengths below are the ones measured in that corpus, which is why they are
 * spelled out rather than rounded.
 *
 * The second was a fold cache that outlived one call. Re-reading a message it
 * had already seen returned nothing but pointers at itself, so a model checking
 * its own work concluded the message was empty. `foldRepeatedBlocks` owns its
 * cache for exactly one invocation, and "called twice, same bytes" is the test
 * that says so.
 */

const NBSP = ' ';
const ZERO_WIDTH = '​';

/** Lengths taken from the measured corpus; both populations, no overlap. */
const REAL_LINK = `https://docs.google.com/spreadsheets/d/${'a'.repeat(44)}/edit`;
const TRACKING_LINK = `https://click.figma.com/ls/click?upn=${'x'.repeat(700)}`;

function outlookReply(newText: string, quoted: string): string {
  return `${newText}\n\nFrom: Someone <someone@example.com>\nSent: Monday\nTo: Other\nSubject: Re: thing\n\n${quoted}`;
}

// ---------------------------------------------------------------------------
// normalisation
// ---------------------------------------------------------------------------

test('invisible characters are removed and counted', () => {
  const out = normalizeBodyText(`hi${ZERO_WIDTH}${ZERO_WIDTH}there`);
  assert.equal(out.text, 'hithere');
  assert.equal(out.invisibleRemoved, 2);
});

test('a non-breaking space becomes a plain space rather than vanishing', () => {
  const out = normalizeBodyText(`보낸${NBSP}사람: 홍길동`);
  assert.equal(out.text, '보낸 사람: 홍길동');
  assert.equal(out.invisibleRemoved, 0, 'NBSP is folded, not deleted');
});

test('a zero-width joiner is kept and a line separator becomes a line break', () => {
  // U+200D holds an emoji sequence together; deleting it turns one glyph into
  // three people. U+2028 is a line break, so removing it welds two lines.
  const family = '\u{1F468}‍\u{1F469}‍\u{1F467}';
  assert.ok(normalizeBodyText(`family ${family}`).text.includes(family));
  assert.equal(normalizeBodyText('a b').text, 'a\nb');
  assert.equal(normalizeBodyText(`a${ZERO_WIDTH}b`).text, 'ab', 'U+200B still goes');
});

test('a Korean quote header survives a non-breaking space', () => {
  // Outlook emits this often enough that a literal space in the separator would
  // silently stop splitting on a whole mailbox.
  const body = normalizeBodyText(`새 내용\n\n보낸${NBSP}사람: 홍길동\n제목: 회신\n\n이전 내용`).text;
  assert.equal(splitQuotedBlocks(body).length, 2);
});

// ---------------------------------------------------------------------------
// links
// ---------------------------------------------------------------------------

test('a long tracking URL folds to a marker that keeps the host', () => {
  const out = foldTrackingLinks(`see ${TRACKING_LINK} ok`);
  assert.equal(out.folded, 1);
  assert.ok(!out.text.includes(TRACKING_LINK), 'the opaque URL is gone');
  assert.ok(out.text.includes('click.figma.com'), 'but the reader still learns where it went');
  assert.ok(out.text.includes(String(TRACKING_LINK.length)));
});

test('ordinary links are returned verbatim however long they look', () => {
  // The regression: every one of these was destroyed by a 60-character rule.
  const links = [
    'https://golfzon-ocr.example.co.kr/api-docs/doc-public.html#overview',
    REAL_LINK,
    `https://www.figma.com/design/${'N'.repeat(22)}/%ED%95%84%EB%93%9C?node-id=12582-5560`,
  ];
  const out = foldTrackingLinks(links.join('\n'));
  assert.equal(out.folded, 0);
  for (const link of links) {
    assert.ok(out.text.includes(link), `${link.length}-char link must survive`);
    assert.ok(link.length < TRACKING_LINK_MIN_CHARS);
  }
});

test('a meeting-join link survives even though it is long', () => {
  // The near-miss that length alone gets wrong. Measured longest opaque run:
  // a Teams join link 125 characters, a SharePoint share 79, the tracking
  // redirector 646-810. Length alone would fold the first two.
  const teams =
    `https://teams.microsoft.com/l/meetup-join/19%3ameeting_${'N'.repeat(52)}` +
    `%40thread.v2/0?context=%7b%22Tid%22%3a%2272f988bf-86f1-41af-91ab-2d7cd011db47%22%2c%22Oid%22%3a%22${'a'.repeat(36)}%22%7d`;
  const sharepoint =
    `https://contoso-my.sharepoint.com/:x:/g/personal/ada_contoso_com/${'E'.repeat(40)}` +
    `?e=5Xy7Qk&nav=${'b'.repeat(60)}&xsdata=${'c'.repeat(72)}`;

  for (const link of [teams, sharepoint]) {
    assert.ok(link.length > TRACKING_LINK_MIN_CHARS, 'the point is that it IS long');
    const out = foldTrackingLinks(link);
    assert.equal(out.folded, 0, `a ${link.length}-char real link must survive`);
    assert.ok(out.text.includes(link));
  }
});

test('a marker cannot be made to name a host that is not the host', () => {
  // `https://teams.microsoft.com@evil.example/…` is a real URL whose host is
  // evil.example. Naming the userinfo would put a trusted name in front of a
  // reader who can no longer see the original.
  const out = foldTrackingLinks(`https://teams.microsoft.com@evil.example/l/${'z'.repeat(400)}`);
  assert.equal(out.folded, 1);
  assert.ok(out.text.includes('evil.example'));
  assert.ok(!/tracking-link: teams\.microsoft\.com,/.test(out.text));
});

test('the threshold is a parameter, so a caller can keep everything', () => {
  const out = foldTrackingLinks(TRACKING_LINK, Infinity);
  assert.equal(out.folded, 0);
  assert.equal(out.text, TRACKING_LINK);
});

test('folding twice is the same as folding once', () => {
  const once = foldTrackingLinks(`a ${TRACKING_LINK} b`).text;
  assert.equal(foldTrackingLinks(once).text, once, 'the marker is not itself a URL');
});

// ---------------------------------------------------------------------------
// splitting
// ---------------------------------------------------------------------------

test('the separator stays at the head of the block it introduces', () => {
  const blocks = splitQuotedBlocks(outlookReply('new words', 'old words'));
  assert.equal(blocks.length, 3, 'new text, the From header, then Subject and body');
  assert.ok((blocks[0] as string).startsWith('new words'));
  assert.ok((blocks[1] as string).trimStart().startsWith('From:'));
});

test('a body with no quote at all is one block', () => {
  assert.deepEqual(splitQuotedBlocks('just a note'), ['just a note']);
});

test('empty and whitespace-only input produce no blocks', () => {
  assert.deepEqual(splitQuotedBlocks(''), []);
  assert.deepEqual(splitQuotedBlocks('   \n\n  '), []);
});

// ---------------------------------------------------------------------------
// folding
// ---------------------------------------------------------------------------

const QUOTED = 'The agreed policy is to trust the printed digits and ignore the pencilled marks.';

test('text repeated from an earlier message folds to a pointer at it', () => {
  const result = foldRepeatedBlocks([
    { body: QUOTED },
    { body: outlookReply('Understood, thanks.', QUOTED) },
  ]);

  assert.equal(result.stats.folded, 1);
  assert.ok((result.bodies[0] as { body: string }).body.includes(QUOTED), 'the original is kept');
  assert.ok((result.bodies[1] as { body: string }).body.includes('[quoted from #1]'));
  assert.ok(!(result.bodies[1] as { body: string }).body.includes(QUOTED), 'the repeat is not');
});

test('the first message can never contain a pointer', () => {
  const result = foldRepeatedBlocks([{ body: QUOTED }, { body: QUOTED }]);
  assert.ok(!(result.bodies[0] as { body: string }).body.includes('[quoted from'));
});

test('every pointer names a message earlier than the one citing it', () => {
  const messages = [
    { body: `${QUOTED} one` },
    { body: outlookReply('second', `${QUOTED} one`) },
    { body: outlookReply('third', `${QUOTED} one\n\nFrom: x\nSubject: y\n\nsecond`) },
  ];
  foldRepeatedBlocks(messages).bodies.forEach((folded, index) => {
    for (const [, cited] of folded.body.matchAll(/#(\d+)/g)) {
      const target = Number(cited);
      assert.ok(target >= 1 && target < index + 1, `#${target} cited from message ${index + 1}`);
    }
  });
});

test('the same input folded twice returns the same bytes', () => {
  // The cache is a local, so a second call cannot see the first one's blocks.
  const messages = [{ body: QUOTED }, { body: outlookReply('ack', QUOTED) }];
  assert.deepEqual(foldRepeatedBlocks(messages), foldRepeatedBlocks(messages));
});

test('a message that is entirely a repeat is labelled rather than left blank', () => {
  // A mailbox holding both the Sent copy and the delivered copy of one mail.
  const result = foldRepeatedBlocks([{ body: QUOTED }, { body: QUOTED }]);
  assert.equal((result.bodies[1] as { duplicateOf?: number }).duplicateOf, 1);
});

test('adjacent pointers collapse into one range', () => {
  // A real three-deep chain: the third message carries the second's header block
  // and the first's body back to back, so two folds land next to each other.
  const header = (who: string): string =>
    `From: ${who} <${who}@example.com>\nSent: Monday 1 September 2026 09:00\nTo: Team <team@example.com>\nCc: Watchers <watchers@example.com>\nSubject: Re: the thing\n\n`;
  const first = `FIRST. ${QUOTED}`;
  const second = `SECOND. Noted, and I have passed it to the vendor for confirmation.`;

  const result = foldRepeatedBlocks([
    { body: first },
    { body: `${second}\n\n${header('ada')}${first}` },
    { body: `THIRD. Confirmed.\n\n${header('bob')}${second}\n\n${header('ada')}${first}` },
  ]);

  const third = (result.bodies[2] as { body: string }).body;
  assert.ok(/\[quoted from #1-#2\]/.test(third), `expected a collapsed range, got:\n${third}`);
});

test('two notices differing only by a short cited link do not fold together', () => {
  // The regression: the comparison key erased every URL, so two build notices
  // became one and the pointer sent the reader to the wrong log.
  const a =
    'Your nightly build finished. Review the full log before you merge anything.\n' +
    'Build log: https://ci.example.com/jobs/1001/console';
  const b = a.replace('1001', '2002');

  const result = foldRepeatedBlocks([{ body: a }, { body: b }]);
  assert.equal(result.stats.folded, 0);
  assert.ok((result.bodies[1] as { body: string }).body.includes('2002'));
});

test('a block repeated inside one message never points at that message', () => {
  // A pointer must resolve to an EARLIER message. A body that quotes the same
  // forwarded mail twice would otherwise emit [quoted from #1] while it IS #1.
  const twice =
    'Please action.\n\nFrom: a@example.com\nSubject: Re: W\n' +
    'Approve the transfer to the account named in the attached statement today.\n\n' +
    'From: a@example.com\nSubject: Re: W\n' +
    'Approve the transfer to the account named in the attached statement today.';

  const result = foldRepeatedBlocks([{ body: twice }]);
  const body = (result.bodies[0] as { body: string }).body;
  assert.ok(!body.includes('[quoted from'), `message #1 must cite nothing, got:\n${body}`);
  assert.equal((result.bodies[0] as { pointers: number }).pointers, 0);
});

test('indentation in authored text is left alone', () => {
  // Alignment is the only structure a pasted table, log or code block has left
  // by the time it reaches here.
  const table = 'col1    col2      col3\na       b         c';
  const result = foldRepeatedBlocks([{ body: table }]);
  assert.equal((result.bodies[0] as { body: string }).body, table);
});

test('short blocks are left alone however often they repeat', () => {
  // Folding "thanks" costs more than it saves and makes the result unreadable.
  const result = foldRepeatedBlocks([{ body: '감사합니다.' }, { body: '감사합니다.' }]);
  assert.equal(result.stats.folded, 0);
  assert.equal((result.bodies[1] as { body: string }).body, '감사합니다.');
});

test('quotedChars reports what folding actually removed', () => {
  const result = foldRepeatedBlocks([{ body: QUOTED }, { body: QUOTED }]);
  assert.ok((result.bodies[1] as { quotedChars: number }).quotedChars >= QUOTED.length);
  assert.equal(result.stats.quotedChars, (result.bodies[1] as { quotedChars: number }).quotedChars);
});

// ---------------------------------------------------------------------------
// budget
// ---------------------------------------------------------------------------

test('a short message is not cut to make room for a long one', () => {
  const grants = shareBudget([40_000, 300, 300, 300], 5_000);
  assert.deepEqual(grants.slice(1), [300, 300, 300], 'the short ones are served whole');
  assert.equal(grants[0], 4_100, 'the long one takes what is left');
  assert.equal(
    grants.reduce((sum, n) => sum + n, 0),
    5_000,
  );
});

test('an even split when everyone wants more than their share', () => {
  assert.deepEqual(shareBudget([10_000, 10_000], 1_000), [500, 500]);
});

test('a budget larger than the need grants the need, not the budget', () => {
  assert.deepEqual(shareBudget([100, 200], 10_000), [100, 200]);
});

test('degenerate budgets are handled rather than thrown', () => {
  assert.deepEqual(shareBudget([], 100), []);
  assert.deepEqual(shareBudget([100], 0), [0]);
  assert.deepEqual(shareBudget([100], -5), [0]);
});
