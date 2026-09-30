import assert from 'node:assert/strict';
import test from 'node:test';

import type {
  GraphBatchRequest,
  GraphBatchResponse,
  GraphClient,
  GraphRequestOptions,
  GraphResponse,
  ServerConfig,
  ToolDefinition,
} from '../contracts.js';
import { serializeResult } from '../util/truncate.js';
import { mailModule } from './mail.js';

/**
 * The first tests in this repo that run a tool handler rather than a helper.
 *
 * That gap is why two bugs reached a live mailbox. A cache that outlived one
 * call made the same request return different bytes the second time, and a
 * link rule keyed on length folded away three references a summary had cited.
 * Neither is visible from a pure function: the first needs the handler called
 * twice, the second needs the whole body path, Exchange rendering included.
 *
 * Every Graph call is stubbed, so this needs no account and no network.
 */

const config: ServerConfig = {
  clientId: '00000000-0000-0000-0000-000000000000',
  tenantId: 'common',
  authority: 'https://login.microsoftonline.com/common',
  graphHost: 'graph.microsoft.com',
  groups: ['mail'],
  readOnly: true,
  graphVersion: 'v1.0',
  allowBeta: false,
  discovery: false,
  orgMode: false,
  maxOutputChars: 60_000,
  verbose: false,
  cacheDir: '/nonexistent',
  authFlow: 'auto',
  authPort: 0,
  allowGenericWrite: false,
};

/** Measured lengths: no genuine link in the sample corpus reached 200 characters. */
const REAL_LINK = 'https://docs.example.com/spreadsheets/d/1mrhAvAFSpBONA5lOUfun/edit?gid=17166';
const TRACKING_LINK = `https://click.example.com/ls/click?upn=${'u'.repeat(720)}`;

const POLICY =
  'Trust only the printed digits. Marks are excluded from the stroke count, and a cell ' +
  'that carries a mark but no digit comes back null and flagged for review.';

/** How Exchange renders a reply: new text, then the quoted message with its header. */
function reply(newText: string, quoted: string, from = 'ada'): string {
  return (
    `${newText}\n\nFrom: ${from} <${from}@example.com>\nSent: Monday 7 September 2026 10:08\n` +
    `To: Team <team@example.com>\nSubject: RE: the thing\n\n${quoted}`
  );
}

const M1 = `Sharing the agreed policy below.\n\n${POLICY}\n\nReference: ${REAL_LINK}`;
const M2 = reply('Confirmed on our side, thank you.', M1);
const M3 = reply('One more question about the 27-hole case.', M2, 'bob');

interface Fixture {
  id: string;
  receivedDateTime: string;
  subject: string;
  body: string;
  contentType?: string;
}

const THREAD: Fixture[] = [
  { id: 'A', receivedDateTime: '2026-09-07T01:08:00Z', subject: 'the thing', body: M1 },
  { id: 'B', receivedDateTime: '2026-09-07T07:32:00Z', subject: 'RE: the thing', body: M2 },
  { id: 'C', receivedDateTime: '2026-09-07T08:10:00Z', subject: 'RE: the thing', body: M3 },
];

const NOTIFICATION: Fixture = {
  id: 'N',
  receivedDateTime: '2026-09-14T04:16:00Z',
  subject: 'someone mentioned you',
  body: `"please take a look"\n[Design] <${TRACKING_LINK}>\nOpen it here: ${REAL_LINK}`,
};

/** A sentence far enough into a long body that a per-message cap would eat it. */
const BURIED = 'The vendor will bear the API cost until closed beta.';

const LONG: Fixture = {
  id: 'LONG',
  receivedDateTime: '2026-09-22T07:15:00Z',
  subject: 'a long one',
  body: `${'Background that nobody will cite. '.repeat(800)}\n\n${BURIED}`,
};

/** Exchange ignored the Prefer header and sent markup; stripHtml eats every href. */
const HTML_FIXTURE: Fixture = {
  id: 'HTML',
  receivedDateTime: '2026-09-23T00:21:00Z',
  subject: 'design feedback',
  contentType: 'html',
  body: `<p>Reflected your notes.</p><p><a href="${REAL_LINK}">the file</a></p>`,
};

/**
 * A thread whose FIRST message is long and whose reply quotes it. Folding
 * replaces the quote with a pointer at #1, so a budget that then trims #1 would
 * leave the pointer naming text no longer in the response.
 */
const DEEP = 'Agreed scope for the integration. ' + 'Detail the vendor will quote back. '.repeat(400);
const DEEP_THREAD: Fixture[] = [
  { id: 'D1', receivedDateTime: '2026-09-01T01:00:00Z', subject: 'scope', body: DEEP },
  {
    id: 'D2',
    receivedDateTime: '2026-09-02T01:00:00Z',
    subject: 'RE: scope',
    body: reply('Understood, we will confirm the estimate this week.', DEEP),
  },
];

const ALL = [...THREAD, NOTIFICATION, LONG, HTML_FIXTURE, ...DEEP_THREAD];

function payload(fixture: Fixture): Record<string, unknown> {
  return {
    id: fixture.id,
    subject: fixture.subject,
    from: { emailAddress: { name: 'Ada', address: 'ada@example.com' } },
    toRecipients: [{ emailAddress: { name: 'Team', address: 'team@example.com' } }],
    receivedDateTime: fixture.receivedDateTime,
    conversationId: ['A', 'B', 'C'].includes(fixture.id)
      ? 'conv-1'
      : ['D1', 'D2'].includes(fixture.id)
        ? 'conv-deep'
        : `conv-${fixture.id}`,
    body: { contentType: fixture.contentType ?? 'text', content: fixture.body },
    // ~250 chars in real Outlook: the id again, URL-encoded, wrapped in an owa query.
    webLink: `https://outlook.office365.com/owa/?ItemID=${encodeURIComponent(fixture.id)}${'A'.repeat(180)}&exvsurl=1&viewmodel=ReadMessageItem`,
  };
}

function stubGraph(stub: { nextLink?: string } = {}): GraphClient & { readonly requests: string[] } {
  const requests: string[] = [];
  return {
    requests,
    async request<T = unknown>(opts: GraphRequestOptions): Promise<GraphResponse<T>> {
      requests.push(opts.path);
      const single = /^\/me\/messages\/(.+)$/.exec(opts.path);
      if (single !== null) {
        const id = decodeURIComponent(single[1] as string);
        const found = ALL.find((entry) => entry.id === id);
        if (found === undefined) throw new Error(`stub has no message ${id}`);
        return { status: 200, data: payload(found) as T };
      }
      if (opts.path === '/me/messages') {
        const filter = String(opts.query?.['$filter'] ?? '');
        const conversation = /conversationId eq '([^']*)'/.exec(filter)?.[1];
        // No conversation filter means a plain list or a $search — serve everything,
        // which is what mail_list_messages and mail_search_messages project down.
        const rows =
          conversation === undefined
            ? [...ALL]
            : ALL.filter((entry) => payload(entry)['conversationId'] === conversation);
        // Deliberately newest-first, so a handler that forgets to sort fails here.
        const response: GraphResponse<T> = {
          status: 200,
          data: { value: [...rows].reverse().map(payload) } as T,
        };
        // Exchange picks its own page size: a short page with a live nextLink is
        // normal, and is exactly the case a row count misreads as "complete".
        if (stub.nextLink !== undefined) response.nextLink = stub.nextLink;
        return response;
      }
      throw new Error(`stub has no route for ${opts.path}`);
    },
    async batch(reqs: GraphBatchRequest[]): Promise<GraphBatchResponse[]> {
      return reqs.map((req) => {
        const id = decodeURIComponent(
          /\/me\/messages\/([^?]+)/.exec(req.url)?.[1] ?? '',
        );
        const found = ALL.find((entry) => entry.id === id);
        return found === undefined
          ? { id: req.id, status: 404 }
          : { id: req.id, status: 200, body: payload(found) };
      });
    },
    async follow<T = unknown>(): Promise<GraphResponse<T>> {
      throw new Error('stub does not page');
    },
  };
}

function tool(name: string, stub: { nextLink?: string } = {}): ToolDefinition {
  const graph = stubGraph(stub);
  const found = mailModule.build({ graph, config }).find((entry) => entry.name === name);
  assert.ok(found !== undefined, `${name} should exist`);
  return found;
}

type ThreadResult = {
  count: number;
  reason: string;
  messages: Array<Record<string, unknown>>;
  meta: { fold: { folded: number; links: number }; participants: string[] };
};

// ---------------------------------------------------------------------------
// mail_fetch_thread
// ---------------------------------------------------------------------------

test('the same call twice returns the same bytes', async () => {
  // The regression: a fold cache that outlived one call returned a message's
  // full text first and nothing but pointers at itself the second time.
  const fetchThread = tool('mail_fetch_thread');
  const first = await fetchThread.handler({ fromMessageId: 'C' });
  const second = await fetchThread.handler({ fromMessageId: 'C' });
  assert.deepEqual(second, first);
});

test('a fresh tool instance folds identically to the first', async () => {
  // Same assertion one layer out: no state may survive in module scope either.
  const a = await tool('mail_fetch_thread').handler({ fromMessageId: 'C' });
  const b = await tool('mail_fetch_thread').handler({ fromMessageId: 'C' });
  assert.deepEqual(b, a);
});

test('ids in any order return the same result', async () => {
  const fetchThread = tool('mail_fetch_thread');
  const forward = await fetchThread.handler({ ids: ['A', 'B', 'C'] });
  const backward = await fetchThread.handler({ ids: ['C', 'B', 'A'] });
  assert.deepEqual(backward, forward);
});

test('messages come back oldest-first however the mailbox ordered them', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'C',
  })) as ThreadResult;
  assert.deepEqual(
    result.messages.map((message) => message['id']),
    ['A', 'B', 'C'],
  );
  assert.deepEqual(
    result.messages.map((message) => message['n']),
    [1, 2, 3],
  );
});

test('a quoted repeat folds and the original survives exactly once', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'C',
  })) as ThreadResult;

  assert.ok(result.meta.fold.folded > 0, 'something should fold in a three-deep chain');
  const bodies = result.messages.map((message) => String(message['body'] ?? ''));
  const carrying = bodies.filter((body) => body.includes(POLICY));
  assert.equal(carrying.length, 1, 'the policy text appears once, not three times');
  assert.ok(bodies[0]?.includes(POLICY), 'and it is the message that first said it');
  assert.ok(bodies[2]?.includes('[quoted from'), 'the reply points back instead of repeating');
});

test('every pointer names an earlier message in this same response', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'C',
  })) as ThreadResult;

  for (const message of result.messages) {
    const position = Number(message['n']);
    for (const [, cited] of String(message['body'] ?? '').matchAll(/#(\d+)/g)) {
      const target = Number(cited);
      assert.ok(
        target >= 1 && target < position,
        `message ${position} cites #${target}, which is not behind it`,
      );
    }
  }
});

test('fold:"off" returns the repeats and is larger', async () => {
  const fetchThread = tool('mail_fetch_thread');
  const folded = (await fetchThread.handler({
    fromMessageId: 'C',
    totalBodyChars: 120_000,
  })) as ThreadResult;
  const whole = (await fetchThread.handler({
    fromMessageId: 'C',
    fold: 'off',
    totalBodyChars: 120_000,
  })) as ThreadResult;

  const size = (r: ThreadResult): number =>
    r.messages.reduce((sum, m) => sum + String(m['body'] ?? '').length, 0);
  assert.ok(size(whole) > size(folded), 'unfolded must be the bigger of the two');
  assert.equal(whole.meta.fold.folded, 0);
});

test('a genuine link survives and a tracking URL does not', async () => {
  // The regression, at the layer it happened: a length-only rule folded away a
  // design file, a spreadsheet and an API doc because they were merely long.
  const result = (await tool('mail_fetch_thread').handler({ ids: ['N'] })) as ThreadResult;
  const body = String(result.messages[0]?.['body'] ?? '');

  assert.ok(body.includes(REAL_LINK), 'a link a reader would follow must be returned intact');
  assert.ok(!body.includes(TRACKING_LINK), 'the opaque redirector must not be');
  assert.ok(body.includes('click.example.com'), 'but the reader still learns it was there');
  assert.equal(result.meta.fold.links, 1);
});

test('foldLinks:false keeps the tracking URL verbatim', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    ids: ['N'],
    foldLinks: false,
  })) as ThreadResult;
  assert.ok(String(result.messages[0]?.['body'] ?? '').includes(TRACKING_LINK));
});

test('an unreadable id is reported, not fatal', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    ids: ['A', 'missing'],
  })) as ThreadResult & { meta: { missing?: Array<{ id: string; status: number }> } };

  assert.equal(result.count, 1);
  assert.deepEqual(result.meta.missing, [{ id: 'missing', status: 404 }]);
});

test('ids a caller named are never dropped by maxMessages', async () => {
  // The regression: the newest-N slice ran in both modes, so 25 ids came back as
  // 20 with the five oldest gone and nothing in meta.missing to say so.
  const result = (await tool('mail_fetch_thread').handler({
    ids: ['A', 'B', 'C', 'N'],
    maxMessages: 2,
  })) as ThreadResult & { meta: { missing?: unknown } };

  assert.equal(result.count, 4, 'every id asked for must come back');
  assert.equal(result.reason, 'complete');
  assert.equal(result.meta.missing, undefined);
});

test('a live nextLink is reported, not read as a finished conversation', async () => {
  // Exchange answers $top=50 with a short page and a live cursor. Counting rows
  // reads that as "complete" and the model summarises part of a thread as all
  // of it — the tool's own description promises reason says which.
  const truthful = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'C',
  })) as ThreadResult;
  assert.equal(truthful.reason, 'complete');

  const partial = (await tool('mail_fetch_thread', {
    nextLink: 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=p2',
  }).handler({ fromMessageId: 'C' })) as ThreadResult;
  assert.equal(partial.reason, 'maxPages');
});

test('a body that arrived as HTML says so and counts the links it lost', async () => {
  // stripHtml has always discarded every href with its tag. The change does not
  // fix that; it refuses to let it pass silently.
  const result = (await tool('mail_fetch_thread').handler({
    ids: ['HTML'],
  })) as ThreadResult & { meta: { strippedHtmlMessages?: number } };

  const message = result.messages[0] as Record<string, unknown>;
  assert.equal(message['bodyRendering'], 'stripped-html');
  assert.equal(message['linksLost'], 1);
  assert.equal(result.meta.strippedHtmlMessages, 1);
  assert.ok(!String(message['body'] ?? '').includes(REAL_LINK), 'the link is genuinely gone');
});

test('an ordinary body says nothing about rendering', async () => {
  const result = (await tool('mail_fetch_thread').handler({ ids: ['A'] })) as ThreadResult;
  const message = result.messages[0] as Record<string, unknown>;
  assert.equal(message['bodyRendering'], undefined);
  assert.equal(message['linksLost'], undefined);
});

test('no pointer ever names a message the budget truncated', async () => {
  // The regression, and the worst one found: folding ran first and the budget
  // cut second, so [quoted from #1] survived while the text it named was trimmed
  // off the end of #1. Two load-bearing facts left the response that way, with
  // the marker still reading as a recoverable reference.
  const fetchThread = tool('mail_fetch_thread');

  for (const totalBodyChars of [1_000, 2_500, 6_000, 20_000, 120_000]) {
    const result = (await fetchThread.handler({
      fromMessageId: 'D2',
      totalBodyChars,
    })) as ThreadResult;

    const cut = new Set(
      result.messages.filter((m) => m['bodyTruncated'] === true).map((m) => Number(m['n'])),
    );
    for (const message of result.messages) {
      for (const [, cited] of String(message['body'] ?? '').matchAll(/#(\d+)/g)) {
        assert.ok(
          !cut.has(Number(cited)),
          `at totalBodyChars=${totalBodyChars}, #${message['n']} points at truncated #${cited}`,
        );
      }
    }
  }
});

test('when the budget cannot hold a pointer target, folding is dropped and said so', async () => {
  // Correctness over compression: a cut body announces itself, a dangling
  // pointer does not. The fallback must be visible, not silent.
  const result = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'D2',
    totalBodyChars: 1_000,
  })) as ThreadResult & { meta: { fold: { droppedBecause?: string; folded: number } } };

  assert.equal(result.meta.fold.droppedBecause, 'budget');
  assert.equal(result.meta.fold.folded, 0);
  for (const message of result.messages) {
    assert.ok(!String(message['body'] ?? '').includes('[quoted from'));
  }
});

test('a roomy budget still folds and keeps the target whole', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'D2',
    totalBodyChars: 120_000,
  })) as ThreadResult & { meta: { fold: { droppedBecause?: string; folded: number } } };

  assert.equal(result.meta.fold.droppedBecause, undefined);
  assert.ok(result.meta.fold.folded > 0, 'the quote should fold');
  assert.ok(String(result.messages[1]?.['body'] ?? '').includes('[quoted from #1]'));
  assert.equal(result.messages[0]?.['bodyTruncated'], undefined);
});

test('messages is the only array at the top level', async () => {
  // The output serialiser shortens whichever top-level array has the MOST
  // elements. A thread cc'd to thirty people would otherwise have its
  // participant list trimmed while the messages it was asked for stayed whole.
  const result = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'C',
  })) as unknown as Record<string, unknown>;

  const arrays = Object.entries(result)
    .filter(([, value]) => Array.isArray(value))
    .map(([key]) => key);
  assert.deepEqual(arrays, ['messages']);
});

test('the result still parses as JSON when the serialiser has to cut it', async () => {
  // Not a formality: serializeResult shortens whichever TOP-LEVEL array has the
  // most ELEMENTS, and a thread cc'd to thirty people has more participants than
  // messages. Picking the wrong array leaves a character-cut fragment that is no
  // longer JSON. The cap here is deliberately small enough to force the cut.
  const result = await tool('mail_fetch_thread').handler({
    ids: ['LONG', 'A', 'B', 'C'],
    totalBodyChars: 120_000,
  });

  const whole = serializeResult(result, 10_000_000);
  assert.equal(whole.truncated, false, 'the fixture must be under a huge cap');

  const cut = serializeResult(result, 4_000);
  assert.equal(cut.truncated, true, 'the fixture must be over a small one');
  const parsed = JSON.parse(cut.text) as { truncated?: { field?: string } };
  assert.equal(parsed.truncated?.field, 'messages', 'messages is what should shrink');
});

test('asking for neither or both of the two selectors says so plainly', async () => {
  const fetchThread = tool('mail_fetch_thread');
  await assert.rejects(() => fetchThread.handler({}), /fromMessageId.*ids|ids.*fromMessageId/s);
  await assert.rejects(
    () => fetchThread.handler({ fromMessageId: 'A', ids: ['A'] }),
    /both/,
  );
});

test('no per-message ceiling by default, so the long message survives', async () => {
  // The regression: maxBodyChars defaulted to 4000 and quietly clipped 17 of 24
  // messages in a real thread, losing five facts the fold itself had preserved.
  // The per-message cap defeats the whole point — the message that matters in a
  // chain is usually the long one.
  const result = (await tool('mail_fetch_thread').handler({
    ids: ['LONG'],
    totalBodyChars: 120_000,
  })) as ThreadResult;

  const message = result.messages[0] as Record<string, unknown>;
  assert.equal(message['bodyTruncated'], undefined, 'nothing should have been cut');
  assert.ok(
    Number(message['bodyChars']) > 20_000,
    `expected the whole body, got ${message['bodyChars']} chars`,
  );
  assert.ok(String(message['body'] ?? '').includes(BURIED), 'text near the end must survive');
});

test('maxBodyChars is honoured when a caller actually asks for it', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    ids: ['LONG'],
    maxBodyChars: 1_000,
    totalBodyChars: 120_000,
  })) as ThreadResult;

  const message = result.messages[0] as Record<string, unknown>;
  assert.equal(message['bodyTruncated'], true);
  assert.ok(Number(message['bodyChars']) < 1_200, 'the cap plus its marker, no more');
});

test('a bigger total budget never returns less text', async () => {
  // Monotonicity: the bug above showed up first as a larger budget producing a
  // SMALLER result, which is the signature of a cap applied in the wrong place.
  const fetchThread = tool('mail_fetch_thread');
  const sizes: number[] = [];
  for (const totalBodyChars of [2_000, 20_000, 120_000]) {
    const result = (await fetchThread.handler({ ids: ['LONG'], totalBodyChars })) as ThreadResult;
    sizes.push(Number(result.messages[0]?.['bodyChars'] ?? 0));
  }
  for (let i = 1; i < sizes.length; i += 1) {
    assert.ok(
      (sizes[i] as number) >= (sizes[i - 1] as number),
      `expected non-decreasing sizes, got ${sizes.join(' -> ')}`,
    );
  }
});

test('budget is shared so a short message is not cut for a long one', async () => {
  const result = (await tool('mail_fetch_thread').handler({
    fromMessageId: 'C',
    totalBodyChars: 1_000,
  })) as ThreadResult;

  const truncated = result.messages.filter((message) => message['bodyTruncated'] === true);
  assert.ok(truncated.length < result.messages.length, 'not everything should be cut');
});

// ---------------------------------------------------------------------------
// mail_list_messages / mail_search_messages
// ---------------------------------------------------------------------------

test('a listing does not carry webLink unless asked', async () => {
  // Measured at 23.7% of a search response, for the message id all over again.
  const rows = (await tool('mail_list_messages').handler({})) as {
    messages: Array<Record<string, unknown>>;
  };
  assert.ok(rows.messages.length > 0);
  for (const row of rows.messages) {
    assert.ok(!('webLink' in row), 'webLink must be absent by default');
    assert.ok(typeof row['id'] === 'string', 'the id — the usable handle — stays');
    assert.ok('bodyPreview' in row || 'subject' in row, 'the readable fields stay');
  }
});

test('includeWebLink brings it back', async () => {
  const rows = (await tool('mail_list_messages').handler({ includeWebLink: true })) as {
    messages: Array<Record<string, unknown>>;
  };
  assert.ok('webLink' in (rows.messages[0] as Record<string, unknown>));
});

test('naming webLink in select still works, whatever the flag says', async () => {
  // An existing caller that spelled it out meant it. `extras` filters against
  // SUMMARY_FIELDS, so webLink has to stay in that set or this silently breaks.
  const rows = (await tool('mail_list_messages').handler({
    select: ['subject', 'webLink'],
  })) as { messages: Array<Record<string, unknown>> };
  assert.ok('webLink' in (rows.messages[0] as Record<string, unknown>));
});

test('search drops webLink by default too', async () => {
  const off = (await tool('mail_search_messages').handler({ query: 'policy' })) as {
    messages: Array<Record<string, unknown>>;
  };
  const on = (await tool('mail_search_messages').handler({
    query: 'policy',
    includeWebLink: true,
  })) as { messages: Array<Record<string, unknown>> };

  assert.ok(off.messages.length > 0, 'the stub must return something to judge');
  assert.ok(!('webLink' in (off.messages[0] as Record<string, unknown>)));
  assert.ok('webLink' in (on.messages[0] as Record<string, unknown>));
});

// ---------------------------------------------------------------------------
// mail_get_message
// ---------------------------------------------------------------------------

test('mail_get_message keeps the keys it always returned', async () => {
  const result = (await tool('mail_get_message').handler({ id: 'A' })) as Record<string, unknown>;
  for (const key of ['id', 'subject', 'from', 'receivedDateTime', 'body', 'bodyFormat']) {
    assert.ok(key in result, `${key} must still be returned`);
  }
  assert.equal(result['bodyFormat'], 'text');
});

test('mail_get_message says which rendering produced the body', async () => {
  const result = (await tool('mail_get_message').handler({ id: 'A' })) as Record<string, unknown>;
  assert.equal(result['bodyRendering'], 'exchange-text');
});

test('mail_get_message folds a tracking URL but keeps a real one', async () => {
  const result = (await tool('mail_get_message').handler({ id: 'N' })) as Record<string, unknown>;
  const body = String(result['body'] ?? '');
  assert.ok(body.includes(REAL_LINK));
  assert.ok(!body.includes(TRACKING_LINK));
  assert.equal(result['foldedLinks'], 1);
});

test('mail_get_message leaves format:"html" untouched', async () => {
  // Folding a URL inside an href would corrupt the markup the caller asked for.
  const result = (await tool('mail_get_message').handler({
    id: 'N',
    format: 'html',
    maxBodyChars: 100_000,
  })) as Record<string, unknown>;
  assert.equal(result['bodyRendering'], 'html');
  assert.ok(String(result['body'] ?? '').includes(TRACKING_LINK));
});
