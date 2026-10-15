import assert from 'node:assert/strict';
import net from 'node:net';
import test from 'node:test';

import {
  canBindLoopback,
  canOpenBrowser,
  canReachEntra,
  defaultPreflight,
} from './preflight.js';

test('macOS and Windows need no display variable to open a browser', async () => {
  assert.equal((await canOpenBrowser('darwin', {})).verdict, 'possible');
  assert.equal((await canOpenBrowser('win32', {})).verdict, 'possible');
});

test('a Linux box without a display cannot open a browser', async () => {
  const result = await canOpenBrowser('linux', {});
  assert.equal(result.verdict, 'impossible');
  assert.match(result.reason, /DISPLAY.*WAYLAND_DISPLAY/);
});

test('a Linux display variable keeps the attempt but proves nothing', async () => {
  // `unknown` keeps the current behaviour: the browser is attempted as before.
  assert.equal((await canOpenBrowser('linux', { DISPLAY: ':0' })).verdict, 'unknown');
  assert.equal(
    (await canOpenBrowser('linux', { WAYLAND_DISPLAY: 'wayland-0' })).verdict,
    'unknown',
  );
});

test('the loopback check reports whether 127.0.0.1 binds', async () => {
  // A sandboxed runner may refuse the bind outright (EPERM): that is a genuine
  // "the browser flow cannot serve its redirect" verdict, not a test failure.
  const result = await canBindLoopback(0);
  if (result.verdict === 'possible') {
    assert.match(result.reason, /127\.0\.0\.1/);
  } else {
    assert.equal(result.verdict, 'impossible');
    assert.match(result.reason, /loopback bind failed/);
  }
});

test('a taken pinned port rules the browser flow out by name', async (t) => {
  const occupant = net.createServer();
  const bound = await new Promise<boolean>((resolve) => {
    occupant.once('error', () => resolve(false));
    occupant.listen(0, '127.0.0.1', () => resolve(true));
  });
  if (!bound) {
    occupant.close();
    t.skip('this runner cannot bind loopback at all (sandboxed)');
    return;
  }
  const address = occupant.address();
  assert.ok(typeof address === 'object' && address !== null);
  try {
    const result = await canBindLoopback(address.port);
    assert.equal(result.verdict, 'impossible');
    assert.match(result.reason, /already in use/);
    assert.match(result.reason, /--auth-port/);
  } finally {
    occupant.close();
  }
});

test('the loopback check never throws, even for nonsense input', async () => {
  const result = await canBindLoopback(99999);
  assert.equal(result.verdict, 'unknown');
  assert.ok(result.reason.length > 0);
});

test('Entra reachability reports what the fetch says, nothing more', async () => {
  const ok = await canReachEntra((async () => new Response(null, { status: 400 })) as typeof fetch);
  assert.equal(ok.verdict, 'possible');

  const down = await canReachEntra(
    (async () => {
      throw new Error('getaddrinfo ENOTFOUND login.microsoftonline.com');
    }) as typeof fetch,
  );
  assert.equal(down.verdict, 'impossible');
  assert.match(down.reason, /ENOTFOUND/);
});

test('the reachability probe times out instead of hanging', async () => {
  // A fetch that never answers but honours the abort signal, like the real
  // one. The long timer stands in for the socket a real fetch holds: without
  // something ref'd on the loop, the abort timer itself cannot fire either.
  const hanging = ((_url: string, init?: { signal?: AbortSignal }) =>
    new Promise<Response>((_resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('probe took too long')), 5000);
      init?.signal?.addEventListener('abort', () => {
        clearTimeout(timer);
        reject(new Error('This operation was aborted'));
      });
    })) as typeof fetch;
  const result = await canReachEntra(hanging, 'https://example.invalid/', 20);
  assert.equal(result.verdict, 'impossible');
  assert.ok(result.reason.length > 0);
});

test('the default preflight exposes the three checks', () => {
  assert.equal(typeof defaultPreflight.canOpenBrowser, 'function');
  assert.equal(typeof defaultPreflight.canBindLoopback, 'function');
  assert.equal(typeof defaultPreflight.canReachEntra, 'function');
});
