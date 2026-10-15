/**
 * Pre-sign-in environment checks for the browser flow.
 *
 * The browser grant needs three things this process can verify cheaply before
 * MSAL opens anything: a desktop session to open the browser in, a loopback
 * listener to catch the redirect on, and a route to Microsoft Entra. When the
 * first two are verifiably absent the provider skips the browser attempt
 * instead of failing through it; Entra reachability is warn-only, because a
 * failed probe must never stop a sign-in that would have worked.
 *
 * Every check is non-throwing and offline-safe except `canReachEntra`, which
 * performs one short-timeout fetch. Tests inject fakes through
 * `createAuthProvider({ preflight })`; nothing here is read from the network
 * in the suite.
 */

import net from 'node:net';

/** What one check concluded. `unknown` keeps the current behaviour. */
export type PreflightVerdict = 'possible' | 'impossible' | 'unknown';

export interface PreflightResult {
  verdict: PreflightVerdict;
  /** One line a warning can quote. Always present, whatever the verdict. */
  reason: string;
}

/** The checks `login` consults, injectable as a unit for tests. */
export interface PreflightChecks {
  canOpenBrowser(): Promise<PreflightResult>;
  canBindLoopback(port: number): Promise<PreflightResult>;
  canReachEntra(): Promise<PreflightResult>;
}

/**
 * Whether a browser opened from this process could be seen.
 *
 * macOS and Windows need no display variable: `open` and `start` target the
 * console session. Everywhere else a missing `DISPLAY`/`WAYLAND_DISPLAY` means
 * headless (SSH, a container, CI) and the browser attempt cannot work. A set
 * variable proves little — the display may still be unreachable — so that
 * stays `unknown` and the attempt goes ahead as before.
 */
export async function canOpenBrowser(
  platform: NodeJS.Platform = process.platform,
  env: NodeJS.ProcessEnv = process.env,
): Promise<PreflightResult> {
  try {
    if (platform === 'darwin' || platform === 'win32') {
      return { verdict: 'possible', reason: `${platform} opens the browser in the console session` };
    }
    const display = env['DISPLAY'] ?? env['WAYLAND_DISPLAY'] ?? '';
    if (display.trim() === '') {
      return {
        verdict: 'impossible',
        reason: 'no desktop session (neither DISPLAY nor WAYLAND_DISPLAY is set)',
      };
    }
    return { verdict: 'unknown', reason: `a display is set (${display.trim()}) but no browser was checked` };
  } catch (error) {
    return {
      verdict: 'unknown',
      reason: `the browser check itself failed: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}

/**
 * Whether the loopback redirect listener can bind.
 *
 * Listens on 127.0.0.1 — the address MSAL's `LoopbackClient` binds — and
 * closes again immediately. A refused bind (a firewall, a VPN that swallows
 * loopback, or the pinned `--auth-port` already taken) is what surfaces
 * minutes later as `loopback_server_timeout`, so failing here saves the wait.
 * Never throws: an unexpected error is `unknown`, not `impossible`.
 */
export async function canBindLoopback(port = 0): Promise<PreflightResult> {
  try {
    await new Promise<void>((resolve, reject) => {
      const server = net.createServer();
      server.once('error', reject);
      server.listen(port, '127.0.0.1', () => {
        server.close((error) => (error ? reject(error) : resolve()));
      });
    });
    return {
      verdict: 'possible',
      reason: port > 0 ? `127.0.0.1:${port} accepts a loopback listener` : '127.0.0.1 accepts a loopback listener',
    };
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    const code = error instanceof Error && 'code' in error ? String(error.code) : '';
    // A listen refusal genuinely means the redirect cannot be served; anything
    // else (no permission to create a socket at all, say) is inconclusive.
    const listenRefusal = ['EADDRINUSE', 'EADDRNOTAVAIL', 'EACCES', 'EPERM', 'EAFNOSUPPORT'].includes(code);
    return {
      verdict: listenRefusal ? 'impossible' : 'unknown',
      reason:
        port > 0 && code === 'EADDRINUSE'
          ? `127.0.0.1:${port} is already in use; the pinned --auth-port is taken`
          : `loopback bind failed: ${detail}`,
    };
  }
}

/** How long the reachability probe waits before calling it unreachable. */
export const ENTRA_PROBE_TIMEOUT_MS = 3000;

/**
 * Whether Microsoft Entra answers at all. Warn-only: a negative here warns
 * before the sign-in attempt, never instead of it.
 *
 * Any HTTP response counts as reachable — even an error status proves the
 * route exists. Injectable `fetchImpl` and `url` keep the suite offline.
 */
export async function canReachEntra(
  fetchImpl: typeof fetch = fetch,
  url = 'https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration',
  timeoutMs: number = ENTRA_PROBE_TIMEOUT_MS,
): Promise<PreflightResult> {
  try {
    await fetchImpl(url, { method: 'HEAD', signal: AbortSignal.timeout(timeoutMs) });
    return { verdict: 'possible', reason: 'Microsoft Entra answered' };
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    return { verdict: 'impossible', reason: `no answer from Microsoft Entra: ${detail}` };
  }
}

/** The production checks, used when `createAuthProvider` gets no override. */
export const defaultPreflight: PreflightChecks = {
  canOpenBrowser: () => canOpenBrowser(),
  canBindLoopback: (port: number) => canBindLoopback(port),
  canReachEntra: () => canReachEntra(),
};
