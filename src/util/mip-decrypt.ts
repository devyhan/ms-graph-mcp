/**
 * MIP decryptor sidecar launcher behind `files_parse_document`.
 *
 * Purview-protected (AIP/RMS) documents cannot be decrypted in-process: the
 * Rights Management crypto lives in the native MIP SDK, which needs a binary
 * this package does not ship. The sidecar (`native/mip-decrypt`, built
 * separately) does the unprotect step; this module resolves its path, feeds it
 * the wrapper bytes through temp files, and maps its outcome back onto the
 * `ParseDeps.decrypt` hook. Unconfigured servers resolve to `undefined` and
 * keep the refusal path, so existing users see no behaviour change.
 *
 * Sidecar protocol (the only coupling between the two processes): argv is
 * `<sidecar> <inPath> <outPath>`; exit 0 means `<outPath>` holds the
 * plaintext (or a byte copy when the input was not actually protected);
 * exit 2 is a usage error; exit 1 is any decrypt failure. Diagnostics go to
 * stderr prefixed `[mip]` and never carry token material.
 */

import { spawn } from 'node:child_process';
import { accessSync, constants as fsConstants, promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

/** Environment fallback for the sidecar path, read when no explicit path is given. */
export const MIP_DECRYPTOR_ENV = 'MS365_MCP_MIP_DECRYPTOR';

/** Binary name looked up on PATH and beside the package when nothing is configured. */
export const MIP_DECRYPTOR_BINARY = 'mip-decrypt';

/**
 * How long one decrypt may run. RMS round-trips plus a cold MIP engine load
 * take seconds, not minutes, but a first run downloads templates; 120s covers
 * that without letting a wedged helper hang a parse forever.
 */
export const MIP_DECRYPT_TIMEOUT_MS = 120_000;

/** Most stderr kept on an error: enough for the `[mip]` diagnosis, never a dump. */
const MAX_STDERR_CHARS = 2000;

/** Machine-readable reason on a sidecar failure, so callers can branch without parsing text. */
export type MipDecryptFailureKind =
  | 'no-runtime'
  | 'usage'
  | 'timeout'
  | 'no-rights'
  | 'failed';

export class MipDecryptError extends Error {
  override readonly name = 'MipDecryptError';
  readonly kind: MipDecryptFailureKind;
  readonly fileName: string;

  constructor(fileName: string, kind: MipDecryptFailureKind, message: string) {
    super(message);
    this.kind = kind;
    this.fileName = fileName;
  }
}

export interface MipDecryptorOptions {
  /**
   * Explicit sidecar path (`ServerConfig.mipDecryptor`, from `--mip-decryptor`
   * or its environment fallback). Wins over the environment lookup, so the
   * configured flag takes effect once the caller wires it through here.
   */
  explicitPath?: string | undefined;
  /** Per-decrypt timeout in milliseconds. Defaults to MIP_DECRYPT_TIMEOUT_MS. */
  timeoutMs?: number | undefined;
  /** Environment to read. Defaults to process.env; tests inject a stub. */
  env?: NodeJS.ProcessEnv | undefined;
}

function isExecutable(file: string): boolean {
  try {
    // accessSync throws when the path is missing or not executable; either
    // way the candidate is unusable, so a boolean is the whole answer.
    accessSync(file, fsConstants.X_OK);
    return true;
  } catch {
    return false;
  }
}

function bundledCandidates(): string[] {
  // A `mip-decrypt` built from native/mip-decrypt and dropped next to the
  // installed package works with zero configuration; running from source it
  // sits at the repo root instead. Both are just guesses — PATH is the
  // documented lookup — so misses are silently skipped.
  const here = path.dirname(fileURLToPath(import.meta.url));
  return [
    path.resolve(here, '..', '..', MIP_DECRYPTOR_BINARY),
    path.resolve(here, '..', '..', 'native', 'mip-decrypt', 'build', MIP_DECRYPTOR_BINARY),
  ];
}

function pathCandidates(env: NodeJS.ProcessEnv): string[] {
  const raw = env['PATH'] ?? '';
  if (raw === '') return [];
  const exts = process.platform === 'win32' ? ['', '.exe', '.cmd', '.bat'] : [''];
  const out: string[] = [];
  for (const dir of raw.split(path.delimiter)) {
    if (dir === '') continue;
    for (const ext of exts) out.push(path.join(dir, `${MIP_DECRYPTOR_BINARY}${ext}`));
  }
  return out;
}

/**
 * Resolves the sidecar path in flag > env > bundled > PATH order. Returns
 * `undefined` when nothing resolves: the caller then omits `decrypt` and the
 * parse lane refuses protected documents instead of decrypting them.
 */
export function resolveMipDecryptorPath(options: MipDecryptorOptions = {}): string | undefined {
  const env = options.env ?? process.env;
  const explicit = options.explicitPath ?? env[MIP_DECRYPTOR_ENV];
  if (explicit !== undefined && explicit !== '') return explicit;
  for (const candidate of [...bundledCandidates(), ...pathCandidates(env)]) {
    if (isExecutable(candidate)) return candidate;
  }
  return undefined;
}

/**
 * Builds the `ParseDeps.decrypt` hook, or `undefined` when no sidecar
 * resolves. The hook shape matches `decrypt?: (bytes, fileName) => Promise<Uint8Array>`.
 */
export function createMipDecryptor(
  options: MipDecryptorOptions = {},
): ((bytes: Uint8Array, fileName: string) => Promise<Uint8Array>) | undefined {
  const sidecar = resolveMipDecryptorPath(options);
  if (sidecar === undefined) return undefined;
  const timeoutMs = options.timeoutMs ?? MIP_DECRYPT_TIMEOUT_MS;
  return (bytes, fileName) => decryptWithSidecar(sidecar, bytes, fileName, timeoutMs);
}

function stderrTail(stderr: string): string {
  const trimmed = stderr.trim();
  if (trimmed.length <= MAX_STDERR_CHARS) return trimmed;
  return `…${trimmed.slice(trimmed.length - MAX_STDERR_CHARS)}`;
}

/**
 * True when the sidecar's stderr says the identity lacks usage rights rather
 * than the file being corrupt or the runtime broken. The C++ sidecar prints an
 * explicit `[mip] NO_RIGHTS` marker; the keyword fallback covers older builds.
 */
function isNoRights(stderr: string): boolean {
  const text = stderr.toLowerCase();
  return (
    text.includes('no_rights') ||
    text.includes('no rights') ||
    text.includes('access denied') ||
    text.includes('access-denied') ||
    text.includes('not authorized') ||
    text.includes('no permission')
  );
}

function removeQuietly(file: string): Promise<void> {
  return fs.unlink(file).then(
    () => undefined,
    () => undefined,
  );
}

/**
 * Runs `<sidecar> <inPath> <outPath>` over one document. Wrapper bytes travel
 * through a 0600 temp file inside a fresh mkdtemp directory; both paths and
 * the directory are removed in a finally, including on timeout kill, so
 * protected bytes are never left behind in /tmp.
 */
export async function decryptWithSidecar(
  sidecar: string,
  bytes: Uint8Array,
  fileName: string,
  timeoutMs: number = MIP_DECRYPT_TIMEOUT_MS,
): Promise<Uint8Array> {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'mip-decrypt-'));
  const inPath = path.join(dir, 'input');
  const outPath = path.join(dir, 'output');
  try {
    // 0600 from creation: the wrapper may itself be sensitive, and the
    // directory default plus umask is not a guarantee worth relying on.
    await fs.writeFile(inPath, bytes, { mode: 0o600 });
    // Success diagnostics stay with the sidecar: stdout is forbidden in serve
    // mode and the parse result has no diagnostics channel, so only failures
    // (which throw with the stderr tail) surface anything.
    await runSidecar(sidecar, inPath, outPath, fileName, timeoutMs);
    let plain: Uint8Array;
    try {
      plain = new Uint8Array(await fs.readFile(outPath));
    } catch {
      throw new MipDecryptError(
        fileName,
        'failed',
        `"${fileName}" was sent to the MIP decryptor but it produced no output file. ${sidecarHint(sidecar)}`,
      );
    }
    if (plain.length === 0) {
      throw new MipDecryptError(
        fileName,
        'failed',
        `"${fileName}" decrypted to an empty file, so the sidecar run is treated as failed. ${sidecarHint(sidecar)}`,
      );
    }
    return plain;
  } finally {
    await removeQuietly(inPath);
    await removeQuietly(outPath);
    await fs.rmdir(dir).then(
      () => undefined,
      () => undefined,
    );
  }
}

function formatDuration(ms: number): string {
  return ms >= 1000 ? `${Math.round(ms / 1000)}s` : `${ms}ms`;
}

function sidecarHint(sidecar: string): string {
  return `Sidecar: ${sidecar}. Re-run it by hand as \`${sidecar} <input> <output>\` to see its [mip] diagnostics.`;
}

function runSidecar(
  sidecar: string,
  inPath: string,
  outPath: string,
  fileName: string,
  timeoutMs: number,
): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    let settled = false;
    const finish = (fn: () => void): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      fn();
    };
    let child: ReturnType<typeof spawn>;
    try {
      child = spawn(sidecar, [inPath, outPath], { stdio: ['ignore', 'ignore', 'pipe'] });
    } catch (error) {
      finish(() =>
        reject(
          new MipDecryptError(
            fileName,
            'no-runtime',
            `"${fileName}" needs the MIP decryptor but it could not be started (${error instanceof Error ? error.message : String(error)}). ` +
              'Build native/mip-decrypt and point MS365_MCP_MIP_DECRYPTOR at the binary.',
          ),
        ),
      );
      return;
    }
    let stderr = '';
    child.stderr?.on('data', (chunk: Buffer | string) => {
      stderr += typeof chunk === 'string' ? chunk : chunk.toString('utf8');
      if (stderr.length > MAX_STDERR_CHARS * 4) stderr = stderr.slice(-MAX_STDERR_CHARS * 4);
    });
    const timer = setTimeout(() => {
      // SIGKILL, not SIGTERM: a helper waiting on device-code input ignores
      // SIGTERM, and a parse must not hang past its budget either way.
      child.kill('SIGKILL');
      finish(() =>
        reject(
          new MipDecryptError(
            fileName,
            'timeout',
            `"${fileName}" was still decrypting after ${formatDuration(timeoutMs)}, so the sidecar was killed. ` +
              'A first run that fetches templates can be slow; otherwise the auth helper may be waiting on input. ' +
              sidecarHint(sidecar),
          ),
        ),
      );
    }, timeoutMs);
    // A timed-out parse must not keep the event loop alive on the timer alone.
    timer.unref?.();
    child.on('error', (error: Error & { code?: string }) => {
      // ENOENT covers both a missing binary and (on Unix, via the shebang) a
      // missing interpreter; either way there is no usable runtime.
      finish(() =>
        reject(
          new MipDecryptError(
            fileName,
            'no-runtime',
            `"${fileName}" needs the MIP decryptor at "${sidecar}" but it could not be started (${error.message}). ` +
              'Build native/mip-decrypt and point MS365_MCP_MIP_DECRYPTOR at the binary.',
          ),
        ),
      );
    });
    child.on('close', (code: number | null) => {
      if (code === 0) {
        finish(() => resolve(stderr));
        return;
      }
      if (code === 2) {
        finish(() =>
          reject(
            new MipDecryptError(
              fileName,
              'usage',
              `"${fileName}" was sent to the MIP decryptor but it reported a usage error (exit 2): ${stderrTail(stderr) || 'no diagnostics'}. ` +
                'The binary is probably not the mip-decrypt sidecar this server expects. ' +
                sidecarHint(sidecar),
            ),
          ),
        );
        return;
      }
      if (isNoRights(stderr)) {
        finish(() =>
          reject(
            new MipDecryptError(
              fileName,
              'no-rights',
              `"${fileName}" is protected and the signed-in identity has no usage rights on it, so the MIP decryptor refused it. ` +
                'Ask the document owner for access, or sign the helper in as an identity that has it. ' +
                `Sidecar said: ${stderrTail(stderr) || 'no diagnostics'}`,
            ),
          ),
        );
        return;
      }
      finish(() =>
        reject(
          new MipDecryptError(
            fileName,
            'failed',
            `"${fileName}" could not be decrypted (sidecar exit ${code ?? 'unknown'}). ` +
              `Sidecar said: ${stderrTail(stderr) || 'no diagnostics'}`,
          ),
        ),
      );
    });
  });
}
