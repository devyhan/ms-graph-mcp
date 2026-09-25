/**
 * stderr-only logging.
 *
 * This process speaks MCP over stdio, so stdout is reserved for JSON-RPC frames.
 * Anything written there corrupts the protocol stream, which is why every line
 * here goes through `process.stderr.write` and nothing calls `console.log`.
 */

/** Structured context appended to a log line as JSON. */
export type LogMeta = Record<string, unknown>;

export interface Logger {
  debug(message: string, meta?: LogMeta): void;
  info(message: string, meta?: LogMeta): void;
  warn(message: string, meta?: LogMeta): void;
  error(message: string, meta?: LogMeta): void;
}

const PREFIX = '[graph-mcp]';

/**
 * Meta may hold values JSON cannot represent (a BigInt from a Graph quota, an
 * Error, a cyclic response object). Logging must never throw, so unsupported
 * values degrade to strings instead of aborting the call.
 */
function safeJson(meta: LogMeta): string {
  const seen = new WeakSet<object>();
  try {
    return JSON.stringify(meta, (_key, value: unknown) => {
      if (typeof value === 'bigint') return `${value.toString()}n`;
      if (value instanceof Error) return { name: value.name, message: value.message };
      if (typeof value === 'object' && value !== null) {
        if (seen.has(value)) return '[Circular]';
        seen.add(value);
      }
      return value;
    });
  } catch {
    return String(meta);
  }
}

function write(level: string, message: string, meta?: LogMeta): void {
  const suffix = meta && Object.keys(meta).length > 0 ? ` ${safeJson(meta)}` : '';
  try {
    process.stderr.write(`${PREFIX} ${level} ${message}${suffix}\n`);
  } catch {
    // A closed or broken stderr must not take the server down.
  }
}

/** Builds a logger whose `debug` output is gated on `verbose`. */
export function createLogger(verbose: boolean): Logger {
  return {
    debug(message: string, meta?: LogMeta): void {
      if (!verbose) return;
      write('DEBUG', message, meta);
    },
    info(message: string, meta?: LogMeta): void {
      write('INFO', message, meta);
    },
    warn(message: string, meta?: LogMeta): void {
      write('WARN', message, meta);
    },
    error(message: string, meta?: LogMeta): void {
      write('ERROR', message, meta);
    },
  };
}

/**
 * Process-wide logger. Modules import this binding directly; `index.ts` swaps in
 * a configured instance once flags are parsed, and the ESM live binding means
 * importers see the replacement without re-importing.
 */
export let logger: Logger = createLogger(false);

/** Replaces the module-level logger. Call once, during startup. */
export function setLogger(next: Logger): void {
  logger = next;
}
