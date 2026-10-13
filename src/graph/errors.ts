/**
 * Turns a Microsoft Graph failure into a `GraphError` the model can act on.
 *
 * Graph documents its errors as `{ error: { code, message, innerError } }`, but
 * the wire never guarantees it: gateways answer with HTML, throttling responses
 * sometimes carry only headers, and token-endpoint style bodies use a flat
 * `{ error, error_description }`. Everything here parses defensively and always
 * produces a GraphError rather than throwing while building one.
 */

import { GraphError } from '../contracts.js';
import { SCOPE_GUESSES } from '../util/scopes.js';

/**
 * Statuses that clear on their own. 502 is included because the Graph gateway
 * emits short-lived 502s that behave exactly like a 503.
 */
const RETRYABLE_STATUSES: ReadonlySet<number> = new Set([429, 502, 503, 504]);

export function isRetryableStatus(status: number): boolean {
  return RETRYABLE_STATUSES.has(status);
}

/** Codes Graph uses when the body is missing, so hints stay keyed on something real. */
const DEFAULT_CODES: ReadonlyMap<number, string> = new Map([
  [400, 'badRequest'],
  [401, 'unauthenticated'],
  [403, 'accessDenied'],
  [404, 'itemNotFound'],
  [405, 'methodNotAllowed'],
  [409, 'conflict'],
  [410, 'resyncRequired'],
  [413, 'requestEntityTooLarge'],
  [429, 'activityLimitReached'],
  [500, 'generalException'],
  [501, 'notImplemented'],
  [502, 'serviceNotAvailable'],
  [503, 'serviceNotAvailable'],
  [504, 'gatewayTimeout'],
  [507, 'quotaLimitReached'],
]);

/**
 * Builds a GraphError from a failed Graph response. The 403 hint names the
 * scopes the path normally needs via the table in `util/scopes.ts`, shared
 * with `graph_request`.
 */
export function parseGraphError(
  status: number,
  path: string,
  bodyText: string,
  headers: Headers,
): GraphError {
  const parsed = readErrorBody(bodyText);
  const code = parsed.code ?? DEFAULT_CODES.get(status) ?? 'unknownError';
  const message =
    parsed.message ?? `Microsoft Graph returned HTTP ${status} for ${path} with no error body.`;
  const requestId =
    parsed.requestId ??
    headers.get('request-id') ??
    headers.get('client-request-id') ??
    undefined;

  return new GraphError({
    message,
    status,
    code,
    path,
    requestId,
    hint: buildHint({
      status,
      code: parsed.code,
      innerCode: parsed.innerCode,
      message,
      path,
      headers,
    }),
  });
}

interface ParsedErrorBody {
  code?: string | undefined;
  innerCode?: string | undefined;
  message?: string | undefined;
  requestId?: string | undefined;
}

function readErrorBody(bodyText: string): ParsedErrorBody {
  const trimmed = bodyText.trim();
  if (trimmed.length === 0) return {};

  let parsed: unknown;
  try {
    parsed = JSON.parse(trimmed);
  } catch {
    // Gateways and proxies answer with HTML or plain text. One line is enough
    // context; the rest is noise in a model's context window.
    return { message: firstLine(trimmed) };
  }

  const root = asRecord(parsed);
  if (!root) return { message: firstLine(trimmed) };

  const errorValue = root['error'];
  if (typeof errorValue === 'string') {
    // OAuth-shaped body, e.g. { error: 'invalid_grant', error_description: '...' }.
    return {
      code: errorValue,
      message: asString(root['error_description']) ?? errorValue,
    };
  }

  // A few endpoints answer without the `error` wrapper; fall back to the root.
  const error = asRecord(errorValue) ?? root;
  const inner = asRecord(error['innerError']);

  return {
    code: asString(error['code']),
    innerCode: asString(inner?.['code']),
    message: asString(error['message']),
    requestId:
      asString(inner?.['request-id']) ??
      asString(inner?.['requestId']) ??
      asString(inner?.['client-request-id']),
  };
}

function buildHint(ctx: {
  status: number;
  code: string | undefined;
  innerCode: string | undefined;
  message: string;
  path: string;
  headers: Headers;
}): string | undefined {
  const codes = [ctx.code, ctx.innerCode]
    .filter((value): value is string => typeof value === 'string')
    .map((value) => value.toLowerCase());
  const has = (...candidates: string[]): boolean =>
    candidates.some((candidate) => codes.includes(candidate.toLowerCase()));
  const message = ctx.message.toLowerCase();

  if (ctx.status === 401 || has('InvalidAuthenticationToken', 'unauthenticated', 'tokenExpired')) {
    return (
      'The Microsoft 365 session expired or the access token was rejected. ' +
      'Ask the user to sign in again (re-run the login tool, or restart the server to trigger device-code sign-in). ' +
      'Repeating this call will keep failing until a new token exists.'
    );
  }

  if (
    ctx.status === 403 ||
    has('accessDenied', 'Authorization_RequestDenied', 'authorizationRequestDenied', 'ErrorAccessDenied')
  ) {
    const scopes = guessScopes(ctx.path);
    return (
      `The signed-in user has no consent for this call. It normally needs ${scopes}. ` +
      'Confirm the tool group that owns this scope is enabled, then have the user sign in again so the new scope is consented. ' +
      'Scopes marked as needing tenant admin consent require an administrator to grant them first. ' +
      'If consent is already in place, the account itself lacks rights on this resource.'
    );
  }

  if (ctx.status === 404 || has('itemNotFound', 'ResourceNotFound', 'Request_ResourceNotFound')) {
    return (
      `No resource exists at ${ctx.path}. Graph ids are opaque and case-sensitive, so re-read the id from a list or search call ` +
      'rather than composing it, and check the resource is addressed in the right container (/me/... vs /users/{id}/..., the right drive, site or team). ' +
      'A deleted or moved item returns this too.'
    );
  }

  if (ctx.status === 405) {
    return (
      'This resource is not exposed on the Graph endpoint version that was used. ' +
      'Resources present only on beta return 405 on v1.0 and vice versa: retry against the other version (beta needs the --beta flag). ' +
      'Also confirm the HTTP method, since some Graph resources accept only POST or PATCH.'
    );
  }

  if (ctx.status === 410 || has('resyncRequired', 'syncStateNotFound')) {
    return (
      'The delta token expired or is no longer valid (directory tokens last about 7 days). ' +
      'Drop the stored token and run the delta query again without one to perform a full resync.'
    );
  }

  if (ctx.status === 429 || has('activityLimitReached')) {
    const retryAfter = ctx.headers.get('retry-after');
    const wait = retryAfter === null ? '' : ` Graph asked for a ${retryAfter}s pause.`;
    return (
      `Microsoft Graph throttled this request and the client already exhausted its retries.${wait} ` +
      'Narrow the query before trying again: lower $top, add $select, add a $filter, or fetch fewer pages. ' +
      'Throttling is scoped per mailbox, per tenant and per resource, so a smaller or later request usually succeeds.'
    );
  }

  if (ctx.status === 400 && (message.includes('$search') || ctx.path.includes('$search'))) {
    return (
      'Graph $search needs the term inside double quotes, e.g. $search="\\"quarterly budget\\"" (KQL syntax for mail and files). ' +
      'It supports no ranges or comparisons: express a date or numeric bound with $filter instead, ' +
      'e.g. $filter=receivedDateTime ge 2026-01-01T00:00:00Z. ' +
      'Outlook resources also refuse $search and $filter on the same request.'
    );
  }

  if (ctx.status === 507 || has('quotaLimitReached')) {
    return (
      'A mailbox, drive or service quota was reached. This is usually transient on the service side: retry in a few minutes, ' +
      'or ask the user to free space if the target is a mailbox or drive.'
    );
  }

  if (isRetryableStatus(ctx.status)) {
    return (
      'Microsoft Graph is temporarily unavailable and the client already retried. ' +
      'Wait a minute and try again; nothing about the request needs to change.'
    );
  }

  return undefined;
}

function guessScopes(path: string): string {
  // The table is shared with `graph_request` (see `util/scopes.ts`); only the
  // prose half is read here, so a 403 names something concrete instead of
  // "permission denied".
  for (const guess of SCOPE_GUESSES) {
    if (guess.test.test(path)) return guess.hint;
  }
  return 'the delegated scope documented for this resource';
}

function firstLine(text: string): string {
  const line = text.split('\n', 1)[0] ?? text;
  return line.length > 500 ? `${line.slice(0, 500)}…` : line;
}

function asRecord(value: unknown): Record<string, unknown> | undefined {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

function asString(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}
