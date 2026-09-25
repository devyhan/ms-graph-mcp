/**
 * Path and host validation for the generic Graph caller.
 *
 * SECURITY: `graph_request` lets a model choose the path, and `@odata.nextLink`
 * lets a *server response* choose a full URL. Both are untrusted. Every request
 * this server sends carries a bearer token, so any path or link that escapes the
 * Graph host leaks that token. These checks are the only thing preventing that;
 * they fail closed and reject anything they do not positively understand.
 */

/** Percent-encoded separators and dots — the standard way to smuggle a traversal. */
const ENCODED_TRICKS = /%(?:2e|2f|5c|00|25)/i;
const WHITESPACE_OR_CONTROL = /[\s\u0000-\u001f\u007f]/;
const VERSION = /^[A-Za-z0-9.]{1,16}$/;

/**
 * Validates a Graph-relative path and returns a cleaned copy.
 *
 * Accepts only paths like `/me/messages` or `/users/abc-123`. Absolute URLs,
 * protocol-relative references, traversal, and encoded separators all throw.
 */
export function normalizeGraphPath(path: string): string {
  if (typeof path !== 'string' || path.length === 0) {
    throw new Error('Graph path must be a non-empty string, e.g. "/me/messages".');
  }

  if (WHITESPACE_OR_CONTROL.test(path)) {
    throw new Error('Graph path must not contain whitespace or control characters; percent-encode them instead.');
  }

  if (path.includes('\\')) {
    throw new Error('Graph path must not contain a backslash.');
  }

  // A ':' before the first '/' means a scheme ("https://evil.com/x", "javascript:x").
  // Colons later in the path are legitimate — OneDrive uses `/drive/root:/file.txt:/content`.
  const colon = path.indexOf(':');
  const firstSlash = path.indexOf('/');
  if (colon !== -1 && (firstSlash === -1 || colon < firstSlash)) {
    throw new Error('Graph path must be relative; absolute URLs and URI schemes are not allowed.');
  }

  if (!path.startsWith('/')) {
    throw new Error('Graph path must start with "/", e.g. "/me/messages".');
  }

  if (path.startsWith('//')) {
    throw new Error('Graph path must not start with "//" (protocol-relative references are not allowed).');
  }

  if (ENCODED_TRICKS.test(path)) {
    throw new Error('Graph path must not contain percent-encoded separators or dots (%2e, %2f, %5c, %25, %00).');
  }

  const segments = path.split('/');
  if (segments.some((segment) => segment === '..')) {
    throw new Error('Graph path must not contain a ".." segment.');
  }

  // '@' is only dangerous where it could be read as userinfo, i.e. the first
  // segment. `/users/user@contoso.com` is an everyday Graph path and stays legal.
  const firstSegment = segments[1];
  if (firstSegment !== undefined && firstSegment.includes('@')) {
    throw new Error('Graph path must not contain "@" in its first segment.');
  }

  // Collapse repeated separators; a trailing slash carries no meaning to Graph.
  let cleaned = path.replace(/\/{2,}/g, '/');
  if (cleaned.length > 1 && cleaned.endsWith('/')) {
    cleaned = cleaned.slice(0, -1);
  }

  return cleaned;
}

/** Accepts either a bare hostname or a full base URL such as `https://graph.microsoft.com`. */
function expectedHostname(expectedHost: string): string {
  if (typeof expectedHost !== 'string' || expectedHost.trim().length === 0) {
    throw new Error('Expected Graph host must be a non-empty string.');
  }
  const value = expectedHost.trim();
  if (value.includes('://')) {
    try {
      return new URL(value).hostname.toLowerCase();
    } catch {
      throw new Error(`Expected Graph host is not a valid URL: ${JSON.stringify(expectedHost)}`);
    }
  }
  return value.replace(/\/+$/, '').toLowerCase();
}

/**
 * Throws unless `url` is HTTPS on exactly the expected host. Applied to every
 * outgoing request and to every `@odata.nextLink` / `@odata.deltaLink` before it
 * is followed, because a redirected link would otherwise receive the token.
 */
export function assertSameGraphHost(url: URL, expectedHost: string): void {
  const expected = expectedHostname(expectedHost);

  if (url.protocol !== 'https:') {
    throw new Error(`Refusing to call ${url.protocol}//${url.host}: Microsoft Graph requests must use https.`);
  }

  if (url.hostname.toLowerCase() !== expected) {
    throw new Error(`Refusing to call host ${JSON.stringify(url.hostname)}: expected ${JSON.stringify(expected)}.`);
  }

  if (url.username !== '' || url.password !== '') {
    throw new Error('Refusing to call a URL carrying embedded credentials.');
  }

  if (url.port !== '' && url.port !== '443') {
    throw new Error(`Refusing to call ${url.hostname} on port ${url.port}: only the default https port is allowed.`);
  }
}

/**
 * Builds the absolute Graph URL for a version and relative path.
 *
 * The URL is assembled through the URL API rather than string concatenation, so
 * a hostile `path` cannot rewrite the origin, and the result is re-checked
 * against the base host before it is returned.
 */
export function resolveGraphUrl(base: string, version: string, path: string): URL {
  if (!VERSION.test(version)) {
    throw new Error(`Invalid Graph version: ${JSON.stringify(version)}. Expected "v1.0" or "beta".`);
  }

  const normalized = normalizeGraphPath(path);

  let url: URL;
  try {
    url = new URL(base);
  } catch {
    throw new Error(`Invalid Graph base URL: ${JSON.stringify(base)}`);
  }

  url.pathname = `/${version}${normalized}`;
  url.search = '';
  url.hash = '';

  assertSameGraphHost(url, base);
  return url;
}
