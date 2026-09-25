/**
 * OData query construction.
 *
 * Tool arguments come from a model, so anything spliced into a `$filter` or
 * `$search` string is untrusted input. These helpers exist so callers never
 * hand-build those strings.
 */

/** Friendly names accepted by `buildQuery`, plus raw pass-through keys. */
export interface ODataParams {
  /** Fields to return, e.g. `['id','subject']` or `'id,subject'`. */
  select?: string | string[] | undefined;
  filter?: string | undefined;
  top?: number | undefined;
  skip?: number | undefined;
  orderby?: string | string[] | undefined;
  search?: string | undefined;
  expand?: string | string[] | undefined;
  count?: boolean | undefined;
  /** Anything else is passed through under its own key, e.g. `$skipToken`. */
  [key: string]: string | string[] | number | boolean | undefined;
}

const FRIENDLY_TO_ODATA: Readonly<Record<string, string>> = {
  select: '$select',
  filter: '$filter',
  top: '$top',
  skip: '$skip',
  orderby: '$orderby',
  search: '$search',
  expand: '$expand',
  count: '$count',
};

/**
 * Escapes a value for use inside an OData single-quoted literal.
 * Doubling the quote is the only escape OData defines, and it is what keeps a
 * value like `O'Brien` from terminating the literal early.
 */
export function escapeODataString(v: string): string {
  return String(v).replace(/'/g, "''");
}

/**
 * Maps friendly parameter names to their OData spelling, drops `undefined`, and
 * joins array values with commas. Unrecognised keys pass through untouched so a
 * caller can supply `$skipToken` or `$deltaToken` directly.
 */
export function buildQuery(params: ODataParams): Record<string, string | number | boolean | undefined> {
  const out: Record<string, string | number | boolean | undefined> = {};

  for (const [key, raw] of Object.entries(params)) {
    if (raw === undefined || raw === null) continue;

    const name = FRIENDLY_TO_ODATA[key] ?? key;
    const value = Array.isArray(raw)
      ? raw.filter((part): part is string => typeof part === 'string' && part.trim().length > 0)
          .map((part) => part.trim())
          .join(',')
      : raw;

    if (typeof value === 'string' && value.length === 0) continue;
    out[name] = value;
  }

  return out;
}

/**
 * Wraps a `$search` term in the double-quoted form Graph requires for mail,
 * files, and most other resources (`$search="subject:budget"`). Embedded quotes
 * and backslashes are escaped so the term cannot break out of the phrase.
 */
export function quoteSearch(term: string): string {
  const escaped = String(term).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  return `"${escaped}"`;
}

// Date, optional time, optional fractional seconds, optional Z or ±HH:MM offset.
const ISO_8601 = /^(\d{4}-\d{2}-\d{2})(?:[T ](\d{2}:\d{2}(?::\d{2}(?:\.\d{1,7})?)?)(Z|[+-]\d{2}:\d{2})?)?$/;

/**
 * Validates an ISO-8601 date or datetime and returns it in canonical `T`-separated
 * form. Rejecting anything else is what stops a model-supplied "date" from
 * carrying `$filter` syntax into a query string.
 */
export function isoDate(v: string): string {
  if (typeof v !== 'string' || v.trim().length === 0) {
    throw new Error('Expected an ISO-8601 datetime string, e.g. 2026-01-31T09:00:00Z.');
  }

  const value = v.trim();
  const match = ISO_8601.exec(value);
  if (!match) {
    throw new Error(
      `Invalid ISO-8601 datetime: ${JSON.stringify(value)}. Use YYYY-MM-DD or YYYY-MM-DDTHH:mm:ss[Z|±HH:MM], e.g. 2026-01-31T09:00:00Z.`,
    );
  }

  const date = match[1] ?? '';
  const time = match[2];
  const offset = match[3] ?? '';

  const invalid = (): never => {
    throw new Error(`Invalid ISO-8601 datetime: ${JSON.stringify(value)} is not a real date or time.`);
  };

  // Ranges are checked component by component: the regex admits shapes like
  // 2026-02-31 or 25:99, and `Date.parse` silently rolls those over rather than
  // failing, which would hand Graph a different day than the caller asked for.
  const year = Number(date.slice(0, 4));
  const month = Number(date.slice(5, 7));
  const day = Number(date.slice(8, 10));
  const probe = new Date(Date.UTC(year, month - 1, day));
  if (probe.getUTCFullYear() !== year || probe.getUTCMonth() !== month - 1 || probe.getUTCDate() !== day) {
    invalid();
  }

  if (time !== undefined) {
    const hour = Number(time.slice(0, 2));
    const minute = Number(time.slice(3, 5));
    const second = time.length > 5 ? Number(time.slice(6)) : 0;
    if (hour > 23 || minute > 59 || second >= 60) invalid();

    if (offset !== '' && offset !== 'Z') {
      const offsetHour = Number(offset.slice(1, 3));
      const offsetMinute = Number(offset.slice(4, 6));
      if (offsetHour > 14 || offsetMinute > 59) invalid();
    }
  }

  return time === undefined ? date : `${date}T${time}${offset}`;
}
