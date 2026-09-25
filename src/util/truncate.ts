/**
 * Turning arbitrary Graph payloads into bounded, model-friendly text.
 *
 * Graph collections and Outlook HTML bodies routinely run to hundreds of
 * kilobytes; handing that to a model wastes the window and often fails outright.
 */

/** Result of serialising a tool return value under a character budget. */
export interface SerializedResult {
  text: string;
  truncated: boolean;
  /** Length of the full serialisation, before any cut. */
  originalChars: number;
}

/**
 * JSON.stringify that survives the values Graph and our own code can produce:
 * cyclic references (a parsed response linked to its request), BigInt, and
 * objects with throwing getters.
 */
function stringifySafe(value: unknown): string {
  const seen = new WeakSet<object>();
  try {
    const json = JSON.stringify(
      value,
      (_key, inner: unknown) => {
        if (typeof inner === 'bigint') return inner.toString();
        if (inner instanceof Error) {
          return { name: inner.name, message: inner.message };
        }
        if (typeof inner === 'object' && inner !== null) {
          if (seen.has(inner)) return '[Circular]';
          seen.add(inner);
        }
        return inner;
      },
      2,
    );
    // `undefined`, a bare function, or a symbol stringify to `undefined`.
    return json === undefined ? String(value) : json;
  } catch {
    return String(value);
  }
}

/**
 * Printed BEFORE the fragment, not after.
 *
 * A reader meeting the payload first will try to parse it, fail somewhere deep
 * inside, and go looking for a malformed value that does not exist. Leading
 * with the warning costs one line and removes that entirely.
 */
function truncationNotice(originalChars: number, keptChars: number): string {
  return (
    `--- OUTPUT TRUNCATED: NOT VALID JSON ---\n` +
    `What follows is the first ${keptChars} of ${originalChars} characters and stops mid-structure, ` +
    `so parsing it will fail. It could not be shortened by dropping rows, which means one item is ` +
    `itself over the limit. Read it as text, or call again for less: fewer fields, a smaller page, ` +
    `a narrower filter, or one item by id.`
  );
}

/**
 * Serialises a tool result to indented JSON, cutting it to `maxChars` and
 * appending an explanation the model can act on. The returned text may exceed
 * `maxChars` by the length of that notice — the notice is the point.
 */
export function serializeResult(value: unknown, maxChars: number): SerializedResult {
  const full = stringifySafe(value);
  const originalChars = full.length;
  const limit = Number.isFinite(maxChars) && maxChars > 0 ? Math.floor(maxChars) : 0;

  if (limit === 0 || originalChars <= limit) {
    return { text: full, truncated: false, originalChars };
  }

  // Drop whole rows before resorting to cutting characters. A tool result is
  // almost always an object with one array in it, and half a row helps nobody:
  // the model parses the JSON, so a result that is still JSON — with fewer
  // items and a field saying so — is worth far more than a prefix of one.
  const trimmed = dropRowsToFit(value, limit);
  if (trimmed !== null) {
    const text = stringifySafe(trimmed);
    if (text.length <= limit) {
      return { text, truncated: true, originalChars };
    }
  }

  // Nothing row-shaped to shed, or the rows are individually too big. Fall back
  // to a character cut, which is not valid JSON — say so first rather than
  // letting the reader discover it from a parse error at some byte offset.
  const kept = full.slice(0, limit);
  return {
    text: truncationNotice(originalChars, kept.length) + '\n' + kept,
    truncated: true,
    originalChars,
  };
}

/**
 * Re-serialises `value` with its longest array shortened until the whole thing
 * fits, returning null when there is no array to shorten.
 *
 * Binary search on the row count rather than dropping one at a time: a chat
 * list can be hundreds of rows and each attempt re-serialises the object.
 */
function dropRowsToFit(value: unknown, limit: number): Record<string, unknown> | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return null;

  const source = value as Record<string, unknown>;
  let key: string | null = null;
  let rows: unknown[] = [];
  for (const [k, v] of Object.entries(source)) {
    if (Array.isArray(v) && v.length > rows.length) {
      key = k;
      rows = v;
    }
  }
  if (key === null || rows.length === 0) return null;

  const build = (count: number): Record<string, unknown> => ({
    ...source,
    [key as string]: rows.slice(0, count),
    truncated: {
      reason: `The full result was over the ${limit}-character output limit.`,
      field: key,
      returned: count,
      total: rows.length,
      advice:
        'These are the first rows only. Ask for fewer fields or a smaller page, ' +
        'filter the query, or page through with the cursor if one is present.',
    },
  });

  // The empty-list case still carries the note, so a caller always learns why.
  if (stringifySafe(build(0)).length > limit) return null;

  let low = 0;
  let high = rows.length;
  while (low < high) {
    const mid = Math.ceil((low + high) / 2);
    if (stringifySafe(build(mid)).length <= limit) low = mid;
    else high = mid - 1;
  }
  return build(low);
}

/** Caps a free-text field (mail body, file preview) with a visible marker. */
export function truncateText(s: string, max: number): string {
  if (typeof s !== 'string') return '';
  const limit = Number.isFinite(max) && max > 0 ? Math.floor(max) : 0;
  if (limit === 0 || s.length <= limit) return s;
  const dropped = s.length - limit;
  return `${s.slice(0, limit)}… [truncated, ${dropped} more characters]`;
}

const BLOCK_TAG = /<\/?(?:p|div|br|tr|li|h[1-6]|blockquote|table|section|article)\b[^>]*>/gi;
const DROPPED_ELEMENT = /<(script|style|head|noscript)\b[\s\S]*?<\/\1\s*>/gi;
const ANY_TAG = /<[^>]*>/g;

/**
 * Cheap HTML-to-text for Outlook bodies. Not a parser and not a sanitiser: the
 * output is plain text for the model, never re-rendered as markup.
 */
export function stripHtml(html: string): string {
  if (typeof html !== 'string' || html.length === 0) return '';

  let text = html
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(DROPPED_ELEMENT, ' ')
    // Keep paragraph structure before flattening everything else.
    .replace(BLOCK_TAG, '\n')
    .replace(ANY_TAG, ' ');

  // Entities are decoded only after tags are gone, so an encoded `&lt;script&gt;`
  // cannot turn back into markup.
  text = text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#0*39;/g, "'")
    .replace(/&apos;/gi, "'")
    .replace(/&amp;/gi, '&');

  return text
    .replace(/[ \t\r\f\v]+/g, ' ')
    .replace(/ ?\n ?/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
