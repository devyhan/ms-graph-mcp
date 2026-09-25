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

function truncationNotice(originalChars: number, keptChars: number): string {
  return (
    `\n\n--- OUTPUT TRUNCATED ---\n` +
    `Showing the first ${keptChars} of ${originalChars} characters; the JSON above is cut mid-structure and is not valid JSON. ` +
    `Narrow the result and call again: pass $select to return fewer fields, a smaller $top, a $filter to reduce rows, ` +
    `or request one specific item by id.`
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

  const kept = full.slice(0, limit);
  return {
    text: kept + truncationNotice(originalChars, kept.length),
    truncated: true,
    originalChars,
  };
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
