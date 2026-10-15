/**
 * Compact TOON-style encoding for tool results.
 *
 * TOON (Token-Oriented Object Notation, spec v4.4, MIT, by Johann Schopplich)
 * writes uniform arrays as one header plus bare data rows instead of
 * repeating every key on every row:
 *
 *   JSON  {"users":[{"id":1,"name":"Ada"},{"id":2,"name":"Bob"}]}
 *   TOON  users[2]{id,name}:
 *           1,Ada
 *           2,Bob
 *
 * This module implements the subset of the spec that tool results need:
 * objects, inline/tabular/list arrays with nesting, and scalar quoting with
 * the spec's escape set. It deliberately omits keyed-tabular form (a dict of
 * uniform objects stays nested objects), explicit delimiter symbols (comma
 * always), and tab indentation. Anything `encodeToon` emits, `decodeToon`
 * reads back exactly; foreign TOON using the omitted forms is rejected
 * rather than misread.
 *
 * Like `stringifySafe` in `./truncate.ts`, the encoder survives hostile
 * values — cycles become `"[Circular]"`, BigInts their decimal text, Errors
 * `{name, message}` — so switching formats cannot turn a payload that used
 * to serialise into a throw.
 */

/** Raised when TOON text is malformed. Never raised by the encoder. */
export class ToonError extends Error {
  override readonly name = 'ToonError';
}

/** A value that fits on one line: no arrays, no objects. */
type Scalar = string | number | boolean | null;

/** JSON number grammar; anything else numeric-looking is a string. */
const NUMBER_RE = /^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$/;

/** Keys that need no quotes. Decoding accepts any unquoted key. */
const BARE_KEY_RE = /^[A-Za-z_][A-Za-z0-9_.]*$/;

/** Marks `undefined`, functions and symbols during sanitising. */
const SKIPPED: unique symbol = Symbol('toon-skipped');

function isScalar(value: unknown): value is Scalar {
  return (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean'
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * Clones `value` into plain JSON data, applying `JSON.stringify` semantics:
 * `undefined`/functions/symbols are dropped from objects and become `null`
 * in arrays, non-finite numbers become `null`, `toJSON` is honoured, and a
 * true ancestor cycle (not a value merely seen twice) becomes `"[Circular]"`.
 */
function toSafeJson(value: unknown, ancestors: object[]): unknown {
  if (value === null) return null;
  switch (typeof value) {
    case 'string':
    case 'boolean':
      return value;
    case 'number':
      return Number.isFinite(value) ? value : null;
    case 'bigint':
      return value.toString();
    case 'undefined':
    case 'function':
    case 'symbol':
      return SKIPPED;
    case 'object': {
      if (value instanceof Error) {
        return { name: value.name, message: value.message };
      }
      const toJson = (value as { toJSON?: unknown }).toJSON;
      if (typeof toJson === 'function') {
        return toSafeJson((toJson as () => unknown).call(value), ancestors);
      }
      if (ancestors.includes(value)) return '[Circular]';
      ancestors.push(value);
      try {
        if (Array.isArray(value)) {
          const out: unknown[] = [];
          for (const item of value) {
            const safe = toSafeJson(item, ancestors);
            out.push(safe === SKIPPED ? null : safe);
          }
          return out;
        }
        const out: Record<string, unknown> = {};
        for (const [key, item] of Object.entries(value)) {
          const safe = toSafeJson(item, ancestors);
          if (safe !== SKIPPED) out[key] = safe;
        }
        return out;
      } finally {
        ancestors.pop();
      }
    }
  }
}

/**
 * True when a string must be quoted: empty, edge whitespace, shaped like a
 * non-string scalar, or carrying a character with a structural meaning
 * (delimiters, brackets, quotes, colon, backslash, controls) or a leading
 * `-`/`#` that would read as a list item or a comment.
 */
function needsQuotes(text: string): boolean {
  if (text === '') return true;
  if (text === 'true' || text === 'false' || text === 'null') return true;
  if (NUMBER_RE.test(text)) return true;
  const first = text[0];
  const last = text[text.length - 1];
  if (first === ' ' || first === '\t' || last === ' ' || last === '\t') return true;
  if (first === '-' || first === '#') return true;
  for (const ch of text) {
    if (ch.charCodeAt(0) <= 0x1f) return true;
    if (
      ch === '"' ||
      ch === '\\' ||
      ch === ':' ||
      ch === ',' ||
      ch === '|' ||
      ch === '[' ||
      ch === ']' ||
      ch === '{' ||
      ch === '}'
    ) {
      return true;
    }
  }
  return false;
}

/** Quotes a string with the spec's escape set (`\u` in lowercase hex). */
function quoteString(text: string): string {
  let out = '"';
  for (const ch of text) {
    switch (ch) {
      case '"':
        out += '\\"';
        break;
      case '\\':
        out += '\\\\';
        break;
      case '\n':
        out += '\\n';
        break;
      case '\r':
        out += '\\r';
        break;
      case '\t':
        out += '\\t';
        break;
      default: {
        const code = ch.charCodeAt(0);
        out += code <= 0x1f ? '\\u' + code.toString(16).padStart(4, '0') : ch;
      }
    }
  }
  return out + '"';
}

function encodeScalar(value: Scalar): string {
  if (value === null) return 'null';
  if (value === true) return 'true';
  if (value === false) return 'false';
  if (typeof value === 'number') return JSON.stringify(value) ?? 'null';
  return needsQuotes(value) ? quoteString(value) : value;
}

function encodeKey(key: string): string {
  return BARE_KEY_RE.test(key) ? key : quoteString(key);
}

function indent(depth: number): string {
  return '  '.repeat(depth);
}

/**
 * The shared field list when `rows` is a non-empty array of flat objects
 * with identical keys in identical order, else null. Nested values stay out:
 * a cell is always one scalar, so anything deeper falls back to list form.
 */
function tabularFields(rows: unknown[]): string[] | null {
  const first = rows[0];
  if (!isRecord(first)) return null;
  const fields = Object.keys(first);
  if (fields.length === 0) return null;
  for (const row of rows) {
    if (!isRecord(row)) return null;
    const keys = Object.keys(row);
    if (keys.length !== fields.length) return null;
    for (let i = 0; i < fields.length; i += 1) {
      if (keys[i] !== fields[i]) return null;
      if (!isScalar(row[keys[i] as string])) return null;
    }
  }
  return fields;
}

function encodeObjectFields(obj: Record<string, unknown>, depth: number): string[] {
  const lines: string[] = [];
  const pad = indent(depth);
  for (const [key, value] of Object.entries(obj)) {
    const encoded = encodeKey(key);
    if (Array.isArray(value)) {
      if (value.length === 0) lines.push(`${pad}${encoded}: []`);
      else lines.push(...encodeArrayLines(encoded, value, depth));
    } else if (isRecord(value)) {
      // A bare `key:` is an empty object, or the head of a nested block when
      // deeper lines follow; the decoder tells the two apart by peeking.
      lines.push(`${pad}${encoded}:`);
      if (Object.keys(value).length > 0) {
        lines.push(...encodeObjectFields(value, depth + 1));
      }
    } else {
      lines.push(`${pad}${encoded}: ${encodeScalar(value as Scalar)}`);
    }
  }
  return lines;
}

function encodeArrayLines(keyPrefix: string, rows: unknown[], depth: number): string[] {
  const pad = indent(depth);
  const head = `${pad}${keyPrefix}[${rows.length}]`;
  if (rows.every(isScalar)) {
    return [`${head}: ${(rows as Scalar[]).map(encodeScalar).join(',')}`];
  }
  const fields = tabularFields(rows);
  if (fields !== null) {
    const rowPad = indent(depth + 1);
    const lines = [`${head}{${fields.map(encodeKey).join(',')}}:`];
    for (const row of rows as Array<Record<string, unknown>>) {
      lines.push(rowPad + fields.map((field) => encodeScalar(row[field] as Scalar)).join(','));
    }
    return lines;
  }
  const lines = [`${head}:`];
  for (const item of rows) lines.push(...encodeListItem(item, depth + 1));
  return lines;
}

function encodeListItem(item: unknown, depth: number): string[] {
  const pad = indent(depth);
  if (isScalar(item)) return [`${pad}- ${encodeScalar(item)}`];
  if (Array.isArray(item)) {
    if (item.length === 0) return [`${pad}- []`];
    // An anonymous header rides on the `- ` line; its rows stay one level
    // deeper, exactly as if the header had stood alone at this depth.
    const nested = encodeArrayLines('', item, depth);
    const first = nested[0];
    if (first === undefined) return [`${pad}- []`];
    return [`${pad}- ${first.slice(pad.length)}`, ...nested.slice(1)];
  }
  if (isRecord(item)) {
    if (Object.keys(item).length === 0) return [`${pad}- {}`];
    return [`${pad}-`, ...encodeObjectFields(item, depth + 1)];
  }
  return [`${pad}- ${encodeScalar(item as Scalar)}`];
}

/**
 * Encodes any value as TOON text. Never throws for hostile input: values
 * that cannot be walked fall back to `String(value)`, mirroring
 * `stringifySafe`.
 */
export function encodeToon(value: unknown): string {
  let safe: unknown;
  try {
    safe = toSafeJson(value, []);
  } catch {
    return String(value);
  }
  if (safe === SKIPPED) return String(value);
  if (isScalar(safe)) return encodeScalar(safe);
  if (Array.isArray(safe)) {
    if (safe.length === 0) return '[]';
    return encodeArrayLines('', safe, 0).join('\n');
  }
  const record = safe as Record<string, unknown>;
  if (Object.keys(record).length === 0) return '';
  return encodeObjectFields(record, 0).join('\n');
}

// ---------------------------------------------------------------------------
// Decoding
// ---------------------------------------------------------------------------

/** One significant line: blanks and full-line comments never reach the parser. */
interface ToonLine {
  depth: number;
  content: string;
  num: number;
}

function splitLines(text: string): ToonLine[] {
  if (text === '') return [];
  const lines: ToonLine[] = [];
  const raw = text.split('\n');
  for (let i = 0; i < raw.length; i += 1) {
    const num = i + 1;
    const line = raw[i];
    if (line === undefined) continue;
    const stripped = line.endsWith('\r') ? line.slice(0, -1) : line;
    if (stripped.trim() === '') continue;
    const match = stripped.match(/^ */);
    const pad = match === null || match[0] === undefined ? '' : match[0];
    if (/^ *\t/.test(stripped)) {
      throw new ToonError(`line ${num}: tabs are not valid indentation`);
    }
    if (pad.length % 2 !== 0) {
      throw new ToonError(`line ${num}: indentation must be a multiple of two spaces`);
    }
    const content = stripped.slice(pad.length);
    if (content.startsWith('#')) continue;
    lines.push({ depth: pad.length / 2, content, num });
  }
  return lines;
}

/** Trims ASCII spaces only: tabs and other whitespace are significant. */
function trimSpaces(text: string): string {
  return text.replace(/^ +| +$/g, '');
}

/** Parses a `"`-quoted string starting at `start`; returns it and the end offset. */
function parseQuoted(text: string, start: number, num: number): { value: string; end: number } {
  let out = '';
  let i = start + 1;
  while (i < text.length) {
    const ch = text[i];
    if (ch === '"') return { value: out, end: i + 1 };
    if (ch !== '\\') {
      out += ch ?? '';
      i += 1;
      continue;
    }
    const esc = text[i + 1];
    if (esc === 'n') out += '\n';
    else if (esc === 'r') out += '\r';
    else if (esc === 't') out += '\t';
    else if (esc === '\\') out += '\\';
    else if (esc === '"') out += '"';
    else if (esc === 'u') {
      const hex = text.slice(i + 2, i + 6);
      if (!/^[0-9a-fA-F]{4}$/.test(hex)) {
        throw new ToonError(`line ${num}: bad \\u escape`);
      }
      const code = Number.parseInt(hex, 16);
      if (code >= 0xd800 && code <= 0xdbff) {
        const low = text.slice(i + 6, i + 12);
        if (!/^\\u[0-9a-fA-F]{4}$/.test(low)) {
          throw new ToonError(`line ${num}: unpaired surrogate in \\u escape`);
        }
        const lowCode = Number.parseInt(low.slice(2), 16);
        if (lowCode < 0xdc00 || lowCode > 0xdfff) {
          throw new ToonError(`line ${num}: unpaired surrogate in \\u escape`);
        }
        out += String.fromCharCode(code, lowCode);
        i += 12;
        continue;
      }
      if (code >= 0xdc00 && code <= 0xdfff) {
        throw new ToonError(`line ${num}: unpaired surrogate in \\u escape`);
      }
      out += String.fromCharCode(code);
      i += 6;
      continue;
    } else {
      throw new ToonError(`line ${num}: unsupported escape "\\${esc ?? ''}"`);
    }
    i += 2;
  }
  throw new ToonError(`line ${num}: unterminated string`);
}

/** Splits on commas, ignoring commas inside quoted strings. */
function splitCells(text: string, num: number): string[] {
  const cells: string[] = [];
  let current = '';
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (quoted) {
      current += ch ?? '';
      if (ch === '\\') {
        i += 1;
        const next = text[i];
        if (next === undefined) throw new ToonError(`line ${num}: dangling escape`);
        current += next;
      } else if (ch === '"') {
        quoted = false;
      }
      continue;
    }
    if (ch === '"') {
      quoted = true;
      current += ch;
    } else if (ch === ',') {
      cells.push(current);
      current = '';
    } else {
      current += ch ?? '';
    }
  }
  if (quoted) throw new ToonError(`line ${num}: unterminated string`);
  cells.push(current);
  return cells;
}

function parseScalar(token: string, num: number): Scalar {
  const text = trimSpaces(token);
  if (text === '') return '';
  if (text.startsWith('"')) {
    const { value, end } = parseQuoted(text, 0, num);
    if (trimSpaces(text.slice(end)) !== '') {
      throw new ToonError(`line ${num}: unexpected content after string`);
    }
    return value;
  }
  if (text === 'true') return true;
  if (text === 'false') return false;
  if (text === 'null') return null;
  if (NUMBER_RE.test(text)) return Number(text);
  return text;
}

/** A field name from a `{...}` header: quoted strings stay strings, anything else is taken literally. */
function parseFieldToken(token: string, num: number): string {
  const text = trimSpaces(token);
  if (!text.startsWith('"')) return text;
  const { value, end } = parseQuoted(text, 0, num);
  if (trimSpaces(text.slice(end)) !== '') {
    throw new ToonError(`line ${num}: unexpected content after field name`);
  }
  return value;
}

/** Splits a field line into its key and the remainder starting at `:` or `[`. */
function parseKey(content: string, num: number): { key: string; rest: string } {
  if (content.startsWith('"')) {
    const { value, end } = parseQuoted(content, 0, num);
    const rest = content.slice(end);
    if (!rest.startsWith(':') && !rest.startsWith('[')) {
      throw new ToonError(`line ${num}: expected ':' or '[' after key`);
    }
    return { key: value, rest };
  }
  const idx = content.search(/[:[]/);
  if (idx === -1) throw new ToonError(`line ${num}: expected a key followed by ':' or '['`);
  return { key: content.slice(0, idx), rest: content.slice(idx) };
}

interface ArrayHeader {
  count: number;
  fields: string[] | null;
  /** Content after the header colon, or null for the block forms. */
  inline: string | null;
}

/** Parses `[N]`, an optional `{fields}` span, and the `:` plus its remainder. */
function parseArrayHeader(source: string, num: number): ArrayHeader {
  const close = source.indexOf(']');
  if (!source.startsWith('[') || close === -1) {
    throw new ToonError(`line ${num}: malformed array header`);
  }
  const countToken = source.slice(1, close);
  if (!/^(0|[1-9]\d*)$/.test(countToken)) {
    throw new ToonError(`line ${num}: array length must be a whole number without leading zeros`);
  }
  let rest = source.slice(close + 1);
  let fields: string[] | null = null;
  if (rest.startsWith('{')) {
    const end = closingBrace(rest, num);
    const inner = trimSpaces(rest.slice(1, end));
    fields = inner === '' ? [] : splitCells(inner, num).map((cell) => parseFieldToken(cell, num));
    rest = rest.slice(end + 1);
  }
  if (!rest.startsWith(':')) {
    throw new ToonError(`line ${num}: array header must end with ':'`);
  }
  const inline = trimSpaces(rest.slice(1));
  return { count: Number(countToken), fields, inline: inline === '' ? null : inline };
}

/** Offset of the first unquoted `}`; quoted keys may carry braces. */
function closingBrace(text: string, num: number): number {
  let quoted = false;
  for (let i = 1; i < text.length; i += 1) {
    const ch = text[i];
    if (quoted) {
      if (ch === '\\') i += 1;
      else if (ch === '"') quoted = false;
      continue;
    }
    if (ch === '"') quoted = true;
    else if (ch === '}') return i;
  }
  throw new ToonError(`line ${num}: unterminated field list`);
}

interface Parsed<T> {
  value: T;
  /** Offset of the first line not consumed. */
  next: number;
}

function parseObjectLines(lines: ToonLine[], index: number, depth: number): Parsed<Record<string, unknown>> {
  const obj: Record<string, unknown> = {};
  let idx = index;
  while (idx < lines.length) {
    const line = lines[idx];
    if (line === undefined || line.depth < depth) break;
    if (line.depth > depth) {
      throw new ToonError(`line ${line.num}: unexpected indentation`);
    }
    // A list item or an anonymous header at this depth belongs to the
    // enclosing array, not to this object.
    if (line.content === '-' || line.content.startsWith('- ') || line.content.startsWith('[')) {
      break;
    }
    const { key, rest } = parseKey(line.content, line.num);
    if (rest.startsWith('[')) {
      const parsed = parseArrayAfterHeader(lines, idx, depth, rest, line.num);
      obj[key] = parsed.value;
      idx = parsed.next;
      continue;
    }
    const after = trimSpaces(rest.slice(1));
    if (after === '[]') {
      obj[key] = [];
      idx += 1;
    } else if (after === '{}') {
      obj[key] = {};
      idx += 1;
    } else if (after !== '') {
      obj[key] = parseScalar(after, line.num);
      idx += 1;
    } else {
      const next = lines[idx + 1];
      if (next !== undefined && next.depth > depth) {
        const nested = parseObjectLines(lines, idx + 1, depth + 1);
        obj[key] = nested.value;
        idx = nested.next;
      } else {
        obj[key] = {};
        idx += 1;
      }
    }
  }
  return { value: obj, next: idx };
}

/**
 * Parses the array whose header sits at `lines[index]` (at `depth`), with
 * the header source passed separately because `- [N]...` items carry it
 * after the dash. Counts are strict: exactly `N` rows or items follow.
 */
function parseArrayAfterHeader(
  lines: ToonLine[],
  index: number,
  depth: number,
  headerSource: string,
  num: number,
): Parsed<unknown[]> {
  const header = parseArrayHeader(headerSource, num);
  if (header.inline !== null) {
    if (header.fields !== null) {
      throw new ToonError(`line ${num}: a header with fields must not carry inline content`);
    }
    const cells = splitCells(header.inline, num);
    if (cells.length !== header.count) {
      throw new ToonError(
        `line ${num}: header says ${header.count} items but carries ${cells.length}`,
      );
    }
    return { value: cells.map((cell) => parseScalar(cell, num)), next: index + 1 };
  }
  if (header.fields !== null) {
    const rows: Array<Record<string, unknown>> = [];
    for (let row = 0; row < header.count; row += 1) {
      const line = lines[index + 1 + row];
      if (line === undefined) {
        throw new ToonError(
          `line ${num}: header says ${header.count} rows but the document ends after ${row}`,
        );
      }
      if (line.depth !== depth + 1) {
        throw new ToonError(`line ${line.num}: expected a row one level deeper than its header`);
      }
      const cells = splitCells(line.content, line.num);
      if (cells.length !== header.fields.length) {
        throw new ToonError(
          `line ${line.num}: expected ${header.fields.length} cells but found ${cells.length}`,
        );
      }
      const record: Record<string, unknown> = {};
      header.fields.forEach((field, at) => {
        record[field] = parseScalar(cells[at] ?? '', line.num);
      });
      rows.push(record);
    }
    return { value: rows, next: index + 1 + header.count };
  }
  const items: unknown[] = [];
  let next = index + 1;
  for (let seen = 0; seen < header.count; seen += 1) {
    const line = lines[next];
    if (line === undefined) {
      throw new ToonError(
        `line ${num}: header says ${header.count} items but the document ends after ${seen}`,
      );
    }
    if (line.depth !== depth + 1) {
      throw new ToonError(`line ${line.num}: expected a list item one level deeper than its header`);
    }
    const parsed = parseListItem(lines, next, depth + 1);
    items.push(parsed.value);
    next = parsed.next;
  }
  return { value: items, next };
}

function parseListItem(lines: ToonLine[], index: number, depth: number): Parsed<unknown> {
  const line = lines[index];
  if (line === undefined || line.depth !== depth) {
    throw new ToonError(`line ${line?.num ?? '?'}: expected a list item`);
  }
  if (line.content === '-') {
    const next = lines[index + 1];
    if (next === undefined || next.depth <= depth) {
      throw new ToonError(`line ${line.num}: a bare '-' needs an object body after it`);
    }
    const nested = parseObjectLines(lines, index + 1, depth + 1);
    return { value: nested.value, next: nested.next };
  }
  if (!line.content.startsWith('- ')) {
    throw new ToonError(`line ${line.num}: expected '- ' followed by an item`);
  }
  const after = trimSpaces(line.content.slice(2));
  if (after === '[]') return { value: [], next: index + 1 };
  if (after === '{}') return { value: {}, next: index + 1 };
  if (after.startsWith('[')) {
    return parseArrayAfterHeader(lines, index, depth, after, line.num);
  }
  return { value: parseScalar(after, line.num), next: index + 1 };
}

/** True when the line opens a `key:` field or a `key[N]...` header. */
function isFieldLine(content: string): boolean {
  if (content.startsWith('"')) {
    try {
      const { end } = parseQuoted(content, 0, 1);
      const rest = content.slice(end);
      return rest.startsWith(':') || rest.startsWith('[');
    } catch {
      return false;
    }
  }
  return /[:[]/.test(content);
}

/**
 * Decodes TOON text back to plain data. Accepts what `encodeToon` emits
 * plus full-line `#` comments and blank lines; anything else malformed
 * raises `ToonError` naming the line.
 */
export function decodeToon(text: string): unknown {
  const lines = splitLines(text);
  if (lines.length === 0) return {};
  const first = lines[0];
  if (first === undefined || first.depth !== 0) {
    throw new ToonError('line 1: the document must start at zero indentation');
  }
  // A truncated character-cut starts with this banner, never with data. It
  // must fail here with the reason: the banner carries a colon, so without
  // this refusal it would decode as a field and return garbage as data. No
  // false positives: the encoder quotes every leading `-`, so it never emits
  // an unquoted line starting this way.
  if (first.content.startsWith('--- OUTPUT TRUNCATED')) {
    throw new ToonError('line 1: output was truncated mid-structure and cannot be decoded');
  }
  if (first.content === '-' || first.content.startsWith('- ')) {
    throw new ToonError('line 1: a list item cannot stand without its array header');
  }
  if (first.content === '[]') {
    if (lines.length > 1) {
      const extra = lines[1];
      throw new ToonError(`line ${extra?.num ?? 2}: unexpected content after the document`);
    }
    return [];
  }
  if (first.content.startsWith('[')) {
    const parsed = parseArrayAfterHeader(lines, 0, 0, first.content, first.num);
    if (parsed.next !== lines.length) {
      const extra = lines[parsed.next];
      throw new ToonError(`line ${extra?.num ?? '?'}: unexpected content after the document`);
    }
    return parsed.value;
  }
  if (isFieldLine(first.content)) {
    const parsed = parseObjectLines(lines, 0, 0);
    if (parsed.next !== lines.length) {
      const extra = lines[parsed.next];
      throw new ToonError(`line ${extra?.num ?? '?'}: unexpected content after the document`);
    }
    return parsed.value;
  }
  if (lines.length > 1) {
    const extra = lines[1];
    throw new ToonError(`line ${extra?.num ?? 2}: unexpected content after the document`);
  }
  return parseScalar(first.content, first.num);
}
