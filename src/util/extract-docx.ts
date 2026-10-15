/**
 * .docx extraction through mammoth, which keeps document structure (headings,
 * lists, tables) where a raw XML scrape would flatten it.
 *
 * Why mammoth and not a hand parser: a .docx paragraph's text is splintered
 * across runs (`<w:r>`) with revision, hyperlink and field-code interleaving,
 * and mammoth is the maintained pure-JS library that resolves all of that to
 * HTML. Units are heading-delimited sections so structure survives the
 * no-omission contract as titles, not prose; each title carries its heading
 * path (`Part > Chapter`) so the manifest outline needs no text fetch.
 *
 * What mammoth does not cover rides a second jszip + fast-xml-parser pass:
 * headers/footers (`word/header*.xml`, `word/footer*.xml`) append as marked
 * units, and tracked revisions (`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` in
 * `word/document.xml`) tag insertions inline as `[+...+]` while deletions —
 * absent from the body by definition — collect into a `tracked-changes` unit
 * as `[-...-]`. Comments ride mammoth itself, but only with a nudge: mammoth
 * 1.13.0 IGNORES comment references by default, and the documented switch is
 * a style mapping, so the engine passes `styleMap: ['comment-reference =>
 * sup']` (option name verified against the installed package's typings) and
 * the appended `<dl>` lands in the last unit.
 */

import { stripHtml } from './truncate.js';
import {
  isProtectionError,
  protectionRefusal,
  toMarkdownTable,
  type ExtractionResult,
} from './extract.js';

export interface DocxEngine {
  convertHtml: (buffer: Buffer, options?: Record<string, unknown>) => Promise<{ html: string; warnings: string[] }>;
}

/** Minimal zip surface for the header/footer/revision pass; jszip behind it. */
export interface DocxZip {
  listParts: () => Promise<string[]>;
  readText: (path: string) => Promise<string | null>;
}

export interface DocxZipDeps {
  openZip: (bytes: Uint8Array) => Promise<DocxZip>;
}

type XmlNode = string | { [key: string]: XmlNode | XmlNode[] | undefined } | XmlNode[];

async function mammothEngine(): Promise<DocxEngine> {
  const mod = (await import('mammoth')) as unknown as {
    default?: {
      convertToHtml: (
        input: { buffer: Buffer },
        options?: Record<string, unknown>,
      ) => Promise<{ value: string; messages?: Array<{ message?: string }> }>;
    };
    convertToHtml: (
      input: { buffer: Buffer },
      options?: Record<string, unknown>,
    ) => Promise<{ value: string; messages?: Array<{ message?: string }> }>;
  };
  const convertToHtml = mod.default?.convertToHtml ?? mod.convertToHtml;
  return {
    convertHtml: async (buffer: Buffer, options?: Record<string, unknown>) => {
      // Without this mapping mammoth drops every comment reference silently;
      // with it, referenced comments append to the end of the HTML.
      const extra = options?.['styleMap'];
      const callerMaps =
        typeof extra === 'string'
          ? [extra]
          : Array.isArray(extra)
            ? extra.filter((entry): entry is string => typeof entry === 'string')
            : [];
      const done = await convertToHtml({ buffer }, { ...options, styleMap: [...callerMaps, 'comment-reference => sup'] });
      return {
        html: done.value,
        warnings: (done.messages ?? [])
          .map((entry) => entry.message)
          .filter((text): text is string => typeof text === 'string' && text.length > 0)
          .slice(0, 5),
      };
    },
  };
}

async function defaultZipDeps(): Promise<
  DocxZipDeps & { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> }
> {
  const zipMod = (await import('jszip')) as unknown as {
    default?: { loadAsync: (bytes: Uint8Array) => Promise<DocxZipFile> };
    loadAsync: (bytes: Uint8Array) => Promise<DocxZipFile>;
  };
  interface DocxZipFile {
    files: Record<string, unknown>;
    file: (path: string) => { async: (kind: 'string') => Promise<string> } | null;
  }
  const loadAsync = zipMod.default?.loadAsync ?? zipMod.loadAsync;
  const xmlMod = (await import('fast-xml-parser')) as unknown as {
    default?: {
      XMLParser: new (
        options: Record<string, unknown>,
      ) => { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> };
    };
    XMLParser: new (
      options: Record<string, unknown>,
    ) => { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> };
  };
  const XMLParser = xmlMod.default?.XMLParser ?? xmlMod.XMLParser;
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', trimValues: false });
  return {
    openZip: async (bytes: Uint8Array) => {
      const zip = await loadAsync(bytes);
      return {
        listParts: async () => Object.keys(zip.files),
        readText: async (path: string) => {
          const entry = zip.file(path);
          if (entry === null) return null;
          return entry.async('string');
        },
      };
    },
    parse: (xml: string) => parser.parse(xml),
  };
}

// mammoth emits h1-h6 for Heading styles and ids them nothing; split on them.
const HEADING = /<h([1-6])\b[^>]*>([\s\S]*?)<\/h\1\s*>/gi;
// Paragraph-ish blocks become one text line each inside a unit. `dt`/`dd` are
// the comment labels and bodies mammoth appends inside its trailing `<dl>`.
const BLOCK = /<(?:p|li|tr|div|blockquote|pre|dt|dd)\b[^>]*>([\s\S]*?)<\/(?:p|li|tr|div|blockquote|pre|dt|dd)\s*>/gi;
const TABLE = /<table\b[^>]*>([\s\S]*?)<\/table\s*>/gi;
const TABLE_ROW = /<tr\b[^>]*>([\s\S]*?)<\/tr\s*>/gi;
const TABLE_CELL = /<(?:td|th)\b[^>]*>([\s\S]*?)<\/(?:td|th)\s*>/gi;

/** One HTML table as Markdown, sharing the shape with the xlsx/pptx units. */
function markdownTable(tableInner: string): string | null {
  const rows: string[][] = [];
  TABLE_ROW.lastIndex = 0;
  let row: RegExpExecArray | null;
  while ((row = TABLE_ROW.exec(tableInner)) !== null) {
    const cells: string[] = [];
    TABLE_CELL.lastIndex = 0;
    let cell: RegExpExecArray | null;
    while ((cell = TABLE_CELL.exec(row[1] ?? '')) !== null) {
      cells.push(stripHtml(cell[1] ?? '').trim());
    }
    if (cells.length > 0) rows.push(cells);
  }
  if (rows.length === 0) return null;
  const [header, ...rest] = rows as [string[], ...string[][]];
  return toMarkdownTable(header, rest);
}

/** Tables become fenced md-table blocks before block-splitting, so `tr` handling never flattens them. */
function tablesToMarkdown(html: string): string {
  TABLE.lastIndex = 0;
  return html.replace(TABLE, (_whole, inner: string) => {
    const table = markdownTable(inner ?? '');
    return table === null ? '' : `\n<div data-md-table>\n${table}\n</div>\n`;
  });
}

function linesOf(html: string): string[] {
  const lines: string[] = [];
  // md-table blocks pass through verbatim as one multi-line entry each.
  const chunks = tablesToMarkdown(html).split(/<div data-md-table>\s*([\s\S]*?)\s*<\/div\s*>/gi);
  for (let i = 0; i < chunks.length; i += 1) {
    const chunk = chunks[i] ?? '';
    if (i % 2 === 1) {
      const table = chunk.trim();
      if (table.length > 0) lines.push(table);
      continue;
    }
    BLOCK.lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = BLOCK.exec(chunk)) !== null) {
      const text = stripHtml(match[1] ?? '').trim();
      if (text.length > 0) lines.push(text);
    }
  }
  // No block tags (a fragment, a single run): the whole chunk is one line.
  if (lines.length === 0) {
    const text = stripHtml(html).trim();
    if (text.length > 0) lines.push(text);
  }
  return lines;
}

/**
 * Splits converted HTML into heading-delimited sections. A document with no
 * headings becomes paragraph chunks of ~6000 characters so no unit is one
 * giant string. Each titled section carries its heading PATH
 * (`Part > Chapter`), not just the leaf, so the manifest titles read as an
 * outline without fetching text.
 */
export function splitSections(html: string): Array<{ title?: string; text: string }> {
  const sections: Array<{ title?: string; text: string }> = [];
  let cursor = 0;
  const stack: Array<{ level: number; text: string }> = [];
  let currentTitle: string | undefined;
  HEADING.lastIndex = 0;
  let match: RegExpExecArray | null;

  const flush = (chunk: string, title: string | undefined): void => {
    const lines = linesOf(chunk);
    if (lines.length === 0 && title === undefined) return;
    if (title === undefined) {
      // Ungrouped prose: chunk so a heading-less document still pages.
      let budget = '';
      for (const line of lines) {
        if (budget.length > 0 && budget.length + line.length + 1 > 6000) {
          sections.push({ text: budget });
          budget = '';
        }
        budget = budget.length === 0 ? line : `${budget}\n${line}`;
      }
      if (budget.length > 0) sections.push({ text: budget });
      return;
    }
    sections.push({ title, text: lines.join('\n') });
  };

  while ((match = HEADING.exec(html)) !== null) {
    flush(html.slice(cursor, match.index), currentTitle);
    const level = Number(match[1]);
    const text = stripHtml(match[2] ?? '').trim();
    while (stack.length > 0 && (stack[stack.length - 1]?.level ?? 0) >= level) stack.pop();
    if (text.length > 0) {
      stack.push({ level, text });
      currentTitle = stack.map((entry) => entry.text).join(' > ');
    } else {
      currentTitle = stack.length > 0 ? stack.map((entry) => entry.text).join(' > ') : undefined;
    }
    cursor = match.index + match[0].length;
  }
  flush(html.slice(cursor), currentTitle);
  return sections;
}

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

function asRecord(node: XmlNode | undefined): Record<string, XmlNode | XmlNode[] | undefined> | undefined {
  if (typeof node === 'object' && node !== null && !Array.isArray(node)) {
    return node as Record<string, XmlNode | XmlNode[] | undefined>;
  }
  return undefined;
}

/** Every `<w:t>` string under the node, in document order. */
function textRuns(node: XmlNode | undefined, out: string[]): void {
  if (typeof node === 'string') return;
  if (Array.isArray(node)) {
    for (const entry of node) textRuns(entry, out);
    return;
  }
  const record = asRecord(node);
  if (record === undefined) return;
  for (const [key, value] of Object.entries(record)) {
    if (key === 'w:t') {
      for (const run of asArray(value)) {
        if (typeof run === 'string' && run.length > 0) out.push(run);
        else if (typeof run === 'object' && run !== null) {
          const text = (run as Record<string, unknown>)['#text'];
          if (typeof text === 'string' && text.length > 0) out.push(text);
        }
      }
    } else if (!key.startsWith('@_') && !key.startsWith('?')) {
      textRuns(value as XmlNode, out);
    }
  }
}

/** One line per `<w:p>`, runs joined; tables flatten to their paragraph lines. */
function ooxmlParagraphs(root: Record<string, XmlNode | XmlNode[] | undefined>): string[] {
  const lines: string[] = [];
  const walk = (node: XmlNode | undefined): void => {
    if (typeof node === 'string') return;
    if (Array.isArray(node)) {
      for (const entry of node) walk(entry);
      return;
    }
    const record = asRecord(node);
    if (record === undefined) return;
    for (const [key, value] of Object.entries(record)) {
      if (key === 'w:p') {
        for (const para of asArray(value)) {
          const runs: string[] = [];
          textRuns(para, runs);
          const line = runs.join('').replace(/\s+/g, ' ').trim();
          if (line.length > 0) lines.push(line);
        }
      } else if (!key.startsWith('@_') && !key.startsWith('?')) {
        walk(value as XmlNode);
      }
    }
  };
  walk(root);
  return lines;
}

export interface DocxRevision {
  kind: 'ins' | 'del';
  text: string;
}

/**
 * Tracked revisions in document order. `w:ins`/`w:moveTo` text is an
 * insertion, `w:del`/`w:moveFrom` a deletion; range markers carry no text and
 * are ignored. Nested opposite runs take the outer kind — Word nests them
 * only when a revision is itself revised.
 */
export function collectRevisions(root: Record<string, XmlNode | XmlNode[] | undefined>): DocxRevision[] {
  const found: DocxRevision[] = [];
  const walk = (node: XmlNode | undefined): void => {
    if (typeof node === 'string') return;
    if (Array.isArray(node)) {
      for (const entry of node) walk(entry);
      return;
    }
    const record = asRecord(node);
    if (record === undefined) return;
    for (const [key, value] of Object.entries(record)) {
      if (key === 'w:ins' || key === 'w:moveTo') {
        for (const run of asArray(value)) {
          const runs: string[] = [];
          textRuns(run, runs);
          const text = runs.join('').replace(/\s+/g, ' ').trim();
          if (text.length > 0) found.push({ kind: 'ins', text });
        }
      } else if (key === 'w:del' || key === 'w:moveFrom') {
        for (const run of asArray(value)) {
          const runs: string[] = [];
          textRuns(run, runs);
          const text = runs.join('').replace(/\s+/g, ' ').trim();
          if (text.length > 0) found.push({ kind: 'del', text });
        }
      } else if (!key.startsWith('@_') && !key.startsWith('?')) {
        walk(value as XmlNode);
      }
    }
  };
  walk(root);
  return found;
}

const HEADER_FOOTER = /^word\/(header|footer)(\d+)\.xml$/;

/**
 * Tags insertions inline (`[+...+]`, first occurrence in document order) and
 * returns the deletion lines for the marked unit. Insertions missing from the
 * body — mammoth dropped the paragraph — join the deletions so nothing is
 * silently lost.
 */
function applyRevisions(
  sections: Array<{ title?: string; text: string }>,
  revisions: DocxRevision[],
): string[] {
  const leftover: string[] = [];
  for (const revision of revisions) {
    if (revision.kind === 'del') {
      leftover.push(`[-${revision.text}-]`);
      continue;
    }
    let tagged = false;
    for (const section of sections) {
      const at = section.text.indexOf(revision.text);
      if (at !== -1) {
        section.text =
          `${section.text.slice(0, at)}[+${revision.text}+]${section.text.slice(at + revision.text.length)}`;
        tagged = true;
        break;
      }
    }
    if (!tagged) leftover.push(`[+${revision.text}+]`);
  }
  return leftover;
}

export async function extractDocx(
  bytes: Uint8Array,
  fileName: string,
  engine?: DocxEngine,
  zipDeps?: DocxZipDeps,
): Promise<ExtractionResult> {
  const impl = engine ?? (await mammothEngine());
  let html: string;
  let warnings: string[];
  try {
    const done = await impl.convertHtml(Buffer.from(bytes));
    html = done.html;
    warnings = done.warnings;
  } catch (error) {
    if (isProtectionError(error)) throw protectionRefusal(fileName);
    throw new Error(`"${fileName}" could not be converted: ${error instanceof Error ? error.message : String(error)}`);
  }
  const sections = splitSections(html);

  // The second pass is best-effort: stub engines in tests feed bytes that are
  // not a zip at all, and a docx without headers/revisions simply has none.
  try {
    const deps = zipDeps ?? (await defaultZipDeps());
    const parse =
      'parse' in deps && typeof (deps as { parse?: unknown }).parse === 'function'
        ? (deps as unknown as { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> }).parse
        : (await defaultZipDeps()).parse;
    const zip = await deps.openZip(bytes);
    const documentXml = await zip.readText('word/document.xml');
    if (documentXml !== null) {
      const leftover = applyRevisions(sections, collectRevisions(parse(documentXml)));
      if (leftover.length > 0) {
        sections.push({ title: 'Tracked changes', text: leftover.join('\n') });
      }
    }
    const parts = (await zip.listParts())
      .map((path) => ({ path, match: HEADER_FOOTER.exec(path) }))
      .filter((entry): entry is { path: string; match: RegExpExecArray } => entry.match !== null)
      .sort((a, b) => {
        const kindOrder = (a.match[1] === 'header' ? 0 : 1) - (b.match[1] === 'header' ? 0 : 1);
        return kindOrder !== 0 ? kindOrder : Number(a.match[2]) - Number(b.match[2]);
      });
    for (const { path, match } of parts) {
      const xml = await zip.readText(path);
      if (xml === null) continue;
      const lines = ooxmlParagraphs(parse(xml));
      if (lines.length === 0) continue;
      const kind = match[1] === 'header' ? 'Header' : 'Footer';
      sections.push({ title: `${kind} ${match[2]}`, text: lines.join('\n') });
    }
  } catch {
    // Not a zip, or no such parts: the mammoth body stands alone.
  }

  if (sections.every((section) => section.text.trim().length === 0)) {
    throw new Error(
      `"${fileName}" converted to no text — it may hold only images or tracked-change deletions. ` +
        'Embedded images are not extracted here.',
    );
  }
  let bodySeen = 0;
  return {
    units: sections.map((section) => {
      // Marked units keep stable greppable ids; body sections stay 1-based.
      const headerFooter =
        section.title === undefined ? null : /^(Header|Footer) (\d+)$/.exec(section.title);
      let id: string;
      if (headerFooter !== null) {
        id = `${headerFooter[1]?.toLowerCase()}-${headerFooter[2]}`;
      } else if (section.title === 'Tracked changes') {
        id = 'tracked-changes';
      } else {
        bodySeen += 1;
        id = `section-${bodySeen}`;
      }
      return {
        id,
        ...(section.title === undefined ? {} : { title: section.title }),
        text: section.text,
      };
    }),
    warnings,
  };
}
