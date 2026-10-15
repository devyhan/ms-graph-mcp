/**
 * .docx extraction through mammoth, which keeps document structure (headings,
 * lists, tables) where a raw XML scrape would flatten it.
 *
 * Why mammoth and not a hand parser: a .docx paragraph's text is splintered
 * across runs (`<w:r>`) with revision, hyperlink and field-code interleaving,
 * and mammoth is the maintained pure-JS library that resolves all of that to
 * HTML. Units are heading-delimited sections so structure survives the
 * no-omission contract as titles, not prose.
 */

import { stripHtml } from './truncate.js';
import { isProtectionError, protectionRefusal, type ExtractionResult } from './extract.js';

export interface DocxEngine {
  convertHtml: (buffer: Buffer) => Promise<{ html: string; warnings: string[] }>;
}

async function mammothEngine(): Promise<DocxEngine> {
  const mod = (await import('mammoth')) as unknown as {
    default?: { convertToHtml: (input: { buffer: Buffer }) => Promise<{ value: string; messages?: Array<{ message?: string }> }> };
    convertToHtml: (input: { buffer: Buffer }) => Promise<{ value: string; messages?: Array<{ message?: string }> }>;
  };
  const convertToHtml = mod.default?.convertToHtml ?? mod.convertToHtml;
  return {
    convertHtml: async (buffer: Buffer) => {
      const done = await convertToHtml({ buffer });
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

// mammoth emits h1-h6 for Heading styles and ids them nothing; split on them.
const HEADING = /<h([1-6])\b[^>]*>([\s\S]*?)<\/h\1\s*>/gi;
// Paragraph-ish blocks become one text line each inside a unit.
const BLOCK = /<(?:p|li|tr|div|blockquote|pre)\b[^>]*>([\s\S]*?)<\/(?:p|li|tr|div|blockquote|pre)\s*>/gi;

function linesOf(html: string): string[] {
  const lines: string[] = [];
  BLOCK.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = BLOCK.exec(html)) !== null) {
    const text = stripHtml(match[1] ?? '').trim();
    if (text.length > 0) lines.push(text);
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
 * giant string.
 */
export function splitSections(html: string): Array<{ title?: string; text: string }> {
  const sections: Array<{ title?: string; text: string }> = [];
  let cursor = 0;
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
    currentTitle = stripHtml(match[2] ?? '').trim() || undefined;
    cursor = match.index + match[0].length;
  }
  flush(html.slice(cursor), currentTitle);
  return sections;
}

export async function extractDocx(
  bytes: Uint8Array,
  fileName: string,
  engine?: DocxEngine,
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
  if (sections.every((section) => section.text.trim().length === 0)) {
    throw new Error(
      `"${fileName}" converted to no text — it may hold only images or tracked-change deletions. ` +
        'Tracked changes and embedded images are not extracted here.',
    );
  }
  return {
    units: sections.map((section, i) => ({
      id: `section-${i + 1}`,
      ...(section.title === undefined ? {} : { title: section.title }),
      text: section.text,
    })),
    warnings,
  };
}
