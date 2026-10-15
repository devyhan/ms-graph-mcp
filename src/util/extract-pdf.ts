/**
 * PDF extraction through pdfjs-dist: one unit per page, text items grouped
 * into rows by y position and ordered by x within the row.
 *
 * Why pdfjs-dist: it is Mozilla's maintained pure-JS PDF renderer, and its
 * `getTextContent` is the only pure-JS route to positioned text plus
 * tagged-PDF marked content (`includeMarkedContent`, which keeps role hints
 * such as headings where the producer tagged them). Row grouping uses the
 * `hasEOL` flags where the content stream sets them and falls back to y
 * proximity otherwise; x ordering inside a row follows the stream order after
 * a y/x sort, which matches reading order for single-column text and degrades
 * gracefully (interleaved columns) rather than crashing for the rest.
 */

import {
  MAX_PDF_PAGES,
  isProtectionError,
  protectionRefusal,
  scannedPdfRefusal,
  type ExtractionResult,
} from './extract.js';

export interface PdfTextItem {
  str: string;
  /** pdfjs transform [a, b, c, d, x, y]; only x/y are read. */
  x: number;
  y: number;
  hasEOL?: boolean;
}

export interface PdfEngine {
  pageCount: (bytes: Uint8Array) => Promise<number>;
  pageItems: (bytes: Uint8Array, pageNumber: number) => Promise<PdfTextItem[]>;
}

/** Items more than this far apart vertically start a new row. */
const ROW_Y_TOLERANCE = 3;

export function rowsOf(items: PdfTextItem[]): string[] {
  // Reading order first: top of the page down, left to right within a row.
  const sorted = [...items]
    .filter((item) => item.str.length > 0)
    .sort((a, b) => (b.y !== a.y && Math.abs(b.y - a.y) > ROW_Y_TOLERANCE ? b.y - a.y : a.x - b.x));
  const rows: string[] = [];
  let current: string[] = [];
  let rowY: number | undefined;
  const flush = (): void => {
    const line = current.join(' ').replace(/\s+/g, ' ').trim();
    if (line.length > 0) rows.push(line);
    current = [];
    rowY = undefined;
  };
  for (const item of sorted) {
    if (rowY !== undefined && Math.abs(item.y - rowY) > ROW_Y_TOLERANCE) flush();
    if (rowY === undefined) rowY = item.y;
    current.push(item.str);
    if (item.hasEOL === true) flush();
  }
  flush();
  return rows;
}

async function pdfjsEngine(): Promise<PdfEngine> {
  const mod = (await import('pdfjs-dist/legacy/build/pdf.mjs')) as unknown as {
    default?: PdfjsLib;
    getDocument: PdfjsLib['getDocument'];
  };
  interface PdfjsPage {
    getTextContent: (options?: Record<string, unknown>) => Promise<{
      items: Array<{
        str?: unknown;
        transform?: unknown;
        hasEOL?: unknown;
      }>;
    }>;
  }
  interface PdfjsDoc {
    numPages: number;
    getPage: (n: number) => Promise<PdfjsPage>;
    destroy: () => Promise<void>;
  }
  interface PdfjsLib {
    getDocument: (options: Record<string, unknown>) => { promise: Promise<PdfjsDoc> };
  }
  const lib: PdfjsLib = mod.default ?? mod;
  const open = async (bytes: Uint8Array): Promise<PdfjsDoc> =>
    lib.getDocument({ data: new Uint8Array(bytes), useSystemFonts: true, isEvalSupported: false }).promise;
  const itemsOf = async (bytes: Uint8Array, pageNumber: number): Promise<PdfTextItem[]> => {
    const doc = await open(bytes);
    try {
      const page = await doc.getPage(pageNumber);
      const content = await page.getTextContent({ includeMarkedContent: true });
      const items: PdfTextItem[] = [];
      for (const raw of content.items) {
        if (typeof raw.str !== 'string') continue;
        const transform = Array.isArray(raw.transform) ? raw.transform : [];
        const x = typeof transform[4] === 'number' ? (transform[4] as number) : 0;
        const y = typeof transform[5] === 'number' ? (transform[5] as number) : 0;
        items.push({ str: raw.str, x, y, hasEOL: raw.hasEOL === true });
      }
      return items;
    } finally {
      await doc.destroy();
    }
  };
  return {
    pageCount: async (bytes: Uint8Array) => {
      const doc = await open(bytes);
      try {
        return doc.numPages;
      } finally {
        await doc.destroy();
      }
    },
    pageItems: itemsOf,
  };
}

export async function extractPdf(
  bytes: Uint8Array,
  fileName: string,
  engine?: PdfEngine,
): Promise<ExtractionResult> {
  const impl = engine ?? (await pdfjsEngine());
  let pageCount: number;
  try {
    pageCount = await impl.pageCount(bytes);
  } catch (error) {
    if (isProtectionError(error)) throw protectionRefusal(fileName);
    if (error instanceof Error && /password/i.test(error.name)) throw protectionRefusal(fileName);
    throw new Error(`"${fileName}" could not be opened: ${error instanceof Error ? error.message : String(error)}`);
  }
  const warnings: string[] = [];
  const kept = Math.min(pageCount, MAX_PDF_PAGES);
  if (pageCount > MAX_PDF_PAGES) {
    warnings.push(
      `Only the first ${MAX_PDF_PAGES} of ${pageCount} pages were extracted; the rest needs a follow-up call.`,
    );
  }
  const units = [];
  for (let page = 1; page <= kept; page += 1) {
    const rows = rowsOf(await impl.pageItems(bytes, page));
    units.push({
      id: `page-${page}`,
      title: `Page ${page}`,
      text: rows.join('\n'),
    });
  }
  if (units.every((unit) => unit.text.trim().length === 0)) {
    throw scannedPdfRefusal(fileName);
  }
  return { units, warnings };
}
