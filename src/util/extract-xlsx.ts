/**
 * .xlsx extraction through exceljs in STREAMING mode: one unit per worksheet,
 * rows as Markdown tables (KV-transposed for single-record blocks), formulas
 * as their SAVED results (never evaluated — evaluation would need the full
 * calc chain and would still disagree with Excel on volatile functions).
 *
 * Why exceljs and not the `xlsx` package: `xlsx` on npm is frozen at 0.18.5
 * with known CVEs and prototype-pollution advisories, while exceljs is
 * maintained and offers a streaming reader, which is what keeps a 100k-row
 * sheet from becoming a heap snapshot. SheetJS is also GPL-licensed for the
 * maintained line; exceljs is MIT.
 *
 * Two fidelity notes for future readers. First, the streaming reader is told
 * to CACHE shared strings and styles: without that, shared-string cells come
 * back as unresolved `{sharedString: index}` stubs and dates stay raw serials
 * (the options object given to `read` REPLACES the defaults, so every needed
 * key is spelled out). Second, the stream drops hidden-row flags and merge
 * ranges entirely, so a small jszip pre-pass over the sheet XML collects them
 * (one sheet string at a time, released before the next) for skip/warn and
 * top-left fill-down at shaping time.
 *
 * One reliability note. exceljs's streaming reader walks the zip through
 * unzipper, which ends the walk early whenever the reader pauses mid-archive
 * to spill a sheet to a temp file — and it spills whenever a sheet streams
 * before the shared strings, which is how Excel itself writes workbooks
 * (upstream exceljs#3064, still open). The pre-pass therefore also REORDERS
 * the archive — rels, workbook, shared strings and styles first — so every
 * sheet streams inline and the spill path never triggers; bytes already in
 * that order pass through untouched. Reader state is pre-seeded with the
 * pre-pass rels (and an empty string table) so workbooks MISSING those parts
 * stream inline too instead of spilling.
 */

import { Readable } from 'node:stream';

import {
  MAX_XLSX_COLS,
  MAX_XLSX_ROWS,
  isProtectionError,
  protectionRefusal,
  toMarkdownTable,
  type ExtractionResult,
} from './extract.js';

export interface XlsxSheetInput {
  name: string;
  /** Row values in order; index 0 is unused like exceljs `row.values`. */
  rows: unknown[][];
  /** Per-cell number formats, aligned with `rows` (index 0 unused). */
  numFmts?: Array<Array<string | undefined>>;
  /** Absolute 1-based row numbers, aligned with `rows`; defaults to index+1. */
  rowNumbers?: number[];
  /** Merge ranges (`A1:B2`) for top-left fill-down. */
  merges?: string[];
  /** A hidden/veryHidden sheet is skipped with a warning, never read. */
  hidden?: boolean;
  /** Hidden rows skipped during the read, warned about at shaping time. */
  skippedHiddenRows?: number;
  /** True when the workbook uses the 1904 date system. */
  date1904?: boolean;
}

export interface XlsxEngine {
  readSheets: (bytes: Uint8Array) => Promise<XlsxSheetInput[]>;
}

/** exceljs `isDateFmt`: strip quoted/bracketed runs, look for date letters. */
function isDateFmt(fmt: string | undefined): boolean {
  if (fmt === undefined || fmt === '') return false;
  const stripped = fmt.replace(/\[[^\]]*]/g, '').replace(/"[^"]*"/g, '');
  return /[ymdhMsb]/.test(stripped);
}

/** First `;`-section of a number format (the positive-number shape). */
function firstSection(fmt: string): string {
  return (fmt.split(';')[0] ?? '').trim();
}

function groupInteger(digits: string): string {
  const sign = digits.startsWith('-') ? '-' : '';
  const body = sign === '' ? digits : digits.slice(1);
  return sign + body.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function decimalsOf(section: string): number {
  const dot = section.match(/\.([0?#]+)/);
  if (dot?.[1] === undefined) return -1;
  // `0.##` shows up to 2 decimals but no trailing zeros: handled by trimming.
  return (dot[1].match(/0/g) ?? []).length;
}

function maxDecimalsOf(section: string): number {
  const dot = section.match(/\.([0?#]+)/);
  return dot?.[1]?.length ?? 0;
}

/** Renders a number the way the format shows it: percent, currency, grouped. */
function formatNumber(value: number, numFmt: string): string {
  if (!Number.isFinite(value)) return String(value);
  const section = firstSection(numFmt);
  if (section === '' || section.toLowerCase() === 'general') return String(value);
  if (section.includes('%')) {
    const decimals = Math.max(0, decimalsOf(section));
    return `${(value * 100).toFixed(decimals)}%`;
  }
  const currency = currencySymbol(section);
  const scientific = /E[+-]00/i.test(section);
  const decimals = decimalsOf(section);
  const maxDecimals = maxDecimalsOf(section);
  const grouped = section.includes('#,##') || section.includes('0,');
  let rendered: string;
  if (scientific) {
    rendered = value.toExponential(Math.max(0, decimals));
  } else if (maxDecimals > 0) {
    rendered = value.toFixed(maxDecimals);
    if (rendered.includes('.')) {
      // `0.##`-style: keep at least `decimals` places, trim the rest.
      let trimmed = rendered;
      while (trimmed.length - trimmed.indexOf('.') - 1 > decimals && trimmed.endsWith('0')) {
        trimmed = trimmed.slice(0, -1);
      }
      rendered = trimmed.endsWith('.') ? trimmed.slice(0, -1) : trimmed;
    }
  } else if (/[0?#]/.test(section)) {
    // Integer shapes (`0`, `#,##0`) round; anything else shows the value raw.
    rendered = value.toFixed(0);
  } else {
    return String(value);
  }
  if (grouped && !scientific) {
    const [head, ...tail] = rendered.split('.');
    rendered = tail.length > 0 ? `${groupInteger(head ?? '')}.${tail.join('.')}` : groupInteger(head ?? '');
  }
  if (currency !== '') {
    return rendered.startsWith('-') ? `-${currency}${rendered.slice(1)}` : `${currency}${rendered}`;
  }
  return rendered;
}

function currencySymbol(section: string): string {
  // `[$€-407]`-style locale blocks name the symbol up to the first `-`.
  const tagged = section.match(/\[\$([^\]$-]+)/);
  if (tagged?.[1] !== undefined && tagged[1] !== '') return tagged[1][0] as string;
  const bare = section.match(/[$€£¥₹₩]/);
  return bare?.[0] ?? '';
}

const MONTHS_FULL = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const MONTHS_SHORT = MONTHS_FULL.map((name) => name.slice(0, 3));

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

/** Excel serial to Date, honouring the workbook's 1904 flag (exceljs maths). */
export function excelSerialToDate(serial: number, date1904: boolean): Date {
  return new Date(Math.round((serial - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000));
}

/**
 * Renders a Date the way the format shows it, in UTC (serials are
 * timezone-free, so local getters would shift the day near midnight).
 * Supports the y/m/d/h/s tokens Excel users actually write; anything exotic
 * falls back to the ISO date rather than guessing.
 */
function formatDate(date: Date, numFmt: string): string | null {
  const section = firstSection(numFmt)
    .replace(/\[[^\]]*]/g, '')
    .replace(/"[^"]*"/g, '')
    .replace(/\\(.)/g, '$1')
    .replace(/_./g, '');
  if (!/[ymdhs]/i.test(section)) return null;
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth();
  const day = date.getUTCDate();
  const hours = date.getUTCHours();
  const minutes = date.getUTCMinutes();
  const seconds = date.getUTCSeconds();
  // Tokenise longest-first so `mmmm` wins over `mm` wins over `m`.
  const tokens = section.match(/yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|am\/pm|A\/P|[^ymdhsaA]+/gi);
  if (tokens === null) return null;
  let sawDate = false;
  let out = '';
  let afterHour = false;
  for (const token of tokens) {
    const word = token.toLowerCase();
    if (/^y+$/.test(word)) {
      sawDate = true;
      afterHour = false;
      out += token.length >= 4 ? String(year) : String(year % 100).padStart(2, '0');
    } else if (/^m+$/.test(word)) {
      if (afterHour) {
        out += token.length >= 2 ? pad2(minutes) : String(minutes);
      } else if (token.length >= 4) {
        sawDate = true;
        out += MONTHS_FULL[month] ?? '';
      } else if (token.length === 3) {
        sawDate = true;
        out += MONTHS_SHORT[month] ?? '';
      } else {
        sawDate = true;
        out += token.length === 2 ? pad2(month + 1) : String(month + 1);
      }
      afterHour = false;
    } else if (/^d+$/.test(word)) {
      sawDate = true;
      afterHour = false;
      out += token.length >= 2 ? pad2(day) : String(day);
    } else if (/^h+$/.test(word)) {
      afterHour = true;
      out += token.length >= 2 ? pad2(hours) : String(hours);
    } else if (/^s+$/.test(word)) {
      afterHour = false;
      out += token.length >= 2 ? pad2(seconds) : String(seconds);
    } else if (word === 'am/pm' || word === 'a/p') {
      afterHour = false;
      out += hours < 12 ? (token.includes('A') ? 'AM' : 'am') : token.includes('A') ? 'PM' : 'pm';
    } else {
      // Literal run: any leftover date letter means an exotic token — bail.
      if (/[ymdhs]/i.test(token)) return null;
      out += token;
      // Separators keep the hour context alive so `hh:mm` reads minutes.
      if (!/^[\s:./\-,]*$/.test(token)) afterHour = false;
    }
  }
  return sawDate || /[0-9]/.test(out) ? out : null;
}

/**
 * Renders one exceljs cell value as text. Saved formula results win; the
 * number format renders what Excel SHOWS (dates, currency, percent) rather
 * than the raw serial, with the 1904 system honoured for date serials.
 */
export function cellText(value: unknown, numFmt?: string, date1904 = false): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'bigint') return String(value);
  if (typeof value === 'number') {
    if (numFmt !== undefined && isDateFmt(numFmt)) {
      const rendered = formatDate(excelSerialToDate(value, date1904), numFmt);
      return rendered ?? excelSerialToDate(value, date1904).toISOString().slice(0, 10);
    }
    if (numFmt !== undefined) return formatNumber(value, numFmt);
    return String(value);
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) {
    if (numFmt !== undefined && isDateFmt(numFmt)) {
      return formatDate(value, numFmt) ?? value.toISOString().slice(0, 10);
    }
    return value.toISOString();
  }
  if (typeof value !== 'object') return '';
  const cell = value as Record<string, unknown>;
  // A formula cell: { formula, result, ... }. The SAVED result is the value;
  // the formula itself is noted so the model knows the cell is computed.
  if ('result' in cell || 'formula' in cell) {
    const result = cell['result'];
    const rendered = result === null || result === undefined ? '' : cellText(result, numFmt, date1904);
    const formula = typeof cell['formula'] === 'string' ? cell['formula'] : '';
    if (rendered !== '') return formula !== '' ? `${rendered} [=${formula}]` : rendered;
    return formula !== '' ? `[=${formula}]` : '';
  }
  if (Array.isArray(cell['richText'])) {
    return (cell['richText'] as Array<{ text?: unknown }>)
      .map((run) => (typeof run.text === 'string' ? run.text : ''))
      .join('');
  }
  if (typeof cell['text'] === 'string') return cell['text'];
  if (typeof cell['error'] === 'string') return cell['error'];
  if (typeof cell['hyperlink'] === 'string' && typeof cell['text'] === 'string') {
    return cell['text'];
  }
  return '';
}

function colLettersToNumber(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

interface MergeRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function parseMergeRef(ref: string): MergeRange | null {
  const match = /^([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$/.exec(ref);
  if (match === null) return null;
  const c1 = colLettersToNumber(match[1] as string);
  const r1 = Number(match[2]);
  const c2 = colLettersToNumber(match[3] as string);
  const r2 = Number(match[4]);
  if (c1 < 1 || r1 < 1 || c2 < c1 || r2 < r1) return null;
  return { r1, c1, r2, c2 };
}

interface FillMatrix {
  rows: unknown[][];
  numFmts: Array<Array<string | undefined>>;
  rowNumbers: number[];
}

/**
 * Fills merged ranges with their top-left value, Excel's own display rule.
 * Only EMPTY cells inside the KEPT region are filled (clamped to the row/column
 * caps, so a whole-column merge cannot allocate a million cells); the map
 * lookup keeps sparse sheets cheap.
 */
function fillMerges(matrix: FillMatrix, merges: string[]): void {
  if (merges.length === 0 || matrix.rows.length === 0) return;
  const indexByRow = new Map<number, number>();
  matrix.rowNumbers.forEach((num, index) => {
    if (!indexByRow.has(num)) indexByRow.set(num, index);
  });
  for (const ref of merges) {
    const range = parseMergeRef(ref);
    if (range === null) continue;
    const topIndex = indexByRow.get(range.r1);
    if (topIndex === undefined) continue;
    if (range.c1 > MAX_XLSX_COLS) continue;
    const topRow = matrix.rows[topIndex] ?? [];
    const topFmt = matrix.numFmts[topIndex] ?? [];
    const topValue = topRow[range.c1];
    if (topValue === null || topValue === undefined || topValue === '') continue;
    const lastCol = Math.min(range.c2, MAX_XLSX_COLS);
    for (let r = range.r1; r <= range.r2; r += 1) {
      const index = indexByRow.get(r);
      if (index === undefined) continue;
      const row = matrix.rows[index] ?? [];
      const fmts = matrix.numFmts[index] ?? [];
      for (let c = range.c1; c <= lastCol; c += 1) {
        const current = row[c];
        if (current === null || current === undefined || current === '') {
          row[c] = topValue;
          if (fmts[c] === undefined) fmts[c] = topFmt[range.c1];
        }
      }
    }
  }
}

interface SheetMeta {
  hiddenRows: Set<number>;
  merges: string[];
}

interface PreparedWorkbook {
  meta: Map<string, SheetMeta>;
  date1904: boolean;
  workbookRels: Array<{ Id: string; Target: string }>;
  /** Input bytes reordered for the streaming walk, or the input untouched. */
  orderedBytes: Uint8Array;
}

/** Parts the streaming reader must see before the first sheet streams. */
const STREAM_FIRST = [
  'xl/_rels/workbook.xml.rels',
  'xl/workbook.xml',
  'xl/sharedStrings.xml',
  'xl/styles.xml',
];

/**
 * The streaming reader's blind spots, collected with one regex pass per sheet
 * part: hidden-row flags, merge ranges, the sheet-name-to-file map and the
 * workbook's 1904 flag — plus the entry reorder that keeps the walk inline
 * (see the header note). Best-effort: a failure here returns the input bytes
 * with empty metadata and must never fail the read itself.
 */
async function prepareWorkbook(bytes: Uint8Array): Promise<PreparedWorkbook> {
  const fallback: PreparedWorkbook = {
    meta: new Map<string, SheetMeta>(),
    date1904: false,
    workbookRels: [],
    orderedBytes: bytes,
  };
  try {
    const zipMod = (await import('jszip')) as unknown as {
      default?: ZipCtor;
      loadAsync: (data: Uint8Array) => Promise<MetaZip>;
    };
    interface MetaZip {
      files: Record<string, { dir: boolean }>;
      file: (path: string) => { async: (kind: string) => Promise<unknown> } | null;
    }
    interface ZipCtor {
      loadAsync: (data: Uint8Array) => Promise<MetaZip>;
      new (): {
        file: (path: string, data: Uint8Array) => void;
        generateAsync: (options: Record<string, unknown>) => Promise<Uint8Array>;
      };
    }
    const ctor = zipMod.default ?? (zipMod as unknown as ZipCtor);
    const zip = await ctor.loadAsync(bytes);
    const names = Object.keys(zip.files).filter((name) => zip.files[name]?.dir !== true);
    const read = async (path: string): Promise<string | null> => {
      const entry = zip.file(path);
      return entry === null ? null : ((await entry.async('string')) as string);
    };
    const workbookRels: Array<{ Id: string; Target: string }> = [];
    const meta = new Map<string, SheetMeta>();
    let date1904 = false;
    const workbookXml = await read('xl/workbook.xml');
    if (workbookXml !== null) {
      date1904 = /<workbookPr\b[^>]*\bdate1904="1"/.test(workbookXml);
      const relsXml = (await read('xl/_rels/workbook.xml.rels')) ?? '';
      const targetById = new Map<string, string>();
      for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) ?? []) {
        const id = /Id="([^"]+)"/.exec(tag)?.[1];
        const target = /Target="([^"]+)"/.exec(tag)?.[1];
        if (id !== undefined && target !== undefined) {
          targetById.set(id, target);
          workbookRels.push({ Id: id, Target: target });
        }
      }
      for (const tag of workbookXml.match(/<sheet\b[^>]*>/g) ?? []) {
        const name = /name="([^"]*)"/.exec(tag)?.[1];
        // XML entities in names are decoded the cheap way; streaming names come
        // decoded already, so this keeps the two maps keyed alike.
        const decoded =
          name === undefined
            ? undefined
            : name
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&apos;/g, "'")
                .replace(/&amp;/g, '&');
        const rid = /(?:r:id|id)="([^"]+)"/.exec(tag)?.[1];
        if (decoded === undefined || rid === undefined) continue;
        const target = targetById.get(rid);
        if (target === undefined) continue;
        const file = target.startsWith('xl/') ? target : `xl/${target.replace(/^\//, '')}`;
        const sheetXml = await read(file);
        const info: SheetMeta = { hiddenRows: new Set(), merges: [] };
        if (sheetXml !== null) {
          for (const row of sheetXml.match(/<row\b[^>]*>/g) ?? []) {
            if (!/\bhidden="1"/.test(row)) continue;
            const num = /\br="(\d+)"/.exec(row)?.[1];
            if (num !== undefined) info.hiddenRows.add(Number(num));
          }
          for (const merge of sheetXml.match(/<mergeCell\b[^>]*ref="([^"]+)"[^>]*>/g) ?? []) {
            const ref = /ref="([^"]+)"/.exec(merge)?.[1];
            if (ref !== undefined) info.merges.push(ref);
          }
        }
        meta.set(decoded, info);
      }
    }
    // Reorder only when a sheet would otherwise stream before its prerequisites.
    const firstSheetAt = names.findIndex((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));
    const presentFirst = STREAM_FIRST.filter((name) => names.includes(name));
    const needsReorder =
      firstSheetAt !== -1 && presentFirst.some((name) => names.indexOf(name) > firstSheetAt);
    let orderedBytes = bytes;
    if (needsReorder) {
      const out = new ctor();
      for (const name of [...presentFirst, ...names.filter((entry) => !presentFirst.includes(entry))]) {
        const data = (await zip.file(name)?.async('uint8array')) as Uint8Array | undefined;
        if (data !== undefined) out.file(name, data);
      }
      // STORE, not DEFLATE: the bytes are transient streaming input, and
      // recompressing would trade seconds of CPU for nothing the reader needs.
      orderedBytes = new Uint8Array(await out.generateAsync({ type: 'uint8array', compression: 'STORE' }));
    }
    return { meta, date1904, workbookRels, orderedBytes };
  } catch {
    return fallback;
  }
}

async function exceljsEngine(): Promise<XlsxEngine> {
  const mod = (await import('exceljs')) as unknown as {
    default?: { stream: { xlsx: { WorkbookReader: new () => XlsxWorkbookReader } } };
    stream: { xlsx: { WorkbookReader: new () => XlsxWorkbookReader } };
  };
  interface XlsxCell {
    numFmt?: unknown;
  }
  interface XlsxRow {
    number?: unknown;
    values: unknown[];
    eachCell: (options: { includeEmpty: boolean }, iteratee: (cell: XlsxCell, col: number) => void) => void;
  }
  interface XlsxWorksheetReader {
    name?: string;
    state?: unknown;
    on: (event: 'row', listener: (row: XlsxRow) => void) => void;
  }
  interface XlsxWorkbookReader {
    properties?: { model?: { date1904?: unknown } };
    sharedStrings?: unknown[];
    workbookRels?: Array<{ Id: string; Target: string }>;
    read: (stream: NodeJS.ReadableStream, options?: Record<string, string>) => Promise<void>;
    on: (event: 'worksheet', listener: (ws: XlsxWorksheetReader) => void) => void;
  }
  const ctor = mod.default?.stream.xlsx.WorkbookReader ?? mod.stream.xlsx.WorkbookReader;
  return {
    readSheets: async (bytes: Uint8Array) => {
      const prepass = await prepareWorkbook(bytes);
      const reader: XlsxWorkbookReader = new ctor();
      // Pre-seeds so sheets stream inline even when their prerequisites are
      // missing from the package; real entries overwrite these when present.
      reader.sharedStrings = [];
      if (prepass.workbookRels.length > 0) reader.workbookRels = prepass.workbookRels;
      const sheets: XlsxSheetInput[] = [];
      let rows = 0;
      reader.on('worksheet', (ws) => {
        const name = ws.name ?? `Sheet${sheets.length + 1}`;
        const state = typeof ws.state === 'string' ? ws.state : 'visible';
        if (state === 'hidden' || state === 'veryHidden') {
          sheets.push({ name, rows: [], hidden: true });
          return;
        }
        const info = prepass.meta.get(name);
        const sheet: XlsxSheetInput = {
          name,
          rows: [],
          numFmts: [],
          rowNumbers: [],
          merges: info === undefined ? [] : [...info.merges],
          skippedHiddenRows: 0,
          date1904: prepass.date1904,
        };
        sheets.push(sheet);
        ws.on('row', (row) => {
          // The cap is enforced again at shaping time; stopping the stream
          // early is the memory win, and 100k rows is the documented cap.
          if (rows >= MAX_XLSX_ROWS) return;
          const num = typeof row.number === 'number' ? row.number : sheet.rows.length + 1;
          if (info !== undefined && info.hiddenRows.has(num)) {
            sheet.skippedHiddenRows = (sheet.skippedHiddenRows ?? 0) + 1;
            return;
          }
          rows += 1;
          sheet.rows.push(Array.isArray(row.values) ? row.values.slice() : []);
          sheet.rowNumbers?.push(num);
          const fmts: Array<string | undefined> = [];
          try {
            row.eachCell({ includeEmpty: true }, (cell, col) => {
              if (col > MAX_XLSX_COLS + 1) return;
              fmts[col] = typeof cell.numFmt === 'string' ? cell.numFmt : undefined;
            });
          } catch {
            // A row that will not enumerate still keeps its values.
          }
          sheet.numFmts?.push(fmts);
        });
      });
      const stream = Readable.from([Buffer.from(prepass.orderedBytes)]);
      // Every key spelled out: `read` REPLACES the constructor defaults, and
      // the defaults that matter here (sharedStrings/styles cached) are what
      // resolve text cells and date formats. Verified on exceljs 4.4.0.
      await reader.read(stream, {
        entries: 'emit',
        worksheets: 'emit',
        sharedStrings: 'cache',
        styles: 'cache',
        hyperlinks: 'ignore',
      });
      const reader1904 = reader.properties?.model?.date1904 === true;
      for (const sheet of sheets) sheet.date1904 = prepass.date1904 || reader1904;
      return sheets;
    },
  };
}

export async function extractXlsx(
  bytes: Uint8Array,
  fileName: string,
  engine?: XlsxEngine,
): Promise<ExtractionResult> {
  const impl = engine ?? (await exceljsEngine());
  let sheets: XlsxSheetInput[];
  try {
    sheets = await impl.readSheets(bytes);
  } catch (error) {
    if (isProtectionError(error)) throw protectionRefusal(fileName);
    throw new Error(`"${fileName}" could not be read: ${error instanceof Error ? error.message : String(error)}`);
  }
  const warnings: string[] = [];
  let totalRows = 0;
  const units = [];
  for (const [index, sheet] of sheets.entries()) {
    if (sheet.hidden === true) {
      warnings.push(`"${sheet.name}" is hidden and was skipped.`);
      continue;
    }
    if ((sheet.skippedHiddenRows ?? 0) > 0) {
      const count = sheet.skippedHiddenRows ?? 0;
      warnings.push(`Skipped ${count} hidden ${count === 1 ? 'row' : 'rows'} in "${sheet.name}".`);
    }
    const matrix: FillMatrix = {
      rows: [],
      numFmts: [],
      rowNumbers: [],
    };
    let cols = 0;
    for (const [rowIndex, row] of sheet.rows.entries()) {
      if (totalRows >= MAX_XLSX_ROWS) break;
      totalRows += 1;
      matrix.rows.push(row);
      matrix.numFmts.push(sheet.numFmts?.[rowIndex] ?? []);
      matrix.rowNumbers.push(sheet.rowNumbers?.[rowIndex] ?? rowIndex + 1);
      cols = Math.max(cols, row.length - 1);
    }
    fillMerges(matrix, sheet.merges ?? []);
    const rendered: string[][] = matrix.rows.map((row, rowIndex) => {
      const fmts = matrix.numFmts[rowIndex] ?? [];
      const cells = row.slice(1, MAX_XLSX_COLS + 1).map((value, offset) => {
        const fmt = fmts[offset + 1];
        return cellText(value, fmt, sheet.date1904 ?? false);
      });
      // Trim trailing empties so ragged rows do not pad pipes to the margin.
      while (cells.length > 0 && cells[cells.length - 1] === '') cells.pop();
      return cells;
    });
    cols = Math.min(cols, MAX_XLSX_COLS);
    const nonEmpty = rendered.filter((cells) => cells.length > 0);
    const text =
      nonEmpty.length === 0 ? '' : toMarkdownTable(nonEmpty[0] as string[], nonEmpty.slice(1));
    const droppedCols = sheet.rows.some((row) => row.length - 1 > MAX_XLSX_COLS);
    units.push({
      id: sheet.name.length > 0 ? sheet.name : `sheet-${index + 1}`,
      title: `${sheet.name.length > 0 ? sheet.name : `Sheet ${index + 1}`} (${matrix.rows.length} rows × ${cols} cols)`,
      text,
      ...(droppedCols ? { droppedCols: true as const } : {}),
    });
  }
  if (units.some((unit) => 'droppedCols' in unit)) {
    warnings.push(`Columns past ${MAX_XLSX_COLS} per row were dropped.`);
  }
  if (totalRows >= MAX_XLSX_ROWS) {
    warnings.push(
      `Stopped after ${MAX_XLSX_ROWS.toLocaleString('en-US')} rows across all sheets; the rest was not read.`,
    );
  }
  if (units.length === 0) {
    if (sheets.length > 0) {
      throw new Error(`"${fileName}" holds only hidden worksheets, so there is nothing to extract.`);
    }
    throw new Error(`"${fileName}" holds no worksheets, so there is nothing to extract.`);
  }
  return {
    units: units.map(({ id, title, text }) => ({ id, title, text })),
    warnings,
  };
}
