/**
 * .xlsx extraction through exceljs in STREAMING mode: one unit per worksheet,
 * rows as tab-separated values, formulas as their SAVED results (never
 * evaluated — evaluation would need the full calc chain and would still
 * disagree with Excel on volatile functions).
 *
 * Why exceljs and not the `xlsx` package: `xlsx` on npm is frozen at 0.18.5
 * with known CVEs and prototype-pollution advisories, while exceljs is
 * maintained and offers a streaming reader, which is what keeps a 100k-row
 * sheet from becoming a heap snapshot. SheetJS is also GPL-licensed for the
 * maintained line; exceljs is MIT.
 */

import { Readable } from 'node:stream';

import {
  MAX_XLSX_COLS,
  MAX_XLSX_ROWS,
  isProtectionError,
  protectionRefusal,
  type ExtractionResult,
} from './extract.js';

export interface XlsxSheetInput {
  name: string;
  /** Row values in order; index 0 is unused like exceljs `row.values`. */
  rows: unknown[][];
}

export interface XlsxEngine {
  readSheets: (bytes: Uint8Array) => Promise<XlsxSheetInput[]>;
}

/** Renders one exceljs cell value as text. Saved formula results win. */
export function cellText(value: unknown): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'bigint') {
    return String(value);
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) return value.toISOString();
  if (typeof value !== 'object') return '';
  const cell = value as Record<string, unknown>;
  // A formula cell: { formula, result, ... }. The SAVED result is the value;
  // the formula itself is noted so the model knows the cell is computed.
  if ('result' in cell || 'formula' in cell) {
    const result = cell['result'];
    const rendered = result === null || result === undefined ? '' : cellText(result);
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

async function exceljsEngine(): Promise<XlsxEngine> {
  const mod = (await import('exceljs')) as unknown as {
    default?: { stream: { xlsx: { WorkbookReader: new () => XlsxWorkbookReader } } };
    stream: { xlsx: { WorkbookReader: new () => XlsxWorkbookReader } };
  };
  interface XlsxRow {
    values: unknown[];
  }
  interface XlsxWorksheetReader {
    name?: string;
    on: (event: 'row', listener: (row: XlsxRow) => void) => void;
  }
  interface XlsxWorkbookReader {
    read: (stream: NodeJS.ReadableStream, options?: Record<string, string>) => Promise<void>;
    on: (event: 'worksheet', listener: (ws: XlsxWorksheetReader) => void) => void;
  }
  const ctor = mod.default?.stream.xlsx.WorkbookReader ?? mod.stream.xlsx.WorkbookReader;
  return {
    readSheets: async (bytes: Uint8Array) => {
      const reader: XlsxWorkbookReader = new ctor();
      const sheets: XlsxSheetInput[] = [];
      let rows = 0;
      reader.on('worksheet', (ws) => {
        const sheet: XlsxSheetInput = { name: ws.name ?? `Sheet${sheets.length + 1}`, rows: [] };
        sheets.push(sheet);
        ws.on('row', (row) => {
          // The cap is enforced again at shaping time; stopping the stream
          // early is the memory win, and 100k rows is the documented cap.
          if (rows >= MAX_XLSX_ROWS) return;
          rows += 1;
          sheet.rows.push(Array.isArray(row.values) ? row.values : []);
        });
      });
      const stream = Readable.from([Buffer.from(bytes)]);
      await reader.read(stream, { entries: 'emit', worksheets: 'emit' });
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
  const units = sheets.map((sheet, index) => {
    const kept: string[] = [];
    let cols = 0;
    for (const row of sheet.rows) {
      if (totalRows >= MAX_XLSX_ROWS) break;
      totalRows += 1;
      // exceljs rows are 1-based: values[0] is always empty.
      const cells = row.slice(1, MAX_XLSX_COLS + 1).map(cellText);
      // Trim trailing empties so ragged rows do not pad tabs to the margin.
      while (cells.length > 0 && cells[cells.length - 1] === '') cells.pop();
      cols = Math.max(cols, cells.length);
      kept.push(cells.join('\t'));
    }
    const droppedCols = sheet.rows.some((row) => row.length - 1 > MAX_XLSX_COLS);
    return {
      id: sheet.name.length > 0 ? sheet.name : `sheet-${index + 1}`,
      title: `${sheet.name.length > 0 ? sheet.name : `Sheet ${index + 1}`} (${kept.length} rows × ${cols} cols)`,
      text: kept.join('\n'),
      ...(droppedCols ? { droppedCols: true as const } : {}),
    };
  });
  if (units.some((unit) => 'droppedCols' in unit)) {
    warnings.push(`Columns past ${MAX_XLSX_COLS} per row were dropped.`);
  }
  if (totalRows >= MAX_XLSX_ROWS) {
    warnings.push(
      `Stopped after ${MAX_XLSX_ROWS.toLocaleString('en-US')} rows across all sheets; the rest was not read.`,
    );
  }
  if (units.length === 0) {
    throw new Error(`"${fileName}" holds no worksheets, so there is nothing to extract.`);
  }
  return {
    units: units.map(({ id, title, text }) => ({ id, title, text })),
    warnings,
  };
}
