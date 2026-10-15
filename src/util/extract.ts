/**
 * Shared shape and guards for Office/PDF text extraction.
 *
 * Every extractor returns UNITS — pages, slides, sheets, sections — never one
 * giant string, so the caller can page, truncate per unit, and account for
 * every unit in a manifest (the no-omission contract: `unitCount` plus a
 * `truncated` flag on each unit says exactly what the model did and did not
 * get).
 *
 * What is deliberately NOT attempted, each mapped to a friendly error at the
 * call site: password/IRM/RMS-protected content, scanned-PDF OCR, legacy OLE
 * `.doc`/`.ppt`/`.xls`, formula evaluation (saved values are returned), and
 * pixel layout, borders, headers/footers and tracked changes.
 */

export type DocumentKind = 'docx' | 'xlsx' | 'pptx' | 'pdf';

/** One page, slide, sheet or section of extracted text. */
export interface ExtractedUnit {
  /** Stable within the document: slide id, page number, sheet name. */
  id: string;
  /** Heading, first line or sheet name, when one exists. */
  title?: string | undefined;
  /** Plain text of the unit. */
  text: string;
}

export interface ExtractionResult {
  units: ExtractedUnit[];
  /** Loss the caller should surface, e.g. capped pages or skipped graphics. */
  warnings: string[];
}

/** Largest document this server will download for parsing. */
export const MAX_DOCUMENT_BYTES = 16 * 1024 * 1024;
/** Most PDF pages extracted; the rest need a second call, not silently read. */
export const MAX_PDF_PAGES = 50;
/** Most spreadsheet rows extracted across all sheets. */
export const MAX_XLSX_ROWS = 100_000;
/** Most columns kept per spreadsheet row. */
export const MAX_XLSX_COLS = 100;

const PDF_MAGIC = [0x25, 0x50, 0x44, 0x46]; // %PDF
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04]; // PK\x03\x04
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0]; // legacy .doc/.ppt/.xls

function hasMagic(bytes: Uint8Array, magic: number[]): boolean {
  return magic.every((byte, i) => bytes[i] === byte);
}

/**
 * True when the OLE container holds an OOXML-encrypted package (MS-OFFCRYPTO)
 * rather than a legacy .doc/.ppt/.xls. OLE directory entries store stream
 * names as UTF-16LE, so the ASCII literal never appears in the bytes; search
 * the encoded stream name instead.
 */
function isEncryptedOle(bytes: Uint8Array): boolean {
  return Buffer.from(bytes).indexOf(Buffer.from('EncryptedPackage', 'utf16le')) !== -1;
}

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  if (dot <= 0 || dot === name.length - 1) return '';
  return name.slice(dot + 1).toLowerCase();
}

/**
 * Picks the extractor from magic bytes first, extension second. Magic wins
 * because a renamed file is still what its bytes say; the extension breaks
 * ties inside zip containers (docx vs xlsx vs pptx share a magic).
 */
export function detectKind(name: string, mimeType: string | undefined, bytes: Uint8Array): DocumentKind {
  if (hasMagic(bytes, OLE_MAGIC)) {
    if (isEncryptedOle(bytes)) throw protectionRefusal(name);
    throw new Error(
      `"${name}" is a legacy OLE document (.doc/.ppt/.xls). Only the modern zip-based formats ` +
        '(.docx/.pptx/.xlsx) and PDF can be parsed here; open the file in its native app and save a copy in the modern format.',
    );
  }
  if (hasMagic(bytes, PDF_MAGIC)) return 'pdf';

  const ext = extensionOf(name);
  if (ext === 'pdf') return 'pdf';
  if (hasMagic(bytes, ZIP_MAGIC) || ext !== '') {
    if (ext === 'docx') return 'docx';
    if (ext === 'xlsx' || ext === 'xlsm') return 'xlsx';
    if (ext === 'pptx') return 'pptx';
    const mime = (mimeType ?? '').toLowerCase();
    if (mime.includes('wordprocessingml')) return 'docx';
    if (mime.includes('spreadsheetml')) return 'xlsx';
    if (mime.includes('presentationml')) return 'pptx';
  }
  if (ext === 'doc' || ext === 'ppt' || ext === 'xls') {
    throw new Error(
      `"${name}" is a legacy .${ext} file. Only .docx/.pptx/.xlsx and PDF can be parsed here; ` +
        'save a copy in the modern format first.',
    );
  }
  throw new Error(
    `"${name}" is not a recognised .docx, .xlsx, .pptx or .pdf file, so it cannot be parsed here. ` +
      'Use files_read_text for text-like files.',
  );
}

/**
 * True when a library failure smells like protection rather than corruption:
 * password, Azure RMS/IRM, or a rights-managed wrapper. Callers map this to
 * the friendly refusal instead of surfacing the library's stack.
 */
export function isProtectionError(error: unknown): boolean {
  const text = `${error instanceof Error ? `${error.name} ${error.message}` : String(error)}`.toLowerCase();
  return (
    text.includes('password') ||
    text.includes('encrypt') ||
    text.includes('decrypt') ||
    text.includes('protected') ||
    text.includes('permission') ||
    text.includes('rms') ||
    text.includes('irm') ||
    text.includes('rights') ||
    text.includes('drm')
  );
}

export function protectionRefusal(name: string): Error {
  return new Error(
    `"${name}" is password-protected or rights-managed (IRM/RMS), so its text cannot be extracted here. ` +
      'Remove the protection in the native app and try again.',
  );
}

export function scannedPdfRefusal(name: string): Error {
  return new Error(
    `"${name}" has no extractable text on any page — it is most likely a scanned PDF. ` +
      'This server does no OCR; run it through an OCR tool (or open it in Word, which converts scans to text) first.',
  );
}
