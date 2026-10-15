/**
 * Referenced-document parsing: the orchestration behind `files_parse_document`.
 *
 * A document reaches the model from three places — a sharing link (Teams
 * message `contentUrl`, mail reference attachment), a drive address
 * (`driveId` + `itemId`, including the `remoteDriveId` stubs the files tools
 * return), or a mail attachment — and always leaves as the same shape: unit
 * ROWS (pages/slides/sheets/sections) plus a manifest that accounts for every
 * unit. That manifest is the no-omission contract: `unitCount` with per-unit
 * `truncated` flags and a `reason` from the paginator's enum, so the model
 * can tell "the document ends here" from "call again with the cursor".
 *
 * The byte fetch and the extractors ride behind `ParseDeps` so tests stub
 * them; `files.ts` wires the real Graph-backed implementations.
 */

import type { GraphClient } from '../contracts.js';
import { resolveShare } from '../graph/shares.js';
import {
  MAX_DOCUMENT_BYTES,
  detectKind,
  type DocumentKind,
  type ExtractionResult,
} from './extract.js';
import { extractDocx } from './extract-docx.js';
import { extractXlsx } from './extract-xlsx.js';
import { extractPptx } from './extract-pptx.js';
import { extractPdf } from './extract-pdf.js';
import { truncateText } from './truncate.js';

// The manifest `reason` reuses the paginator's vocabulary (paginate.ts): 'complete'
// means the document ran out, 'maxItems' means units remain behind `cursor`.
export type DocumentSource =
  | { shareUrl: string }
  | { driveId: string; itemId: string }
  | { messageId: string; attachmentId: string };

export interface ParseOptions {
  source: DocumentSource;
  /** First unit to return. Defaults to 0. */
  unitOffset?: number;
  /** Most units to return. Defaults to 20. */
  unitLimit?: number;
  /** Most characters per unit. Defaults to 6000. */
  maxCharsPerUnit?: number;
}

export interface FetchBytes {
  (opts: { path: string; scopes: string[]; offset: number; byteBudget: number }): Promise<{
    bytes: Uint8Array;
    nextOffset: number;
    totalLength?: number | undefined;
  }>;
}

export interface ParseDeps {
  graph: GraphClient;
  fileScopes: string[];
  mailScopes: string[];
  fetchBytes: FetchBytes;
  extract?: (kind: DocumentKind, bytes: Uint8Array, fileName: string) => Promise<ExtractionResult>;
}

export interface ParsedUnit {
  /** 0-based index into the whole document, stable across pages. */
  i: number;
  id: string;
  title?: string | undefined;
  /** Characters of `text` actually returned (post-truncation). */
  chars: number;
  truncated: boolean;
  text: string;
}

export interface ParseManifest {
  units: Array<{ i: number; id: string; chars: number; truncated: boolean }>;
  /**
   * The whole-document outline: every unit's index, id and title, not just
   * the returned slice, so the model can survey the document and fetch only
   * the units it needs. Additive: nothing else in the manifest moved.
   */
  titles: Array<{ i: number; id: string; title?: string | undefined }>;
  unitCount: number;
  returnedRange: { from: number; to: number } | null;
  reason: 'complete' | 'maxItems';
  cursor?: string | undefined;
}

export interface ParseResult {
  fileName: string;
  mimeType?: string | undefined;
  documentKind: DocumentKind;
  unitKind: 'section' | 'sheet' | 'slide' | 'page';
  driveId?: string | undefined;
  itemId?: string | undefined;
  units: ParsedUnit[];
  manifest: ParseManifest;
  warnings: string[];
}

export interface ParseCursor {
  v: 1;
  source: DocumentSource;
  unitOffset: number;
  unitLimit: number;
  maxCharsPerUnit: number;
}

const DEFAULT_UNIT_LIMIT = 20;
const DEFAULT_MAX_CHARS_PER_UNIT = 6000;
const MAX_UNIT_LIMIT = 50;

const UNIT_KIND: Record<DocumentKind, ParseResult['unitKind']> = {
  docx: 'section',
  xlsx: 'sheet',
  pptx: 'slide',
  pdf: 'page',
};

export function encodeParseCursor(cursor: ParseCursor): string {
  return Buffer.from(JSON.stringify(cursor), 'utf8').toString('base64url');
}

export function decodeParseCursor(cursor: string): ParseCursor {
  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(cursor, 'base64url').toString('utf8'));
  } catch {
    throw new Error('Invalid cursor: pass the `cursor` from a previous files_parse_document result verbatim.');
  }
  const record = parsed as Partial<ParseCursor>;
  if (record === null || typeof record !== 'object' || record.v !== 1 || record.source === undefined) {
    throw new Error('Invalid cursor: pass the `cursor` from a previous files_parse_document result verbatim.');
  }
  return {
    v: 1,
    source: record.source as DocumentSource,
    unitOffset: Math.max(0, Math.trunc(record.unitOffset ?? 0)),
    unitLimit: Math.min(MAX_UNIT_LIMIT, Math.max(1, Math.trunc(record.unitLimit ?? DEFAULT_UNIT_LIMIT))),
    maxCharsPerUnit: Math.max(200, Math.trunc(record.maxCharsPerUnit ?? DEFAULT_MAX_CHARS_PER_UNIT)),
  };
}

export async function defaultExtract(
  kind: DocumentKind,
  bytes: Uint8Array,
  fileName: string,
): Promise<ExtractionResult> {
  switch (kind) {
    case 'docx':
      return extractDocx(bytes, fileName);
    case 'xlsx':
      return extractXlsx(bytes, fileName);
    case 'pptx':
      return extractPptx(bytes, fileName);
    case 'pdf':
      return extractPdf(bytes, fileName);
  }
}

interface DriveItemPayload {
  id?: string | null;
  name?: string | null;
  size?: number | null;
  file?: { mimeType?: string | null } | null;
  folder?: unknown | null;
  package?: { type?: string | null } | null;
}

interface AttachmentPayload {
  id?: string | null;
  name?: string | null;
  contentType?: string | null;
  size?: number | null;
  contentBytes?: string | null;
  sourceUrl?: string | null;
  '@odata.type'?: string | null;
}

function seg(id: string, what: string): string {
  if (/[/\\%]/.test(id)) {
    throw new Error(`${what} ${JSON.stringify(id)} is not a bare id: pass it exactly as a listing tool returned it.`);
  }
  return encodeURIComponent(id);
}

async function driveBytes(
  deps: ParseDeps,
  driveId: string,
  itemId: string,
): Promise<{ fileName: string; mimeType?: string; bytes: Uint8Array; driveId: string; itemId: string }> {
  const meta = await deps.graph.request<DriveItemPayload>({
    path: `/drives/${seg(driveId, 'driveId')}/items/${seg(itemId, 'itemId')}`,
    query: { $select: 'id,name,size,file,folder,package' },
    scopes: deps.fileScopes,
  });
  const item = meta.data ?? {};
  const fileName = item.name ?? itemId;
  if (item.folder !== null && item.folder !== undefined) {
    throw new Error(`"${fileName}" is a folder, not a document. List it with files_list_children instead.`);
  }
  if (item.package !== null && item.package !== undefined) {
    throw new Error(
      `"${fileName}" is a ${item.package.type ?? 'package'} package (such as a OneNote notebook), which has no document text to parse.`,
    );
  }
  const size = typeof item.size === 'number' ? item.size : undefined;
  if (size !== undefined && size > MAX_DOCUMENT_BYTES) {
    throw new Error(
      `"${fileName}" is ${(size / 1048576).toFixed(1)} MiB; documents over ${MAX_DOCUMENT_BYTES / 1048576} MiB cannot be parsed here.`,
    );
  }
  const fetched = await deps.fetchBytes({
    path: `/drives/${seg(driveId, 'driveId')}/items/${seg(itemId, 'itemId')}/content`,
    scopes: deps.fileScopes,
    offset: 0,
    byteBudget: MAX_DOCUMENT_BYTES,
  });
  if (fetched.bytes.length >= MAX_DOCUMENT_BYTES && (fetched.totalLength ?? fetched.bytes.length) > MAX_DOCUMENT_BYTES) {
    throw new Error(`"${fileName}" exceeds the ${MAX_DOCUMENT_BYTES / 1048576} MiB parse limit.`);
  }
  return {
    fileName,
    mimeType: item.file?.mimeType ?? undefined,
    bytes: fetched.bytes,
    driveId,
    itemId: item.id ?? itemId,
  };
}

async function mailBytes(
  deps: ParseDeps,
  messageId: string,
  attachmentId: string,
): Promise<{ fileName: string; mimeType?: string; bytes: Uint8Array }> {
  const meta = await deps.graph.request<AttachmentPayload>({
    path: `/me/messages/${seg(messageId, 'messageId')}/attachments/${seg(attachmentId, 'attachmentId')}`,
    query: { $select: 'id,name,contentType,size' },
    scopes: deps.mailScopes,
  });
  const item = meta.data ?? {};
  const fileName = item.name ?? attachmentId;
  // @odata.type is OData control information, not a property: Graph returns it
  // alongside any $select, so branching on it here is safe (verified live).
  const kind = (item['@odata.type'] ?? '').replace('#microsoft.graph.', '');
  if (kind === 'referenceAttachment') {
    throw new Error(
      `"${fileName}" is a reference attachment (a link, not a file). Pass its sourceUrl as shareUrl to files_parse_document instead — list it with mail_get_attachment_content.`,
    );
  }
  if (kind !== '' && kind !== 'fileAttachment') {
    throw new Error(
      `"${fileName}" is an Outlook ${kind || 'item'} attachment, not a file, so there is no document to parse.`,
    );
  }
  const size = typeof item.size === 'number' ? item.size : undefined;
  if (size !== undefined && size > MAX_DOCUMENT_BYTES) {
    throw new Error(
      `"${fileName}" is ${(size / 1048576).toFixed(1)} MiB; documents over ${MAX_DOCUMENT_BYTES / 1048576} MiB cannot be parsed here.`,
    );
  }
  const fetched = await deps.fetchBytes({
    path: `/me/messages/${seg(messageId, 'messageId')}/attachments/${seg(attachmentId, 'attachmentId')}/$value`,
    scopes: deps.mailScopes,
    offset: 0,
    byteBudget: MAX_DOCUMENT_BYTES,
  });
  if (fetched.bytes.length >= MAX_DOCUMENT_BYTES && (fetched.totalLength ?? fetched.bytes.length) > MAX_DOCUMENT_BYTES) {
    throw new Error(`"${fileName}" exceeds the ${MAX_DOCUMENT_BYTES / 1048576} MiB parse limit.`);
  }
  return { fileName, mimeType: item.contentType ?? undefined, bytes: fetched.bytes };
}

export async function parseDocument(options: ParseOptions, deps: ParseDeps): Promise<ParseResult> {
  const unitOffset = Math.max(0, Math.trunc(options.unitOffset ?? 0));
  const unitLimit = Math.min(MAX_UNIT_LIMIT, Math.max(1, Math.trunc(options.unitLimit ?? DEFAULT_UNIT_LIMIT)));
  const maxCharsPerUnit = Math.max(200, Math.trunc(options.maxCharsPerUnit ?? DEFAULT_MAX_CHARS_PER_UNIT));
  const extract = deps.extract ?? defaultExtract;

  const source = options.source;
  let resolved: { fileName: string; mimeType?: string; bytes: Uint8Array; driveId?: string; itemId?: string };
  if ('shareUrl' in source) {
    const shared = await resolveShare(deps.graph, deps.fileScopes, source.shareUrl);
    if (shared.kind === 'folder') {
      throw new Error('The sharing link points at a folder, not a document. List it with files_list_children instead.');
    }
    resolved = await driveBytes(deps, shared.driveId, shared.itemId);
  } else if ('driveId' in source) {
    resolved = await driveBytes(deps, source.driveId, source.itemId);
  } else {
    resolved = await mailBytes(deps, source.messageId, source.attachmentId);
  }

  const kind = detectKind(resolved.fileName, resolved.mimeType, resolved.bytes);
  const extracted = await extract(kind, resolved.bytes, resolved.fileName);

  const sliced = extracted.units.slice(unitOffset, unitOffset + unitLimit);
  const units: ParsedUnit[] = sliced.map((unit, index) => {
    const text = truncateText(unit.text, maxCharsPerUnit);
    const truncated = text.length !== unit.text.length;
    return {
      i: unitOffset + index,
      id: unit.id,
      ...(unit.title === undefined ? {} : { title: unit.title }),
      chars: text.length,
      truncated: truncated ? true : false,
      text,
    };
  });
  const done = unitOffset + sliced.length >= extracted.units.length;
  const reason = done ? 'complete' : 'maxItems';
  const manifest: ParseManifest = {
    units: units.map(({ i, id, chars, truncated }) => ({ i, id, chars, truncated })),
    titles: extracted.units.map((unit, i) => ({
      i,
      id: unit.id,
      ...(unit.title === undefined ? {} : { title: unit.title }),
    })),
    unitCount: extracted.units.length,
    returnedRange: units.length === 0 ? null : { from: units[0]?.i ?? 0, to: (units[units.length - 1]?.i ?? 0) },
    reason,
    ...(done
      ? {}
      : {
          cursor: encodeParseCursor({
            v: 1,
            source,
            unitOffset: unitOffset + sliced.length,
            unitLimit,
            maxCharsPerUnit,
          }),
        }),
  };
  return {
    fileName: resolved.fileName,
    ...(resolved.mimeType === undefined ? {} : { mimeType: resolved.mimeType }),
    documentKind: kind,
    unitKind: UNIT_KIND[kind],
    ...(resolved.driveId === undefined ? {} : { driveId: resolved.driveId }),
    ...(resolved.itemId === undefined ? {} : { itemId: resolved.itemId }),
    units,
    manifest,
    warnings: extracted.warnings,
  };
}

