/**
 * .pptx extraction through jszip plus fast-xml-parser, one unit per slide in
 * PRESENTATION order (the `sldIdLst` in presentation.xml resolved through the
 * rels), not zip order. All `<a:t>` runs are collected — body text, tables,
 * grouped shapes — and speaker notes are appended with a marker.
 *
 * Why this pair and not a PowerPoint library: there is no maintained pure-JS
 * pptx text library; the format is a zip of small XML parts, and jszip plus a
 * real XML parser is exactly enough to walk it. Chart and SmartArt VALUES are
 * not extracted — their XML holds series references, not numbers — and each
 * occurrence is reported as a warning naming the rel target, so the loss is
 * explicit rather than silent.
 */

import {
  isProtectionError,
  protectionRefusal,
  type ExtractedUnit,
  type ExtractionResult,
} from './extract.js';

export interface PptxZip {
  readText: (path: string) => Promise<string | null>;
}

export interface PptxDeps {
  openZip: (bytes: Uint8Array) => Promise<PptxZip>;
}

type XmlNode = string | { [key: string]: XmlNode | XmlNode[] | undefined } | XmlNode[];

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

/** Every `<a:t>` string under the node, in document order. */
function textRuns(node: XmlNode | undefined, out: string[]): void {
  if (typeof node === 'string') return;
  if (Array.isArray(node)) {
    for (const entry of node) textRuns(entry, out);
    return;
  }
  const record = asRecord(node);
  if (record === undefined) return;
  for (const [key, value] of Object.entries(record)) {
    if (key === 'a:t') {
      for (const run of asArray(value)) {
        if (typeof run === 'string' && run.length > 0) out.push(run);
        else if (typeof run === 'object' && run !== null) {
          // fast-xml-parser may hang the string under #text when attributes mix in.
          const text = (run as Record<string, unknown>)['#text'];
          if (typeof text === 'string' && text.length > 0) out.push(text);
        }
      }
    } else if (!key.startsWith('@_') && !key.startsWith('?')) {
      textRuns(value as XmlNode, out);
    }
  }
}

/** Paragraph breaks: one line per `<a:p>`, runs inside it joined with spaces. */
function paragraphs(slide: Record<string, XmlNode | XmlNode[] | undefined>): string[] {
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
      if (key === 'a:p') {
        for (const para of asArray(value)) {
          const runs: string[] = [];
          textRuns(para, runs);
          const line = runs.join(' ').replace(/\s+/g, ' ').trim();
          if (line.length > 0) lines.push(line);
        }
      } else if (!key.startsWith('@_') && !key.startsWith('?')) {
        walk(value as XmlNode);
      }
    }
  };
  walk(slide);
  return lines;
}

async function defaultDeps(): Promise<PptxDeps & { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> }> {
  const zipMod = (await import('jszip')) as unknown as {
    default?: { loadAsync: (bytes: Uint8Array) => Promise<PptxZipFile> };
    loadAsync: (bytes: Uint8Array) => Promise<PptxZipFile>;
  };
  interface PptxZipFile {
    file: (path: string) => { async: (kind: 'string') => Promise<string> } | null;
  }
  const loadAsync = zipMod.default?.loadAsync ?? zipMod.loadAsync;
  const xmlMod = (await import('fast-xml-parser')) as unknown as {
    default?: { XMLParser: new (options: Record<string, unknown>) => { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> } };
    XMLParser: new (options: Record<string, unknown>) => { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> };
  };
  const XMLParser = xmlMod.default?.XMLParser ?? xmlMod.XMLParser;
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', trimValues: false });
  return {
    openZip: async (bytes: Uint8Array) => {
      const zip = await loadAsync(bytes);
      return {
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

/** `ppt/slides/slide1.xml` -> `ppt/slides/_rels/slide1.xml.rels`. */
function relsFor(part: string): string {
  const slash = part.lastIndexOf('/');
  return slash === -1 ? `_rels/${part}.rels` : `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
}

/** Id -> Target for a `.rels` part. */
function relTargets(
  rels: Record<string, XmlNode | XmlNode[] | undefined> | null,
): Map<string, string> {
  const targets = new Map<string, string>();
  const root = rels === null ? undefined : asRecord(rels['Relationships']);
  for (const rel of asArray(root?.['Relationship'])) {
    const record = asRecord(rel);
    const id = record?.['@_Id'];
    const target = record?.['@_Target'];
    if (typeof id === 'string' && typeof target === 'string') targets.set(id, target);
  }
  return targets;
}

export async function extractPptx(
  bytes: Uint8Array,
  fileName: string,
  deps?: PptxDeps,
): Promise<ExtractionResult> {
  const impl = deps ?? (await defaultDeps());
  const parse =
    'parse' in impl && typeof (impl as { parse?: unknown }).parse === 'function'
      ? ((impl as unknown as { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> }).parse)
      : (await defaultDeps()).parse;

  let zip: PptxZip;
  try {
    zip = await impl.openZip(bytes);
  } catch (error) {
    if (isProtectionError(error)) throw protectionRefusal(fileName);
    throw new Error(`"${fileName}" is not a readable .pptx file: ${error instanceof Error ? error.message : String(error)}`);
  }

  const presentationXml = await zip.readText('ppt/presentation.xml');
  if (presentationXml === null) {
    throw new Error(`"${fileName}" has no ppt/presentation.xml, so it is not a readable .pptx file.`);
  }
  const presentation = parse(presentationXml);
  const sldIds = asArray(asRecord(presentation['p:presentation'])?.['p:sldIdLst']).flatMap((lst) =>
    asArray(asRecord(lst)?.['p:sldId']),
  );
  const presRels = relTargets(parse((await zip.readText('ppt/_rels/presentation.xml.rels')) ?? '<Relationships/>'));

  const warnings: string[] = [];
  const units: ExtractedUnit[] = [];
  let position = 0;
  for (const entry of sldIds) {
    const record = asRecord(entry);
    const rid = record?.['@_r:id'];
    const slideId = record?.['@_id'];
    if (typeof rid !== 'string') continue;
    const target = presRels.get(rid);
    if (target === undefined) continue;
    position += 1;
    const slidePath = target.startsWith('ppt/') ? target : `ppt/${target.replace(/^\//, '')}`;
    const slideXml = await zip.readText(slidePath);
    if (slideXml === null) {
      warnings.push(`Slide ${position} (${target}) is missing from the package and was skipped.`);
      continue;
    }
    const slide = parse(slideXml);
    const lines = paragraphs(slide);

    // Speaker notes ride in a separate part linked from the slide's own rels.
    const slideRels = relTargets(parse((await zip.readText(relsFor(slidePath))) ?? '<Relationships/>'));
    for (const [, relTarget] of slideRels) {
      if (!relTarget.includes('notesSlide')) continue;
      const dir = slidePath.slice(0, slidePath.lastIndexOf('/'));
      const notesPath = `${dir}/${relTarget.replace(/^\//, '')}`.replace(/\/[^/]+\/\.\./g, '');
      const notesXml = await zip.readText(notesPath);
      if (notesXml === null) continue;
      const noteLines = paragraphs(parse(notesXml));
      if (noteLines.length > 0) lines.push(`Speaker notes: ${noteLines.join(' ')}`);
    }

    // Charts and SmartArt hold references where the numbers should be.
    const graphicRefs: string[] = [];
    for (const [, relTarget] of slideRels) {
      if (/\/(charts|diagrams)\//.test(relTarget)) graphicRefs.push(relTarget);
    }
    if (/<c:chart\b/.test(slideXml) || /<dgm:/.test(slideXml) || graphicRefs.length > 0) {
      const ref = graphicRefs.length > 0 ? ` (${graphicRefs.join(', ')})` : '';
      warnings.push(`Slide ${position} references a chart or SmartArt graphic${ref}; values shown in the graphic are not extracted.`);
    }

    units.push({
      id: typeof slideId === 'string' ? slideId : `slide-${position}`,
      ...(lines[0] === undefined ? {} : { title: lines[0] }),
      text: lines.join('\n'),
    });
  }

  if (units.length === 0) {
    throw new Error(`"${fileName}" holds no slides, so there is nothing to extract.`);
  }
  return { units, warnings };
}
