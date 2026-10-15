/**
 * .pptx extraction through jszip plus fast-xml-parser, one unit per slide in
 * PRESENTATION order (the `sldIdLst` in presentation.xml resolved through the
 * rels), not zip order. All `<a:t>` runs are collected — body text, tables,
 * grouped shapes — in spTree document order, and speaker notes are appended
 * with a marker.
 *
 * Why this pair and not a PowerPoint library: there is no maintained pure-JS
 * pptx text library; the format is a zip of small XML parts, and jszip plus a
 * real XML parser is exactly enough to walk it.
 *
 * Graphics resolve instead of warning when they can: a `p:graphicFrame`'s
 * relationship ids are followed through the slide's own `.rels` to
 * `diagrams/data*.xml` (SmartArt `dgm:t` text), `charts/chart*.xml` (cached
 * `c:v` values) and, one hop further through the chart's rels,
 * `embeddings/*.xlsx` (extracted as workbook text, capped). Only genuinely
 * uncached graphics keep the warning naming the rel target, so the loss is
 * explicit rather than silent. Repeated master/footer boilerplate dedupes
 * across slides, and note placeholders (header, slide-image, date) are
 * skipped — notes keep only the speaker's own body text.
 */

import {
  isProtectionError,
  protectionRefusal,
  type ExtractedUnit,
  type ExtractionResult,
} from './extract.js';

export interface PptxZip {
  readText: (path: string) => Promise<string | null>;
  readBytes?: (path: string) => Promise<Uint8Array | null>;
}

export interface PptxDeps {
  openZip: (bytes: Uint8Array) => Promise<PptxZip>;
}

type XmlNode = string | { [key: string]: XmlNode | XmlNode[] | undefined } | XmlNode[];

/** Note placeholder types that are slide furniture, not speaker text. */
const NOTE_SKIP_PH = new Set(['hdr', 'sldImg', 'dt', 'ftr', 'sldNum']);

/** Most embedded-workbook lines kept per graphic; the rest count, not print. */
const MAX_EMBEDDED_LINES = 30;

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

/** Every `<dgm:t>` (SmartArt) or `<c:v>` (cached chart value) string, in order. */
function graphicRuns(node: XmlNode | undefined, tag: string, out: string[]): void {
  if (typeof node === 'string') return;
  if (Array.isArray(node)) {
    for (const entry of node) graphicRuns(entry, tag, out);
    return;
  }
  const record = asRecord(node);
  if (record === undefined) return;
  for (const [key, value] of Object.entries(record)) {
    if (key === tag) {
      for (const run of asArray(value)) {
        if (typeof run === 'string' || typeof run === 'number') {
          const text = String(run).trim();
          if (text.length > 0) out.push(text);
        } else if (typeof run === 'object' && run !== null) {
          const text = (run as Record<string, unknown>)['#text'];
          if (typeof text === 'string' && text.trim().length > 0) out.push(text.trim());
          else if (typeof text === 'number') out.push(String(text));
        }
      }
    } else if (!key.startsWith('@_') && !key.startsWith('?')) {
      graphicRuns(value as XmlNode, tag, out);
    }
  }
}

/** Paragraph breaks: one line per `<a:p>`, runs inside it joined with spaces. */
function paragraphLines(node: XmlNode | undefined): string[] {
  const lines: string[] = [];
  const walk = (current: XmlNode | undefined): void => {
    if (typeof current === 'string') return;
    if (Array.isArray(current)) {
      for (const entry of current) walk(entry);
      return;
    }
    const record = asRecord(current);
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
  walk(node);
  return lines;
}

/** Placeholder types declared anywhere under a shape (`p:ph type="hdr"`). */
function placeholderTypes(shape: XmlNode | undefined): string[] {
  const types: string[] = [];
  const walk = (current: XmlNode | undefined): void => {
    if (typeof current === 'string') return;
    if (Array.isArray(current)) {
      for (const entry of current) walk(entry);
      return;
    }
    const record = asRecord(current);
    if (record === undefined) return;
    for (const [key, value] of Object.entries(record)) {
      if (key === 'p:ph') {
        for (const ph of asArray(value)) {
          const type = asRecord(ph)?.['@_type'];
          if (typeof type === 'string') types.push(type);
        }
      } else if (!key.startsWith('@_') && !key.startsWith('?')) {
        walk(value as XmlNode);
      }
    }
  };
  walk(shape);
  return types;
}

/** Relationship ids referenced under a graphic frame (`r:id`, `r:dm`, ...). */
function frameRelIds(frame: XmlNode | undefined): string[] {
  const ids: string[] = [];
  const walk = (current: XmlNode | undefined): void => {
    if (typeof current === 'string') return;
    if (Array.isArray(current)) {
      for (const entry of current) walk(entry);
      return;
    }
    const record = asRecord(current);
    if (record === undefined) return;
    for (const [key, value] of Object.entries(record)) {
      if ((key === '@_r:id' || key === '@_r:dm' || key === '@_r:lo' || key === '@_r:qs' || key === '@_r:cs') && typeof value === 'string') {
        ids.push(value);
      } else if (!key.startsWith('?')) {
        walk(value as XmlNode);
      }
    }
  };
  walk(frame);
  return [...new Set(ids)];
}

async function defaultDeps(): Promise<PptxDeps & { parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined> }> {
  const zipMod = (await import('jszip')) as unknown as {
    default?: { loadAsync: (bytes: Uint8Array) => Promise<PptxZipFile> };
    loadAsync: (bytes: Uint8Array) => Promise<PptxZipFile>;
  };
  interface PptxZipFile {
    file: (path: string) => { async: (kind: string) => Promise<unknown> } | null;
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
          return (await entry.async('string')) as string;
        },
        readBytes: async (path: string) => {
          const entry = zip.file(path);
          if (entry === null) return null;
          return new Uint8Array((await entry.async('uint8array')) as Uint8Array);
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

/** Resolves a rel target (usually `../`-relative) against its owning part. */
function resolveTarget(part: string, target: string): string {
  if (target.startsWith('ppt/') || target.startsWith('/')) {
    return target.startsWith('ppt/') ? target : `ppt/${target.replace(/^\//, '')}`;
  }
  const dir = part.slice(0, part.lastIndexOf('/'));
  return `${dir}/${target}`.replace(/\/[^/]+\/\.\./g, '');
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

interface GraphicData {
  lines: string[];
  /** Rel targets that pointed at graphics with no extractable values. */
  uncached: string[];
}

/**
 * Follows one graphic frame's rel ids: SmartArt data (`dgm:t`), cached chart
 * values (`c:v`), and embedded workbooks one hop past the chart. Frames with
 * no rel ids (drawingML tables) are not graphics and resolve to nothing.
 */
async function resolveGraphic(
  frame: XmlNode,
  slidePath: string,
  slideRels: Map<string, string>,
  zip: PptxZip,
  parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined>,
): Promise<GraphicData> {
  const lines: string[] = [];
  const uncached: string[] = [];
  const ids = frameRelIds(frame);
  if (ids.length === 0) return { lines, uncached };
  for (const id of ids) {
    const target = slideRels.get(id);
    if (target === undefined) continue;
    const path = resolveTarget(slidePath, target);
    if (/\/diagrams\/data[^/]*\.xml$/.test(path)) {
      const xml = await zip.readText(path);
      const runs: string[] = [];
      if (xml !== null) graphicRuns(parse(xml), 'dgm:t', runs);
      if (runs.length > 0) lines.push(`SmartArt: ${runs.join(' | ')}`);
      else uncached.push(target);
    } else if (/\/charts\/chart[^/]*\.xml$/.test(path)) {
      const xml = await zip.readText(path);
      const runs: string[] = [];
      if (xml !== null) graphicRuns(parse(xml), 'c:v', runs);
      if (runs.length > 0) {
        lines.push(`Chart values: ${runs.join(', ')}`);
      } else {
        uncached.push(target);
      }
      // Embedded workbook behind the chart, when there is one.
      const chartRels = relTargets(parse((await zip.readText(relsFor(path))) ?? '<Relationships/>'));
      for (const [, chartTarget] of chartRels) {
        if (!/embeddings\/[^/]*\.xlsx?m?$/.test(chartTarget)) continue;
        const embeddedPath = resolveTarget(path, chartTarget);
        const data = await zip.readBytes?.(embeddedPath);
        if (data === null || data === undefined) continue;
        try {
          const { extractXlsx } = await import('./extract-xlsx.js');
          const book = await extractXlsx(data, chartTarget.split('/').pop() ?? 'embedded.xlsx');
          const bookLines = book.units.flatMap((unit) => [`[${unit.id}]`, ...unit.text.split('\n')]);
          const kept = bookLines.slice(0, MAX_EMBEDDED_LINES);
          if (bookLines.length > MAX_EMBEDDED_LINES) {
            kept.push(`… [${bookLines.length - MAX_EMBEDDED_LINES} more rows in embedded workbook]`);
          }
          if (kept.length > 0) lines.push(`Embedded workbook:\n${kept.join('\n')}`);
        } catch {
          // A corrupt embedding must not fail the slide; the chart warning stands.
          if (!uncached.includes(target)) uncached.push(target);
        }
      }
    }
  }
  // A frame whose rels point at neither charts nor diagrams (a picture, an
  // OLE object) is not a graphic this pass claims: no lines, no warning.
  return { lines, uncached };
}

/**
 * Slide lines in spTree document order — modulo fast-xml-parser, which groups
 * same-named children, so all `p:sp` paragraphs precede all `p:graphicFrame`
 * resolutions when the two interleave (the old walk had the same grouping).
 * Graphic frames resolve to their cached values; everything else contributes
 * its paragraphs. Tables (`a:tbl` frames without rel ids) keep the old
 * behaviour of flattening to paragraph lines.
 */
async function slideLines(
  slide: Record<string, XmlNode | XmlNode[] | undefined>,
  slidePath: string,
  slideRels: Map<string, string>,
  zip: PptxZip,
  parse: (xml: string) => Record<string, XmlNode | XmlNode[] | undefined>,
): Promise<{ lines: string[]; uncached: string[] }> {
  const lines: string[] = [];
  const uncached: string[] = [];
  const walk = async (node: XmlNode | undefined): Promise<void> => {
    if (typeof node === 'string') return;
    if (Array.isArray(node)) {
      for (const entry of node) await walk(entry);
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
      } else if (key === 'p:graphicFrame' || key === 'p:graphicframe') {
        for (const frame of asArray(value)) {
          const graphic = await resolveGraphic(frame, slidePath, slideRels, zip, parse);
          if (graphic.lines.length > 0) {
            lines.push(...graphic.lines);
          } else {
            // Tables and unresolvable frames keep their paragraph text.
            lines.push(...paragraphLines(frame));
          }
          uncached.push(...graphic.uncached);
        }
      } else if (!key.startsWith('@_') && !key.startsWith('?')) {
        await walk(value as XmlNode);
      }
    }
  };
  await walk(slide);
  return { lines, uncached: [...new Set(uncached)] };
}

/**
 * Note lines minus slide furniture: shapes whose placeholder type is a
 * header, slide image, date, footer or slide number are skipped, so only the
 * speaker's body text survives.
 */
function noteLines(notes: Record<string, XmlNode | XmlNode[] | undefined>): string[] {
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
      if (key === 'p:sp' || key === 'p:cxnSp') {
        for (const shape of asArray(value)) {
          if (placeholderTypes(shape).some((type) => NOTE_SKIP_PH.has(type))) continue;
          lines.push(...paragraphLines(shape));
        }
      } else if (!key.startsWith('@_') && !key.startsWith('?')) {
        walk(value as XmlNode);
      }
    }
  };
  walk(notes);
  return lines;
}

/**
 * Drops boilerplate: a short line repeated on more than half of the slides
 * (three-slide minimum, so a two-slide deck never eats a shared title) is
 * footer/master furniture and is kept only on its first slide. Titles are
 * fixed before this runs, so the outline never loses an entry.
 */
function dedupeBoilerplate(units: ExtractedUnit[]): void {
  if (units.length < 3) return;
  const firstSeen = new Map<string, number>();
  const slidesWith = new Map<string, number>();
  for (const [index, unit] of units.entries()) {
    const seen = new Set<string>();
    for (const line of unit.text.split('\n')) {
      if (line.length === 0 || line.length > 100 || seen.has(line)) continue;
      seen.add(line);
      slidesWith.set(line, (slidesWith.get(line) ?? 0) + 1);
      if (!firstSeen.has(line)) firstSeen.set(line, index);
    }
  }
  const boilerplate = new Set<string>();
  for (const [line, count] of slidesWith) {
    if (count > units.length / 2) boilerplate.add(line);
  }
  if (boilerplate.size === 0) return;
  for (const [index, unit] of units.entries()) {
    const kept = unit.text
      .split('\n')
      .filter((line) => !boilerplate.has(line) || firstSeen.get(line) === index);
    units[index] = { ...unit, text: kept.join('\n') };
  }
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
    const slidePath = resolveTarget('ppt/presentation.xml', target);
    const slideXml = await zip.readText(slidePath);
    if (slideXml === null) {
      warnings.push(`Slide ${position} (${target}) is missing from the package and was skipped.`);
      continue;
    }
    const slide = parse(slideXml);
    const slideRels = relTargets(parse((await zip.readText(relsFor(slidePath))) ?? '<Relationships/>'));
    const { lines, uncached } = await slideLines(slide, slidePath, slideRels, zip, parse);

    // Speaker notes ride in a separate part linked from the slide's own rels.
    for (const [, relTarget] of slideRels) {
      if (!relTarget.includes('notesSlide')) continue;
      const notesPath = resolveTarget(slidePath, relTarget);
      const notesXml = await zip.readText(notesPath);
      if (notesXml === null) continue;
      const noteText = noteLines(parse(notesXml));
      if (noteText.length > 0) lines.push(`Speaker notes: ${noteText.join(' ')}`);
    }

    // Charts and SmartArt without cached values keep the explicit warning.
    const hasGraphicMarkup = /<c:chart\b/.test(slideXml) || /<dgm:/.test(slideXml);
    const graphicRefs = [...slideRels.values()].filter((t) => /\/(charts|diagrams)\//.test(t));
    if (uncached.length > 0) {
      warnings.push(
        `Slide ${position} references a chart or SmartArt graphic (${uncached.join(', ')}); ` +
          'values shown in the graphic are not extracted.',
      );
    } else if (graphicRefs.length > 0 || hasGraphicMarkup) {
      const cached = lines.some(
        (line) =>
          line.startsWith('Chart values:') || line.startsWith('SmartArt:') || line.startsWith('Embedded workbook:'),
      );
      if (!cached) {
        const ref = graphicRefs.length > 0 ? ` (${graphicRefs.join(', ')})` : '';
        warnings.push(
          `Slide ${position} references a chart or SmartArt graphic${ref}; ` +
            'values shown in the graphic are not extracted.',
        );
      }
    }

    units.push({
      id: typeof slideId === 'string' ? slideId : `slide-${position}`,
      title: lines[0] === undefined ? `Slide ${position}` : lines[0],
      text: lines.join('\n'),
    });
  }

  if (units.length === 0) {
    throw new Error(`"${fileName}" holds no slides, so there is nothing to extract.`);
  }
  dedupeBoilerplate(units);
  return { units, warnings };
}
