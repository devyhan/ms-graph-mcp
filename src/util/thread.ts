/**
 * Folding a reply chain down to what is actually new in it.
 *
 * A corporate Outlook thread quotes itself. Message N carries most of messages
 * 1..N-1 inside its own body, so fetching a thread one message at a time sends
 * the same paragraphs through the model's context once per reply. Measured on a
 * real 24-message thread: 552 blocks of text, 393 of them (71%) byte-identical
 * to a block already seen in an earlier message.
 *
 * Everything here is deliberately pure — no Graph, no SDK, no Node built-ins —
 * because the two failures worth guarding against are both testable offline:
 * folding text that was not actually a repeat, and returning different bytes for
 * the same input.
 */

/** The value a folded block is replaced with, and what a caller can search for. */
const POINTER_OPEN = '[quoted from ';
const POINTER_CLOSE = ']';

/**
 * Below this many characters (after normalisation) a block is emitted verbatim
 * and never registered as foldable.
 *
 * Short blocks are greetings, sign-offs and single-line acknowledgements. They
 * repeat constantly, and folding them buys back fewer characters than the
 * pointer costs while making the result unreadable to a person.
 */
const MIN_BLOCK_CHARS = 40;

/**
 * A URL longer than this is treated as machine plumbing rather than something a
 * reader would follow.
 *
 * The threshold is measured, not guessed. In the sample corpus the click-tracking
 * redirector emitted 56 URLs of 498-843 characters while every genuinely cited
 * link — a Figma file, a Google Sheet, an API doc — ran 46-130. The two
 * populations do not overlap, and an earlier 60-character rule destroyed three
 * real links. 200 sits in the empty middle.
 */
export const TRACKING_LINK_MIN_CHARS = 200;

/**
 * Length alone is not enough, and the near-miss is expensive.
 *
 * Measured longest unbroken opaque run, by class: genuinely cited links 26-44
 * characters, a Teams meeting-join link 125, a SharePoint sharing link 79, and
 * the click-tracking redirector 646-810. Length puts Teams (247 chars) and
 * SharePoint (259) uncomfortably close to the 200 above; this separates them by
 * a factor of five. A URL carrying a single unbroken 200-character token is
 * machine plumbing — nobody typed it and nobody reads it.
 */
const OPAQUE_RUN_MIN_CHARS = 200;

/**
 * Zero-width and bidirectional formatting characters. Notification mail padded
 * with these carries hundreds of them; they render as nothing and tokenise as
 * something.
 */
// U+200D (zero-width JOINER) is excluded: it is load-bearing inside an emoji
// sequence, and deleting it turns one glyph into several. U+2028/U+2029 are
// handled separately below — they are line breaks, not decoration.
const INVISIBLE = new RegExp('[\\u00ad\\u034f\\u200b\\u200c\\u200e\\u200f\\u2060\\ufeff]', 'g');
const LINE_SEPARATOR = new RegExp('[\\u2028\\u2029]', 'g');

/** Non-breaking space. Not invisible — it must become a plain space, not vanish. */
const NBSP = new RegExp('\\u00a0', 'g');

/**
 * Quote separators, in the two renderings Exchange actually produces.
 *
 * The set is deliberately NOT the union of every header line. Splitting on
 * `Sent:`, `To:` and `Cc:` as well shatters one header block into four, each too
 * short to register, and measured 5 percentage points WORSE than this set. What
 * works is splitting where a quoted message begins, plus at `Subject:` — which
 * separates the routing header from the quoted body, giving two units that each
 * repeat cleanly.
 *
 * A separator this misses costs compression, never data: the block simply fails
 * to match anything and is emitted whole. That asymmetry is the reason this
 * module folds blocks rather than trimming a quoted tail.
 *
 * Korean patterns use `\s*` between syllables on purpose. Outlook emits
 * `보낸 사람:` with a non-breaking space often enough that a literal space
 * here would silently never match — normalisation converts it first, but the
 * looser pattern costs nothing and survives a caller that skips normalisation.
 */
const SEPARATOR = new RegExp(
  '(?=^[ \\t]*(?:' +
    [
      'From[ \\t]*:',
      '보낸\\s*사람[ \\t]*:',
      'Subject[ \\t]*:',
      '주제[ \\t]*:',
      '_{5,}',
      '-{2,}[ \\t]*Original Message',
      // Gmail and Apple Mail. Absent from the Outlook corpus this was tuned on,
      // and measured to change nothing there, but a mailbox with outside
      // correspondents will see them.
      'On .{0,100}wrote[ \\t]*:',
      '.{0,60}님이[ \\t]*작성',
    ].join('|') +
    '))',
  'm',
);

/** Matches a URL up to the first character that could not be part of one. */
function urlPattern(): RegExp {
  // Built per call: a module-level /g regex carries lastIndex between callers.
  return /https?:\/\/[^\s<>()[\]"'`]+/g;
}

// ---------------------------------------------------------------------------
// Text normalisation
// ---------------------------------------------------------------------------

export interface NormalizedBody {
  text: string;
  /** Invisible characters removed. Reported so a caller can see it happened. */
  invisibleRemoved: number;
}

/**
 * Strips invisible characters and folds non-breaking spaces to plain ones.
 *
 * Runs before both splitting and hashing. Two blocks that differ only by a
 * non-breaking space are the same block to a reader, and must be the same block
 * to the fold or half the repeats go unmatched.
 */
export function normalizeBodyText(input: string): NormalizedBody {
  if (typeof input !== 'string' || input.length === 0) {
    return { text: '', invisibleRemoved: 0 };
  }
  const invisible = input.match(INVISIBLE);
  const text = input.replace(INVISIBLE, '').replace(LINE_SEPARATOR, '\n').replace(NBSP, ' ');
  return { text, invisibleRemoved: invisible === null ? 0 : invisible.length };
}

export interface LinkFoldResult {
  text: string;
  folded: number;
}

/**
 * Replaces opaque tracking URLs with a marker that keeps the host.
 *
 * The host is kept deliberately. "A Figma link was here" is a fact a summariser
 * needs; an anonymous `[link]` throws it away, which is how the first version of
 * this rule lost three cited references.
 */
export function foldTrackingLinks(
  text: string,
  minChars: number = TRACKING_LINK_MIN_CHARS,
): LinkFoldResult {
  if (typeof text !== 'string' || text.length === 0) return { text: '', folded: 0 };
  const limit = Number.isFinite(minChars) && minChars > 0 ? Math.floor(minChars) : Infinity;

  let folded = 0;
  const out = text.replace(urlPattern(), (url) => {
    if (url.length <= limit || longestOpaqueRun(url) < OPAQUE_RUN_MIN_CHARS) return url;
    folded += 1;
    return `[tracking-link: ${hostOf(url)}, ${url.length} chars]`;
  });
  return { text: out, folded };
}

/** The longest run of unbroken token characters after the host. */
function longestOpaqueRun(url: string): number {
  const afterScheme = url.indexOf('://');
  const rest = afterScheme === -1 ? url : url.slice(afterScheme + 3);
  let longest = 0;
  for (const run of rest.match(/[A-Za-z0-9._~%+=-]+/g) ?? []) {
    if (run.length > longest) longest = run.length;
  }
  return longest;
}

/** Host of a URL without constructing a URL object, which throws on odd input. */
function hostOf(url: string): string {
  const afterScheme = url.indexOf('://');
  if (afterScheme === -1) return 'unknown';
  const rest = url.slice(afterScheme + 3);
  const end = rest.search(/[/?#]/);
  const authority = (end === -1 ? rest : rest.slice(0, end)).trim();
  // Drop any userinfo. `https://teams.microsoft.com@evil.example/...` is a real
  // URL whose host is evil.example, and a marker naming the part before the `@`
  // would put a trusted name in front of a reader who cannot see the original.
  const at = authority.lastIndexOf('@');
  const host = at === -1 ? authority : authority.slice(at + 1);
  return host === '' ? 'unknown' : host;
}

// ---------------------------------------------------------------------------
// Block splitting and folding
// ---------------------------------------------------------------------------

/**
 * Splits a body at quote boundaries, keeping the separator at the head of the
 * block it introduces so a `From:`/`Subject:` header stays attached to what it
 * announces. Blocks that are only whitespace are dropped.
 */
export function splitQuotedBlocks(text: string): string[] {
  if (typeof text !== 'string' || text.trim() === '') return [];
  return text.split(SEPARATOR).filter((block) => block.trim() !== '');
}

/**
 * The comparison key for a block.
 *
 * Mail addresses picked up an `<mailto:…>` annotation from Exchange's text
 * rendering, and clients add or drop it as a thread is quoted onward — in a
 * measured audit of 393 folds, every single non-identical pair differed by
 * nothing else.
 *
 * URLs are deliberately NOT stripped. An earlier version dropped them all, on
 * the theory that a tracking link is regenerated per delivery and two copies of
 * one footer would otherwise never match. That was wrong twice over: the
 * tracking population is already rewritten to a marker carrying its length, so
 * those footers do not match either way, and erasing the rest meant two notices
 * differing only by a short cited link — two build logs, two documents — folded
 * into one, sending the reader to the wrong URL.
 *
 * The normalised string IS the key. Hashing it would add a collision that
 * substitutes one block's text for a pointer to an unrelated block — silently
 * wrong data in exchange for nothing at this scale.
 */
function blockKey(block: string): string {
  const stripped = stripLeadingHeaders(block);
  // A block that is nothing BUT routing headers keeps them: it is a real,
  // repeating unit of a quote chain, and stripping it to nothing would put it
  // under the minimum and stop it folding at all.
  const body = stripped.trim() === '' ? block : stripped;
  return body
    .replace(/<mailto:[^>]*>/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const HEADER_LINE = new RegExp(
  '^[ \\t]*(?:From|Sent|To|Cc|Bcc|Subject|보낸\\s*사람|보낸\\s*날짜|받는\\s*사람|참조|주제|제목|날짜)[ \\t]*:',
);

/**
 * Drops the routing header at the top of a quoted block before comparing it.
 *
 * A message quoted into a reply arrives wrapped in `Subject: …`; the original it
 * was quoted from has no such line. Without this the two never match, so the
 * first copy of every message is kept twice — once as itself and once as the
 * earliest quote of itself.
 */
function stripLeadingHeaders(block: string): string {
  const lines = block.split('\n');
  let index = 0;
  while (index < lines.length) {
    const line = lines[index] as string;
    if (line.trim() !== '' && !HEADER_LINE.test(line)) break;
    index += 1;
  }
  return lines.slice(index).join('\n');
}

export interface FoldInput {
  /** Body text, already normalised and link-folded. */
  body: string;
}

export interface FoldedBody {
  body: string;
  /** Characters removed by folding, from this message. */
  quotedChars: number;
  /** Pointer markers emitted in this body. */
  pointers: number;
  /**
   * How much of this body a later message's pointer depends on. Cutting below
   * this leaves those markers naming text that is no longer in the response.
   */
  protectedChars: number;
  /**
   * Set when this message folded away to nothing but pointers at a single
   * earlier message — the mailbox holding both a Sent copy and a delivered copy
   * of one mail, which would otherwise read as an empty message.
   */
  duplicateOf?: number;
}

export interface FoldStats {
  blocks: number;
  folded: number;
  quotedChars: number;
}

export interface FoldResult {
  bodies: FoldedBody[];
  stats: FoldStats;
}

/**
 * Folds text repeated across `messages` into pointers at the earliest message
 * that carried it.
 *
 * The cache is this function's own local, and every pointer names a 1-based
 * index INSIDE the array passed in. Both properties exist for the same reason:
 * the same input must produce the same output, and a pointer must always resolve
 * within the response it appears in. A session-scoped cache satisfied neither —
 * re-reading one message returned pointers at itself.
 *
 * `messages` must be ordered oldest-first, since a pointer may only ever aim
 * backwards.
 */
export function foldRepeatedBlocks(messages: readonly FoldInput[]): FoldResult {
  const seen = new Map<string, number>();
  const staged = new Map<string, number>();
  const stats: FoldStats = { blocks: 0, folded: 0, quotedChars: 0 };

  // Pass one decides what folds. Pass two renders, because a message cannot know
  // which of its own blocks a LATER message will point at until every message
  // has been read — and that is what decides how much of it must survive the
  // character budget.
  const drafts: Array<{ pieces: Piece[]; quotedChars: number }> = [];
  const cited = new Map<number, Set<string>>();

  messages.forEach((message, index) => {
    const position = index + 1;
    const pieces: Piece[] = [];
    let quotedChars = 0;

    for (const block of splitQuotedBlocks(message.body)) {
      stats.blocks += 1;
      const key = blockKey(block);

      if (key.length < MIN_BLOCK_CHARS) {
        pieces.push({ text: block });
        continue;
      }

      const origin = seen.get(key);
      if (origin === undefined) {
        // Staged, not published. A block this message is the first to carry must
        // not become a pointer target for this same message: a body that quotes
        // one forwarded mail twice would otherwise emit `[quoted from #3]` while
        // it IS #3, a reference that resolves nowhere. The repeat stays verbatim,
        // which is the harmless direction.
        staged.set(key, position);
        pieces.push({ text: block, key });
        continue;
      }

      stats.folded += 1;
      quotedChars += block.length;
      pieces.push({ text: '', source: origin });
      let keys = cited.get(origin);
      if (keys === undefined) {
        keys = new Set<string>();
        cited.set(origin, keys);
      }
      keys.add(key);
    }
    for (const [key, at] of staged) if (!seen.has(key)) seen.set(key, at);
    staged.clear();

    drafts.push({ pieces, quotedChars });
  });

  const bodies = drafts.map(({ pieces, quotedChars }, index) => {
    const { body, pointers, sources, protectedChars } = render(pieces, cited.get(index + 1));
    stats.quotedChars += quotedChars;

    const folded: FoldedBody = { body, quotedChars, pointers, protectedChars };
    // Nothing of this message survived except references to one other message.
    if (body.replace(/\[quoted from[^\]]*\]/g, '').trim() === '' && sources.size === 1) {
      folded.duplicateOf = [...sources][0];
    }
    return folded;
  });

  return { bodies, stats };
}

interface Piece {
  text: string;
  /** Set when this piece replaced a repeat; names the message that carried it. */
  source?: number;
  /** Comparison key, present on verbatim blocks so a citation can find them. */
  key?: string;
}

/**
 * Joins pieces back into a body, collapsing a run of adjacent pointers into one
 * marker. A thread twelve replies deep otherwise ends with twelve consecutive
 * `[quoted from …]` lines, which is noise standing in for noise.
 *
 * Also reports `protectedChars`: the offset just past the last block that a
 * later message points at. Truncating this body below that offset would delete
 * the text those pointers name, leaving markers that resolve to nothing — a
 * silent loss dressed up as a reference. The budget uses it as a floor.
 *
 * Blank-line collapsing is applied per piece rather than to the joined string so
 * the offsets stay exact; the difference is cosmetic and confined to piece
 * boundaries.
 */
function render(
  pieces: ReadonlyArray<Piece>,
  citedKeys: ReadonlySet<string> | undefined,
): {
  body: string;
  pointers: number;
  sources: Set<number>;
  protectedChars: number;
} {
  const out: string[] = [];
  const sources = new Set<number>();
  const protectAfter: number[] = [];
  let pointers = 0;
  let run: number[] = [];

  const flush = (): void => {
    if (run.length === 0) return;
    out.push(POINTER_OPEN + describe(run) + POINTER_CLOSE);
    pointers += 1;
    run = [];
  };

  for (const piece of pieces) {
    if (piece.source === undefined) {
      flush();
      // Blank lines are collapsed; runs of spaces and tabs are NOT. Indentation
      // is the only structure a plain-text table, a log excerpt or pasted code
      // has left by the time it reaches here, and squeezing it is not ours to do.
      out.push(piece.text.replace(/\n{3,}/g, '\n\n'));
      if (piece.key !== undefined && citedKeys?.has(piece.key) === true) {
        protectAfter.push(out.length - 1);
      }
    } else {
      sources.add(piece.source);
      run.push(piece.source);
    }
  }
  flush();

  const joined = out.join('\n');
  const lead = joined.length - joined.trimStart().length;
  const body = joined.trim();

  let protectedChars = 0;
  if (protectAfter.length > 0) {
    const last = Math.max(...protectAfter);
    let end = 0;
    for (let i = 0; i <= last; i += 1) end += (out[i] as string).length + (i > 0 ? 1 : 0);
    // Clamped to the body actually emitted. When the protected block is the last
    // piece, `end` still counts whitespace the trim removed, and a floor larger
    // than the text it protects can never be met — which would read as a
    // permanent dangling pointer and drop folding on every call.
    protectedChars = Math.min(body.length, Math.max(0, end - lead));
  }

  return { body, pointers, sources, protectedChars };
}

/** `#3`, `#1-#4` for a contiguous run, or a short list. */
function describe(run: readonly number[]): string {
  const unique = [...new Set(run)].sort((a, b) => a - b);
  if (unique.length === 1) return `#${unique[0]}`;

  const first = unique[0] as number;
  const last = unique[unique.length - 1] as number;
  if (last - first + 1 === unique.length) return `#${first}-#${last}`;

  if (unique.length <= 6) return unique.map((n) => `#${n}`).join(', ');
  return `${unique.slice(0, 5).map((n) => `#${n}`).join(', ')}, … #${last}`;
}

// ---------------------------------------------------------------------------
// Budget
// ---------------------------------------------------------------------------

/**
 * Spreads `total` characters across `needs` so that no message is cut to make
 * room for a longer one (max-min fair share).
 *
 * Equal shares, then whatever a short message does not use is redistributed to
 * those still over their share, repeatedly. A thread of one 40,000-character
 * message and nine 300-character ones therefore returns all nine short ones
 * whole rather than clipping every message to the same length.
 *
 * `floors` are served before any of that. A floor is the part of a body that a
 * later message's pointer depends on: cutting below it would leave markers
 * naming text no longer in the response, which reads as a reference and is a
 * deletion. Correctness outranks fairness, so floors come first.
 */
export function shareBudget(
  needs: readonly number[],
  total: number,
  floors: readonly number[] = [],
): number[] {
  const count = needs.length;
  if (count === 0) return [];
  const budget = Number.isFinite(total) && total > 0 ? Math.floor(total) : 0;
  if (budget <= 0) return needs.map(() => 0);

  const grant = needs.map(() => 0);
  let remaining = budget;

  // Floors first, in index order. A floor is text a later message's pointer
  // names; the earliest message is the one most often pointed at, so serving in
  // order spends the budget where the most references hang off it.
  for (let index = 0; index < count; index += 1) {
    const want = Math.min(floors[index] ?? 0, needs[index] as number, remaining);
    if (want > 0) {
      grant[index] = want;
      remaining -= want;
    }
  }

  // Then max-min over whatever each message still wants.
  const open: number[] = [];
  for (let index = 0; index < count; index += 1) {
    if ((needs[index] as number) > (grant[index] as number)) open.push(index);
  }

  while (open.length > 0 && remaining > 0) {
    const share = Math.floor(remaining / open.length);
    if (share === 0) break;

    const settled: number[] = [];
    for (const index of open) {
      const want = (needs[index] as number) - (grant[index] as number);
      if (want <= share) {
        grant[index] = needs[index] as number;
        remaining -= want;
        settled.push(index);
      }
    }
    if (settled.length === 0) {
      // Everyone still open wants more than an equal share: give each exactly
      // that and stop, so the split is even rather than first-come.
      for (const index of open) {
        grant[index] = (grant[index] as number) + share;
        remaining -= share;
      }
      break;
    }
    for (const index of settled) open.splice(open.indexOf(index), 1);
  }

  // Hand out the rounding remainder in index order so the result is stable.
  for (const index of open) {
    if (remaining <= 0) break;
    const want = (needs[index] as number) - (grant[index] as number);
    const give = Math.min(want, remaining);
    grant[index] = (grant[index] as number) + give;
    remaining -= give;
  }
  return grant;
}
