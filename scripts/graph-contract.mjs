#!/usr/bin/env node
/**
 * Drift detector for the Microsoft Graph surface this server depends on.
 *
 * The 94 tests in `npm test` stub `fetch`. They prove our logic is right and
 * they would keep passing if Microsoft renamed every endpoint tomorrow, because
 * nothing in them ever reaches Graph. This script covers the other half: it
 * asks Microsoft what Graph looks like today and compares that against what the
 * code actually asks for.
 *
 * Two public, unauthenticated sources:
 *   - https://graph.microsoft.com/v1.0/$metadata        the CSDL model
 *   - the permissions reference markdown in microsoft-graph-docs-contrib
 *
 * Deliberately NOT part of `npm test`: it needs the network, it depends on two
 * services we do not control, and a contributor's pull request should not fail
 * because Microsoft shipped something overnight. It runs on a schedule instead,
 * where a failure is a notification rather than a blocked merge.
 *
 * Everything it checks is derived from the built output — GROUPS and the real
 * tool definitions — never from a list maintained here. A hand-kept list would
 * drift from the code, which is the failure this script exists to prevent.
 *
 *   node scripts/graph-contract.mjs            human-readable
 *   node scripts/graph-contract.mjs --json     machine-readable, for CI
 */

import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const CACHE = path.join(ROOT, '.graph-cache');

const METADATA_URL = 'https://graph.microsoft.com/v1.0/$metadata';
const PERMISSIONS_URL =
  'https://raw.githubusercontent.com/microsoftgraph/microsoft-graph-docs-contrib/main/concepts/permissions-reference.md';

const JSON_OUT = process.argv.includes('--json');
const findings = [];

/** A drift finding. `severity` decides the exit code, not the noise level. */
function report(severity, check, subject, detail) {
  findings.push({ severity, check, subject, detail });
}

/**
 * Fetches with an on-disk cache so a local run does not pull 2.3 MB each time.
 * CI passes --no-cache; the cache directory is gitignored.
 */
async function fetchText(url, cacheName) {
  const cached = path.join(CACHE, cacheName);
  const useCache = !process.argv.includes('--no-cache');
  if (useCache && existsSync(cached)) return readFile(cached, 'utf8');

  const res = await fetch(url, { headers: { accept: 'text/plain, application/xml, */*' } });
  if (!res.ok) throw new Error(`${url} returned HTTP ${res.status}`);
  const text = await res.text();
  await mkdir(CACHE, { recursive: true });
  await writeFile(cached, text);
  return text;
}

// ---------------------------------------------------------------------------
// What the code actually depends on
// ---------------------------------------------------------------------------

/**
 * Every delegated scope the server can ask for, and which group asked for it.
 * Read from the built catalogue plus every tool definition, because a tool may
 * declare a scope its group's meta does not list — `me_get_mailbox_settings`
 * needs `MailboxSettings.Read`, for instance — and a scope nobody declares at
 * sign-in is a scope that 403s at call time.
 */
async function declaredScopes() {
  const { GROUPS } = await import(path.join(ROOT, 'dist/tools/groups.js'));
  const { collectTools } = await import(path.join(ROOT, 'dist/server.js'));

  const owners = new Map();
  const note = (scope, owner) => {
    if (!owners.has(scope)) owners.set(scope, new Set());
    owners.get(scope).add(owner);
  };

  for (const [name, meta] of Object.entries(GROUPS)) {
    for (const s of meta.readScopes) note(s, name);
    for (const s of meta.writeScopes) note(s, name);
  }

  // A throwaway config just rich enough to enumerate every tool.
  const config = {
    clientId: '', tenantId: 'common', authority: '', graphHost: 'graph.microsoft.com',
    groups: Object.keys(GROUPS), readOnly: false, graphVersion: 'v1.0', allowBeta: false,
    discovery: false, orgMode: true, maxOutputChars: 60000, verbose: false, cacheDir: '',
    allowGenericWrite: true, authFlow: 'auto', authPort: 0,
  };
  const refuse = () => { throw new Error('contract check makes no Graph calls'); };
  const tools = collectTools({ graph: { request: refuse, batch: refuse, follow: refuse }, config });
  for (const t of tools) for (const s of t.scopes) note(s, t.group);

  return { owners, groups: GROUPS, toolCount: tools.length };
}

/**
 * Graph paths reachable from the tool modules, harvested from the source.
 *
 * A regex over template literals rather than a call graph: the paths are built
 * inline at each call site, and the goal is coverage of the named segments, not
 * an exact route table. Placeholders are stripped so `/chats/${id}/messages`
 * becomes the segment set {chats, messages}.
 */
async function referencedSegments() {
  const { readdir } = await import('node:fs/promises');
  const dir = path.join(ROOT, 'src/tools');
  const files = (await readdir(dir)).filter((f) => f.endsWith('.ts') && !f.endsWith('.test.ts'));
  const segments = new Map();

  for (const file of files) {
    const text = await readFile(path.join(dir, file), 'utf8');
    for (const line of text.split('\n')) {
      // Comments carry worked examples — `/me/drive/root:/a b.txt:/content` —
      // whose fragments are not path segments. Reading them produced exactly
      // the false positives this skip exists to remove.
      const trimmed = line.trim();
      if (trimmed.startsWith('*') || trimmed.startsWith('//') || trimmed.startsWith('/**')) continue;

      for (const m of line.matchAll(/[`'"](\/(?:me|users|groups|chats|teams|sites|planner|drives|search|deviceManagement|servicePrincipals|applications)[^`'"]*)[`'"]/g)) {
        for (const raw of m[1].split('/')) {
          const seg = raw.replace(/\$\{[^}]*\}/g, '').replace(/[():'"].*$/, '').trim();
          // Named segments only: no placeholders, no OData operators, and long
          // enough that a stray letter from an example cannot qualify.
          if (seg.length >= 3 && !seg.startsWith('$') && /^[a-zA-Z][a-zA-Z0-9]*$/.test(seg)) {
            if (!segments.has(seg)) segments.set(seg, new Set());
            segments.get(seg).add(file.replace('.ts', ''));
          }
        }
      }
    }
  }
  return segments;
}

// ---------------------------------------------------------------------------
// Checks
// ---------------------------------------------------------------------------

/**
 * Every scope we may request must still exist, must still be delegated, and its
 * admin-consent flag must still match what the group catalogue assumes. That
 * last one matters more than it looks: `requiresAdminConsent` decides whether a
 * group is hidden without `--org-mode`, so if Microsoft reclassifies a
 * permission we would start showing users a consent prompt they cannot satisfy.
 */
function checkScopes(perms, { owners, groups }) {
  // Split the reference into "### <PermissionName>" blocks.
  const blocks = new Map();
  const parts = perms.split(/^### /m);
  for (const part of parts.slice(1)) {
    const name = part.split('\n', 1)[0].trim();
    if (name) blocks.set(name, part);
  }

  // Scopes that are OIDC, not Graph permissions, and so are absent from the
  // reference by design.
  const OIDC = new Set(['offline_access', 'openid', 'profile', 'email']);

  const adminByGroup = new Map(
    Object.entries(groups).map(([n, m]) => [n, m.requiresAdminConsent === true]),
  );

  for (const [scope, groupNames] of [...owners].sort()) {
    if (OIDC.has(scope)) continue;

    const block = blocks.get(scope);
    if (!block) {
      report('error', 'scope-exists', scope,
        `Not found in the permissions reference. Requested by: ${[...groupNames].join(', ')}. ` +
        'Either it was renamed or removed, or the reference changed shape.');
      continue;
    }

    if (/\bdeprecated\b/i.test(block.split('\n').slice(0, 14).join('\n'))) {
      report('warn', 'scope-deprecated', scope,
        `Marked deprecated in the permissions reference. Requested by: ${[...groupNames].join(', ')}.`);
    }

    // "| AdminConsentRequired | <application> | <delegated> |"
    const row = block.match(/^\|\s*AdminConsentRequired\s*\|([^|]*)\|([^|]*)\|/m);
    if (!row) {
      report('warn', 'scope-shape', scope, 'No AdminConsentRequired row; the reference format may have changed.');
      continue;
    }
    const delegated = row[2].trim();
    if (delegated === '-') {
      report('error', 'scope-delegated', scope,
        'No longer offered as a DELEGATED permission. This server never uses application ' +
        'permissions, so a tool requesting it would fail for every user.');
      continue;
    }

    const needsAdmin = /^yes$/i.test(delegated);
    // A group is flagged admin-consent if ANY of its scopes needs one; a scope
    // needing admin consent inside a group not flagged for it is the drift.
    const anyOwnerFlagged = [...groupNames].some((g) => adminByGroup.get(g) === true);
    if (needsAdmin && !anyOwnerFlagged) {
      report('error', 'scope-admin-consent', scope,
        `Now requires admin consent (delegated), but every group requesting it ` +
        `(${[...groupNames].join(', ')}) is exposed without --org-mode. Users would hit a ` +
        'consent prompt they cannot satisfy. Flag the group requiresAdminConsent, or drop the scope.');
    }
    if (!needsAdmin && [...groupNames].every((g) => adminByGroup.get(g) === true)) {
      report('info', 'scope-admin-consent', scope,
        'No longer requires admin consent, but is only reachable behind --org-mode. ' +
        'It could move to a group ordinary users can enable.');
    }
  }
}

/** Every named path segment we build a URL from must still exist in the model. */
function checkSegments(metadata, segments) {
  const known = new Set();
  for (const m of metadata.matchAll(/<(?:EntitySet|Singleton|NavigationProperty|Property|Action|Function)\s+Name="([^"]+)"/g)) {
    known.add(m[1]);
  }
  // Actions and functions appear bound, under their namespace.
  for (const m of metadata.matchAll(/<(?:Action|Function)\s+Name="([^"]+)"/g)) known.add(m[1]);

  for (const [seg, files] of [...segments].sort()) {
    if (!known.has(seg)) {
      report('warn', 'path-segment', seg,
        `Not present in the v1.0 $metadata as an entity set, navigation property, action or ` +
        `function. Used by: ${[...files].join(', ')}. It may have been renamed, moved to beta, ` +
        'or it may be a segment this check cannot see (a cast, or a key).');
    }
  }
}

/** The shipped client ID must stay empty until a real registration is published. */
async function checkShippedClientId() {
  const { SHIPPED_CLIENT_ID } = await import(path.join(ROOT, 'dist/config.js'));
  if (SHIPPED_CLIENT_ID !== '') {
    report('info', 'shipped-client-id', SHIPPED_CLIENT_ID,
      'A client ID is shipped. Confirm the registration is multi-tenant and still live: a ' +
      'single-tenant or deleted app returns AADSTS700016 to every other tenant.');
  }
}

// ---------------------------------------------------------------------------

async function main() {
  if (!existsSync(path.join(ROOT, 'dist/server.js'))) {
    console.error('dist/ is missing. Run `npm run build` first.');
    process.exit(2);
  }

  const [metadata, perms] = await Promise.all([
    fetchText(METADATA_URL, 'metadata.xml'),
    fetchText(PERMISSIONS_URL, 'permissions-reference.md'),
  ]);

  const scopes = await declaredScopes();
  checkScopes(perms, scopes);
  checkSegments(metadata, await referencedSegments());
  await checkShippedClientId();

  const counts = findings.reduce((a, f) => ({ ...a, [f.severity]: (a[f.severity] ?? 0) + 1 }), {});
  const summary = {
    checkedAt: new Date().toISOString(),
    scopes: scopes.owners.size,
    tools: scopes.toolCount,
    metadataBytes: metadata.length,
    errors: counts.error ?? 0,
    warnings: counts.warn ?? 0,
    info: counts.info ?? 0,
    findings,
  };

  if (JSON_OUT) {
    console.log(JSON.stringify(summary, null, 2));
  } else {
    console.log(`Graph contract check — ${scopes.owners.size} scopes, ${scopes.toolCount} tools\n`);
    if (findings.length === 0) {
      console.log('No drift. Every declared scope and path segment still matches Microsoft.');
    } else {
      for (const f of findings) {
        console.log(`[${f.severity.toUpperCase()}] ${f.check}: ${f.subject}`);
        console.log(`  ${f.detail}\n`);
      }
      console.log(`${summary.errors} error(s), ${summary.warnings} warning(s), ${summary.info} info.`);
    }
  }

  // Only an error fails the run. A warning is drift worth a look, not a break.
  process.exit(summary.errors > 0 ? 1 : 0);
}

main().catch((err) => {
  console.error(`Contract check could not run: ${err.message}`);
  // Exit 2 separates "we could not check" from "we checked and found drift",
  // so a GitHub outage does not read as a Graph breaking change.
  process.exit(2);
});
