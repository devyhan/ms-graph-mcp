# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.7.0] - 2026-10-19

### Added

- Login-time identity snapshot: every successful sign-in caches who signed
  in (username, display name, tenant) encrypted beside the token cache, so
  status output can name the session without a Graph call. `session_status`
  returns it as `cachedIdentity`, and tools can opt in to a one-line header
  above their output. The snapshot is dropped with `logout`, when the
  session can no longer mint tokens, or when another process signs in over
  it; a missing or corrupt file reads as absent, never as an error.

## [0.6.1] - 2026-10-19

### Added

- Proxy support: when `HTTPS_PROXY`/`https_proxy`/`ALL_PROXY`/`all_proxy`
  is set, the global fetch is wrapped at startup (CONNECT tunneling for
  HTTPS, absolute-form for HTTP, `NO_PROXY` honored), so Entra, Graph and
  the update check work behind a proxy. One stderr line names the proxy
  host when it activates. A malformed proxy value warns on stderr and falls
  back to a direct connection instead of failing startup.

### Fixed

- Behind a proxy, sign-in and every Graph call failed with `ENOTFOUND`
  (Node's `fetch` ignores proxy variables); requests now route through
  the proxy automatically. See [Proxy](docs/configuration.md#proxy).

## [0.6.0] - 2026-10-16

### Added

- Per-group usage examples: every tool group now has a prompt-plus-call
  example in `docs/examples.md`.
- Boolean flags accept an explicit value: `--read-only=false` switches off
  `MS365_MCP_READ_ONLY=true`, using the same vocabulary as the environment
  variables.
- `permissions` annotates cross-group scopes with a one-clause reason, so
  `files -> Mail.Read` and the search index scopes explain themselves.

### Changed

- README slimmed to an overview; detail sections moved to `docs/`
  (auth, configuration, tools, examples, troubleshooting, security,
  agents, why).
- `status` reports the exposed tool count (36 under `--read-only`,
  6 under `--discovery`) instead of the raw catalogue total.
- `status` names dropped groups inline with the `--org-mode` remedy
  instead of printing an empty group list.
- `logout` with no session says nothing was removed rather than
  claiming it deleted files.
- `--groups` is case-insensitive, like `--preset`.
- `--help` documents that `--preset` and `--groups` combine by union.
- `permissions` closing line now distinguishes Microsoft's classification
  from the tenant consent policy gate.

### Fixed

- Help examples used the nonexistent `microsoft-graph-mcp` command;
  every user-facing command string now uses the real `ms-graph-mcp` bin.
- A bad environment value no longer defeats `--help`/`--version`;
  real commands still fail fast on it.
- Non-`login` commands warn on stderr when `--scopes`/`--full` would
  be silently ignored.

## [0.5.0] - 2026-10-15

### Added

- CLI update notice: human commands print a one-line stderr notice when a
  newer release is on npm (24h cache, silent offline, `NO_UPDATE_NOTIFIER`
  respected, `serve` stdout untouched).
- README install prompt block at the top for one-paste AI setup.

### Changed

- Raise extraction fidelity: unit text renders tables as Markdown and the
  manifest carries a `titles[]` outline, so outline-first reads need no text
  fetch. docx gains header/footer units, tracked-change tagging and comments;
  xlsx renders number formats, skips hidden rows/sheets with warnings, and
  fills merged cells; pptx follows SmartArt/chart references with warning
  fallback, dedupes master text, and skips note placeholders.
- Releases now also cut a GitHub Release from the CHANGELOG section.

## [0.4.0] - 2026-10-15

### Added

- New `files_parse_document` tool: resolves a share URL, drive item or mail
  attachment and returns parsed text as unit rows (slides/pages/sheets) with
  a manifest, so the model can verify nothing was missed.
- Binary fetch layer (manual 302, Range on the redirect target), `/shares`
  resolution, channel `filesFolder`, and mail attachment content support.
- Pure-JS extractors for docx/xlsx/pptx/pdf. Protected, legacy and oversize
  documents fail with specific guidance instead of raw errors.

## [0.3.0] - 2026-10-15

### Added

- Opt-in TOON compact output: new `--output toon|json` option
  (`MS365_MCP_OUTPUT`) serializes tool results in a compact tabular form,
  cutting uniform multi-row output to roughly a third of its size.

### Changed

- The default stays `json` with byte-identical output, and truncation notices
  adapt to the active format. No new dependencies.

## [0.2.1] - 2026-10-15

### Fixed

- `status` and `session_status` now probe the cached session instead of
  trusting the account record, so a tokenless cache shows "token unusable"
  rather than looking signed in.
- Conditional Access (53003), network, and loopback-timeout failures get
  actionable guidance instead of raw MSAL text.

### Added

- `login` gains a preflight check that skips flows the environment cannot run.
- `auth_begin_login` gains `mode: 'full'` for upfront user-grantable consent.

[Unreleased]: https://github.com/devyhan/ms-graph-mcp/compare/v0.5.0...HEAD
[0.5.0]: https://github.com/devyhan/ms-graph-mcp/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/devyhan/ms-graph-mcp/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/devyhan/ms-graph-mcp/compare/v0.2.1...v0.3.0
[0.2.1]: https://github.com/devyhan/ms-graph-mcp/compare/v0.2.0...v0.2.1
