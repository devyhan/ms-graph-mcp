# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- README slimmed to an overview; detail sections moved to `docs/`
  (auth, configuration, tools, examples, troubleshooting, security,
  agents, why) with per-group usage examples.

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
