# ms-graph-mcp

[![npm version](https://img.shields.io/npm/v/@devyhan/ms-graph-mcp)](https://www.npmjs.com/package/@devyhan/ms-graph-mcp)
[![CI](https://github.com/devyhan/ms-graph-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/devyhan/ms-graph-mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A Model Context Protocol server that gives an AI assistant access to your own Microsoft 365 data — Outlook mail and calendar, OneDrive, SharePoint, To Do, Planner, Teams chat, Entra ID and Intune — through the Microsoft Graph API. It runs locally over stdio, signs in as you with delegated permissions, and stores its token cache encrypted on your machine. ([Why an MCP server](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/why.md))

Published on npm as [`@devyhan/ms-graph-mcp`](https://www.npmjs.com/package/@devyhan/ms-graph-mcp).

> **Not affiliated with Microsoft.** This is an independent community project. Microsoft, Microsoft 365, Microsoft Graph, Outlook, OneDrive, SharePoint, Teams, Entra and Intune are trademarks of the Microsoft group of companies, used here only to describe what this software connects to.

## Features

- **111 tools across 14 groups** — mail, calendar, files, chat, teams, tasks, contacts, search, directory and device inventory, each enabled independently
- **Delegated sign-in only** — acts as you, with your permissions; the token never leaves the server process or enters the conversation
- **Incremental consent** — a bare login asks for the minimal set, and each tool call authenticates with just the scopes it needs
- **Model-shaped responses** — Graph results are projected down to sender, timestamp and plain text with paging cursors, not raw markup
- **Document parsing** — `.docx`, `.xlsx`, `.pptx` and `.pdf` extracted as unit rows from OneDrive, SharePoint, Teams or mail attachments
- **Guardrails** — host-pinned generic requests, `--read-only` mode, opt-in generic writes, retry with `Retry-After`
- **Discovery mode** — `--discovery` collapses the catalogue to `discover_tools` + `call_tool` so clients pay for schemas only on use

## Install with one prompt

Paste this to your AI assistant, replacing the two IDs with your Entra app registration's values:

```text
https://github.com/devyhan/ms-graph-mcp

Install the MCP above.
• CLIENT_ID: <APPLICATION_ID>
• TENANT_ID: <DIRECTORY_ID>
```

Find both IDs in the [Microsoft Entra admin center](https://entra.microsoft.com) under **App registrations → your app → Overview**: `Application (client) ID` and `Directory (tenant) ID`. The app must allow public client flows with an `http://localhost` redirect URI (under **Authentication → Mobile and desktop applications**), and an administrator must grant consent for the Graph permissions it requests.

## Quick start

You need Node.js 20 or later. Sign in once, in a terminal:

```
npx @devyhan/ms-graph-mcp login
```

The server is meant to sign in through a shared multi-tenant Entra application published by this project — except that the shared registration is not published yet (`SHIPPED_CLIENT_ID` in `src/config.ts` is still empty), so until it ships, register an application of your own and name it on each command:

```
npx @devyhan/ms-graph-mcp login --client-id <your-client-id>
```

[Register your own Entra application](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/auth.md#register-your-own-entra-application) has the steps.

Your system browser opens on the Microsoft sign-in page; approve the permissions it asks for and the terminal finishes on its own. On a machine with no browser to open (SSH, container, headless server), use `login --auth-flow device` and sign in from another device. Then check it worked:

```
npx @devyhan/ms-graph-mcp status
```

`status` prints the application in use and where it came from, the resolved account, and the scopes this configuration may request. How the flows, consent, and `AADSTS` errors work is in [Authentication and consent](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/auth.md).

## Configuration

The fastest path is the Claude Code plugin — no config file to edit:

```
/plugin marketplace add devyhan/ms-graph-mcp
/plugin install ms-graph-mcp@devyhan
```

For other clients, add an MCP server entry (here with the `work` preset):

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["-y", "@devyhan/ms-graph-mcp", "--preset", "work"]
    }
  }
}
```

Every setting is also an `MS365_MCP_*` environment variable or a command-line flag. See [Configuration](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/configuration.md) for the plugin settings, per-client files, and every command and flag.

## Tools

111 tools across 14 groups. `--read-only` removes every write tool, leaving 80. Enable only the groups you need.

| Group | Product | Tools |
| --- | --- | --- |
| `me` | Entra ID profile, Outlook mailbox settings | 2 |
| `mail` | Outlook mail | 14 |
| `calendar` | Outlook calendar | 9 |
| `files` | OneDrive | 13 |
| `todo` | Microsoft To Do | 8 |
| `planner` | Microsoft Planner | 9 |
| `contacts` | Outlook contacts | 6 |
| `chat` | Teams one-to-one and group chats | 7 |
| `sharepoint` | SharePoint sites, libraries, lists | 11 |
| `search` | Microsoft Search across mail, files, sites | 1 |
| `teams` * | Teams teams, channels, channel messages | 10 |
| `directory` * | Entra ID user and group lookups | 9 |
| `intune` * | Intune managed device inventory | 6 |
| `generic` | Signing in, any Graph endpoint, introspection | 6 |

\* Needs `--org-mode`: these groups request scopes only a tenant administrator can consent to.

The full reference — scopes, admin-consent flags, and every tool with its arguments — is in [Tools](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/tools.md). Under `--discovery` the client sees six tools instead: `discover_tools` and `call_tool` plus the session, sign-in and permission probes.

## Usage examples

> "Summarize the reply chain about the Q3 invoice."

Search first, then read the whole thread in one call instead of looping `mail_get_message`:

```json
{ "query": "Q3 invoice", "after": "2026-09-01" }
```

```json
{ "fromMessageId": "<id from the search>" }
```

> "Find a 30-minute slot with Ana and Ben next week."

```json
{
  "attendees": ["ana@contoso.com", "ben@contoso.com"],
  "durationMinutes": 30,
  "start": "2026-10-19T09:00:00",
  "end": "2026-10-23T18:00:00",
  "timeZone": "UTC"
}
```

More per-group examples — files, chat, search, Planner, SharePoint, Teams, directory, Intune, and the generic escape hatch — are in [Usage examples](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/examples.md).

## Instructions for AI agents

If you are an AI assistant working with this server, follow [Instructions for AI agents](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/agents.md): call `session_status` once first, sign in with `auth_begin_login` plus `auth_login_status` when needed, orient via `discover_tools`, and read the Hint and Blocked-by lines on errors.

## Troubleshooting

Something failing? Start with [Troubleshooting](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/troubleshooting.md) — the common failures and their fixes — then [Authentication and consent](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/auth.md) for `AADSTS` error codes.

## Security

- **Delegated permissions only.** The server acts as you, holds no client secret, and the token never enters the conversation.
- **Host pinning.** `graph_request` paths and every `@odata.nextLink` are checked against the Graph hostname before the request goes out.
- **Read-only mode.** `--read-only` removes write tools from the catalogue; generic writes need `--allow-generic-write` on top.
- **Encrypted token cache.** The refresh token lives AES-256-GCM-encrypted in your config directory (`--cache-dir` overrides).

Details: [Security](https://github.com/devyhan/ms-graph-mcp/blob/main/docs/security.md). To report a vulnerability privately, see [SECURITY.md](SECURITY.md).

## What this is not

This is a local, delegated-permission MCP server. It acts as you, with the permissions you have, over a stdio transport on your own machine. It is not a replacement for Microsoft's own Agent 365 MCP servers, which are remote, run inside Microsoft's tenant boundary, use application permissions and admin-managed agent identities, and require a Microsoft 365 Copilot licence. If you need agent identity, tenant-wide governance, or auditing that a tenant administrator controls, that is the product to look at.

It is also not a general Microsoft Graph client library. The tools cover the common read and write paths for each product area and project their results down to what a model can usefully read; `graph_request` exists for everything else, deliberately behind a flag.

## Development

Requires Node.js 20 or later.

```
git clone https://github.com/devyhan/ms-graph-mcp.git
cd ms-graph-mcp
npm install
npm test
```

| Command | What it does |
| --- | --- |
| `npm test` | Builds, then runs the suite. This is the one to run. |
| `npm run typecheck` | Types only, no emit. |
| `npm run build` | Compile to `dist/`. |
| `npm run contract` | Checks scopes and paths against live Microsoft data. Needs the network. |

The test suite needs no Microsoft account, no tenant and no network — every HTTP call is stubbed. `npm run contract` is separate on purpose and runs weekly in CI instead. See [CONTRIBUTING.md](CONTRIBUTING.md) for the code layout, the rules for adding a tool, and how to report a bug. Changes are tracked in [CHANGELOG.md](CHANGELOG.md).

## Licence

MIT — see [LICENSE](LICENSE).
