# Security policy

This server holds a Microsoft 365 refresh token and reads mailboxes, files and
chat on a real person's behalf. A bug here is not an inconvenience, so please
report one privately rather than in a public issue.

## Reporting a vulnerability

Use GitHub's private vulnerability reporting:
**[Report a vulnerability](https://github.com/devyhan/ms-graph-mcp/security/advisories/new)**
(Security tab → Report a vulnerability). It is private to the maintainers until
an advisory is published.

If that is unavailable to you, open a public issue saying only that you have a
security report and how to reach you — no details.

**Please do not include real credentials or real data.** No access tokens,
refresh tokens, client secrets, `token-cache.enc` or `cache.key` files, real
mail or chat content, or a tenant ID you would not publish. A redacted
reproduction is enough; if a real value is genuinely load-bearing, say so and we
will arrange somewhere to put it. Anything pasted into an issue or an advisory
comment is retained even after editing.

Expect an acknowledgement within a few days. This is a volunteer project with no
paid support, so a fix may take longer than an acknowledgement — you will be told
which is happening. Credit in the advisory is offered unless you decline.

## In scope

Anything that lets someone reach data the signed-in user did not intend to
expose, or that reaches this machine:

- **Escaping the Graph host.** `graph_request` lets the model choose a request
  path and every request carries a bearer token, so a path that resolves
  anywhere but the configured Graph host leaks that token. The guard is
  `src/util/paths.ts` plus the host check in `src/graph/client.ts`, applied to
  outgoing requests and to every `@odata.nextLink` before it is followed.
- **Token disclosure.** A token appearing in tool output, on stdout, in an error
  message, or in a log line. The model must never see one.
- **The token cache.** `src/auth/token-cache.ts` — the AES-256-GCM encryption,
  the key handling, and the file modes.
- **Privilege beyond consent.** Anything that reaches data outside the scopes the
  user actually consented to, or that bypasses `--read-only`.
- **Prompt injection with real consequence.** Content inside a mail or chat
  message is untrusted input. A crafted message that induces a destructive tool
  call is interesting if the server's own guards were meant to stop it — for
  example a write reachable while `--read-only` is set, or a destructive tool
  that misreports itself as read-only to the client.
- **Supply chain.** A dependency or build step that could introduce any of the
  above.

## Out of scope

- **Microsoft Graph and Entra themselves.** Report those to
  [MSRC](https://msrc.microsoft.com/report). That includes tenant consent policy
  behaviour: that an ordinary user cannot self-consent to `Mail.Read` is a
  Microsoft policy decision, not a bug here.
- **A user's own configuration.** Over-broad scopes, an application registered
  with permissions it did not need, or admin consent granted too generously.
  The server asks only for what the enabled groups require; what an
  administrator grants is their decision.
- **The model calling a tool the user did not want.** A model deciding to read
  your mail when you asked something else is a client-side authorization
  question. Use `--read-only` and a narrow `--groups` set. A tool that *lies*
  about whether it writes is in scope; a tool honestly labelled and then invoked
  is not.
- **Anything requiring an attacker who already has your machine.** The token
  cache key sits beside the ciphertext at mode `0600`. That protects against a
  stray backup or a synced folder, not against someone with read access to your
  home directory, and it is documented as such.

## What this server does with your data

It sends requests from your machine to `graph.microsoft.com` and returns the
results. There is no service in between: no telemetry, no analytics, no
crash reporting, and no network destination other than Microsoft's Graph and
login endpoints. The only thing written to disk is the encrypted token cache, in
your own config directory.

## Supported versions

Pre-1.0. Fixes land on the latest release only. Once 1.0 ships, this section will
name a support window.
