# Contributing

Thanks for looking. This is a small project and issues are welcome even when you
have no fix in mind.

## Getting set up

```
git clone https://github.com/devyhan/ms-graph-mcp.git
cd ms-graph-mcp
npm install
npm test
```

That is the whole setup. **The test suite needs no Microsoft account, no tenant
and no network** — every HTTP call is stubbed. If a change you make requires a
credential to test, that is a signal the change is in the wrong layer.

| Command | What it does |
| --- | --- |
| `npm test` | Builds, then runs the suite. This is the one to run. |
| `npm run typecheck` | Types only, no emit. |
| `npm run build` | Compile to `dist/`. |
| `npm run contract` | Checks our scopes and paths against live Microsoft data. Needs the network. |

`npm run contract` is separate from `npm test` on purpose: it depends on two
services we do not control, and your pull request should not fail because
Microsoft shipped something overnight. CI runs it weekly instead.

## Shape of the code

```
src/contracts.ts   shared types; everything is built against these
src/config.ts      flags, environment, clouds, the tool group catalogue
src/index.ts       the executable: serve | login | logout | status | permissions
src/server.ts      MCP registration, the tool wrapper, discovery mode
src/auth/          MSAL public-client sign-in, the encrypted token cache
src/graph/         the HTTP client, error mapping, the paginator
src/tools/         one module per product area, plus groups/generic/discovery
src/util/          logger, truncation, OData helpers, the path guard
```

A tool module exports a `ToolModule` and returns `ToolDefinition[]`. It never
touches the MCP SDK — registration happens in `server.ts` — which is what keeps
the modules testable without a protocol in the way.

## Adding a tool

1. Put it in the module for its product area, or add a module and register it in
   `collectTools`.
2. Declare the **least-privileged** scopes that tool actually needs. Per-tool
   scopes are load-bearing: the consent request is the union of them, so a scope
   you add here is a scope every user is asked to grant.
3. Set `write: true` if it changes anything. That drives `--read-only` and the
   `destructiveHint` a client uses to decide whether to confirm.
4. Project the result down. Graph returns far more than a model can read, and
   clients truncate tool output — return the fields that answer the question,
   plus a cursor when more remains.
5. Write the description for the model, not for a person. If the endpoint has a
   trap, the description is where it goes. Several already say things like
   "`$filter` is silently ignored unless `$orderby` names the same property",
   because that is cheaper than the model rediscovering it.

## Things that will be asked in review

- **Never write to stdout in `serve` mode.** It carries the JSON-RPC framing; a
  stray `console.log` corrupts the session. Diagnostics go to stderr. The CLI
  subcommands are the only place `out()` is called.
- **Build URLs through `resolveGraphUrl`**, never by concatenation, and do not
  weaken `src/util/paths.ts`. It is the guard that keeps a model-chosen path from
  reaching a host that is not Graph — with your bearer token attached.
- **Sign-in never happens inside a request.** `getToken` refreshes silently and
  throws `InteractionRequiredError` otherwise. It must not open a browser.
- **No new dependencies** without a reason in the pull request. The dependency
  list is three packages and staying small is deliberate for something that holds
  a mailbox token.
- Match the surrounding style. Comments explain why, not what.

## Commit messages

A short subject line, then a body explaining *why* if it is not obvious. Long,
specific messages are fine and preferred over short vague ones.

## Reporting a bug

Include what you ran, what you expected, what happened, and the output of
`npx ms-graph-mcp status` **with the client ID and tenant ID redacted**.

Never paste a token, a `token-cache.enc`, a `cache.key`, or real mail or chat
content into an issue. For anything that looks like a security problem, see
[SECURITY.md](SECURITY.md) and report it privately instead.

## Licence

Contributions are accepted under the [MIT licence](LICENSE).
