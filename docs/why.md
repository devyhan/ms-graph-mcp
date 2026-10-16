# Why an MCP server

Microsoft Graph is an ordinary HTTP API. Everything this server does could, in principle, be a
`curl` command or a fifty-line script, and for a question you ask once it should be. This project
exists for the other case: an assistant that reaches the same API many times a day, on your
behalf, without being handed your credentials and without being told how Graph works on every
turn.

Five things move from the conversation into the server.

## The assistant never handles your token

For a model to call Graph directly, an access token has to reach the model — pasted into the
conversation, returned by a shell command, or read out of a file the model can open. Once it is
there it is in the transcript, in the context window, and in whatever logs that client keeps.
Access tokens are bearer credentials: anyone holding one is you, for an hour, to the full extent
of the scopes it carries.

Here the token never leaves the server process. The model sees tool names, JSON Schemas and
results; the refresh token sits encrypted on disk and the access token exists only inside a
request the model cannot observe. That boundary is the reason `getToken` refuses to become
interactive mid-request, and the reason no tool returns a `@microsoft.graph.downloadUrl` — a
pre-authenticated link that would hand file contents to anyone who read the transcript.

## Graph's sharp edges are encoded once, in the tool descriptions

Graph is not uniform, and its inconsistencies are the kind a model rediscovers by getting them
wrong. A few that this server already knows about, each of which cost a real debugging session
to establish:

- `GET /chats/{id}/messages` does not support `$select`. You cannot ask Graph for fewer fields;
  you get the whole `chatMessage` and trim it yourself.
- On that same endpoint `$filter` is **silently ignored** unless `$orderby` names the same
  property. Not rejected — ignored. A date-bounded query returns the wrong window and no error,
  which is why `chat_fetch_history` pins the ordering rather than exposing it.
- Channel messages are narrower still: no `$filter`, no `$orderby`, and an order derived from the
  reply chain rather than creation time. A date window there cannot be a server-side query or an
  early stop, so `teams_fetch_channel_history` filters while paging and tells you the window may
  be incomplete.
- Page size caps at 50, and Teams allows roughly one request per second per chat or channel. A
  naive loop over a long history is throttled; the paginator paces itself.
- Planner rejects every update and delete without an `If-Match` etag from a prior read.
- Outlook `$search` takes a quoted term and supports no date range at all; a date-bounded mail
  query has to use `$filter`.
- Directory queries using `$count` or `$search` need a `ConsistencyLevel: eventual` header that
  nothing in the URL suggests.

A model driving `curl` relearns these from error responses, or worse, from a 200 that quietly
returned the wrong rows. Written into a tool description once, they are in front of the model
before it makes the call.

## Responses are projected down to something a model can read

A single `chatMessage` carries an HTML body, reactions, mentions, attachments and hosted-content
references. Fifty of them is far more than a useful answer needs and more than many clients will
accept — Claude Code truncates tool output near 25,000 tokens. Raw Graph responses spend that
budget on markup.

Every tool here returns a projection instead: for a message, the sender, the timestamp, the body
stripped of HTML and truncated to a caller-controlled budget, and a count of attachments. Long
results carry a `nextLink` cursor so the model can continue deliberately rather than by
re-requesting and hoping.

## Guardrails a script does not have

- **Host pinning.** `graph_request` lets the model choose a path, and every request carries your
  bearer token, so a path that escapes the Graph host leaks that token. Three checks sit between
  the two: the path is validated, the URL is built through the URL API rather than by
  concatenation, and the resulting host is compared against the pinned Graph hostname. The same
  check runs on every `@odata.nextLink` before it is followed, because a paging link comes from a
  server response rather than from you. This is not hypothetical: a comparable Graph MCP server
  shipped a string-concatenated path that allowed an `@attacker.com` host escape.
- **Read-only mode.** `--read-only` removes write tools from the catalogue entirely, so a client
  never sees them.
- **Least-privilege consent.** Sign-in is incremental: a bare login asks only for the minimal
  set, and each tool call authenticates with just the scopes its own index needs.
- **Retry and backoff.** 429 and 503 are retried with `Retry-After` honoured; a mutating request is
  replayed only when Graph said it did not process it.

## One server, many clients

The same process serves Claude Code, Claude Desktop, VS Code and any other MCP client, over both
the 2026-07-28 revision and the 2025-11-25 one, from a single factory. A shell script is bound to
the one place it runs.

## When you do not need this

If you want one answer once, use `curl`. If you are writing a program that talks to Graph, use the
Microsoft Graph SDK — this is not a client library, and `graph_request` is a deliberate escape
hatch rather than a general interface. The server earns its keep when the caller is a model, the
access is ongoing, and the credential must not be in the conversation.

Worth being plain about the split: rather more than half of the code here is not MCP at all. It is
the part that calls Graph correctly — retries, batching, pagination, path validation, scope
derivation — and you would need it whatever the caller was. MCP is the thin, typed surface over
the top that lets a model use it safely.
