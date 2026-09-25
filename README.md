# ms-graph-mcp

A Model Context Protocol server that gives an AI assistant access to your own Microsoft 365 data — Outlook mail and calendar, OneDrive, SharePoint, To Do, Planner, Teams chat, Entra ID and Intune — through the Microsoft Graph API. It runs locally over stdio, signs in as you with delegated permissions, and stores its token cache encrypted on your machine.

> **Not affiliated with Microsoft.** This is an independent community project. Microsoft, Microsoft 365, Microsoft Graph, Outlook, OneDrive, SharePoint, Teams, Entra and Intune are trademarks of the Microsoft group of companies, used here only to describe what this software connects to.

## Why an MCP server

Microsoft Graph is an ordinary HTTP API. Everything this server does could, in principle, be a
`curl` command or a fifty-line script, and for a question you ask once it should be. This project
exists for the other case: an assistant that reaches the same API many times a day, on your
behalf, without being handed your credentials and without being told how Graph works on every
turn.

Five things move from the conversation into the server.

### The assistant never handles your token

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

### Graph's sharp edges are encoded once, in the tool descriptions

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

### Responses are projected down to something a model can read

A single `chatMessage` carries an HTML body, reactions, mentions, attachments and hosted-content
references. Fifty of them is far more than a useful answer needs and more than many clients will
accept — Claude Code truncates tool output near 25,000 tokens. Raw Graph responses spend that
budget on markup.

Every tool here returns a projection instead: for a message, the sender, the timestamp, the body
stripped of HTML and truncated to a caller-controlled budget, and a count of attachments. Long
results carry a `nextLink` cursor so the model can continue deliberately rather than by
re-requesting and hoping.

### Guardrails a script does not have

- **Host pinning.** `graph_request` lets the model choose a path, and every request carries your
  bearer token, so a path that escapes the Graph host leaks that token. Three checks sit between
  the two: the path is validated, the URL is built through the URL API rather than by
  concatenation, and the resulting host is compared against the pinned Graph hostname. The same
  check runs on every `@odata.nextLink` before it is followed, because a paging link comes from a
  server response rather than from you. This is not hypothetical: a comparable Graph MCP server
  shipped a string-concatenated path that allowed an `@attacker.com` host escape.
- **Read-only mode.** `--read-only` removes write tools from the catalogue entirely, so a client
  never sees them.
- **Least-privilege consent.** The scopes requested at sign-in are derived from the enabled groups
  and the tools inside them, so `--groups mail --read-only` asks for `Mail.Read` and nothing else.
- **Retry and backoff.** 429 and 503 are retried with `Retry-After` honoured; a mutating request is
  replayed only when Graph said it did not process it.

### One server, many clients

The same process serves Claude Code, Claude Desktop, VS Code and any other MCP client, over both
the 2026-07-28 revision and the 2025-11-25 one, from a single factory. A shell script is bound to
the one place it runs.

### When you do not need this

If you want one answer once, use `curl`. If you are writing a program that talks to Graph, use the
Microsoft Graph SDK — this is not a client library, and `graph_request` is a deliberate escape
hatch rather than a general interface. The server earns its keep when the caller is a model, the
access is ongoing, and the credential must not be in the conversation.

Worth being plain about the split: rather more than half of the code here is not MCP at all. It is
the part that calls Graph correctly — retries, batching, pagination, path validation, scope
derivation — and you would need it whatever the caller was. MCP is the thin, typed surface over
the top that lets a model use it safely.

## Quick start

You need Node.js 20 or later. Sign in once, in a terminal:

```
npx ms-graph-mcp login
```

There is no application to register first: the server signs in with a shared multi-tenant Entra application published by this project. See [Which application am I signing in to](#which-application-am-i-signing-in-to) for what that means for you and for your tenant.

**The shared registration is not published yet.** The constant it will occupy, `SHIPPED_CLIENT_ID` in `src/config.ts`, is still empty, so this build has no default application to fall back on. `login` says so and stops before it contacts Entra, and `status` prints the same registration walkthrough. Until the real ID ships, register an application of your own — or use one your organisation already has — and name it on each command:

```
npx ms-graph-mcp login --client-id <your-client-id>
```

[Register your own Entra application](#register-your-own-entra-application) has the steps. When the shared ID lands, nothing else in this README changes: `--client-id` keeps working and keeps taking precedence over the shipped default.

Either way, your system browser opens on the Microsoft sign-in page. Pick the account you want the server to act as, approve the permissions it asks for, and the tab reports that sign-in is complete. Back in the terminal, the command has caught the redirect, exchanged it for tokens, and written the refresh token to an encrypted cache under your config directory.

On a machine with no browser to open — an SSH session, a container, a headless server — ask for the device code flow instead:

```
npx ms-graph-mcp login --auth-flow device
```

That prints a short code and a URL to open on any other device. Enter the code there and the terminal finishes on its own. The default, `--auth-flow auto`, already falls back to this when the browser cannot be opened; passing `device` explicitly skips the attempt that would fail. Both are OAuth 2.0 sign-ins and both end with the same delegated token — see [How sign-in works](#how-sign-in-works).

Either way, check it worked:

```
npx ms-graph-mcp status
```

`status` prints the application ID in use and where it came from, the account it resolved, and the scopes this configuration will request. Then point an MCP client at the server.

### As a Claude Code plugin

The plugin wires the server into Claude Code and asks you for its settings, so there
is no MCP config to edit and nothing to export in a shell profile.

```
/plugin marketplace add devyhan/ms-graph-mcp
/plugin install ms-graph-mcp@devyhan
```

Claude Code prompts for four settings when the plugin is enabled, and fills in
sensible values for three of them:

| Setting | Default | What it does |
| --- | --- | --- |
| Application (client) ID | none, and required | The Entra application to sign in with. The dialog will not accept an empty value, because there is nothing to fall back on. |
| Directory (tenant) ID | `common` | Suits most work and school accounts. Use `consumers` for a personal Microsoft account. |
| Tool groups | the personal set | Profile, mail, calendar, files, To Do, contacts and search. A shorter list means a smaller consent prompt. |
| Read-only | on | Hides every tool that writes and drops the write scopes from the consent request. Turn it off when you want the server to act, not just read. |

Read-only defaults to on here, and to off on the command line. The plugin is the
path someone takes without reading this file first, and an assistant that can send
mail and delete files on a fresh install is the wrong first state; on the command
line the choice is already explicit.

You can change any of them later from `/config`.

You still need your own Entra application — see
[Register your own Entra application](#register-your-own-entra-application) — because
this project publishes no shared one. Paste its **Application (client) ID** into the
first prompt.

Sign-in remains a terminal step, once:

```
npx ms-graph-mcp login
```

The server will not open a browser from inside a Claude Code session. The MCP
transport owns stdout, and a sign-in prompt in the middle of a tool call would
corrupt the JSON-RPC stream, so it returns an error telling you to run `login`
instead. See [How sign-in works](#how-sign-in-works).

Everything works before you configure anything: the server starts, lists its tools,
and `status` explains what is missing. An installed-but-unconfigured plugin shows a
working server rather than a broken one.

**If a setting does not take effect**, run `npx ms-graph-mcp status` and read
the `Client ID from:` line — it names which source won. Should a value arrive as the
literal text `${user_config.client_id}`, the placeholder was never substituted;
the server refuses it and says so rather than passing it to Entra, where it would
fail much later and blame something else.

### Without the plugin

Every setting is also an `MS365_MCP_*` environment variable, which is what the plugin
sets under the hood, and every one has a command-line flag — see
[Commands and flags](#commands-and-flags). Use these when you run the server directly:

```sh
export MS365_MCP_CLIENT_ID="<your application (client) ID>"
export MS365_MCP_TENANT_ID="<your tenant ID>"
export MS365_MCP_GROUPS="mail,calendar,files"
export MS365_MCP_READ_ONLY=1
```

A client launched from a desktop icon rather than a terminal does not read your shell
profile, so a variable exported there will not reach it. That is the failure the
plugin's prompts exist to avoid; if you are configuring by hand, prefer the flags.

### Claude Code

```
claude mcp add microsoft-graph -- npx -y ms-graph-mcp --preset work
```

Or add it to `.mcp.json` in the project root:

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["-y", "ms-graph-mcp", "--preset", "work"]
    }
  }
}
```

### Claude Desktop

Edit `claude_desktop_config.json` (Settings > Developer > Edit Config):

```json
{
  "mcpServers": {
    "microsoft-graph": {
      "command": "npx",
      "args": ["-y", "ms-graph-mcp", "--preset", "work"]
    }
  }
}
```

### VS Code

Add `.vscode/mcp.json` to the workspace:

```json
{
  "servers": {
    "microsoft-graph": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "ms-graph-mcp", "--preset", "work"]
    }
  }
}
```

While the shared registration is unpublished, give each of these your own application ID as well: add `"env": { "MS365_MCP_CLIENT_ID": "<your-client-id>" }` beside `args`, or append `"--client-id", "<your-client-id>"` to `args`.

The server starts whether or not an application is configured. With none it writes a warning to stderr, lists its tools as normal, and fails every Graph call until you sign in.

## Which application am I signing in to

**A shared application published by this project.** Unless you override it, the client ID the server presents to Entra is a multi-tenant registration this project publishes — the same one every other user of this package signs in with. The server picks one in this order, and stops at the first that is set:

1. `--client-id <id>`
2. `MS365_MCP_CLIENT_ID`
3. the application shipped with the package
4. nothing, in which case sign-in refuses to start and prints the registration steps

`status` reports which of the four it landed on, and `login` names the authority and the application on stderr before it opens the browser, so you can see whose application you are about to consent to while you can still stop. Shipping an application ID is the model Lokka and the Softeria MS-365 MCP server use, and it exists for one reason: registering an application is the step most people never get past.

**First use in an organisation creates an enterprise application in that tenant.** The first time anyone in a tenant consents to a multi-tenant app, Entra creates a service principal for it in that directory — an entry under **Enterprise applications** carrying this project's application name. Administrators can see it there, review the delegated permissions it has been granted, and disable or block it. From then on every request the server makes is attributed to that application, alongside your user, in the tenant's sign-in and Graph activity logs. This is a consequence of signing in with a shared application, not a footnote: if your organisation would object to an unfamiliar third-party application appearing in its directory, register your own before the first sign-in.

**The application holds no secret and no data.** It is a public client: there is no client secret anywhere in this package, and there could not be one, because a secret shipped to every user is not a secret. Every token Entra issues is issued to the signed-in user, is scoped to what that user can already reach, and is cached encrypted on that user's own machine (see [Security](#security)). No server of ours sits between you and Microsoft, and neither the package nor the owner of the registration ever receives your tokens, mail, files or directory records. What the owner of a registration does control is the registration object itself, which is what the next two paragraphs are about.

**Using your own registration instead.** Pass `--client-id` (or set `MS365_MCP_CLIENT_ID`) and follow [Register your own Entra application](#register-your-own-entra-application). Reasons to do it:

- An enterprise policy that permits only applications registered in its own tenant, or that requires an internal owner and a review for every enterprise application.
- A sovereign cloud. The shared registration exists in the global cloud only, so the server refuses to offer it to `--cloud usgov`, `usgovdod` or `china` and says why, rather than letting you discover it after typing your password. Register an application in that cloud and name it with `--client-id`.
- A tenant that has blocked the shared application, or an administrator who will grant consent only to something they own.

**If the shared registration is ever deleted or blocked**, sign-in starts failing with `AADSTS700016`, `AADSTS7000112` or `AADSTS700054`, and `--client-id <your-id>` is the entire fix. Nothing else changes: same flows, same scopes, same tool catalogue, same cache location. Run `login` once after switching, because the cached refresh token belongs to the application that obtained it.

## Consent

**Shipping an application ID removes the registration step. It does not remove the consent step.** Entra still asks the signed-in user to approve every delegated scope the server requests, and the tenant's consent policy still decides which of those a user is allowed to approve without an administrator.

**Work and school accounts in a default tenant.** Under the Microsoft-managed default consent policy that new tenants get, an ordinary user cannot self-consent to `Mail.Read`, `Calendars.Read`, `Calendars.ReadWrite`, `Calendars.Read.Shared`, `Chat.Read`, `Tasks.Read`, `Tasks.ReadWrite`, `Files.Read.All`, `Sites.Read.All` or `MailboxSettings.Read`, among others — even though Microsoft's permissions reference marks every one of them as not requiring admin consent. That covers most of what the `personal` and `work` presets ask for. Sign-in fails at the consent screen, and the server prints which of your requested scopes fall under that policy, however small the request was. An administrator clears it once: **Enterprise applications > (the application) > Permissions > Grant admin consent** for the shared app, or **App registrations > (your app) > API permissions > Grant admin consent** for your own.

**Personal Microsoft accounts** consent for themselves — no administrator, no tenant policy. In exchange, a large part of the catalogue is unavailable to them at any consent level: the Teams chat and channel APIs are documented as not supported for delegated personal-account access, and Planner, SharePoint, the directory tools and Intune do not exist for a personal account at all. Use a narrow group set, and sign in against the personal-account authority with `--tenant-id consumers`; a refresh token issued through the `common` authority is rejected at its first refresh.

**Ask for less, then widen.** The scopes requested at sign-in are derived from the enabled groups, so start narrow and add:

```
npx ms-graph-mcp login --groups me,mail
npx ms-graph-mcp login --groups me,mail,calendar,files
```

Entra prompts again only for what you added; what was already granted stays granted. `--read-only` drops every write scope from the request. `npx ms-graph-mcp permissions --preset work` prints the exact scope list for a configuration, formatted for pasting into an admin consent request.

## Register your own Entra application

Once the shared registration ships, most users can skip this section. It remains the supported fallback for the reasons listed in [Which application am I signing in to](#which-application-am-i-signing-in-to), and until then it is the only way to sign in.

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com) and go to **Applications > App registrations > New registration**.
2. Give it a name. Under **Supported account types**, pick the audience you need. A personal Microsoft account (outlook.com, hotmail.com) requires "Accounts in any organizational directory and personal Microsoft accounts"; a work or school account used only in its own tenant needs nothing more than the single-tenant option.
3. Leave the redirect URI empty for now and click **Register**. Copy the **Application (client) ID** from the Overview page — that is the value for `--client-id`.
4. Go to **Authentication > Add a platform > Mobile and desktop applications** and tick the redirect URI `http://localhost`. The browser flow, which is the default, redirects the finished sign-in there — to a listener on 127.0.0.1 — and fails without it.
5. On the same page, scroll to **Advanced settings** and set **Allow public client flows** to **Yes**. Neither sign-in flow works without it.
6. Optionally, under **API permissions**, add the delegated Microsoft Graph permissions for the tool groups you plan to enable (see the table below). This is not required for user-consentable scopes — you will be prompted to consent at first sign-in — but it is required for the admin-consent groups, and for anything the tenant's consent policy withholds from ordinary users, where an administrator must click **Grant admin consent**.

`http://localhost` is a wildcard: Entra accepts a redirect back to it on whatever port the machine happened to be listening on. Some tenants have policies that reject it and require an exact port. In that case register `http://localhost:53682` (any free port will do) instead, and start the server with `--auth-port 53682` so it binds the port the registration names.

No client secret is needed, and you should not create one. This is a public client running on your machine: it authenticates you interactively and stores a refresh token, and a secret shipped alongside it would not be a secret.

A personal Microsoft account also needs `--tenant-id consumers` at sign-in, whatever the registration looks like.

### Registering a multi-tenant application

Step 2 above is enough for a registration you alone use. An organisation cloning this project's setup — one registration, many users, more than one tenant — needs the same shape the shared application has:

| Manifest property | Value | Why |
| --- | --- | --- |
| `signInAudience` | `AzureADMultipleOrgs`, or `AzureADandPersonalMicrosoftAccount` to include personal accounts | `AzureADMyOrg` exists only in the tenant that registered it and returns `AADSTS700016` in every other one. |
| `api.requestedAccessTokenVersion` | `2` | Required by any audience that includes personal Microsoft accounts. |
| `isFallbackPublicClient` | `true`, shown as "Allow public client flows: Yes" | Both sign-in flows are public-client flows. |
| Redirect URI | `http://localhost` under **Mobile and desktop applications** | Where the browser flow's authorization code lands. |

The ordering gotcha: if you change the audience to one that includes personal accounts while the token version is still 1, the portal rejects the save with **"Property api.requestedAccessTokenVersion is invalid"**. Set `api.requestedAccessTokenVersion` to `2` in the manifest and save that on its own first, then change `signInAudience`.

A multi-tenant registration also means the enterprise application described in [Which application am I signing in to](#which-application-am-i-signing-in-to) appears in every tenant whose users sign in, including yours.

## How sign-in works

Both sign-in flows are OAuth 2.0, and both end with the same delegated access token. They differ in how you prove who you are, not in what they are.

**`browser`** is the authorization code grant with PKCE (RFC 7636). The server binds a one-shot listener on the loopback address, 127.0.0.1, opens the system browser at the Entra authorize endpoint, and Entra sends the finished sign-in back to that listener as a redirect carrying an authorization code. The server exchanges the code, together with the proof key it generated before opening the browser, for tokens. The code is valid only on this machine and only for that one exchange.

**`device`** is the device authorization grant (RFC 8628). The server asks Entra for a user code, prints it with a URL, and polls the token endpoint while you enter the code in a browser somewhere else. It is no less an OAuth flow than the browser one; it exists for the case where the machine running the server has no browser it can open, and the sign-in has to happen on a different device.

**`auto`** is the default. It tries `browser`, and when that fails — no desktop session, no way to launch a browser, no loopback port — it writes the reason to stderr and runs `device` instead. Pass `--auth-flow browser` when you would rather see that error than a fallback, and `--auth-flow device` when you never want a browser opened.

**Delegated permissions only.** The server signs in as you and calls Graph as you. It never authenticates as an application, holds no client secret, and requests no application permissions, so it reaches exactly what your account reaches and nothing more. Every request it makes is attributable to your user in the tenant's sign-in logs.

**The session is a refresh token in an encrypted cache.** What sign-in leaves behind is the MSAL cache blob, encrypted with AES-256-GCM into `token-cache.enc` in your config directory (`~/.config/microsoft-graph-mcp` on macOS and Linux, `%APPDATA%\microsoft-graph-mcp` on Windows), with the 32-byte key beside it in `cache.key` at mode `0600`. See [Security](#security) for the details and for `--cache-dir`.

**`getToken` never prompts.** In `serve` mode the server refreshes access tokens silently and does nothing else. When the refresh token has expired, been revoked, or no longer covers a scope, the tool call returns an error telling you to run `login` again rather than starting a sign-in. That is deliberate: the MCP transport owns stdout, so a prompt or a browser launched in the middle of a request would corrupt the JSON-RPC stream. Sign-in happens out of band, in a terminal.

**This is a local server.** It does not implement the MCP authorization specification, and per that specification it should not: the spec covers HTTP transports and directs stdio servers to take credentials from the environment instead. So the server obtains user credentials itself, as an ordinary public OAuth client, and the MCP client is not involved in authentication at all. A remote HTTP mode — where the MCP client performs OAuth against Entra and the server exchanges that token for a Graph token on the user's behalf — is not implemented.

### When the browser flow does not work

**Nothing opens, or the command sits there.** There is no desktop session to open a browser in, which is normal over SSH and inside containers. Use `--auth-flow device` and sign in from another machine. Under the default `auto` flow this resolves itself, with a line on stderr saying what failed.

**The browser opens but the redirect is rejected** (`AADSTS50011`, redirect URI mismatch). The app registration has no matching redirect URI. Add the **Mobile and desktop applications** platform with `http://localhost`, as in step 4 above.

**The loopback listener cannot bind.** With the default `--auth-port 0` the operating system picks a free port, so this only comes up when a port has been pinned and something else already holds it, or when a local firewall blocks the listener. Choose another port with `--auth-port <n>`. Pin a port deliberately, and register `http://localhost:<n>` alongside it, when the tenant rejects the bare `http://localhost` wildcard.

## Sign-in errors

Entra reports a failure as an `AADSTS` code inside the error text. These are the ones that come from the application identity or from consent, rather than from the browser leg.

| Code | What it means | What fixes it |
| --- | --- | --- |
| `AADSTS700016` | The application exists, but not in the tenant you signed in to. A single-tenant registration returns this to every tenant except its own; so does one the tenant has blocked, or whose enterprise application an administrator deleted. | Sign in to the tenant that owns the registration (`--tenant-id <tenant-id>`), or use a multi-tenant registration. For the shared application, `--client-id` with your own is the fallback. |
| `AADSTS700038` | Entra did not accept the value as an application identifier at all. A mistyped GUID, or one that was never registered. | Check what `status` prints against the **Application (client) ID** on the registration's Overview page, and correct `--client-id` or `MS365_MCP_CLIENT_ID`. |
| `AADSTS50020` | The account comes from a different identity provider than the application accepts: a personal Microsoft account at a work-only application, or the reverse. | `--tenant-id consumers` for a personal Microsoft account; otherwise sign in with a work or school account. The server names the `--tenant-id` the failed run used. |
| `AADSTS9002332` | The application accepts Microsoft Entra work or school users only, and you signed in with a personal Microsoft account. | Sign in with a work or school account, or pass `--client-id` for a registration whose audience includes personal accounts and add `--tenant-id consumers`. |
| `AADSTS7000112`, `AADSTS700054` | The application is disabled in that tenant — its enterprise application was disabled, or the service principal was blocked from sign-in. | Ask an administrator to re-enable it under **Enterprise applications > Properties > "Enabled for users to sign-in"**, or pass `--client-id` with your own registration. |
| `AADSTS65004` | The consent screen was declined. Nothing was granted and no session was created. | Ask for less: a narrower `--groups`, or `--read-only` to drop the write scopes. |
| `AADSTS90093` | The requested permissions need an administrator, and the signed-in user is not allowed to grant them. | Drop the admin-consent groups (do not pass `--org-mode`) and narrow `--groups`, or have an administrator grant consent once, as in [Consent](#consent). |

The server sorts these into three kinds, and only one of them falls back.

**Application rejections** — the first five rows — are reported as `Sign-in failed (<code>)`, followed by the likely cause, the flag that fixes it, and the first line of what Entra actually said. `auto` does not retry them on the device code path: an application that is absent, disabled, or wrong for your account type is exactly as absent on another device. The MSAL stack goes to stderr separately, prefixed `[auth]`.

**Consent decisions** — the last two rows, plus `AADSTS65001` and `AADSTS900941` — print `Sign-in was refused at the consent screen`, the number of scopes requested, and the subset of them the default tenant consent policy withholds from ordinary users. `auto` does not fall back here either, because the device code flow shows the same consent screen and would be refused the same way.

**Browser failures** are everything else. Under `auto` the server writes `Browser sign-in failed: ...` to stderr and continues on the device code path. Under `--auth-flow browser` it stops and prints the three usual causes: no desktop session, a loopback redirect that could not be served, and a registration with no matching redirect URI.

## Tools

101 tools across 14 groups. `--read-only` removes every write tool, leaving 64. The reference for each group is below; the descriptions here are the first line of what the model itself is shown.

### Tool groups

Tools are grouped by product area. Enable only the groups you need — the scopes requested at sign-in are derived from the enabled groups, so a narrower selection means a narrower consent prompt.

| Group | Product | Delegated scopes | Admin consent | Tools |
| --- | --- | --- | --- | --- |
| `me` | Entra ID profile, Outlook mailbox settings | `User.Read`, `MailboxSettings.Read` | No | 2 |
| `mail` | Outlook mail | `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` | No | 12 |
| `calendar` | Outlook calendar | `Calendars.Read`, `Calendars.Read.Shared`, `Calendars.ReadWrite` | No | 9 |
| `files` | OneDrive | `Files.Read`, `Files.Read.All`, `Files.ReadWrite` | No | 10 |
| `todo` | Microsoft To Do | `Tasks.Read`, `Tasks.ReadWrite` | No | 8 |
| `planner` | Microsoft Planner | `Tasks.Read`, `Tasks.ReadWrite` | No | 9 |
| `contacts` | Outlook contacts | `Contacts.Read`, `Contacts.ReadWrite` | No | 6 |
| `chat` | Teams one-to-one and group chats | `Chat.Read`, `Chat.ReadWrite` | No | 6 |
| `sharepoint` | SharePoint sites, libraries, lists | `Sites.Read.All`, `Sites.ReadWrite.All` | No | 11 |
| `search` | Microsoft Search across mail, files, sites | `Mail.Read`, `Files.Read`, `Sites.Read.All`, `Calendars.Read`, `Chat.Read` | No | 1 |
| `teams` | Teams teams, channels, channel messages | `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `TeamMember.Read.All` | **Yes** | 9 |
| `directory` | Entra ID user and group lookups | `User.Read.All`, `Group.Read.All` | **Yes** | 9 |
| `intune` | Intune managed device inventory | `DeviceManagementManagedDevices.Read.All` | **Yes** | 6 |
| `generic` | Any Graph endpoint, plus introspection | reuses the union of the above | No | 3 |

`offline_access` and `User.Read` are always requested: the first keeps refresh tokens working, the second identifies the account.

The three groups marked **Yes** request scopes that no ordinary user can consent to, so they stay disabled unless you pass `--org-mode`. Enabling them in a tenant where an administrator has not granted consent means sign-in fails outright, not that those tools quietly return errors.

The **Admin consent** column is Microsoft's own classification of each permission. A tenant's consent policy is a second, separate gate, and it withholds more than this column shows — `Mail.Read` and `Calendars.Read` among them. See [Consent](#consent).

The `generic` group is always on and cannot be selected or removed. It holds `graph_request` (the escape hatch for endpoints no other tool covers), `graph_schema` (fetches one item from a path and reports its property names and types), and `graph_list_permissions` (reports the enabled groups, their scopes, and the signed-in account).

### Tool reference

<details>
<summary><b><code>me</code></b> — My profile: 2 tools, all read-only</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `me_get_profile` |  | Returns the signed-in user's own Entra ID profile from /me: by default id, displayName, mail, userPrincipalName, jobTitle, officeLocation and preferredLanguage. | `select` |
| `me_get_mailbox_settings` |  | Returns the signed-in user's Outlook mailbox configuration: time zone, date and time format, locale, working hours, and the automatic-reply (out-of-office) setting. | `replyBodyChars` |

</details>

<details>
<summary><b><code>mail</code></b> — Mail: 12 tools, 7 of them writes</summary>

Bodies are the expensive part of a mailbox, so list calls return `bodyPreview` only and `mail_get_message` fetches one message in full, HTML stripped to text by default.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `mail_list_messages` |  | Lists Outlook messages newest-first, 10 per call by default (50 max), as a compact projection: id, subject, from, toRecipients, receivedDateTime, isRead, hasAttachments, bodyPreview and webLink. | `folderId`, `top`, `skip`, `filter`, `orderby`, `select`, `unreadOnly` |
| `mail_search_messages` |  | Finds messages by text, returning the same compact projection as mail_list_messages, 10 per call by default (50 max). | **`query`**, `top`, `from`, `to`, `after`, `before`, `hasAttachments` |
| `mail_get_message` |  | Returns one message with its body plus sender, recipients, timestamps, importance, conversationId and webLink. | **`id`**, `format`, `maxBodyChars` |
| `mail_list_folders` |  | Lists the top-level Outlook mail folders with their ids, unread counts and total counts, 50 per call by default. | `top`, `includeChildren` |
| `mail_list_attachments` |  | Lists attachment metadata for one message — id, name, contentType, size in bytes, whether it is inline, and its kind (file, item, or reference) — 20 per call by default. | **`id`**, `top` |
| `mail_send` | **write** | Sends a message immediately from the signed-in mailbox — there is no undo and no confirmation step, so confirm the recipients and text with the user first. | **`to`**, **`subject`**, **`body`**, `cc`, `bcc`, `contentType`, `saveToSentItems` |
| `mail_create_draft` | write | Creates an unsent draft in the Drafts folder and returns its id and webLink. | `to`, **`subject`**, **`body`**, `cc`, `bcc`, `contentType` |
| `mail_reply` | **write** | Sends a reply to an existing message immediately — it does not create a draft, and there is no undo, so confirm the text with the user first. | **`id`**, **`comment`**, `replyAll` |
| `mail_forward` | **write** | Forwards an existing message immediately, attachments included — it does not create a draft and there is no undo, so confirm the recipients with the user first. | **`id`**, **`to`**, `comment` |
| `mail_move_message` | write | Moves a message to another mail folder. | **`id`**, **`destinationId`** |
| `mail_mark_read` | write | Sets the read flag on one message. | **`id`**, `isRead` |
| `mail_delete_message` | **write** | Deletes a message. | **`id`** |

</details>

<details>
<summary><b><code>calendar</code></b> — Calendar: 9 tools, 4 of them writes</summary>

`calendar_list_events` switches to `/me/calendarView` when you give it a date range, which is the only form that expands recurring series. `calendar_get_schedule` and `calendar_find_meeting_times` are POSTs that read rather than write.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `calendar_list_events` |  | Lists events with subject, start, end, location, organizer and attendee count, plus a nextLink when more pages exist. | `start`, `end`, `calendarId`, `top`, `select`, `orderby`, `timeZone` |
| `calendar_get_event` |  | Returns one event in full: subject, start, end, location, organizer, the attendee list with each response, the description body (HTML converted to text and capped), recurrence summary, and the Outlook web link. | **`id`**, `timeZone` |
| `calendar_list_calendars` |  | Lists the signed-in user's calendars with id, name, colour, owner and whether they can be edited or shared. | `top` |
| `calendar_get_schedule` |  | Returns free/busy availability for up to 20 people or rooms over a time window, as an availabilityView string where each character covers one interval (0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere), plus working hours and any visible busy blocks. | **`schedules`**, **`start`**, **`end`**, `availabilityViewInterval`, `timeZone` |
| `calendar_find_meeting_times` |  | Suggests meeting slots that work for the signed-in user and the given attendees, ranked by confidence, with the attendees who are unavailable in each slot. | **`attendees`**, **`durationMinutes`**, `start`, `end`, `maxCandidates`, `timeZone` |
| `calendar_create_event` | write | Creates an event on the primary calendar and returns the created event. | **`subject`**, **`start`**, **`end`**, `timeZone`, `attendees`, `body`, `bodyType`, `location`, `isAllDay`, `isOnlineMeeting` |
| `calendar_update_event` | write | Patches the given fields on an event and returns the updated event; omitted fields are left alone. | **`id`**, `subject`, `start`, `end`, `timeZone`, `attendees`, `body`, `bodyType`, `location`, `isAllDay`, `isOnlineMeeting`, `showAs`, `categories`, `reminderMinutesBeforeStart` |
| `calendar_delete_event` | **write** | Deletes an event, moving it to Deleted Items. | **`id`** |
| `calendar_respond_event` | **write** | Accepts, declines, or tentatively accepts a meeting invitation, optionally with a comment to the organizer. | **`id`**, **`response`**, `comment`, `sendResponse` |

</details>

<details>
<summary><b><code>files</code></b> — Files: 10 tools, 4 of them writes</summary>

Path-addressed calls encode each segment separately, so spaces and `&` in filenames work. `files_read_text` refuses binary types by extension before downloading, and no tool returns `@microsoft.graph.downloadUrl`, which is a pre-authenticated link to the file contents.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `files_list_children` |  | Lists the direct children of a folder in the signed-in user's OneDrive, returning id, name, kind (file or folder), size, mime type, child count, last-modified time and webUrl per entry, plus count and nextLink when more pages exist. | `itemId`, `path`, `top`, `orderby` |
| `files_get_item` |  | Reads the metadata of one OneDrive file or folder: name, kind, size, mime type, created and last-modified times and authors, parent folder path and ids, and webUrl. | `itemId`, `path` |
| `files_search` |  | Searches the whole of the signed-in user's OneDrive for free text, matching file names and indexed file contents, and returns the same compact fields as files_list_children plus count and nextLink. | **`query`**, `top` |
| `files_list_recent` |  | Lists files the signed-in user recently opened or edited, most recent first, with count and nextLink. | `top` |
| `files_list_shared` |  | Lists files and folders other people have shared with the signed-in user, with count and nextLink. | `top` |
| `files_read_text` |  | Returns the contents of a text-like OneDrive file (plain text, Markdown, CSV, JSON, XML, source code, config) as a string, truncated to maxChars, which defaults to 20000. | **`itemId`**, `maxChars` |
| `files_create_folder` | write | Creates a folder in the signed-in user's OneDrive and returns the created item. | **`name`**, `parentItemId`, `parentPath` |
| `files_upload_text` | write | Writes a text file to the given path in the signed-in user's OneDrive and returns the stored item. | **`path`**, **`content`** |
| `files_delete` | **write** | Deletes a file or folder from the signed-in user's OneDrive. | **`itemId`** |
| `files_create_link` | write | Creates a sharing link for a OneDrive file or folder and returns its URL, permission type and scope. | **`itemId`**, `type`, `scope` |

</details>

<details>
<summary><b><code>todo</code></b> — To Do: 8 tools, 5 of them writes</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `todo_list_lists` |  | Lists the signed-in user's Microsoft To Do lists with their id, display name and well-known name (for example defaultList for the built-in Tasks list). | none |
| `todo_list_tasks` |  | Lists tasks in one To Do list, returning id, title, status, importance, due and reminder times and a short body preview per task, plus count and nextLink when more pages exist. | **`listId`**, `top`, `filter`, `includeCompleted` |
| `todo_get_task` |  | Reads one To Do task in full: title, status, importance, reminder state, start/due/completed times, categories and the notes body (truncated). | **`listId`**, **`taskId`** |
| `todo_create_task` | write | Creates a task in a To Do list and returns the created task. | **`listId`**, **`title`**, `body`, `dueDateTime`, `reminderDateTime`, `importance`, `timeZone` |
| `todo_update_task` | write | Patches an existing To Do task and returns the updated task. | **`listId`**, **`taskId`**, `title`, `body`, `status`, `importance`, `dueDateTime`, `reminderDateTime`, `clearDueDateTime`, `clearReminderDateTime`, `timeZone` |
| `todo_complete_task` | write | Marks a To Do task completed by patching its status to 'completed', and returns the updated task including the completedDateTime Graph stamps on it. | **`listId`**, **`taskId`** |
| `todo_delete_task` | **write** | Deletes a To Do task permanently and returns a confirmation. | **`listId`**, **`taskId`** |
| `todo_create_list` | write | Creates a new To Do list and returns its id and display name. | **`displayName`** |

</details>

<details>
<summary><b><code>planner</code></b> — Planner: 9 tools, 3 of them writes</summary>

Planner requires an `If-Match` etag on every update and delete. The tools fetch the item first and supply the etag themselves.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `planner_list_my_plans` |  | Lists the Planner plans the signed-in user can see, with id, title, owning Microsoft 365 group id and container. | `limit` |
| `planner_list_my_tasks` |  | Lists Planner tasks assigned to the signed-in user across every plan, with title, planId, bucketId, due date, percentComplete and assignee ids. | `limit` |
| `planner_get_plan` |  | Returns one Planner plan: id, title, owning Microsoft 365 group id, container and creation time. | **`planId`** |
| `planner_list_buckets` |  | Lists the buckets (columns) of a plan with id, name and orderHint. | **`planId`**, `limit` |
| `planner_list_plan_tasks` |  | Lists every task in a plan regardless of assignee, with title, bucketId, due date, percentComplete and assignee ids. | **`planId`**, `limit` |
| `planner_get_task` |  | Returns one Planner task. | **`taskId`**, `includeDetails` |
| `planner_create_task` | write | Creates a task in a plan and returns it. | **`planId`**, **`title`**, `bucketId`, `assigneeIds`, `dueDateTime`, `percentComplete`, `priority` |
| `planner_update_task` | write | Updates a Planner task and returns it. | **`taskId`**, `title`, `bucketId`, `assigneeIds`, `dueDateTime`, `clearDueDateTime`, `percentComplete`, `priority` |
| `planner_delete_task` | **write** | Permanently deletes a Planner task; there is no recycle bin. | **`taskId`** |

</details>

<details>
<summary><b><code>contacts</code></b> — Contacts: 6 tools, 3 of them writes</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `contacts_list` |  | Lists personal contacts from the signed-in user's default Outlook Contacts folder, 25 per page by default (max 100). | `top`, `skip`, `filter`, `orderby`, `select` |
| `contacts_get` |  | Returns one personal contact by id, with names, company, job title, department, every email address, phone numbers, postal addresses, birthday, categories and timestamps. | **`id`** |
| `contacts_search` |  | Full-text search over the default Contacts folder using Outlook `$search`, returning the same compact rows as contacts_list, 25 per page by default. | **`query`**, `top`, `skip` |
| `contacts_create` | write | Creates a personal contact in the signed-in user's default Contacts folder and returns the stored contact, including the new id. | **`givenName`**, `surname`, `emailAddresses`, `mobilePhone`, `companyName`, `jobTitle` |
| `contacts_update` | write | Updates the given fields of one personal contact and returns the stored result. | **`id`**, `givenName`, `surname`, `displayName`, `emailAddresses`, `mobilePhone`, `companyName`, `jobTitle` |
| `contacts_delete` | **write** | Deletes one personal contact. | **`id`** |

</details>

<details>
<summary><b><code>chat</code></b> — Chat: 6 tools, 1 of them writes</summary>

Graph offers no `$select` and no date filter on chat messages, so projection happens here and date windows are applied while paging. `chat_fetch_history` is the one to reach for beyond a single page.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `chat_list_chats` |  | Lists the Teams chats the signed-in user belongs to (one-to-one, group, and meeting chats — never channel conversations, which the teams group covers). | `top`, `filter`, `expandMembers` |
| `chat_get_chat` |  | Returns one Teams chat by id: topic, chatType, created and last-updated timestamps, the Teams deep link, whether the signed-in user has hidden it, how far they have read, and the meeting join URL for a meeting chat. | **`chatId`** |
| `chat_list_messages` |  | Reads recent messages from one Teams chat, newest first. | **`chatId`**, `top`, `orderby` |
| `chat_list_members` |  | Lists everyone in one Teams chat: the per-chat membership id, display name, email, directory user id, and roles (an owner of a group chat has ["owner"]; ordinary members have none). | **`chatId`** |
| `chat_send_message` | **write** | Posts a message to a Teams chat as the signed-in user. | **`chatId`**, **`content`**, `contentType` |
| `chat_fetch_history` |  | Reads a long run of one Teams chat, newest first, paging automatically so you do not have to follow nextLink yourself. | **`chatId`**, `since`, `until`, `maxMessages`, `bodyChars`, `includeSystem`, `cursor` |

</details>

<details>
<summary><b><code>sharepoint</code></b> — SharePoint: 11 tools, 3 of them writes</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `sharepoint_search_sites` |  | Searches SharePoint sites and returns id, display name, description and web URL for each match, plus a count and a nextLink when more pages exist. | **`query`**, `top` |
| `sharepoint_get_site` |  | Returns one site: its composite id, display name, description, web URL, hostname and timestamps. | **`siteId`** |
| `sharepoint_list_drives` |  | Lists the document libraries (drives) on a site with id, name, description, drive type and web URL, plus a count and a nextLink when more pages exist. | **`siteId`**, `top` |
| `sharepoint_list_drive_items` |  | Lists the children of a document-library folder: name, whether it is a file or a folder, size, child count, MIME type, last modified time and author, and the web URL, plus a count and a nextLink when more pages exist. | **`siteId`**, **`driveId`**, `itemId`, `folderPath`, `top` |
| `sharepoint_list_lists` |  | Lists the SharePoint lists on a site with id, display name, description, template and web URL, plus a count and a nextLink when more pages exist. | **`siteId`**, `includeHidden`, `top` |
| `sharepoint_get_list` |  | Returns one list: id, display name, description, template, web URL and timestamps. | **`siteId`**, **`listId`**, `includeColumns` |
| `sharepoint_list_list_items` |  | Lists items in a SharePoint list with their column values, plus a count and a nextLink when more pages exist. | **`siteId`**, **`listId`**, `top`, `filter`, `fields` |
| `sharepoint_get_list_item` |  | Returns one list item with its column values, creator, last editor, timestamps and web URL. | **`siteId`**, **`listId`**, **`itemId`**, `fields` |
| `sharepoint_create_list_item` | write | Creates an item in a SharePoint list and returns its new id, web URL and stored column values. | **`siteId`**, **`listId`**, **`fields`** |
| `sharepoint_update_list_item` | write | Updates column values on an existing list item and returns the stored values. | **`siteId`**, **`listId`**, **`itemId`**, **`fields`** |
| `sharepoint_delete_list_item` | **write** | Deletes one item from a SharePoint list. | **`siteId`**, **`listId`**, **`itemId`** |

</details>

<details>
<summary><b><code>search</code></b> — Search: 1 tools, all read-only</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `search_query` |  | Runs a relevance-ranked Microsoft Search query across Microsoft 365 and returns a flattened list of hits, each with its rank, a text summary and the resource's id, name or subject, web URL and last-modified time. | **`query`**, `entityTypes`, `from`, `size`, `fields` |

</details>

<details>
<summary><b><code>teams</code></b> — Teams: 9 tools, 1 of them writes</summary>

Channel messages have an even narrower query surface than chats: no `$filter`, no `$orderby`, and an order based on the reply chain rather than creation time. `teams_fetch_channel_history` filters client-side for that reason and says so.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `teams_list_joined` |  | Lists the Microsoft Teams teams the signed-in user is a member of, returning id, displayName, description, visibility and archived state. | none |
| `teams_list_channels` |  | Lists the channels of one team, returning id, displayName, description, membershipType and webUrl for each. | **`teamId`**, `membershipType` |
| `teams_get_channel` |  | Returns one channel: displayName, description, membershipType, creation time, the channel email address (empty unless the channel has one provisioned) and the deep link webUrl. | **`teamId`**, **`channelId`** |
| `teams_list_channel_messages` |  | Lists the top-level messages of a channel, newest first, with sender, timestamp, plain-text body (HTML stripped and truncated to `bodyChars`) and attachment, mention and reaction counts. | **`teamId`**, **`channelId`**, `top`, `bodyChars` |
| `teams_get_channel_message` |  | Returns one channel message in full: sender, timestamps, plain-text body truncated to `bodyChars`, attachment names, mention text and reaction types. | **`teamId`**, **`channelId`**, **`messageId`**, `bodyChars` |
| `teams_list_message_replies` |  | Lists the replies to one top-level channel message, newest first, with the same compact projection as teams_list_channel_messages. | **`teamId`**, **`channelId`**, **`messageId`**, `top`, `bodyChars` |
| `teams_fetch_channel_history` |  | Walks a channel's message history and returns a bounded page of root messages with sender, creation and modification times, plain-text body (HTML stripped, truncated to `bodyChars`) and attachment count. | **`teamId`**, **`channelId`**, `since`, `until`, `maxMessages`, `bodyChars`, `includeReplies`, `cursor` |
| `teams_list_members` |  | Lists the members of a team with displayName, email, the directory userId and their roles ("owner" for owners, empty for ordinary members, "guest" for guests). | **`teamId`**, `top` |
| `teams_send_channel_message` | **write** | Posts a new top-level message to a channel as the signed-in user and returns the created message id, timestamp and deep link. | **`teamId`**, **`channelId`**, **`content`**, `contentType`, `subject`, `importance` |

</details>

<details>
<summary><b><code>directory</code></b> — Directory: 9 tools, all read-only</summary>

Advanced queries against the directory need the `ConsistencyLevel: eventual` header; the client sets it automatically when a request uses `$count` or a directory `$search`.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `directory_list_users` |  | Lists users in the Microsoft Entra ID tenant, 25 per page by default (max 999). | `top`, `filter`, `select`, `orderby`, `skipToken` |
| `directory_search_users` |  | Finds users by name or email address, returning the same compact rows as directory_list_users, 25 per page by default. | **`query`**, `top`, `select`, `skipToken` |
| `directory_get_user` |  | Returns one user by object id or userPrincipalName, with names, contact details, job title, department, office, employee id, account state, usage location and on-premises sync fields. | **`id`** |
| `directory_get_user_manager` |  | Returns the user's manager as a compact row (id, displayName, userPrincipalName, mail, jobTitle, department). | **`id`** |
| `directory_list_direct_reports` |  | Lists the users who report directly to the given user, 25 per page by default (max 999), as compact rows. | **`id`**, `top`, `skipToken` |
| `directory_list_groups` |  | Lists groups in the Microsoft Entra ID tenant, 25 per page by default (max 999). | `top`, `filter`, `select`, `skipToken` |
| `directory_get_group` |  | Returns one group by object id, with its description, mail nickname, type flags, visibility, classification, lifecycle dates, dynamic membership rule and on-premises sync fields. | **`id`** |
| `directory_list_group_members` |  | Lists the direct members of a group, 25 per page by default (max 999). | **`id`**, `top`, `skipToken` |
| `directory_list_user_groups` |  | Returns the ids of every group the user belongs to, including memberships inherited through nested groups. | **`id`** |

</details>

<details>
<summary><b><code>intune</code></b> — Intune: 6 tools, all read-only</summary>

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `intune_list_managed_devices` |  | Lists Intune-managed devices in the tenant with id, name, operating system and version, compliance state, last sync time, primary user, manufacturer and model, plus count and nextLink when more pages exist. | `top`, `filter`, `select` |
| `intune_get_managed_device` |  | Reads one Intune-managed device in detail: hardware identity, primary user, enrollment and management state, encryption and supervision flags, Exchange access state, storage and memory in GB, and the last sync time. | **`deviceId`** |
| `intune_list_device_apps` |  | Lists the applications Intune has detected on one managed device, returning display name, version, size in MB and the tenant-wide deviceCount for each, with count and nextLink. | **`deviceId`**, `top` |
| `intune_list_compliance_policies` |  | Lists the tenant's Intune device compliance policies with id, display name, platform-specific policyType (from the @odata.type discriminator, e.g. | `top` |
| `intune_get_device_summary` |  | Returns the tenant-wide Intune device counts in one call: total enrolled devices, MDM-enrolled and dual-enrolled counts, a per-operating-system breakdown, and the Exchange access state breakdown. | none |
| `intune_list_noncompliant` |  | Lists managed devices whose compliance state is exactly 'noncompliant', with the usual device fields plus the compliance grace period expiry, count and nextLink. | `top` |

</details>

<details>
<summary><b><code>graph</code></b> — Generic: 3 tools, 1 of them writes</summary>

Always enabled and not selectable. `graph_request` is the escape hatch for endpoints no purpose-built tool covers.

| Tool | | What it does | Arguments |
| --- | --- | --- | --- |
| `graph_request` | **write** | Sends an arbitrary request to Microsoft Graph and returns the raw JSON response. | **`path`**, `method`, `version`, `query`, `body`, `maxPages` |
| `graph_schema` |  | Fetches one item from a Graph path ($top=1) and reports the property names it carries with their JSON types, plus the @odata.context that names the resource type. | **`entityPath`**, `version` |
| `graph_list_permissions` |  | Reports how this server is configured: which tool groups are enabled, the delegated Microsoft Graph scopes each one uses, which of them need a tenant administrator to consent, whether write tools are suppressed, and the signed-in account. | none |

</details>

### Presets

| Preset | Groups |
| --- | --- |
| `personal` | me, mail, calendar, files, todo, contacts, search |
| `work` | personal plus chat, planner, sharepoint |
| `admin` | directory, intune, teams |
| `all` | every group |

`personal` is the default when neither `--groups` nor `--preset` is given. `--groups` and `--preset` can be combined; their union is taken.

## Commands and flags

```
ms-graph-mcp [command] [options]
```

| Command | Effect |
| --- | --- |
| `serve` | Run the MCP server on stdio. The default when no command is given. |
| `login` | Sign in and write the token cache. Opens a browser by default; see `--auth-flow`. |
| `logout` | Clear the stored token cache. |
| `status` | Print the signed-in account and the resolved configuration. |
| `permissions` | Print the delegated scopes this configuration requests, grouped. |

Every option has an `MS365_MCP_*` environment variable fallback; the flag wins. Boolean environment variables accept `1`/`0`, `true`/`false`, `yes`/`no`, `on`/`off`. An unrecognised flag is a hard error rather than a silent no-op.

| Flag | Environment variable | Default | Effect |
| --- | --- | --- | --- |
| `--client-id <id>` | `MS365_MCP_CLIENT_ID` | the shipped application | Entra application (client) ID. Overrides the shipped default, and is required in a build that has none, in a sovereign cloud, or wherever your tenant will not accept the shared app. |
| `--tenant-id <id>` | `MS365_MCP_TENANT_ID` | `common` | Tenant ID, or `common` / `organizations` / `consumers`. Personal Microsoft accounts need `consumers`. |
| `--cloud <name>` | `MS365_MCP_CLOUD` | `global` | Sovereign cloud: `global`, `usgov`, `usgovdod`, `china`. Anything but `global` needs `--client-id`. |
| `--auth-flow <name>` | `MS365_MCP_AUTH_FLOW` | `auto` | Interactive sign-in flow: `auto`, `browser`, `device`. |
| `--auth-port <n>` | `MS365_MCP_AUTH_PORT` | `0` | Loopback port for the browser flow's redirect. `0` lets the OS pick. |
| `--groups <a,b,c>` | `MS365_MCP_GROUPS` | see presets | Comma-separated tool groups to enable. |
| `--preset <name>` | `MS365_MCP_PRESET` | `personal` | Group bundle: `personal`, `work`, `admin`, `all`. |
| `--read-only` | `MS365_MCP_READ_ONLY` | off | Hide every tool that writes. |
| `--org-mode` | `MS365_MCP_ORG_MODE` | off | Enable the groups whose scopes need tenant admin consent. |
| `--discovery` | `MS365_MCP_DISCOVERY` | off | Expose `discover_tools` and `call_tool` instead of every tool. |
| `--beta` | `MS365_MCP_ALLOW_BETA` | off | Allow requests against the Graph beta endpoint. |
| `--allow-generic-write` | `MS365_MCP_ALLOW_GENERIC_WRITE` | off | Let `graph_request` send POST, PATCH, PUT and DELETE. |
| `--max-output-chars <n>` | `MS365_MCP_MAX_OUTPUT_CHARS` | `60000` | Truncate serialised tool output beyond this many characters. |
| `--cache-dir <path>` | `MS365_MCP_CACHE_DIR` | OS config directory | Directory holding the token cache. |
| `--verbose`, `-v` | `MS365_MCP_VERBOSE` | off | Log Graph requests to stderr. |
| `--help`, `-h` | | | Print help and exit. |
| `--version` | | | Print the package version and exit. |

`MS365_MCP_AUTH_FLOW=interactive`, the undocumented value that used to select the browser flow, is still accepted as an alias for `browser`.

`--discovery` is worth knowing about: with every group enabled the server exposes 101 tools, and 101 tool schemas is a large, permanent cost on the first turn of every conversation. In discovery mode the client sees three tools instead — `discover_tools` searches the catalogue and returns names and descriptions without schemas, `call_tool` invokes one by name, and `graph_list_permissions` reports the configuration. The model pays for a tool's schema only when it actually calls it.

## Security

**Version 1.0 by default.** Requests go to the `v1.0` Graph endpoint. The `beta` endpoint is refused unless the server is started with `--beta`, because beta resources change without notice and without deprecation.

**Host pinning on the generic tool.** `graph_request` lets the model choose the request path, and every request this server sends carries your bearer token, so a path that escapes the Graph host leaks that token. Three independent checks stand between the two: the path is validated before use (absolute URLs, URI schemes, `..` segments, percent-encoded separators and protocol-relative references are all rejected), the URL is then built through the URL API rather than string concatenation, and the resulting host is compared against the pinned Graph hostname for the selected cloud before the request goes out. The same host check is applied to every `@odata.nextLink` before it is followed, because a paging link is chosen by the server response rather than by you.

**Read-only mode.** `--read-only` removes every write tool from the catalogue, so a client never sees them and cannot call them. `graph_request` stays available but refuses every method except GET, and `call_tool` re-checks the target in case a write tool reached it another way.

**Generic writes are opt-in separately.** Even without `--read-only`, `graph_request` refuses POST, PATCH, PUT and DELETE unless `--allow-generic-write` is passed. A purpose-built tool states in its description exactly what it changes; an arbitrary write to an arbitrary path does not.

**Least-privilege groups.** The scopes requested at sign-in are derived from the enabled groups and their tools, so enabling `--groups mail` asks for mail permissions and nothing else. `--read-only` also drops the write scopes from the consent request, rather than consenting to permissions the server will never exercise. The three admin-consent groups are excluded unless `--org-mode` is passed, so a personal-account user is never shown a consent prompt they cannot satisfy.

**A shared client ID is an identifier, not a credential.** The application ID the server ships with is public by construction — it appears in every authorize URL — and it grants nothing on its own. Access still requires an interactive sign-in by a user, consent recorded in that user's tenant, and a refresh token held on that user's machine. What it does mean for a tenant is a service principal visible to its administrators; see [Which application am I signing in to](#which-application-am-i-signing-in-to).

**Where the token cache lives.** The MSAL cache blob is encrypted with AES-256-GCM and written to `token-cache.enc` in your config directory, with the 32-byte key beside it in `cache.key`. Both files are created with mode `0600` inside a directory with mode `0700`.

| Platform | Default location |
| --- | --- |
| macOS, Linux | `$XDG_CONFIG_HOME/microsoft-graph-mcp`, or `~/.config/microsoft-graph-mcp` |
| Windows | `%APPDATA%\microsoft-graph-mcp` |

Override with `--cache-dir`. Run `npx ms-graph-mcp status` to see the resolved path, and `logout` to delete both files. Keeping the key next to the ciphertext protects against a stray backup or a file synced to the cloud, not against an attacker who already has read access to your home directory.

**No secrets in configuration.** The server is a public client and takes no client secret. The only credential it holds is the refresh token in the cache described above.

## What this is not

This is a local, delegated-permission MCP server. It acts as you, with the permissions you have, over a stdio transport on your own machine. It is not a replacement for Microsoft's own Agent 365 MCP servers, which are remote, run inside Microsoft's tenant boundary, use application permissions and admin-managed agent identities, and require a Microsoft 365 Copilot licence. If you need agent identity, tenant-wide governance, or auditing that a tenant administrator controls, that is the product to look at.

It is also not a general Microsoft Graph client library. The tools cover the common read and write paths for each product area and project their results down to what a model can usefully read; `graph_request` exists for everything else, deliberately behind a flag.

## Licence

MIT — see [LICENSE](LICENSE).
