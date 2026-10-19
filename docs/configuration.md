# Configuration
How to point an MCP client at the server, and every command and flag.
## As a Claude Code plugin

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
| Directory (tenant) ID | `common` | Suits most work and school accounts. Use `consumers` for a personal Microsoft account. A single-tenant app registration needs your tenant ID instead — `common` is refused for it (`AADSTS50194`). |
| Tool groups | the personal set | Profile, mail, calendar, files, To Do, contacts and search. Sign-in asks only for the minimal set plus the scopes you name. |
| Read-only | on | Hides every tool that writes and drops the write scopes from the consent request. Turn it off when you want the server to act, not just read. |

Read-only defaults to on here, and to off on the command line. The plugin is the
path someone takes without reading this file first, and an assistant that can send
mail and delete files on a fresh install is the wrong first state; on the command
line the choice is already explicit.

You can change any of them later from `/config`.

You still need your own Entra application — see
[Register your own Entra application](auth.md#register-your-own-entra-application) — because
this project publishes no shared one. Paste its **Application (client) ID** into the
first prompt.

Then sign in without leaving the conversation. Ask Claude to sign you in, or say
so when a tool reports there is no session: it calls `auth_begin_login`, which
opens the system browser for an OAuth 2.0 authorization-code + PKCE sign-in and
returns immediately. Finish in the browser, and Claude confirms the result with
`auth_login_status`.

This works safely from a tool call because the finished authorization code
returns to MSAL over a one-shot loopback HTTP listener; the MCP JSON-RPC stream
on stdout remains untouched. The device authorization grant remains available
from the CLI for SSH sessions, containers, and other headless environments.

The terminal route still exists and is better when you have a terminal open
anyway, because it opens the browser for you:

```
npx @devyhan/ms-graph-mcp login
```

Either way the session lands in the same encrypted cache, so signing in one way
and using the other is fine. See [How sign-in works](auth.md#how-sign-in-works).

Everything works before you configure anything: the server starts, lists its tools,
and `status` explains what is missing. An installed-but-unconfigured plugin shows a
working server rather than a broken one.

**If a setting does not take effect**, run `npx @devyhan/ms-graph-mcp status` and read
the `Client ID from:` line — it names which source won. Should a value arrive as the
literal text `${user_config.client_id}`, the placeholder was never substituted;
the server refuses it and says so rather than passing it to Entra, where it would
fail much later and blame something else.

## Without the plugin

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

## Claude Code

```
claude mcp add microsoft-graph -- npx -y @devyhan/ms-graph-mcp --preset work
```

Or add it to `.mcp.json` in the project root:

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

## Claude Desktop

Edit `claude_desktop_config.json` (Settings > Developer > Edit Config):

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

## VS Code

Add `.vscode/mcp.json` to the workspace:

```json
{
  "servers": {
    "microsoft-graph": {
      "type": "stdio",
      "command": "npx",
      "args": ["-y", "@devyhan/ms-graph-mcp", "--preset", "work"]
    }
  }
}
```

While the shared registration is unpublished, give each of these your own application ID as well: add `"env": { "MS365_MCP_CLIENT_ID": "<your-client-id>" }` beside `args`, or append `"--client-id", "<your-client-id>"` to `args`.

The server starts whether or not an application is configured. With none it writes a warning to stderr, lists its tools as normal, and fails every Graph call until you sign in.
## Commands and flags

```
ms-graph-mcp [command] [options]
```

| Command | Effect |
| --- | --- |
| `serve` | Run the MCP server on stdio. The default when no command is given. |
| `login` | Sign in and write the token cache. Asks only for the minimal scopes unless `--scopes` / `--full` widen it. Opens a browser by default; see `--auth-flow`. |
| `logout` | Clear the stored token cache. |
| `status` | Print the signed-in account and the resolved configuration. |
| `permissions` | Print the delegated scopes this configuration may request, grouped. |

Every option has an `MS365_MCP_*` environment variable fallback; the flag wins. Boolean environment variables accept `1`/`0`, `true`/`false`, `yes`/`no`, `on`/`off`. Boolean flags accept the same values after `=` (for example `--read-only=false` switches off `MS365_MCP_READ_ONLY=true`); a bare flag still means true. Group names in `--groups` / `MS365_MCP_GROUPS` are case-insensitive, like `--preset`. An unrecognised flag is a hard error rather than a silent no-op.

| Flag | Environment variable | Default | Effect |
| --- | --- | --- | --- |
| `--client-id <id>` | `MS365_MCP_CLIENT_ID` | the shipped application | Entra application (client) ID. Overrides the shipped default, and is required in a build that has none, in a sovereign cloud, or wherever your tenant will not accept the shared app. |
| `--tenant-id <id>` | `MS365_MCP_TENANT_ID` | `common` | Tenant ID, or `common` / `organizations` / `consumers`. Personal Microsoft accounts need `consumers`. |
| `--cloud <name>` | `MS365_MCP_CLOUD` | `global` | Sovereign cloud: `global`, `usgov`, `usgovdod`, `china`. Anything but `global` needs `--client-id`. |
| `--auth-flow <name>` | `MS365_MCP_AUTH_FLOW` | `auto` | Interactive sign-in flow: `auto`, `browser`, `device`. |
| `--auth-port <n>` | `MS365_MCP_AUTH_PORT` | `0` | Loopback port for the browser flow's redirect. `0` lets the OS pick. |
| `--scopes <a,b,c>` | | minimal set | Delegated scopes for `login` to request on top of the minimal set. |
| `--full` | | off | Ask `login` for every scope the configuration may need. |
| `--groups <a,b,c>` | `MS365_MCP_GROUPS` | see presets | Comma-separated tool groups to enable. Combines with `--preset`; their union is taken. |
| `--preset <name>` | `MS365_MCP_PRESET` | `personal` | Group bundle: `personal`, `work`, `admin`, `all`. Combines with `--groups`; their union is taken. |
| `--read-only` | `MS365_MCP_READ_ONLY` | off | Hide every tool that writes. |
| `--org-mode` | `MS365_MCP_ORG_MODE` | off | Enable the groups whose scopes need tenant admin consent. |
| `--discovery` | `MS365_MCP_DISCOVERY` | off | Expose `discover_tools` and `call_tool` instead of every tool. |
| `--beta` | `MS365_MCP_ALLOW_BETA` | off | Allow requests against the Graph beta endpoint. |
| `--allow-generic-write` | `MS365_MCP_ALLOW_GENERIC_WRITE` | off | Let `graph_request` send POST, PATCH, PUT and DELETE. |
| `--max-output-chars <n>` | `MS365_MCP_MAX_OUTPUT_CHARS` | `60000` | Truncate serialised tool output beyond this many characters. |
| `--output <name>` | `MS365_MCP_OUTPUT` | `json` | Tool result format: `json` (indented) or `toon` (compact tables: headers once plus bare rows). |
| `--cache-dir <path>` | `MS365_MCP_CACHE_DIR` | OS config directory | Directory holding the token cache. |
| `--verbose`, `-v` | `MS365_MCP_VERBOSE` | off | Log Graph requests to stderr. |
| `--help`, `-h` | | | Print help and exit. |
| `--version` | | | Print the package version and exit. |

`MS365_MCP_AUTH_FLOW=interactive`, the undocumented value that used to select the browser flow, is still accepted as an alias for `browser`.

`--discovery` is worth knowing about: with every group enabled the server exposes 111 tools, and 111 tool schemas is a large, permanent cost on the first turn of every conversation. In discovery mode the client sees six tools instead — `discover_tools` searches the catalogue and returns names and descriptions without schemas, `call_tool` invokes one by name, `graph_list_permissions` reports the configuration, `session_status` probes the session, and `auth_begin_login` with `auth_login_status` sign you in. The last three are exposed here deliberately: a caller with no session cannot reach a tool through `call_tool` that it has not discovered yet, and signing in is the first thing such a caller needs. The model pays for a tool's schema only when it actually calls it.

## Proxy

Node's `fetch` ignores proxy variables, so without help every Entra and Graph
call fails behind a proxy while curl and npm work. The server wraps the global
`fetch` once at startup, which routes sign-in (MSAL), Graph calls, and the
update check through the proxy — there is no flag to pass. Set the variables
in the environment that launches the server process: for a desktop-launched
MCP client that means the client's `env` map, not your shell profile.

The first usable value wins, in this order: `HTTPS_PROXY`, `https_proxy`,
`ALL_PROXY`, `all_proxy`. `HTTP_PROXY` is deliberately ignored — only the
HTTPS-capable variables route. A bare `host:port` gains the `http://` scheme
curl would assume, and credentials in the URL (`http://user:pass@host:port`)
are sent as `Proxy-Authorization`.

`NO_PROXY` (or lowercase `no_proxy`) is a comma-separated list of hostnames
that bypass the proxy. `*` bypasses everything; otherwise a rule matches its
exact host and every subdomain of it, case-insensitively, with an optional
leading dot ignored — `example.com` covers `api.example.com` but not
`example.com.evil.net`. Rules are hostnames only: ports and CIDR ranges are
not matched.

Anything outside the proxy's reach — no proxy variable set, a `NO_PROXY`
match, a non-HTTP(S) URL — goes direct through the native fetch, as do
redirect hops that land on a bypassed host. A malformed proxy value, or a
scheme that cannot serve a tunnel such as `socks5:`, warns on stderr and falls
back to a direct connection rather than failing startup; proxy setup can never
take the server down.

When proxying activates, one line on stderr names the proxy host and the
variable it came from (never the credentials):

```
[graph-mcp] INFO Proxying fetch requests via 127.0.0.1:8080 (from $HTTPS_PROXY).
```
