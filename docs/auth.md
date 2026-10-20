# Authentication and consent
## Which application am I signing in to

**A shared application published by this project.** Unless you override it, the client ID the server presents to Entra is a multi-tenant registration this project publishes — the same one every other user of this package signs in with. The server picks one in this order, and stops at the first that is set:

1. `--client-id <id>`
2. `MS365_MCP_CLIENT_ID`
3. the application shipped with the package
4. nothing, in which case sign-in refuses to start and prints the registration steps

`status` reports which of the four it landed on, and `login` names the authority and the application on stderr before it opens the browser, so you can see whose application you are about to consent to while you can still stop. Shipping an application ID is the model Lokka and the Softeria MS-365 MCP server use, and it exists for one reason: registering an application is the step most people never get past.

**First use in an organisation creates an enterprise application in that tenant.** The first time anyone in a tenant consents to a multi-tenant app, Entra creates a service principal for it in that directory — an entry under **Enterprise applications** carrying this project's application name. Administrators can see it there, review the delegated permissions it has been granted, and disable or block it. From then on every request the server makes is attributed to that application, alongside your user, in the tenant's sign-in and Graph activity logs. This is a consequence of signing in with a shared application, not a footnote: if your organisation would object to an unfamiliar third-party application appearing in its directory, register your own before the first sign-in.

**The application holds no secret and no data.** It is a public client: there is no client secret anywhere in this package, and there could not be one, because a secret shipped to every user is not a secret. Every token Entra issues is issued to the signed-in user, is scoped to what that user can already reach, and is cached encrypted on that user's own machine (see [Security](../README.md#security)). No server of ours sits between you and Microsoft, and neither the package nor the owner of the registration ever receives your tokens, mail, files or directory records. What the owner of a registration does control is the registration object itself, which is what the next two paragraphs are about.

**Using your own registration instead.** Pass `--client-id` (or set `MS365_MCP_CLIENT_ID`) and follow [Register your own Entra application](#register-your-own-entra-application). Reasons to do it:

- An enterprise policy that permits only applications registered in its own tenant, or that requires an internal owner and a review for every enterprise application.
- A sovereign cloud. The shared registration exists in the global cloud only, so the server refuses to offer it to `--cloud usgov`, `usgovdod` or `china` and says why, rather than letting you discover it after typing your password. Register an application in that cloud and name it with `--client-id`.
- A tenant that has blocked the shared application, or an administrator who will grant consent only to something they own.

**If the shared registration is ever deleted or blocked**, sign-in starts failing with `AADSTS700016`, `AADSTS7000112` or `AADSTS700054`, and `--client-id <your-id>` is the entire fix. Nothing else changes: same flows, same scopes, same tool catalogue, same cache location. Run `login` once after switching, because the cached refresh token belongs to the application that obtained it.

## Consent

**Shipping an application ID removes the registration step. It does not remove the consent step.** Entra still asks the signed-in user to approve every delegated scope the server requests, and the tenant's consent policy still decides which of those a user is allowed to approve without an administrator.

**Work and school accounts in a default tenant.** Under the Microsoft-managed default consent policy that new tenants get, an ordinary user cannot self-consent to `Mail.Read`, `Calendars.Read`, `Calendars.ReadWrite`, `Calendars.Read.Shared`, `Chat.Read`, `Tasks.Read`, `Tasks.ReadWrite`, `Files.Read.All`, `Sites.Read.All` or `MailboxSettings.Read`, among others — even though Microsoft's permissions reference marks every one of them as not requiring admin consent. That covers most of what the `personal` and `work` presets ask for. Sign-in fails at the consent screen, and the server prints which of your requested scopes fall under that policy, however small the request was. An administrator clears it once: **Enterprise applications > (the application) > Permissions > Grant admin consent** for the shared app, or **App registrations > (your app) > API permissions > Grant admin consent** for your own.

**Personal Microsoft accounts** consent for themselves — no administrator, no tenant policy. In exchange, a large part of the catalogue is unavailable to them at any consent level: the Teams chat and channel APIs are documented as not supported for delegated personal-account access, and Planner, SharePoint, the directory tools and Intune do not exist for a personal account at all. Use a narrow group set, and sign in against the personal-account authority with `--tenant-id consumers`; a refresh token issued through the `common` authority is rejected at its first refresh.

**Sign-in is incremental.** A bare `login` asks only for the minimal set (`User.Read`, `offline_access`), and consent grows as tools are first used: when a call fails for missing consent, its error names the scopes, and signing in again with just those prompts only for what was added. From a conversation, call `auth_begin_login` with `scopes` (or `mode: "full"` for everything grantable up front); from a terminal:

```
npx @devyhan/ms-graph-mcp login
npx @devyhan/ms-graph-mcp login --scopes Mail.Read,Calendars.Read
```

What was already granted stays granted. `login --full` restores the old ask-everything behaviour for scripts and admin pre-consent flows, and `--read-only` drops every write scope from a `--full` request. `npx @devyhan/ms-graph-mcp permissions --preset work` prints the exact scope list a configuration may need, formatted for pasting into an admin consent request.

## Search permissions

`search_query` is the cheapest way to find something, because it asks Microsoft's index instead of pulling whole mailboxes and chat rooms through your context. What it can reach depends on one scope per index, and a call asks only for the scopes the entity types in it actually need.

| `entityTypes` | Scopes that call requests | Tenant admin? |
| --- | --- | --- |
| `message` | `Mail.Read` | No |
| `event` | `Calendars.Read` | No |
| `driveItem` | `Files.Read`, `Sites.Read.All` | No |
| `listItem`, `site` | `Sites.Read.All` | No |
| `chatMessage` | `Chat.Read` **and** `ChannelMessage.Read.All` | **Yes** ¹ |

¹ Searching Teams chat is the one that needs an administrator. Graph refuses the `chatMessage` index without `ChannelMessage.Read.All` — it says so itself on a 403 — and that scope needs tenant admin consent. It belongs to the `teams` group, so a chat search needs `teams` enabled alongside `chat`, which means `--org-mode`. It is deliberately **not** in the `search` group's own scope list: `search` is in the personal preset, and putting an admin-consent scope there would put every personal install behind an administrator it does not otherwise need.

Until that consent exists, the chat index is unavailable and finding something in Teams means reading rooms with `chat_fetch_history` instead. On one real mailbox that difference was measured at about 110,000 characters of tool output versus 900 for the same question.

### Getting admin consent

`npx @devyhan/ms-graph-mcp permissions` prints, for your configuration, exactly which scopes need an administrator and a paragraph to send them. Scopes needing consent are marked `[admin]` per scope, not just per group, because a group can be user-consentable overall and still carry one scope that is not.

```
npx @devyhan/ms-graph-mcp permissions --groups search,chat,teams,me --org-mode
```

The tools say the same thing at the point of failure. A call blocked by a missing admin scope answers with **"Blocked by missing tenant admin consent — not by a missing sign-in"**, names the scope, and says which group to enable afterwards — rather than telling you to sign in again, which for these scopes changes nothing however many times you do it. `graph_list_permissions` returns the same detail as structured fields (`adminConsentScopes`, `howToGetAdminConsent`) so an assistant can relay it.

Once an administrator has granted the scope, sign in again for just it (with the owning group enabled so the tool exists):

```
npx @devyhan/ms-graph-mcp login --scopes ChannelMessage.Read.All --groups search,chat,teams,me --org-mode
```

### Rights Management consent

Reading a Purview-protected (AIP/RMS) document needs no extra Graph scope: `files_parse_document` downloads with the scopes it already has, and the unprotect step happens outside Graph entirely, inside the `mip-decrypt` sidecar. That is why no RMS scope appears in any group, in `permissions`, or in the consent prompt — `https://aadrm.com/user_impersonation` is not a Graph permission and this server never asks for it.

The sidecar authenticates on its own: on each RMS challenge it runs the helper named by `MIP_DECRYPT_AUTH_HELPER`, which mints the token in whatever flow it implements (the example defaults to MSAL device code, `MIP_AUTH_FLOW=browser` where Conditional Access blocks the device grant). Any consent that resource needs happens inside that helper's flow, not in this server's sign-in. What the helper's identity does need is usage rights on the document itself; without them the decrypt fails with a `NO_RIGHTS` diagnosis rather than corrupt output. See `native/mip-decrypt/README.md` for the sidecar's own configuration.

## Register your own Entra application

Once the shared registration ships, most users can skip this section. It remains the supported fallback for the reasons listed in [Which application am I signing in to](#which-application-am-i-signing-in-to), and until then it is the only way to sign in.

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com) and go to **Applications > App registrations > New registration**.
2. Give it a name. Under **Supported account types**, pick the audience you need. A personal Microsoft account (outlook.com, hotmail.com) requires "Accounts in any organizational directory and personal Microsoft accounts"; a work or school account used only in its own tenant needs nothing more than the single-tenant option. If you keep the single-tenant option, every command must name that tenant (`--tenant-id <tenant-id>`): the default `common` endpoint is refused for single-tenant applications (`AADSTS50194`).
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

**The session is a refresh token in an encrypted cache.** What sign-in leaves behind is the MSAL cache blob, encrypted with AES-256-GCM into `token-cache.enc` in your config directory (`~/.config/microsoft-graph-mcp` on macOS and Linux, `%APPDATA%\microsoft-graph-mcp` on Windows), with the 32-byte key beside it in `cache.key` at mode `0600`. See [Security](../README.md#security) for the details and for `--cache-dir`.

**Login also captures who signed in.** Every successful sign-in snapshots the identity from the MSAL account — username, display name, tenant, plus when and how it was captured — into `identity-cache.enc` beside the token cache, encrypted with the same 32-byte key. Status output reads that file from disk, so it can name the session without a Graph call: `session_status` returns it as `cachedIdentity` (the account and client IDs stay in the file), and tools that opt in print it as a one-line header above their output.

**The snapshot lives and dies with the session.** `logout` deletes it together with the token cache and the key. It is also dropped when no cached account remains, when a silent token refresh fails, and when its account no longer matches the live one because another process signed in since — in each case readers fall back to a live lookup rather than quoting a stale name. A missing, corrupt, or undecryptable file reads as absent, never as an error, and a snapshot that cannot be written warns on stderr without failing the sign-in in flight.

**`getToken` never prompts.** In `serve` mode the server refreshes access tokens silently and does nothing else. When the refresh token has expired, been revoked, or no longer covers a scope, the data tool returns an error rather than unexpectedly opening a browser and holding that request open. Interactive authentication starts only when the caller explicitly invokes `auth_begin_login`; that tool launches the browser in the background without writing to the MCP-owned stdout stream.

**Signing in from a tool call.** A sign-in the caller asked for is a different thing from one that starts itself mid-request, and `auth_begin_login` is that first thing. It explicitly starts the browser grant in the background, returns while the user completes the Microsoft page, and `auth_login_status` reports when the redirect has been exchanged for a session. `session_status` probes the resulting session. None of the three ever returns a token, all three survive `--read-only` — signing in changes nothing in the tenant — and all three stay visible under `--discovery`, because a caller with no session cannot invoke a tool it has not discovered. The tool and CLI paths ask for exactly the same scopes and leave the session in the same encrypted cache, so you can start in the conversation and carry on from the terminal or the other way round.

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
| `AADSTS50194` | The application is registered single-tenant but the server signs in through `common` (or `organizations`), which Entra refuses for such applications. | Sign in to the tenant that owns the registration (`--tenant-id <tenant-id>`), and restart the server; the cached session is reused, so retry before signing in again. |
| `AADSTS50020` | The account comes from a different identity provider than the application accepts: a personal Microsoft account at a work-only application, or the reverse. | `--tenant-id consumers` for a personal Microsoft account; otherwise sign in with a work or school account. The server names the `--tenant-id` the failed run used. |
| `AADSTS9002332` | The application accepts Microsoft Entra work or school users only, and you signed in with a personal Microsoft account. | Sign in with a work or school account, or pass `--client-id` for a registration whose audience includes personal accounts and add `--tenant-id consumers`. |
| `AADSTS7000112`, `AADSTS700054` | The application is disabled in that tenant — its enterprise application was disabled, or the service principal was blocked from sign-in. | Ask an administrator to re-enable it under **Enterprise applications > Properties > "Enabled for users to sign-in"**, or pass `--client-id` with your own registration. |
| `AADSTS65004` | The consent screen was declined. Nothing was granted and no session was created. | Ask for less: sign in bare for the minimal set and add scopes with `--scopes` as tools need them, or `--read-only` to drop the write scopes from a `--full` request. |
| `AADSTS90093` | The requested permissions need an administrator, and the signed-in user is not allowed to grant them. | Drop the admin-consent groups (do not pass `--org-mode`) and sign in for fewer scopes, or have an administrator grant consent once, as in [Consent](#consent). |

The server sorts these into three kinds, and only one of them falls back.

**Application rejections** — the first six rows — are reported as `Sign-in failed (<code>)`, followed by the likely cause, the flag that fixes it, and the first line of what Entra actually said. `auto` does not retry them on the device code path: an application that is absent, disabled, or wrong for your account type is exactly as absent on another device. The MSAL stack goes to stderr separately, prefixed `[auth]`.

**Consent decisions** — the last two rows, plus `AADSTS65001` and `AADSTS900941` — print `Sign-in was refused at the consent screen`, the number of scopes requested, and the subset of them the default tenant consent policy withholds from ordinary users. `auto` does not fall back here either, because the device code flow shows the same consent screen and would be refused the same way.

**Browser failures** are everything else. Under `auto` the server writes `Browser sign-in failed: ...` to stderr and continues on the device code path. Under `--auth-flow browser` it stops and prints the three usual causes: no desktop session, a loopback redirect that could not be served, and a registration with no matching redirect URI.
