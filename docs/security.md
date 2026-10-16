# Security

**Version 1.0 by default.** Requests go to the `v1.0` Graph endpoint. The `beta` endpoint is refused unless the server is started with `--beta`, because beta resources change without notice and without deprecation.

**Host pinning on the generic tool.** `graph_request` lets the model choose the request path, and every request this server sends carries your bearer token, so a path that escapes the Graph host leaks that token. Three independent checks stand between the two: the path is validated before use (absolute URLs, URI schemes, `..` segments, percent-encoded separators and protocol-relative references are all rejected), the URL is then built through the URL API rather than string concatenation, and the resulting host is compared against the pinned Graph hostname for the selected cloud before the request goes out. The same host check is applied to every `@odata.nextLink` before it is followed, because a paging link is chosen by the server response rather than by you.

**Read-only mode.** `--read-only` removes every write tool from the catalogue, so a client never sees them and cannot call them. `graph_request` stays available but refuses every method except GET, and `call_tool` re-checks the target in case a write tool reached it another way.

**Generic writes are opt-in separately.** Even without `--read-only`, `graph_request` refuses POST, PATCH, PUT and DELETE unless `--allow-generic-write` is passed. A purpose-built tool states in its description exactly what it changes; an arbitrary write to an arbitrary path does not.

**Least-privilege groups.** Each tool call authenticates with just the scopes its own index needs, and sign-in is incremental — a bare login asks only for the minimal set, so enabling `--groups mail` never prompts for permissions outside mail. `--read-only` also drops the write tools from the catalogue, rather than exposing operations the server will never exercise. The three admin-consent groups are excluded unless `--org-mode` is passed, so a personal-account user is never shown a consent prompt they cannot satisfy.

**A shared client ID is an identifier, not a credential.** The application ID the server ships with is public by construction — it appears in every authorize URL — and it grants nothing on its own. Access still requires an interactive sign-in by a user, consent recorded in that user's tenant, and a refresh token held on that user's machine. What it does mean for a tenant is a service principal visible to its administrators; see [Which application am I signing in to](auth.md#which-application-am-i-signing-in-to).

**Where the token cache lives.** The MSAL cache blob is encrypted with AES-256-GCM and written to `token-cache.enc` in your config directory, with the 32-byte key beside it in `cache.key`. Both files are created with mode `0600` inside a directory with mode `0700`.

| Platform | Default location |
| --- | --- |
| macOS, Linux | `$XDG_CONFIG_HOME/microsoft-graph-mcp`, or `~/.config/microsoft-graph-mcp` |
| Windows | `%APPDATA%\microsoft-graph-mcp` |

Override with `--cache-dir`. Run `npx @devyhan/ms-graph-mcp status` to see the resolved path, and `logout` to delete both files. Keeping the key next to the ciphertext protects against a stray backup or a file synced to the cloud, not against an attacker who already has read access to your home directory.

**No secrets in configuration.** The server is a public client and takes no client secret. The only credential it holds is the refresh token in the cache described above.
