# Troubleshooting

| Symptom | Likely cause | Fix |
| --- | --- | --- |
| `login` stops before contacting Entra | No application ID: the shared registration is unpublished and no `--client-id` was given | [Register your own Entra application](auth.md#register-your-own-entra-application) and pass `--client-id` |
| Consent screen refuses, or `AADSTS90093` / `AADSTS900941` | The tenant's consent policy withholds scopes ordinary users cannot grant | Ask an administrator for one-time consent — see [Consent](auth.md#consent) |
| `AADSTS50194` | Single-tenant registration with the default `common` endpoint | Pass `--tenant-id <tenant-id>` — see [Sign-in errors](auth.md#sign-in-errors) |
| `Blocked by missing tenant admin consent` on a tool call | An admin-consent scope (`teams`, `directory`, `intune` groups) was never granted | Grant it once, then sign in again for just that scope — see [Search permissions](auth.md#search-permissions) |
| Browser opens but redirect is rejected (`AADSTS50011`) | Registration has no `http://localhost` redirect URI | Add the **Mobile and desktop applications** platform — step 4 of registration |
| Nothing opens, command sits there | No desktop session (SSH, container) | `login --auth-flow device` |
| A plugin setting does not take effect | Placeholder was never substituted, or another source won | Run `status` and read the `Client ID from:` line |
| `token unusable` / `ENOTFOUND` behind a proxy | Requests going direct instead of through the proxy | Fixed automatically: the server routes `fetch` through `HTTPS_PROXY`/`https_proxy`/`ALL_PROXY`/`all_proxy` at startup. If it persists, check the variable is set where the server runs (client `env`, not the shell profile), that `NO_PROXY` is not bypassing the host, and that stderr shows the `Proxying fetch requests via …` line — see [Proxy](configuration.md#proxy) |
| Tool output looks cut off | Output exceeded `--max-output-chars` (default 60000) | Follow the `nextLink` / cursor, or raise the flag |
