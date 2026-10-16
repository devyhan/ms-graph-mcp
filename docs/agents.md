# Instructions for AI agents

If you are an AI assistant working with this server, paste the block below into your context. It is the short form of [Consent](auth.md#consent), [Search permissions](auth.md#search-permissions) and [Sign-in errors](auth.md#sign-in-errors).

```
1. Call session_status once first: cached account (Entra ID, tenant, consented scopes) plus live GET /me reachability.
2. If no session is signed in, call auth_begin_login (opens the system browser, returns immediately), ask the user to finish there, then call auth_login_status. Never ask for credentials or an authorization code.
3. In discovery mode, orient via discover_tools and invoke what it names through call_tool.
4. On errors, read the Hint and Blocked-by lines: they name the fix. Re-running login with only the missing scopes prompts only for what was added.
5. For scope and consent questions see Consent, Search permissions and Sign-in errors in this README.
```

The server also offers an `agent_start` prompt with the same steps, and sends every client the same orientation as a one-paragraph server instruction on connect.
