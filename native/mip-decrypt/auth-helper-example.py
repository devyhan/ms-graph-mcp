#!/usr/bin/env python3
"""Minimal auth helper showing the mip-decrypt token contract.

Usage: auth-helper-example.py <client-id> <authority> <scope>

Prints an access token for <scope> on stdout (last non-empty line wins);
diagnostics and the sign-in prompt go to stderr. Tokens come from an MSAL
flow with a file cache beside this script, so the first run is interactive
and later runs are silent until the refresh token dies.

`MIP_AUTH_FLOW` selects the first-run flow: `device` (default) or `browser`.
Tenants with Conditional Access commonly block the device-code grant
(AADSTS53003 observed live); where that happens, use `browser`, which can
complete the same interactive challenges as the server's own sign-in.
`MIP_AUTH_LOGIN_HINT` optionally presets the browser account.

This is a starting shape, not production auth: bring your own cache
protection, broker/conditional-access story, and headless strategy.
Requires `pip install msal`. Never log the token this prints.
"""

import json
import os
import sys

CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".mip-token-cache.json")


def main(argv: list[str]) -> int:
    if len(argv) != 4:
        print("usage: auth-helper-example.py <client-id> <authority> <scope>", file=sys.stderr)
        return 2
    client_id, authority, scope = argv[1], argv[2], argv[3]

    try:
        import msal
    except ImportError:
        print("[auth] msal package missing: pip install msal", file=sys.stderr)
        return 1

    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        # Pre-existing cache may predate the 0600 write below; tighten it.
        os.chmod(CACHE_PATH, 0o600)
        with open(CACHE_PATH, encoding="utf-8") as handle:
            cache.deserialize(handle.read())

    app = msal.PublicClientApplication(
        client_id, authority=authority, token_cache=cache
    )
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent([scope], account=accounts[0])
    if result is None:
        if os.environ.get("MIP_AUTH_FLOW", "device").lower() == "browser":
            kwargs = {"scopes": [scope]}
            hint = os.environ.get("MIP_AUTH_LOGIN_HINT")
            if hint:
                kwargs["login_hint"] = hint
            result = app.acquire_token_interactive(**kwargs)
        else:
            flow = app.initiate_device_flow(scopes=[scope])
            if "user_code" not in flow:
                print(f"[auth] device flow failed: {json.dumps(flow)}", file=sys.stderr)
                return 1
            # The person completes this out of band; the sidecar just waits.
            print(f"[auth] sign in: {flow['verification_uri']} code {flow['user_code']}",
                  file=sys.stderr)
            result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        # 0600 from creation: os.open applies the mode atomically so no
        # umask instant leaves the refresh token world-readable. The mode
        # is ignored for a pre-existing file, hence the chmod after.
        fd = os.open(CACHE_PATH, os.O_CREAT | os.O_TRUNC | os.O_WRONLY, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(cache.serialize())
        os.chmod(CACHE_PATH, 0o600)

    if "access_token" not in result:
        print(f"[auth] token failed: {result.get('error')}: {result.get('error_description')}",
              file=sys.stderr)
        return 1
    # Token on stdout, and nothing else: the sidecar reads the last line.
    print(result["access_token"])
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
