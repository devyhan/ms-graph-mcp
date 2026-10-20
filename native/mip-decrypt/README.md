# mip-decrypt sidecar

Native companion that unprotects one Purview-protected (AIP/RMS) file so
`files_parse_document` can parse it. The Node launcher
(`src/util/mip-decrypt.ts`) spawns this binary; nothing else in the server
touches it, and no test here executes it (the suite drives a fake sidecar).

## Why a sidecar

Rights-management crypto lives in Microsoft's closed-source MIP SDK, which
only ships as native libraries. This ~300-line C++ file is the thinnest
wrapper that turns "wrapper bytes in" into "plaintext bytes out" without
pulling the SDK's link closure into the Node process. The binary is
deliberately not shipped with the npm package (SDK licensing plus ~100 MB
of platform natives): build it once per platform and point the server at it.

## Protocol (caller contract)

- `mip-decrypt <input> <output>` — the only argv shape.
- Exit `0`: `<output>` holds plaintext (a byte copy when the input was not
  actually protected).
- Exit `2`: usage error — wrong argc or missing `MIP_DECRYPT_*` config.
- Exit `1`: decrypt failure. A `[mip] NO_RIGHTS: ...` stderr line marks a
  rights refusal; anything else is corruption, auth failure, or SDK breakage.
- stderr carries `[mip]` diagnostics; stdout is unused. Token material is
  never printed.

## Prerequisites

- A C++17 compiler (`g++` or `clang++`).
- The MIP SDK bundles: download from <https://aka.ms/mipsdkbins> and extract
  somewhere; that tree is `MIP_SDK_ROOT`. Setup background is at
  <https://learn.microsoft.com/en-us/information-protection/develop/setup-configure-mip>.
- An Entra app registration (client ID) the engine identifies as, and a user
  identity with usage rights on the documents. No IDs or secrets live in this
  directory — they arrive by environment at runtime.

## Build

```sh
MIP_SDK_ROOT=/path/to/mipsdk ./build.sh
# -> ./build/mip-decrypt
```

Point the server at the binary:

```sh
export MS365_MCP_MIP_DECRYPTOR=/path/to/ms-graph-mcp/native/mip-decrypt/build/mip-decrypt
```

## Runtime configuration

| Variable | Required | Default | Meaning |
| --- | --- | --- | --- |
| `MIP_DECRYPT_CLIENT_ID` | yes | — | Entra application (client) ID. |
| `MIP_DECRYPT_USERNAME` | yes | — | UPN the engine acts as. |
| `MIP_DECRYPT_AUTH_HELPER` | yes | — | Token helper executable (contract below). |
| `MIP_DECRYPT_AUTHORITY` | no | `https://login.microsoftonline.com/common` | Entra authority. |
| `MIP_DECRYPT_STORAGE` | no | per-user temp subdir | Writable MIP storage dir. |

MIP storage (`MIP_DECRYPT_STORAGE`, default `$TMPDIR/mip-decrypt`) persists
engine state, cached templates, and policy between runs — no document
plaintext, but identity-adjacent state worth clearing on teardown or
identity rotation: `rm -rf "${MIP_DECRYPT_STORAGE:-$TMPDIR/mip-decrypt}"`.

## Auth-helper token contract

The sidecar never does interactive auth itself: on each RMS challenge it runs

```sh
<MIP_DECRYPT_AUTH_HELPER> '<client-id>' '<authority>' '<scope>'
```

- The helper prints the access token for that scope on **stdout** (last
  non-empty line wins) and exits zero. Diagnostics go to **stderr**, which
  the sidecar leaves on the terminal — never merged into the token.
- A non-zero exit, empty output, or a non-JWT-shaped line fails the decrypt
  loudly instead of sending garbage to RMS (which would answer 401).
- Scope selection is the sidecar's job: `https://aadrm.com` challenges map to
  `https://aadrm.com/user_impersonation`; anything else maps to
  `<resource>/.default` when the challenge scope is not already a v2 scope.
  (MIP sends the tenant id as "scope" for RMS, which MSAL would misread as a
  Graph scope — hence the mapping.)
- `auth-helper-example.py` is a minimal MSAL helper showing the shape
  (device code by default, `MIP_AUTH_FLOW=browser` where Conditional Access
  blocks the device grant with AADSTS53003); bring your own cache/refresh
  strategy for production. Never log the token it returns. The helper keeps
  its own token cache — the server's MSAL cache is not shared with it.

## Platform notes

- **macOS (arm64/x64)**: supported; use the SDK's `macos` binaries. First run
  may prompt to allow the downloaded libraries; clear quarantine with
  `xattr -dr com.apple.quarantine`.
- **Linux (x64)**: supported; use the SDK's `generic` binaries. Needs
  `libgsf`, `libsecret`, OpenSSL and friends per the MIP Linux prerequisites.
- **Windows**: not scripted — `build.sh` refuses it. The source is portable
  C++17/MIP API; build it with MSVC against the SDK's `win32`/`win64`
  binaries and point `MS365_MCP_MIP_DECRYPTOR` at the `.exe`.

## Troubleshooting

- `exit 2, missing configuration`: one of the three required env vars is
  unset in the sidecar's environment (note: launchd/systemd services do not
  inherit your shell's exports).
- `NO_RIGHTS`: the helper's identity lacks usage rights. Sign the helper in
  as someone who has them, or ask the owner for access.
- `auth helper exited with status N`: run the helper by hand with the same
  arguments. The usual cause is a broken helper environment — e.g. `msal`
  missing for the python its shebang resolves to — whose traceback the
  sidecar refuses to mistake for a token.
- Link errors about `mip_*`: `MIP_SDK_ROOT` points at the wrong tree or the
  wrong platform's `bin/` dir.
- A decrypt that hangs past 120s is killed by the launcher; a helper waiting
  on device-code input is the usual cause — pre-authenticate it first.
