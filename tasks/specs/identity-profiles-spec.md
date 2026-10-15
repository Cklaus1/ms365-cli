# DRAFT — User and admin profiles, and credential storage

**Status:** DRAFT — must pass its own Step 1 adversarial review before any build. Not scheduled.
**As of:** 2026-10-15 · **Base:** `build/chief-of-staff` at `9cfa22c` · **Prefix:** ID-
**Scope reduced 2026-10-15 (owner):** user and admin profiles only. The earlier, broader version is archived as `identity-profiles-spec.v0.1-deep-id.md`; its app-only identities, cross-user access, consent records and notice/opt-out items are not part of this spec.
**Resolves:** needs-human N6 (encryption at rest) from docs/ms365-findings-spec.md; builds on D21 (profiles), D26 (`.default`), D28 (keytar removed), lesson 6 (no hanging native code in the CLI process).

## Summary

ms365 already caches several accounts and has named profiles (`--profile`, `--account`), but credentials sit in one plaintext 0600 file and every profile is the same kind. This spec:
- **Credential storage.** Moves the refresh-token cache into the OS keychain where one works, with a silent fallback to the 0600 file, so sign-in stays one-time.
- **Two profile types.** `user` (default, today's behaviour) and `admin` (token borrowed from the Azure CLI on demand, never stored by ms365).

Every identity remains the signed-in person acting for themselves. No friction is added: after one sign-in, commands stay silent until Microsoft itself requires re-authentication.

## Current state (verified in code)

| Area | Today | File |
|---|---|---|
| Profiles | `{clientId?, tenantId?, scopes?, orgMode?, account?}` in `configDir()/profiles/<name>.json` | `src/profiles.ts:7-13` |
| Token storage | `FileTokenCachePlugin`: one `token-cache.json` (0600, atomic, locked) holding every cached account's refresh tokens in plaintext | `src/auth.ts:168` |
| Keychain | None. keytar was removed (D28) after it hung the process on WSLg. | — |
| Admin work | `auth grant-scopes` borrows a Graph token from `az account get-access-token`, used in memory only | `src/grant-scopes.ts` |

## Findings

| ID | Sev | Finding | Why it matters | Fix |
|---|---|---|---|---|
| ID-1 | P2 | Refresh tokens for every cached account sit in plaintext in one file. | Anyone or anything running as your OS user can copy `token-cache.json` and act as every account in it, until revocation or about 90 days of inactivity. Adding admin or app identities makes this worse. | Keychain-backed storage with a file fallback (ID-2 to ID-5). |
| ID-2 | P2 | There is no storage abstraction: the cache plugin reads and writes the file directly. | Keychain support, per-profile storage and tests all need one seam. | A `CredentialStore` interface with `get(key)`, `set(key, value)`, `delete(key)` and `available()`. Implementations: `file` (today's behaviour) and `keychain`. The cache plugin uses the store. |
| ID-3 | P2 | A keychain must never hang or block exit (lesson 6). | libsecret without a D-Bus session hangs inside native code, and a stuck native thread blocks `process.exit`. | No native module in the CLI process. The keychain backend shells out to each OS's own tool with a hard timeout and SIGKILL: <br>• **macOS:** `security find-generic-password` / `add-generic-password` <br>• **Windows:** PowerShell DPAPI (`ConvertFrom-SecureString`) storing a blob in the config directory, or `cmdkey` <br>• **Linux:** `secret-tool lookup` / `store`, only when `DBUS_SESSION_BUS_ADDRESS` is set |
| ID-4 | P2 | The storage choice must be silent and predictable. | Friction comes from prompts and surprise failures. | Per-profile setting `storage: auto | keychain | file`, default `auto`. `auto` probes the keychain once per process with a timeout of about 1.5 s; on failure it uses the file and prints one line the first time ("keychain unavailable; using the 0600 file"), recorded so it isn't repeated. Under `keychain`, a failure is an error (exit 3) instead of a fallback. |
| ID-5 | P3 | Moving an existing file cache into the keychain needs a migration. | Avoid a re-login when someone switches to keychain storage. | The first successful keychain write imports `token-cache.json`, then replaces the file with a marker naming the store. `ms365 auth storage [show|use keychain|use file]` switches explicitly and migrates in both directions. |
| ID-6 | P3 | Keychain item size limits. | One MSAL cache holding many accounts can exceed some backends' limits (Windows Credential Manager blobs are about 2.5 KB). | Store one item per account (key `ms365:<profile>:<homeAccountId>`) plus a small index item. Shared app metadata stays in the file, which then holds no refresh tokens. |
| ID-7 | P2 | Profiles have no identity type. | `user` and `admin` need different sign-in and storage rules. | Add `type: 'user' | 'admin'` (default `user`) to `Profile`, validated on load. |
| ID-8 | P2 | Admin identities shouldn't be stored at all. | An admin refresh token at rest is the highest-value thing on the machine. | `type: admin` profiles never write tokens. Each command borrows a token from `az account get-access-token --tenant <t>`, as `grant-scopes` does, held in memory only. Without az → exit 3 with `az login --tenant <t>`. |
| ID-12 | P3 | Fewer keychain reads for long-running processes. | Each command start pays for one keychain read (milliseconds; macOS may prompt once). | `mcp serve` and `watch --follow` read once and keep tokens in memory. The macOS keychain item is created with the CLI binary in its access list, so "Always Allow" sticks across upgrades when the binary is signed. |

## Open questions

| # | Question | Options | Recommendation | Tag |
|---|---|---|---|---|
| OQ-1 | Default storage on macOS and Windows | (a) `auto` (keychain, file fallback); (b) file stays default, keychain opt-in | **(a)** | engineering |
| OQ-2 | How the Windows backend stores data | (a) DPAPI-encrypted blob file via PowerShell; (b) Credential Manager via `cmdkey` (size limits) | **(a)** | engineering |

## Acceptance criteria

1. On macOS or Windows with `storage: auto`, after `ms365 auth login`, `token-cache.json` contains no refresh tokens, and 50 consecutive commands make no keychain prompts after the first "Always Allow".
2. On WSL or headless Linux with no D-Bus session, commands behave exactly as today (file storage), print the fallback notice once per config directory, and add no more than 1.5 s to the first command.
3. A keychain backend that hangs (simulated by a stub tool that sleeps) never delays exit beyond its timeout (spawned test).
4. `ms365 auth storage use keychain` then `use file` round-trips every cached account with no re-login (byte-identical MSAL cache).
5. A `type: admin` profile writes nothing to any store (the store's `set` is never called); without `az` it exits 3 with the `az login` hint.
6. Full suite, typecheck, `docs:check` and smoke stay green.

## Suggested task breakdown

| Task | Scope | Covers | Depends on |
|---|---|---|---|
| IDT-1 | `src/credential-store.ts` (interface + file backend); cache plugin uses it; no behaviour change | ID-2 | — |
| IDT-2 | Keychain backends (macOS `security`, Linux `secret-tool`, Windows DPAPI via PowerShell) behind a timeout + SIGKILL child runner; `auto` probing and the fallback notice | ID-1, ID-3, ID-4, ID-6 | IDT-1 |
| IDT-3 | Migration and the `auth storage` command | ID-5 | IDT-2 |
| IDT-4 | `Profile.type`; admin profiles borrow from az | ID-7, ID-8 | — |
| IDT-5 | Long-running processes read credentials once | ID-12 | IDT-2 |
| IDT-6 | Docs (README auth section, skills), snapshot regeneration, smoke checks for fallback and no-hang | AC 1–6 | all |

## Out of scope

- App-only (certificate) identities, acting on behalf of other users, consent records, and monitoring notices. See the archived v0.1 for the previous design; none of it is planned here.
