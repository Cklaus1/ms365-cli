# DRAFT — Multi-identity profiles and credential storage

**Status:** DRAFT — must pass its own Step 1 adversarial review before any build. Not scheduled.
**As of:** 2026-10-15 · **Base:** `main` at `c8c9b53` (plus the uncommitted `auth grant-scopes` work) · **Prefix:** ID-
**Resolves:** needs-human N6 (encryption at rest) from docs/ms365-findings-spec.md; builds on D21 (profiles), D26 (`.default`), D28 (keytar removed), lesson 6 (no hanging native code in the CLI process).

## Summary

ms365 already caches several accounts and has named profiles (`--profile`, `--account`), but every identity is a person signing in for themselves, and credentials are stored in one plaintext 0600 file. This spec does three things:
- **Identity types.** Each profile gets an explicit type: `user` (default), `admin` (borrowed from the Azure CLI, never stored) or `app` (certificate, no person).
- **Credential storage.** The refresh-token cache moves into the OS keychain where one works, with a silent fallback to the 0600 file, so sign-in stays one-time.
- **Consent records for shared access.** When a person opts in to let another tool or person act for them, the grant is recorded, listable and revocable by that person.

- **Notice with opt-out (decided).** The company tells employees, in Teams, exactly what is monitored (metadata and content), and each person can decline with **one** choice that covers everything. The decline list is admin-only. Declines are honored technically, not just recorded. Opt-in grants (above) stay available for anything outside the notice.

The goal is no extra friction: after one sign-in, commands stay silent until Microsoft itself requires re-authentication (about 90 days of inactivity, or tenant Conditional Access policy).

## Current state (verified in code)

| Area | Today | File |
|---|---|---|
| Profiles | `{clientId?, tenantId?, scopes?, orgMode?, account?}` in `configDir()/profiles/<name>.json` | `src/profiles.ts:7-13` |
| Token storage | `FileTokenCachePlugin`: one `token-cache.json` (0600, atomic, locked) holding every cached account's refresh tokens in plaintext | `src/auth.ts:168` |
| Keychain | None. keytar was removed (D28) after it hung the process on WSLg. | — |
| Admin work | `auth grant-scopes` borrows a Graph token from `az account get-access-token`, used in memory only | `src/grant-scopes.ts` |
| App-only sign-in | Not implemented (N2). No confidential client, no certificate support. | — |
| Commands | About 150 `/me` endpoint references across command files; only `resolve` takes `--user` | `src/commands/*.ts` |

## Findings

| ID | Sev | Finding | Why it matters | Fix |
|---|---|---|---|---|
| ID-1 | P2 | Refresh tokens for every cached account sit in plaintext in one file. | Anyone or anything running as your OS user can copy `token-cache.json` and act as every account in it, until revocation or about 90 days of inactivity. Adding admin or app identities makes this worse. | Keychain-backed storage with a file fallback (ID-2 to ID-5). |
| ID-2 | P2 | There is no storage abstraction: the cache plugin reads and writes the file directly. | Keychain support, per-profile storage and tests all need one seam. | A `CredentialStore` interface with `get(key)`, `set(key, value)`, `delete(key)` and `available()`. Implementations: `file` (today's behaviour) and `keychain`. The cache plugin uses the store. |
| ID-3 | P2 | A keychain must never hang or block exit (lesson 6). | libsecret without a D-Bus session hangs inside native code, and a stuck native thread blocks `process.exit`. | No native module in the CLI process. The keychain backend shells out to each OS's own tool with a hard timeout and SIGKILL: <br>• **macOS:** `security find-generic-password` / `add-generic-password` <br>• **Windows:** PowerShell DPAPI (`ConvertFrom-SecureString`) storing a blob in the config directory, or `cmdkey` <br>• **Linux:** `secret-tool lookup` / `store`, only when `DBUS_SESSION_BUS_ADDRESS` is set |
| ID-4 | P2 | The storage choice must be silent and predictable. | Friction comes from prompts and surprise failures. | Per-profile setting `storage: auto | keychain | file`, default `auto`. `auto` probes the keychain once per process with a timeout of about 1.5 s; on failure it uses the file and prints one line the first time ("keychain unavailable; using the 0600 file"), recorded so it isn't repeated. Under `keychain`, a failure is an error (exit 3) instead of a fallback. |
| ID-5 | P3 | Moving an existing file cache into the keychain needs a migration. | Avoid a re-login when someone switches to keychain storage. | The first successful keychain write imports `token-cache.json`, then replaces the file with a marker naming the store. `ms365 auth storage [show|use keychain|use file]` switches explicitly and migrates in both directions. |
| ID-6 | P3 | Keychain item size limits. | One MSAL cache holding many accounts can exceed some backends' limits (Windows Credential Manager blobs are about 2.5 KB). | Store one item per account (key `ms365:<profile>:<homeAccountId>`) plus a small index item. Shared app metadata stays in the file, which then holds no refresh tokens. |
| ID-7 | P2 | Profiles have no identity type. | `user`, `admin` and `app` need different sign-in, storage and safety rules. | Add `type: 'user' | 'admin' | 'app'` (default `user`) to `Profile`, validated on load. |
| ID-8 | P2 | Admin identities shouldn't be stored at all. | An admin refresh token at rest is the highest-value thing on the machine. | `type: admin` profiles never write tokens. Each command borrows a token from `az account get-access-token --tenant <t>`, as `grant-scopes` does, held in memory only. Without az → exit 3 with `az login --tenant <t>`. |
| ID-9 | P2 | App-only sign-in is missing (N2). | Unattended automation (CI jobs, a server running ms365 for itself) can't use device-code sign-in. | `type: app` with `clientId`, `tenantId` and `certificate: keychain:<name> | file:<path.pem> | env:<VAR>`, using MSAL `ConfidentialClientApplication` with certificate credentials (never client secrets). Commands that need `/me` fail clearly under an app identity (exit 2, "this command needs a user identity") until they support a target user; see ID-11. |
| ID-10 | P2 | App identities need guardrails. | A certificate with tenant-wide application permissions is powerful. | A `type: app` profile must declare the permissions it expects. `ms365 auth whoami` shows the app's granted application roles, and ms365 warns (to stderr) when the token holds roles beyond the declared set. The docs require scoping such apps with Exchange RBAC for Applications or Teams application access policies to only the mailboxes or users that have opted in. ms365 doesn't add bulk or org-wide collection commands. |
| ID-11 | P3 | Commands hard-wire `/me`. | A non-`/me` target is needed for app identities, and for users who opted in to delegate access (ID-13). | A shared `userPath(opts)` helper returns `/me` or `/users/{id}`, plus a global `--as-user <upn>` that's accepted only when the identity can act for that user. Move the about 150 call sites group by group (mail, calendar, todo first), with the dry-run snapshot as the safety net. |
| ID-12 | P3 | Fewer keychain reads for long-running processes. | Each command start pays for one keychain read (milliseconds; macOS may prompt once). | `mcp serve` and `watch --follow` read once and keep tokens in memory. The macOS keychain item is created with the CLI binary in its access list, so "Always Allow" sticks across upgrades when the binary is signed. |
| ID-13 | P2 | Opt-in delegation has no consent record. | When a person lets a tool or another person act for them, they must be able to see and withdraw that. | `ms365 consent grant --to <profile|tool> --scopes <list> [--expires <t>]` is run by the person themselves, under their own sign-in. It records `{grantor, grantee, scopes, createdAt, expiresAt}` locally and in the grantor's own OneDrive app folder (`/me/drive/special/approot/consents.json`), so the grantor always holds the authoritative copy. This goes through the Microsoft Graph API, like the existing `drive` commands. No OneDrive desktop or sync client is needed, so it works on Linux (verified 2026-10-15: `ms365 drive quota` succeeds from WSL). `ms365 consent list` shows grants; `ms365 consent revoke <id>` removes the grant and the grantee's cached tokens for that person. Commands acting for someone else (`--as-user`) check for an unexpired grant covering the command's catalog scopes; otherwise exit 4 `consent_missing`. |
| ID-14 | P3 | Shared access needs an audit trail. | Accountability for acting on someone else's behalf. | Every command run with `--as-user` writes an audit line naming the actor, the target user and the consent ID, reusing `src/audit.ts`. |
| ID-15 | P2 | There is no way to tell employees what is monitored, or to take a "decline" back. | Transparent notice is the condition that makes an opt-out model defensible, and it must be verifiable afterwards. | `ms365 notice send` delivers the notice and tracks the outcome (ID-16 to ID-20). |
| ID-16 | P2 | The notice text and the enforced scope can drift apart. | If the notice says "metadata" but the tool reads content (or the reverse), employees were not actually informed. | One config file, `configDir()/notice.json`, is the **single source of truth**: `{ version, effectiveDate, owner, contact, retentionDays, workloads: { mail: 'none'|'metadata'|'content', teams: …, calendar: …, files: …, meetings: … }, purpose, declineUrl }`. The notice text is **generated** from it, and every command that reads another person's data (`--as-user`, IDT-6/7) checks its target workload against the same file. A workload not listed is `none`. Participation is **opt-out for every workload and level**, and the notice offers **one** Decline: it excludes the person from all of them (no per-workload or per-level choices). Changing the config bumps `version` and re-notifies (ID-19). |
| ID-17 | P2 | Delivery to many people must be reliable and traceable. | Nobody should get two notices, and the company must be able to show who was notified, when and how. | `notice send --to <upn,… | --group <id>>` creates a 1:1 Teams chat per person (existing `chat create` + `chat send`) with `--idempotency-key notice:<version>:<upn>`, paced to Teams throttling limits (resumable on 429), and records `{upn, version, sentAt, chatId, messageId}` in the notice ledger (`configDir()/notice-ledger.json`, 0600, plus an audit line per send). A failed or unknown-outcome delivery is reported and never silently retried (R-7 rules). `notice status` lists delivered / pending / failed. |
| ID-18 | P2 | A decline needs a mechanism, and the mechanism must be easy. | A hard-to-find decline makes the opt-out meaningless. | The notice carries two paths: **(a)** a link to a Microsoft Form (or Power Automate adaptive card with Continue / Decline buttons) that appends to a SharePoint list; **(b)** `ms365 notice decline` run by the employee themselves. Both end up in one SharePoint list (`Monitoring Opt-outs`), readable with `ms365 sp list-items`. A decline takes effect on the next run and can be reversed by the employee with `ms365 notice resume`. |
| ID-19 | P2 | Declines must actually exclude the person. | Recording a decline but still reading their data would be worse than having no notice. | Before any command reads another person's data, the tool loads the decline list (cached 5 minutes) and refuses with exit 4 `participant_declined` for anyone on it, whatever the workload. If the list can't be read, the tool **fails closed** (exit 4 `opt_out_list_unavailable`), never open. Because the list is admin-only (ID-22), only an admin identity (the `type: admin` profile from ID-8, or an app identity with the list permission) can read it, and therefore only those identities can read other people's data at all. The decline list is used only for exclusion: it is not an input to any report, score or `--fields`. Changing the notice config (ID-16) re-opens the question: previous declines stay in force, and everyone else is re-notified. |
| ID-20 | P3 | Notices need follow-up and a way to ask questions. | People miss messages; some will want to discuss it. | `notice remind` re-sends to pending recipients only, once per N days (default 7, max 3 reminders). The notice always names a human contact and where the full policy lives. No one is marked "agreed" by silence beyond the effective date: the ledger records `notified` and `no_response` separately, and reporting uses those words. |
| ID-21 | P3 | A notice sent from one person's account is easy to dismiss as spam or phishing. | Low read rates would undermine the whole model. | The text is plain, short and specific; it links only to the company's own SharePoint or Forms domain; and `notice send --preview <upn>` lets the owner send it to themselves first. A `--dry-run` shows every recipient and the exact text (no messages sent). |
| ID-22 | P2 | The decline list must be readable by admins only, yet employees must be able to add their own decline and withdraw it. | An open list would show who declined; a closed list would stop people declining. | The `Monitoring Opt-outs` SharePoint list is created with **item-level permissions**: *Create and Edit access: create items and edit items that were created by the user*; *Read access: read items that were created by the user*. Employees therefore see and edit only their own entry; the site/list owners group (tenant admins) sees all. `notice init-list` creates the list with exactly these settings (and refuses to continue if it cannot verify them via the list's `readSecurity` / `writeSecurity` properties). A test asserts a non-admin token cannot enumerate the list. The decline list is also excluded from `describe`, `--fields` and MCP resources. |

## Open questions

| # | Question | Options | Recommendation | Tag |
|---|---|---|---|---|
| OQ-1 | Default storage on macOS and Windows | (a) `auto` (keychain, file fallback); (b) the file stays the default and keychain is opt-in | **(a)** | engineering |
| OQ-2 | How the Windows backend stores data | (a) a DPAPI-encrypted blob file through PowerShell; (b) Credential Manager via `cmdkey` (size limits) | **(a)**: no size limits, and still tied to the user's Windows login | engineering |
| OQ-3 | Where the consent record lives (ID-13) | (a) local only; (b) local plus the grantor's OneDrive app folder | **(b)**: the grantor keeps control even if the grantee's machine disappears. It's API-only, so it works on Linux without a OneDrive client. Option (a) stays available for tenants without OneDrive. | needs-human |
| OQ-4 | Whether app identities (ID-9/ID-10) ship at all | (a) yes, with the guardrails; (b) defer | Decide based on a concrete automation need; this spec adds no bulk-collection commands either way | needs-human |
| OQ-5 | Certificate source for app profiles | keychain / file / env var / Key Vault reference | Keychain and env var first; a Key Vault reference via `az keyvault secret show` later | engineering |
| OQ-6 | Should `--as-user` work with delegated tokens (Outlook shared mailbox and delegate permissions)? | (a) yes, when Exchange delegate access exists; (b) app identities only | **(a)**: it matches how Outlook delegation already works, and still needs a consent record | engineering |
| OQ-7 | Opt-out vs opt-in as the default (ID-15, ID-18, ID-19) | (a) opt-out for everything; (b) opt-in; (c) per-workload mix | **RESOLVED by the owner (2026-10-15): (a)** — opt-out for both metadata and content, and **one single decline covers all workloads and levels**. There is no separate per-workload or per-level choice. | decided |
| OQ-8 | Who may read the decline list | (a) the notice owner and tenant admins only; (b) also managers | **RESOLVED by the owner (2026-10-15): (a)** — admins only, used only for enforcement, never in a report. Enforced by SharePoint permissions (ID-22). | decided |
| OQ-9 | How the Continue/Decline buttons are built | (a) Microsoft Form + SharePoint list (no code); (b) Power Automate adaptive card; (c) a Teams bot | **(a)** to start. It works with no extra infrastructure, and ms365 only needs to read the list. | engineering |

## Acceptance criteria

1. On macOS or Windows with `storage: auto`, after `ms365 auth login`, `token-cache.json` contains no refresh tokens, and 50 consecutive commands make no keychain prompts after the first "Always Allow".
2. On WSL or headless Linux with no D-Bus session, commands behave exactly as today (file storage), print the fallback notice once per config directory, and never take more than 1.5 s extra on the first command.
3. A keychain backend that hangs (simulated by a stub tool that sleeps) never delays exit beyond its timeout, and never hangs the process (spawned test).
4. `ms365 auth storage use keychain` then `use file` round-trips every cached account with no re-login (byte-identical MSAL cache after the round trip).
5. A `type: admin` profile writes nothing to any store; a test asserts the store's `set` is never called.
6. A `type: app` profile signs in with a PEM certificate (mocked token endpoint). `/me` commands fail with exit 2 and a clear message. `auth whoami` lists roles and warns on roles that weren't declared.
7. `--as-user bob@…` without an unexpired consent grant from bob exits 4 `consent_missing` and makes no Graph call. With a grant it works and writes an audit line naming the consent ID.
8. `ms365 consent revoke <id>` run by the grantor removes the grant; the grantee's next `--as-user` call exits 4.
9. Full suite, typecheck, `docs:check` and smoke stay green; the dry-run snapshot shows no unexplained changes from the `userPath` migration.
10. `notice send --dry-run` lists every recipient and prints the generated text, sends nothing, and writes no ledger entry. The text names exactly the workloads and levels in `notice.json` (a test asserts that changing a level changes the text, and an unlisted workload is described as not monitored).
11. Sending the same notice version twice to the same person delivers once (idempotency key), and a 429 mid-run resumes without duplicates (mocked).
12. A person on the decline list makes any `--as-user` call exit 4 `participant_declined` with no Graph call to their data, for every workload and level (one decline covers all). An unreadable decline list exits 4 `opt_out_list_unavailable` (fail closed). `notice resume` by that person removes the exclusion.
13. No command or report output includes the decline list (guard test: grep plus a runtime check on `--fields` and `describe`).
14. Bumping the notice `version` re-notifies everyone not declined and keeps existing declines.
15. `notice init-list` sets item-level permissions so a non-admin sees only their own entry; a test with a non-admin token cannot list others' entries, and a non-admin identity calling `--as-user` fails closed (exit 4 `opt_out_list_unavailable`).

## Suggested task breakdown

| Task | Scope | Covers | Depends on |
|---|---|---|---|
| IDT-1 | `src/credential-store.ts` (interface plus file backend); cache plugin uses it; no behaviour change | ID-2 | — |
| IDT-2 | Keychain backends (macOS `security`, Linux `secret-tool`, Windows DPAPI via PowerShell) behind a timeout and SIGKILL child runner; `auto` probing and the fallback notice | ID-3, ID-4, ID-6 | IDT-1 |
| IDT-3 | Migration and the `auth storage` command | ID-5 | IDT-2 |
| IDT-4 | `Profile.type`; admin profiles borrow from az | ID-7, ID-8 | — |
| IDT-5 | App profiles: ConfidentialClientApplication with a certificate, `auth whoami`, declared-roles check | ID-9, ID-10 | IDT-4, OQ-4 |
| IDT-6 | `userPath()` + `--as-user`, migrated group by group (mail → calendar → todo → the rest) | ID-11 | IDT-4 |
| IDT-7 | Consent records: `consent grant/list/revoke`, enforcement in the preAction hook, audit lines | ID-13, ID-14 | IDT-6, OQ-3 |
| IDT-8 | Long-running processes read credentials once | ID-12 | IDT-2 |
| IDT-9 | Docs (README auth section, skills), snapshot regeneration, smoke checks for fallback and no-hang | AC 1–9 | all |
| IDT-10 | `notice.json` schema, generated notice text, `notice preview` | ID-16, ID-21 | — |
| IDT-11 | `notice send` / `status` / `remind`: Teams delivery with idempotency, pacing, ledger, audit lines | ID-15, ID-17, ID-20 | IDT-10 |
| IDT-12 | Decline list reader (SharePoint list), cache, fail-closed check in the preAction hook; `notice decline` / `resume` | ID-18, ID-19, ID-22 | IDT-6, IDT-10 |
| IDT-13 | Guard tests (decline list never in output), docs, snapshot, smoke | AC 10–15 | IDT-10 – IDT-12 |

## Out of scope

- Commands that collect data across many users at once (bulk mailbox or Teams export). `notice` and the decline list deliver and honor choices; they do not add any way to read other people's content. Any such feature would need its own spec, built on this model, and its scope must match `notice.json`.
- Changing how Microsoft decides re-authentication (Conditional Access sign-in frequency, MFA prompts). ms365 only stops adding friction of its own.
