# Opportunities (deferred work)

Things discovered mid-build that have no assigned task. Each entry gives a proposed severity, and none of them are acted on during this run.

- **[P3] `--dry-run mail send` prints "Email sent successfully."** Found in the Step 0 smoke run. The dry-run message misleads both agents and people. Proposal: print "(dry-run) not sent" instead.
- **[P3] Warm in-process MCP server** (alternative to D16). Spawning the CLI per call costs process start-up each time. Proposal: an embedded mode where `handleErrors` doesn't exit.
- **[P3] `drive copy --wait`** (D20). There's no copy command. Add `drive copy`, which returns a monitor URL, and poll it when `--wait` is given.
- **[P3] MCP notifications on delta changes** (§6 event-driven, "later").
- **[P3] Lazy-load command modules and MSAL** (P-7, measured during T05). `--version` takes 70 ms; the potential saving is about 30 ms. It needs stub command registration so `--help` still works.
- **[P2] KQL quoting gaps outside mail.** `contacts list --search`, `search people` and `groups list --search` (`groups.ts:28`) build `"…"` without `kqlQuote`, so a term containing `"` breaks the query. (Found by T19.)
- **[P3] Teams commands with no `--all`.** `teams members`, `channels`, `list` and `chat get` don't paginate. (T14)
- **[P2] More user.ts status lines on stdout.** `set-presence`, `set-timezone` and `set-language` print to stdout (C-19 pattern). The presence commands throw a plain `Error` for a missing client ID. `set-ooo --status`/`--audience` and `set-timezone` values aren't validated. The `upload-photo` `status >= 400` check never runs. (T17)
- **[P3] Planner usage errors exit 1 instead of 2.** `planner update-task --percent` validation and an unguarded `JSON.parse` in `update-plan-details --categories` both exit 1. The G-9 Premium hint also shows on a genuine 404. (T16)
- **[P2] Dry-run leaks the share password.** `GraphClient` dry-run prints request bodies to stdout, including `drive share --password` values. Redact known secret fields in dry-run output. (T15)
- **[P3] Mail folder names in move/triage.** `move --folder` and `triage --folder` don't resolve folder names. Attachment upload-session chunks have no retry or resume, unlike `uploadFile`. (T10)
- **[P2] Calendar time-zone defaults.** `calendar.ts` `--timezone` defaults to `'UTC'`, which T11 is fixing (M-9). Logged by T12; no action needed if T11 covers it.
- **[P3] free-busy --emails isn't validated**, and dry-run `cal update` prints the zone-read GET. (T11)
- **[P3] Setup wizard with piped stdin.** With non-TTY stdin, the second prompt never gets its input and the wizard exits 0 silently (this predates the run). `setup.ts` `prompt()` duplicates the readline code in `confirm()`. (T18)
- **[P3] `workflow digest --dry-run` crashes** with "Cannot read properties of undefined (reading 'dateTime')", because dry-run data has no `value` array. Found by the T20 snapshot.
- **[P2] `user set-presence`/`clear-presence` need a client ID even under `--dry-run`**, and throw a plain `Error` (exit 1, not 3). Found by the T20 snapshot; overlaps with T17's discovery.
- **[P3] `mail reply --html` is ignored**; the comment is always sent as plain text. (T35)
- **[P3] `auth accounts` marks "(active)" from the saved selection** even when `--account` is passed. (T40)
- **[P3] Catalog scope approximations.** `search query` and `watch` list the union of their resource types' scopes, and `user list` uses `User.Read.All`. Per-option scopes would be more precise. (T30)
- **[P3] `mail shared-send` has no `--idempotency-key`.** A crash mid-send leaves a pending claim that blocks the key for 24 h; the error explains this. (T38)
- **[P3] `validateEmail` throws a plain Error**, so calendar attendees and similar callers exit 1 instead of 2. (T43)
- **[P3] Completion gaps.** `--version` is offered below the top level, `--x=value` forms aren't completed, and the PowerShell completer covers only `ms365`. Column choices for chats, teams, contacts and planner are unverified against real data. (T44)

## From the follow-up specs run

- **[P2] `auth scopes`/`auth status` under `--dry-run` still build an AuthManager** and may refresh the token cache (a local side effect, DR-22 class). (core-1)
- **[P3] MCP dry-run results return only the stdout placeholder**; the stderr trace isn't included in the tool result. (core-1, for the B10 MCP track)
- **[P3] `drive upload` "Uploading…" is written with `console.error`**, so `--quiet` doesn't suppress it. (B4)
- **[P3] `digest` dry-run trace under-counts the per-list GETs.** (B7)
- **[P3] `user set-presence --duration` takes ISO 8601 (`PT30M`)** while the other duration options take `30m`. Unify on one form. (FIN)
- **[P3] The snapshot records only the first error line** to stay reviewable. A full-stderr variant could catch more regressions. (FIN)
- **[P3] `auth grant-scopes` under an admin profile** needs `MS365_CLI_CLIENT_ID` from env, because the profile's clientId is ignored. (IDT-4)
- **[P3] File-mode cache is still one `token-cache.json`, last writer wins** across processes; only keychain mode is split per account with read-modify-write. A merge in file mode would conflict with IDT-1's read-once tests. (IDT-2)
- **[P4] Tests write 0644 `token-cache.json` fixtures**, so the suite prints the ID-19 chmod notice to stderr. Write fixtures 0600. (IDT-2)
- **[P3] A crashed writer leaves `token-cache.json.lock`;** stale-lock breaking waits 30 s but the lock wait is 10 s, so a command right after a crash exits `lock_timeout`. Check the lock's PID for liveness before waiting out the age. (IDT-3)
- **[P4] `invalid_grant` message lacks the AADSTS detail** (MSAL's errorMessage only links to aka.ms); surface `subError`/`errorMessage` from the server response. Pre-existing. (IDT-5)
- **[P3] Admin-profile az hint shows az's own tenant GUID** instead of the profile's tenant (seen with `--tenant contoso.onmicrosoft.com`). (IDT-6)
- **[P4] `MS365_CLI_STORAGE` missing from the generated ms365-shared skill env list** (src/docs-gen.ts:275). (IDT-6)
- **[P4] `audit-trail changes` filters activities after paging,** so a page can be empty yet carry `nextCursor`. Push the activity filter server-side once the `or` list in `$filter` is verified live (AOT-7). (AOT-3)
- **[P3] `audit-trail apps` reads the tenant ID through a cast to GraphClient's private `auth`.** Add a public `tokenInfo()`/`tenantId()` to GraphClient. (AOT-5)
- **[P4] `audit-trail apps --with-consenters`** (AO-10 optional join with directoryAudits by principalId) not built; no AC requires it. (AOT-5)
- **[P4] Snapshot sampler gives `--app x`,** so the `audit-trail access` snapshot exits 2 before the POST. Teach the sampler GUID-shaped values for `--app`. (AOT-5)
- **[P4] `audit-trail import` stops at the first malformed row (exit 1).** Consider `--skip-bad-rows` with a count on stderr. (AOT-6)
- **[P4] `audit-trail changes` shows role targets as `?`.** The role's display name is in `modifiedProperties` (Role.DisplayName), not in `targetResources[].displayName`. (AOT-7 live)

## From the agent-host-integration run
- **[P3] `search_cap` is not reported across `--cursor` pages:** `complete` counts only the current page, so a mail `$search` hitting the 250 cap over several cursor pages ends `complete: true`. Carry the running count in the encoded cursor. (review #7)
- **[P3] A send applied from a plan cannot be reconciled:** `apply` records no `correlation` (internetMessageId) and rethrows raw Graph errors, so a lost `/send` in a plan has no `mail reconcile` path. Record the draft's internetMessageId and map a failure after the draft-create to `outcome_unknown`. (review #8)
- **[P3] Smoke cadence vs request-shape changes:** smoke ran per tier and missed an AHT-8 regression for five tasks (lesson 13). Cheap fix: a guard test that each smoke `grep` pattern for a request URL still appears in the dry-run snapshot. (dev-loop)
- **[P4] Cross-spec contract has no machine check:** Chief parses `describe --json`, `auth status -o json`, health records and error envelopes; a shared JSON-schema fixture (or a recorded `describe --json` golden that both repos test) would have caught the AHT-11 `scopes` change before the peer did (lesson 14). (contract)
- **[P4] `requestCommand` is a second description of each plan command's request shapes:** a catalog field per plan command (method + path pattern) would let the recorder and `apply` share one table. (E27)
- **[P4] `--follow` window-roll gap line repeats on a retried roll:** harmless for idempotent consumers, but a host counting gaps sees two. (review #2 fix)

## From the 2026-10-17 follow-up run (AHT-15, AHT-9 non-TTY)
- **[done] Cross-spec contract has no machine check** → AHT-15 `test/fixtures/contract/` + `npm run contract:check` (9a3773c).
- **[P3] Node 22 reproduction of the contract fixture is proven only in CI:** no Node 22 locally; the CI matrix step `node scripts/contract.mjs --check` is the evidence. Install a second Node locally (or a `nvm`-free tarball under /opt) so the gate can run both.
- **[P4] `apply` cannot say who approved:** a host-approved delete is logged like any apply; an optional `--approved-by <ref>` recorded in the audit line would let a host's audit trail and ms365's join without timestamps.
- **[P4] Fixture cases are hand-listed in `scripts/contract.mjs`:** a new error code or `watch` line kind needs a manual case; a test that every `code` thrown in `src/` has an `error-*` case would make coverage mechanical.
- **[P2] `requestCommand` (E27) matches `/events|calendar(s)` and `/messages|mailFolders` anywhere in a `/me/` path:** `DELETE /me/drive/root:/events/x` classifies as `calendar delete` (ceiling judged as Calendars.ReadWrite) and `DELETE /me/calendars/C1` as an event delete. The non-TTY host path no longer relies on it (E29 (3) checks the exact request), but the TTY prompt and the ceiling still do. Anchor the patterns at `^/me/(events|calendars?)` / `^/me/(messages|mailFolders)` and give each plan command its exact request shapes (see the P4 catalog item above). (implementation review 2026-10-17, finding 1)
