# DRAFT — Test depth and verification gaps

**Status:** DRAFT — must pass its own Step 1 adversarial review before any build. Not scheduled.

**As of:** 2026-10-14 · **Scope:** `test/**`, `scripts/smoke.sh`, `.github/workflows/ci.yml`, `src/mcp-server.ts` at `ed5bcff` (branch `build/followup-specs`) · **Format/decisions:** `docs/ms365-findings-spec.md` (D1–D30; D24 "build to documented shape, mark †" and D27–D30 "no legacy consumers, no shims" apply). **Severity:** as in the findings spec (P0–P3). Theme priority: P3.

## Summary

The offline suite is broad, but it checks less than it appears to in four places:
- **Snapshot.** The dry-run snapshot records method + URL and drops every request body. The findings spec's interim acceptance test for workstream 5 is "method + URL + **body** per command".
- **MCP.** The server is only tested with a hand-rolled JSON-RPC harness and fake runners. No real MCP client ever talks to it, and no spawned session makes a `tools/call`.
- **Live tenant.** Roughly 45 † items across six specs wait for a live tenant. The `MS365_SMOKE_LIVE=1` mode that N7 names does not exist.
- **Completion.** The zsh, fish and PowerShell completion tests only check strings, and their syntax checks pass silently when the shell is missing.

This spec closes each gap with test-only changes plus one justified devDependency.

## Findings

| ID | Sev | Location | Finding | Failure scenario | Fix |
|---|---|---|---|---|---|
| TD-1 | P3 | `test/snapshot.test.ts:39,53` | The brief's premise was "only the first error line is recorded". `[VERIFIED: false positive as stated]`: `handleErrors` writes the error as one line (`src/errors.ts:86-88`, `Error: msg (fix: …)`), so `find(/^error:/i)` loses nothing. **The real gap:** the trace filter at `:39` drops `Body:`, `Headers:`, `File:` and `would ask` lines. Every non-trace, non-error stderr line is dropped too (`Uploading …`, `Warning: …`). **Measured** (scratch copy of the test, run twice): 177 dropped lines across 153 of 362 entries (120 `Body`, 35 `Headers`, 15 `would ask`, 5 `File`, 2 `Uploading`), about 15 KB. The two runs were byte-identical once the tmp dir was normalised. | A change that turns `saveToSentItems:true` into `false`, drops `startDateTime` from `mail flag` (M-19), or sends the `owner` field (M-27) leaves the snapshot green. These are exactly the regressions that workstream 5's interim acceptance (findings spec §5, row 5 `[REVISED]`) relies on the snapshot to catch. | Add a `detail` array to each snapshot entry. It holds every stderr line that is not a request line, the trailer or the error, after normalisation (OQ-1). Keep `requests` as the scan-friendly endpoint view. Keep the trailer assertions as they are. |
| TD-2 | P3 | `test/snapshot.test.ts:52` `[ADDED]` | `clean()` normalises only the tmp path. Nothing removes timing or retry noise (`Retrying in Ns`, `graph-client.ts:282-304`), `request-id` suffixes (`errors.ts:88` `rid`), or `$HOME`. None of these appears under dry-run today (measured), but TD-1 widens what gets recorded. | A future warning that carries a duration or a request ID makes the snapshot flaky. | Add one `normalize()` with a fixed rule list (tmp → `<tmp>`, `$HOME` → `<home>`, `/\d+(\.\d+)?m?s\b/` → `<dur>` inside `Retrying`/`resuming` lines, `request-id: \S+` → `<rid>`). Unit-test the rule list. |
| TD-3 | P2 | `test/mcp.test.ts:71-410`; `src/mcp-server.ts:415-446` | Every protocol test uses the in-repo `session()` harness. Nothing checks results against the MCP schema or drives the server with a real client. Unchecked today: `initialize` (`capabilities.resources.subscribe`), `tools/list` (each `inputSchema` is valid JSON Schema, tool names fit the allowed charset), `tools/call` results (`content`/`structuredContent`/`isError`), and that `notifications/resources/updated` has a valid shape. | A field rename or a shape slip, such as `structuredContent` set to an array or a `resources/read` without `contents[].uri`, passes the suite. Claude Desktop or an SDK client then rejects the response. | Add a conformance test with `@modelcontextprotocol/sdk` as a **devDependency** (OQ-2). Its `Client` + `StdioClientTransport` spawns `ms365 mcp serve`, and the SDK validates every response with its own zod schemas. Also validate each `inputSchema` with the SDK's JSON-Schema validator. |
| TD-4 | P2 | `test/mcp.test.ts:428-441`; `src/mcp-server.ts:207-225` | The only spawned test sends `initialize` and `tools/list`. `spawnRunner(process.argv[1])` re-launches the CLI with `process.execArgv`, and that path has never run end to end under tsx (loader flags) or under `dist/`. No spawned session makes a `tools/call` or exercises `resources/subscribe` → `notifications/resources/updated` → `resources/read`. | A child that can't find its entry point (for example, `execArgv` missing the tsx loader once bundled) makes every `tools/call` return `isError` with a stack trace. The unit tests stay green because they inject `runner`. | Spawned E2E session against `dist/index.js` (built in `beforeAll` or by the CI step), with `MS365_CLI_CONFIG_DIR` set to a temp dir. It runs `initialize`, `tools/list`, `tools/call mail_send` with `--dry-run` as a global (trace item present, `isError:false`), and `tools/call` of an offline tool (`describe`). Watch resources: `subscribe` with `--watch-interval 15` (the minimum the server accepts). Each `watch` pass is a real Graph call, so the full subscribe→notify loop only runs in the live suite (TD-7); offline, assert just the `resources/list` shape and that an unknown URI returns `-32602`. |
| TD-5 | P3 | `test/mcp.test.ts:421-429` `[ADDED]` | The spawned test is `it.skipIf(!registered)`. `registered` probes `--help` for `mcp` and swallows every error. `mcp` has shipped since then (`src/index.ts:161`). | If tsx breaks or the help layout changes, the only spawned MCP test quietly reports *skipped* instead of failing. | Drop the probe and make the test unconditional. |
| TD-6 | P3 | `test/mcp.test.ts:251-252` `[ADDED]` | The dry-run fixture uses `[dry-run] body:` and `Dry run: 1 request previewed.` Neither is a real format. The real ones are `[dry-run] Body:` (`graph-client.ts:109`) and the trailer `[dry-run] N request(s) previewed; nothing was sent.` (`errors.ts:122`). Because the real trailer starts with `[dry-run]`, `dryRunContent` (`mcp-server.ts:228-231`) **includes it** in the trace item, and no test asserts that. | Changing the trailer or `Body:` format breaks what MCP clients see, and the fixture-based test still passes. | Build the fixture from `dryRunTrace()` output (or capture it from a real `--dry-run` run), and assert that the trailer is the trace item's last line. TD-4's E2E covers the real path. |
| TD-7 | P2 | `docs/ms365-findings-spec.md` N7, D24; `tasks/specs/DECISIONS.md` N3; `scripts/smoke.sh` | N7 says to run "the smoke suite with `MS365_SMOKE_LIVE=1`", but no such mode exists (`grep -r SMOKE_LIVE` finds nothing, and `smoke.sh` is offline by design). None of the † items has a test that could run against a tenant. | When a tenant becomes available, nothing is ready to run. G-1 (2026-12-31) and G-2 (2027-04-01) are dated, so their readiness can't be confirmed in time. | Add an opt-in live suite (design below) plus the † checklist that maps each item to a test. |
| TD-8 | P3 | `.github/workflows/ci.yml:31-42` `[ADDED]` | CI never runs `scripts/smoke.sh`, the build-loop smoke gate, and never runs anything against `dist/` beyond building it. | A bundling regression, such as a dynamic import tsup didn't include, passes CI because vitest runs from source. | Add a CI step `SMOKE_NO_BUILD=1 scripts/smoke.sh` after `npm run build`, with the TD-4 E2E pointed at `dist/`. |
| TD-9 | P3 | `test/completion.test.ts:122-151`; CI `ubuntu-latest` | zsh and fish are checked by string `toContain` only. Their syntax checks end in `if (!has(…)) return;` (`:131`, `:141`) and pass silently when the shell is absent, which is always true locally and in CI. PowerShell gets no parse check at all. No zsh, fish or pwsh test checks what actually completes. The bash tests (`it.runIf(has('bash'))`, `:77-118`) are real. | A quoting bug in the generated zsh `case` table or a fish `string match` typo ships green. Users get no completions or a shell error on `source`. | Add `MS365_REQUIRE_SHELLS=1`, which turns a missing shell into a test failure (CI sets it). Add behavioural checks that mirror the bash ones (design below), and add a CI step that installs the shells. |

† depends on documented Graph behaviour and has not been run against a live tenant.

## Live test suite design (TD-7)

- **Location:** `test/live/*.live.test.ts`. The default vitest `include` excludes it, so it runs only via `npm run test:live` (its own `vitest.live.config.ts`). It spawns `node dist/index.js` for real. Nothing is injected, and the user's real token cache is used through `MS365_CLI_CONFIG_DIR=$MS365_LIVE_CONFIG_DIR`.
- **Gates:**
  - `MS365_SMOKE_LIVE=1` is required, or the whole suite is `describe.skip` with a single reason line.
  - `MS365_SMOKE_LIVE_WRITE=1` additionally enables the write tests.
  - `MS365_LIVE_ADMIN=1` enables the few tests that need tenant-admin toggles (G-1 and G-2).
  - `scripts/smoke.sh` gains a `MS365_SMOKE_LIVE=1` branch that runs `npm run test:live`, so N7's instruction works literally.
- **Safety:**
  - Before anything else, a preflight asserts that the signed-in UPN's domain matches `MS365_LIVE_TENANT_DOMAIN`. The suite aborts on a mismatch, so it can never run against a production mailbox by accident.
  - The read-only tier passes `--read-only` on every call, which makes the CLI's own gate refuse any write.
- **Write hygiene:**
  - Every object a test creates carries the `ms365-smoke-<runId>` prefix (subject, file name, plan title, todo title).
  - Each test registers its cleanup in `afterEach` (delete the message, file, plan, task or event), and changed settings are restored to the value captured beforehand (timezone, presence).
  - A `sweep` test that runs first deletes leftover `ms365-smoke-*` objects older than 24 h.
  - Mail goes only to the test account itself.
- **Fixtures:** `test/live/fixtures.md` lists what the test tenant must hold beforehand, because these can't be created cheaply per run: a team with more than 1 page of members, a OneNote section with more than 20 pages, a folder with more than 200 children, more than 1 page of groups that match `smoke`, a room mailbox, a sharing link, a recorded meeting and a second account (for M-23).
- **Output:** each test is named after its † IDs (for example `[M-3] mail triage default filter`), so the checklist below can be checked off from the vitest report. A † item whose expectation fails is reported as **"Graph disagrees with docs"**, with the response body attached. The fix then goes through the owning spec.

### † checklist (item → test; R = read-only tier, W = write tier, A = admin, F = needs fixture)

| † item (spec) | Tier | Test asserts |
|---|---|---|
| C-8 (findings) | R F | `groups list --search smoke --all --top 1` returns more than 1 page and page 2 succeeds (has `ConsistencyLevel`) |
| M-3 (findings) | R | plain `mail triage` exits 0 (no InefficientFilter) |
| M-6, D10 (findings) | W | `teams chat reply` posts to `/chats/{id}/messages` with a `messageReference` attachment, and the reply is listed |
| M-19 (findings) | W | `mail flag --due` on a message sent to self exits 0, and `dueDateTime` reads back |
| M-23 (findings) | W F | `cal decline --no-notify --propose-start` gets a usage error offline. With notify (second account's invite), the proposal is accepted |
| M-27, X-14 (findings, perf-dx) | W | `planner create-plan` with `container` succeeds, and the `container` column is populated |
| M-30, Q-4 (findings, query-pagination) | R | `mail list --search x` returns ≤ 250 results and no `nextCursor`. `--search` + `--filter` is rejected before any request |
| M-4, M-5 (findings) | W | on an uploaded `smoke.xlsx`: `excel sort-range` and `format-range` (font only, fill only) exit 0 and read back |
| M-18, Q-20 (findings, query-pagination) | R F | `cal rooms` / `room-lists` list the fixture room, and `--all` pages |
| M-20 (findings) | R F | `drive download-shared <https link>` writes the file (`u!` encoding) |
| G-1 (findings) | W A | after 2026-12-31: `mail update-draft <non-draft>` gives the CLI's clear message, not a raw 403 |
| G-2 (findings) | R A | with the tenant's tempauth control on: `drive download` of PDF/DOCX gets byte-equal output (sha256 against upload) |
| G-3 (findings) | — | none: no code uses the retiring endpoints. The offline grep guard stays |
| G-7 (findings) | R F | `meetings all-transcripts --api-version beta` lists the fixture meeting's transcript |
| G-8 (findings) | R | `/me/notes` shape check (only if PM-9's command exists, else skip with reason) |
| G-9, Q4 (findings, validation) | R F | Premium plan ID: which of 403/404 Graph returns, so Q4's wording can be pinned |
| DR-17 (dry-run) | W | `cal update --start` without `--timezone` keeps `originalStartTimeZone` (read back) |
| X-7 (perf-dx) | W | `drive copy --wait` handles 202 + monitor (no `Authorization` on monitor host), `--conflict rename` |
| X-13 (perf-dx) | R | `teams chats` with `$expand=members`: 1:1 rows show member names |
| X-15 (perf-dx) | W | `drive share` without a password creates a link without one (no `"password":""` sent) |
| Q-1, Q-2 (query-pagination) | R | `contacts list --search 'Bob "B"'` / `search people 'O"Brien'` exit 0 |
| Q-3, F16 (query-pagination, DECISIONS) | R | `groups list --search 'R&D "core"'` exits 0 (directory `\"` escaping) |
| Q-6, OQ-3 (query-pagination) | R F | `teams members <big> --all` exceeds 1 page. `joinedTeams` returns no `nextLink` |
| Q-8 (query-pagination) | R F | `onenote pages <section> --all` returns more than 20. `drive ls <folder> --all` returns more than 200 |
| Q-11 (query-pagination) | W | `mail move <id> --folder ms365-smoke-<run>` (non-well-known name) succeeds |
| Q-18 (query-pagination) | R F | with an app granted only `User.ReadBasic.All`: `user list` exits 0 |
| R-1 (robustness) | W | `mail create-draft --attach 5MB.bin`: 3 MiB chunks, final 201 with an empty body accepted, attachment size reads back |
| R-2 (robustness) | W | `mail reply --html` to self: the received reply is HTML and keeps the quoted thread |
| R-5 (robustness) | — | not live-testable (it states that a server-side dedupe ID is absent). The offline ledger test covers it |
| V-5, Q2 (validation) | W | `set-timezone` accepts a Windows name and an IANA name, then restores the original |
| V-6, F25 (validation) | W | each documented `set-presence` pair is accepted, then cleared. `--duration PT4H` passes and `PT5H` is a CLI usage error |
| V-12 (validation) | R | `cal free-busy --emails self --interval 5` and `1440` exit 0 |
| V-19, Q6 / V-8 (validation) | W | every allowed `todo --status` value and `planner --percent` value (per Q6) is accepted by Graph |
| V-22 (validation) | R | `mail list --top 1000` (the documented max) exits 0 |
| WS-5, WS-8 (findings §5) | R/W | rolled up: every workstream-5 row above passes, and the G-2 row passes before 2027-04-01 |
| X-12 (perf-dx) | — | Windows, not a tenant: covered by the CI matrix (OQ-5), not here |

## Completion checks design (TD-9)

- **fish:** `fish --no-execute` on the script, then `fish -c 'source <script>; complete -C "ms365 mail li"'` and so on. Assert the same cases as the bash block (`:79-96`, X-10, X-11).
- **zsh:** `zsh -n`, then a non-interactive driver. It runs `autoload -U compinit && compinit -u`, sources the script, stubs `compadd` to print the candidates that match `$PREFIX`, sets `words`/`CURRENT` and calls `_ms365`. Add one `zsh/zpty` interactive smoke (`ms365 ma<TAB>` → `mail`) as a guard against the stub drifting from real zsh (OQ-4).
- **pwsh:** a parse check with `[System.Management.Automation.Language.Parser]::ParseInput`, then dot-source the script and run `TabExpansion2 -inputScript 'ms365 mail li' -cursorColumn 13`, asserting on `.CompletionMatches.CompletionText`.
- **CI:** add a `completion` job on `ubuntu-latest`. It runs `sudo apt-get update && sudo apt-get install -y zsh fish`, checks `pwsh -v` (on GitHub's Ubuntu images pwsh is expected to be preinstalled; otherwise install it from Microsoft's apt repo), then runs `MS365_REQUIRE_SHELLS=1 npx vitest run test/completion.test.ts`. Optionally add `macos-latest` (zsh native, `brew install fish`) and `windows-latest` (pwsh native; also covers X-12) per OQ-5.

## Open questions

- **OQ-1 (engineering): Snapshot shape for TD-1.**
  - (a) Replace `requests` and `error` with one normalised `stderr` array.
  - (b) Keep `requests` and `error`, and add `detail` holding the remaining normalised lines.
  - (c) Write bodies to a separate file snapshot (`toMatchFileSnapshot`, one per module).
  - **Recommend (b).** The diff stays per-entry and the endpoint view stays scannable. It adds about 177 lines to a file of 4,063 (measured). (c) splits one command's behaviour across two files, which makes review harder.
- **OQ-2 (engineering): Conformance tooling for TD-3.**
  - (a) `@modelcontextprotocol/sdk` as an exact-pinned devDependency.
  - (b) Vendor the official `schema/2025-06-18/schema.json` and add `ajv` as a devDependency.
  - (c) Run the official MCP Inspector CLI through `npx` in CI only.
  - **Recommend (a).** One devDependency provides both a real client (handshake, `notifications/initialized`, capability enforcement with `enforceStrictCapabilities`) and the spec's result schemas. Nothing enters `dependencies`, and `dist/` is unaffected because `src/` never imports the SDK. Pin the version and assert that it still lists `2025-06-18` in its supported versions, so an SDK bump that drops the version fails loudly. (b) validates shapes but not behaviour. (c) is a heavy, network-dependent web-UI package and is hard to assert on.
- **OQ-3 (needs-human): The test tenant.**
  - (a) A Microsoft 365 Developer Program sandbox (eligibility is restricted now).
  - (b) A small paid tenant (one E3/E5 seat, plus a free second account for M-23).
  - (c) A dedicated account inside the GeneHealth tenant.
  - **Recommend (b).** It is isolated, admin toggles (G-1, G-2) are allowed, and nothing ever touches production. (c) only if a CA/compliance owner signs off, and never with `MS365_LIVE_ADMIN`.
- **OQ-4 (engineering): The zsh harness.**
  - (a) A `compadd` stub only.
  - (b) zpty only.
  - (c) The stub for every case, plus one zpty smoke.
  - **Recommend (c).** The stub is deterministic and fast, and zpty proves the stub matches real zsh.
- **OQ-5 (engineering): The CI shell matrix.**
  - (a) Ubuntu job only.
  - (b) Add macOS and Windows.
  - **Recommend (a) now.** Add a `windows-latest` pwsh leg only once X-12 is scheduled, because it doubles CI time for one P3.
- **OQ-6 (needs-human): Where the live suite runs.**
  - (a) Manually, on a developer machine.
  - (b) Also on a weekly `schedule:` workflow, with a GitHub Environment that holds a refresh-token secret.
  - **Recommend (a) first, then (b) after OQ-3.** A stored refresh token must be from the isolated tenant and rotated, which is the reason N7 stalled (90-day expiry).

## Acceptance criteria

1. Changing any request body field in a command, for example `saveToSentItems:true` → `false` in `mail send`, fails `test/snapshot.test.ts`. Running the suite twice in a row produces no snapshot diff.
2. The `normalize()` rule list has unit tests for tmp, `$HOME`, durations and request IDs.
3. `npm test` includes an SDK-client test that spawns `mcp serve`. It completes `initialize`, `listTools`, `callTool` (dry-run `mail_send` and offline `describe`) and `listResources`, and the SDK raises no schema-validation errors. Every tool's `inputSchema` compiles as JSON Schema.
4. The spawned MCP test has no `skipIf`. A `--dry-run` `tools/call` returns two content items, and the last line of the second is the real trailer.
5. The `test/mcp.test.ts` dry-run fixtures contain no strings that `src/` cannot produce. Each fixture is generated by `dryRunTrace()`, or a test asserts that it matches `src/errors.ts`'s trailer regex.
6. CI runs `scripts/smoke.sh` against `dist/` with 0 FAIL and 0 SKIP.
7. With `MS365_SMOKE_LIVE` unset, `npm run test:live` exits 0 and reports one skipped suite with its reason. The default `npm test` runs no live file.
8. With `MS365_SMOKE_LIVE=1` and a mismatched `MS365_LIVE_TENANT_DOMAIN`, the suite aborts before making any Graph call. Without `MS365_SMOKE_LIVE_WRITE=1`, no write test runs, and every read test passes `--read-only`.
9. Every † ID listed in `docs/ms365-findings-spec.md` and `tasks/specs/*.md` appears in the † checklist (a doc-lint test greps the specs for `†` IDs and diffs them against `test/live/*.live.test.ts` test names plus the "—" rows).
10. CI's `completion` job runs with `MS365_REQUIRE_SHELLS=1`. The zsh, fish and pwsh behavioural cases (the bash set plus X-10 and X-11) pass, and removing any of the three shells fails the job.

## Suggested task breakdown

| Task | Scope | Depends on |
|---|---|---|
| TD-T1 snapshot `detail` + `normalize()` (TD-1, TD-2), one snapshot regeneration reviewed as a diff | `test/snapshot.test.ts`, `test/helpers/normalize.ts`, `test/__snapshots__/` | — |
| TD-T2 add the `@modelcontextprotocol/sdk` devDependency and the conformance test (TD-3) | `package.json`, `test/mcp-conformance.test.ts` | OQ-2 |
| TD-T3 spawned E2E against `dist/`, drop `skipIf`, real dry-run fixtures (TD-4, TD-5, TD-6) | `test/mcp.test.ts`, `test/mcp-conformance.test.ts` | TD-T2 |
| TD-T4 CI: smoke step + build-before-E2E (TD-8) | `.github/workflows/ci.yml` | TD-T3 |
| TD-T5 completion harnesses for fish, zsh and pwsh, plus the `MS365_REQUIRE_SHELLS` gate (TD-9) | `test/completion.test.ts`, `test/helpers/shells.ts` | OQ-4 |
| TD-T6 CI `completion` job | `.github/workflows/ci.yml` | TD-T5, TD-T4 (same file; serialise) |
| TD-T7 live harness: config, gates, tenant preflight, cleanup registry, sweep, `smoke.sh` branch (TD-7) | `vitest.live.config.ts`, `test/live/harness.ts`, `scripts/smoke.sh`, `package.json` | — |
| TD-T8 live read tier (all R rows) + `fixtures.md` | `test/live/read.live.test.ts`, `test/live/fixtures.md` | TD-T7 |
| TD-T9 live write and admin tiers (W and A rows) | `test/live/write.live.test.ts` | TD-T7 |
| TD-T10 † doc-lint (AC 9) | `test/live-checklist.test.ts` | TD-T8, TD-T9 |

**Out of scope:** running the live suite. That still waits on OQ-3 and N7, and stays † until it runs.
