# DRAFT — Validation, exit-code and stdout consistency

**Status:** DRAFT — must pass its own Step 1 adversarial review before any build. Not scheduled.

**As of:** 2026-10-14 · **Scope:** `src/**` at `41c41f3` (branch `build/ms365-findings-spec`) · **Source:** `tasks/opportunities.md` lines 12, 13, 16, 17, 18, 20, 22, 25 · **Theme priority:** P2

## Summary

D2 set an exit-code contract: 2 for usage, 3 for auth, 5 for not found. D2 and C-19 also say stdout carries data only. Many command paths still break both rules. User-input mistakes throw a plain `Error`, which exits 1. Enumerated and numeric flags go to Graph unchecked, or they silently produce wrong output with exit 0. A few commands still print status lines on stdout. This spec routes every user-input error through `CliError` with `ExitCode.USAGE` (or `NOT_FOUND`), puts commander `.choices()` on the enumerated flags, and moves the remaining status lines to `info()`.

Verification method: every row was read against HEAD. Rows marked **(ran)** were also reproduced with `tsx src/index.ts --profile p --dry-run …`, run from a scratch directory with an isolated `MS365_CLI_CONFIG_DIR` and no `MS365_CLI_CLIENT_ID`. The profile `p` holds only a client ID. † means the row depends on documented Graph behaviour that hasn't been checked against a live tenant (D24).

## Findings

| ID | Sev | Location | Finding | Failure scenario | Fix |
|---|---|---|---|---|---|
| V-1 | P2 | `src/commands/user.ts:138,238,250` | `set-presence`, `set-timezone` and `set-language` print their confirmation with `console.log` (the C-19 pattern). | `ms365 user set-timezone UTC -q \| jq` gets `Timezone set to UTC.` on stdout. `-q` doesn't suppress it. **(ran)** | Use `info()` for these lines, as `clear-presence` (`user.ts:152`) already does. |
| V-2 | P2 | `src/commands/user.ts:126-128,146-147` | The presence commands read `process.env.MS365_CLI_CLIENT_ID` directly and throw a plain `Error`. They ignore a client ID supplied by `--profile` (`index.ts:119`). The message ("Client ID not configured") doesn't match `AUTH_PATTERN` (`errors.ts:52`), so it exits 1. | `ms365 --profile p user set-presence …` fails with exit 1 even though the client ID is configured and every other command works. **(ran)** | Expose the resolved ID (`AuthManager.clientId`, read-only) through a public `GraphClient` accessor and use it as `sessionId`. Delete the env read and the local check: the `AuthManager` constructor already throws `no_client_id` with exit 3 (`auth.ts:261-265`). |
| V-3 | P3 `[ADDED]` | `src/index.ts:116-128` → `src/auth.ts:259-265` | The opportunity said presence needs a client ID under `--dry-run`. That's true, but it isn't specific to presence. `getClient()` always constructs an `AuthManager`, which throws without a client ID, and the dry-run path in `GraphClient.request` never calls `getToken()` (`graph-client.ts:165-175`). | `ms365 --dry-run mail list` with no client ID exits 3, so a fresh install can't preview anything. **(ran)** | See Q1. If (b) is chosen, build `AuthManager` lazily on the first `getToken()`, and under dry-run use the placeholder `sessionId` `"<client-id>"` when none is configured. |
| V-4 | P3 | `src/commands/user.ts:60,63` | `set-ooo --status` and `--audience` take any string. | `--status bogus --audience nobody` sends the PATCH and exits 0 under dry-run. Live, Graph returns a 400 and the command exits 1. **(ran)** | Use `new Option(...).choices(['alwaysEnabled','scheduled','disabled'])` and `.choices(['all','contactsOnly','none'])`. Commander's override already exits 2 and lists the allowed values (`index.ts:77-80`). |
| V-5 | P3 | `src/commands/user.ts:229-251` | `set-timezone` and `set-language` (`[ADDED]`) send any value. | `set-timezone Not/AZone` prints "Timezone set" and exits 0 under dry-run **(ran)**; live, the PATCH fails with exit 1 †. `set-language xx_bad` has the same problem. | Validate the time zone per Q2. Validate the language tag with `Intl.getCanonicalLocales(locale)`, which throws `RangeError` on a malformed tag; map that to a `usage` error. |
| V-6 | P3 `[ADDED]` | `src/commands/user.ts:121-123` | `set-presence --availability`, `--activity` and `--duration` aren't validated. The help lists `BeRightBack`, `Offline` and `OffWork`, which belong to `setUserPreferredPresence`, not `setPresence` †. `--duration` is described as "in minutes" but must be ISO 8601 (`PT5M`–`PT4H` †). | `--availability Offline` is rejected by Graph with exit 1. `--duration 30` is sent as-is. | Use choices per Q3, check `--duration` against `^PT(\d+H)?(\d+M)?$` with a range of 5–240 minutes, and fix the help text. |
| V-7 | P3 | `src/commands/user.ts:196-214` | `upload-photo` has four problems. (1) The `response.status >= 400` check is dead code: `request()` throws on HTTP errors and dry-run returns `status: 0`. (2) "File not found" and "Photo too large" are plain `Error`s, so they exit 1. (3) `[ADDED]` An unknown extension silently falls back to `image/jpeg` (`:203`). (4) The size pre-check duplicates `uploadFile`'s `file_too_large` check (`graph-client.ts:342-344`). | `upload-photo /nonexistent.jpg` exits 1 **(ran)**. `upload-photo me.webp` is uploaded labelled as JPEG. | Delete lines 212-214. Throw `usage` errors for a missing file, an oversized file and an unsupported extension, matching `mail.ts:136`. Allowed extensions: `.jpg`/`.jpeg`/`.png`/`.gif`. Keep the pre-check, since it gives the clearer message. |
| V-8 | P3 | `src/commands/planner.ts:206-211` | The `update-task --percent` check throws a plain `Error`, so it exits 1. `parseInt` also accepts `50abc`. | `--percent 30` exits 1 **(ran)**. | Throw `CliError` with `usage` and require `/^(0\|50\|100)$/` (see Q6). Validate before `getClient()` and before any request. |
| V-9 | P3 | `src/commands/planner.ts:234-236` | `update-plan-details --categories` calls `JSON.parse` with no guard. The plan-name resolve request runs before the parse, and the shape isn't checked. | `--categories '{bad'` sends a GET first, then exits 1 with a raw V8 parse message **(ran)**. | Parse before `resolve()`. On failure, or when the value isn't an object with keys `category1`–`category25` and string or null values, throw a `usage` error, e.g. `--categories must be a JSON object like {"category1":"Red"}`. |
| V-10 | P3 | `src/commands/planner.ts:8-22` | The G-9 Premium hint is attached to every 403 and 404 on `/planner/(plans\|tasks)`, including a mistyped ID. | `planner task WRONGID` reports "Premium plans … aren't available", which sends the user down the wrong path. | Reword per Q4: `If this is a Premium plan (Project for the web), Graph's Planner API can't reach it.` |
| V-11 | P2 | `src/utils.ts:66-79` | `validateEmail`, and `parseRecipients` built on it, throw a plain `Error`, so they exit 1. Callers: `mail.ts:445,457,461,567,579,647,650,812,965,1036` and `calendar.ts:159,397`. `[VERIFIED: false positive for teams chat create — teams.ts:265-273 already wraps it in usageEmail(), exit 2]` | `mail send --to bad` exits 1, and so does `cal create --attendees bad` **(ran)**. An agent can't tell a typo from a server failure. | Throw `CliError` with `{code:'usage', exitCode: USAGE}` in `validateEmail`. Then delete `usageEmail` (`teams.ts:267-273`) and call `validateEmail` directly. |
| V-12 | P3 | `src/commands/calendar.ts:367,376` | `free-busy --emails` isn't validated, and `--interval abc` becomes `availabilityViewInterval: null` (`[ADDED]`). | `--emails notanemail,x --interval abc` sends `"schedules":["notanemail","x"]` with a null interval and exits 0 under dry-run **(ran)**. | Map the list through `validateEmail`. Require `--interval` to be an integer from 5 to 1440 †, or throw `usage`. |
| V-13 | P3 `[ADDED]` | `src/utils.ts:52-58`; caller `workflow.ts:250` | `parseDate` throws a plain `Error`, so it exits 1. It also accepts impossible dates: `new Date('2026-02-30T00:00:00Z')` rolls over to 03-02, and `2026-02-30` is still sent to Graph. In `email-to-task` it runs only after the To Do list GET (`workflow.ts:222-231`). | `wf email-to-task ID --due 2026-02-30` creates a task due on a date that doesn't exist, or exits 1 after a network call. | Check `/^\d{4}-\d{2}-\d{2}$/` and round-trip through `toISOString().slice(0,10) === input`, throwing `usage` on failure. In `email-to-task`, validate `--due` before any request. |
| V-14 | P3 | `src/commands/auth.ts:80,88`; `src/auth.ts:340-358` | `auth accounts` marks "(active)" only from the saved selection. It ignores `--account`, `MS365_CLI_ACCOUNT` and `profile.account`. `[ADDED]` When nothing is saved, no account is marked, although `getCurrentAccount()` falls back to `accounts[0]` (`auth.ts:357`). | `ms365 --account bob@x auth accounts` shows alice as active, while every other command runs as bob. | Add `AuthManager.activeAccountId()`, using the same precedence as `getCurrentAccount()` but returning `null` rather than throwing when no account matches, and mark that account. If an `--account` value matches nothing, print an `info()` warning. |
| V-15 | P3 | `src/commands/setup.ts:8-16,112,134` | Each `prompt()` opens a new readline and closes it. With piped stdin, the first interface buffers all the input and discards it on `close()`. The second `question()` never resolves, the event loop empties, and the process exits 0 without reaching `handleErrors`' exit handler. | `printf '2\nn\nabc\n' \| ms365 auth setup` stops at "Open Azure Portal in browser? [Y/n]" with exit 0, and no client ID is asked for. **(ran)** | Use one line reader for the whole wizard (see V-16). When stdin closes before an answer arrives, throw a `usage` error: `setup needs an answer for "<question>" (stdin closed)`. |
| V-16 | P3 | `src/commands/setup.ts:8-16` vs `src/utils.ts:38-44` | `setup.ts` `prompt()` duplicates the readline code in `confirm()`. | It's a maintenance hazard and the root cause of V-15. | Add `createPrompter()` to `utils.ts` (one `readline.Interface` plus an async line iterator that rejects on `close`). Build `confirm()` and the wizard on it. |
| V-17 | — | `src/commands/calendar.ts:89-93,116,144,266,368,393` | `--timezone` defaults to `'UTC'`. `[VERIFIED: already fixed in 40c1d59 (T11/M-9) — every --timezone resolves through resolveTimeZone() (MS365_CLI_TIMEZONE → system zone); dry-run free-busy shows America/New_York; the only 'UTC' literal left is calendar.ts:28, which maps Graph's UTC label]` | — | None. |
| V-18 | P3 `[ADDED]` | `drive.ts:114,322`; `mail.ts:529,604,768`; `workflow.ts:229` | The plain `Error`s that remain in `src/commands` all exit 1: "File not found" (local path) in `drive upload` and `drive cp`; "Folder not found" and "Drafts folder not found"; "No content found"; "No To Do lists found". | `drive cp ./missing.pdf /x.pdf` exits 1, while `mail send --attach ./missing.pdf` exits 2 (`mail.ts:136`). `wf email-to-task --dry-run` always fails "No To Do lists" because the dry-run data has no `value`. | Local file missing: `usage` (2). Folder or list not found: `not_found` (5). Attachment with no bytes: `CliError` code `unsupported_attachment` (1). In `workflow.ts:226-229`, use the placeholder `LIST_ID` under dry-run, like `mail.ts:178`. Out of scope: `docs-gen.ts:108`, which is a build-time tool. |
| V-19 | P3 `[ADDED]` | `workflow.ts:207`; `todo.ts:73,103,104`; `mail.ts:1012`; `drive.ts:198-199`; `calendar.ts:461`; `excel.ts:45` | Enumerated flags list their values in the help text but don't use `.choices()`: `--importance`, `--status`, `--type`, `--scope`, `--color` and `--series-by`. | `todo update … --status done` reaches Graph and fails with a 400 (exit 1) †. | Use `new Option(...).choices([...])` with the listed values. It exits 2 and prints the allowed values. |
| V-20 | P3 `[ADDED]` | `src/commands/workflow.ts:459-460` | `focus-time --work-start` and `--work-end` aren't validated. `NaN` makes every slot look busy. | `wf focus-time --work-start abc` prints "No focus time available today. Your calendar is fully booked." and exits 0 **(ran)**. | Require integers 0–24 with start < end, or throw `usage`. |
| V-21 | P3 `[ADDED]` | `drive.ts:221-223`; `workflow.ts:257-263` | `drive share` and `wf email-to-task` print human text for any `-o` value except `json`, so `table`/`csv`/`yaml`/`ndjson` get prose. | `ms365 drive share … -o yaml` prints `Sharing link (view, organization):` instead of YAML. | Use human text only for `-o text`, and `formatOutput(data, opts.output)` for every other format. |

## Open questions

| # | Question | Options | Recommendation | Tag |
|---|---|---|---|---|
| Q1 | Should `--dry-run` work with no client ID (V-3)? | (a) keep requiring one, exit 3; (b) build `AuthManager` lazily so dry-run never needs one | **(b)**. Dry-run never sends a request, and previewing before setup helps both agents and new users. | engineering |
| Q2 | What should `set-timezone` accept (V-5)? Graph `mailboxSettings.timeZone` takes Windows names and also IANA names †. | (a) IANA only, via `resolveTimeZone`; (b) an `Intl`-valid IANA name, or a Windows-shaped name (`/^[A-Z][A-Za-z .+()-]* Time$\|^UTC$/`); (c) no validation | **(b)**. It catches typos without rejecting the Windows names that Outlook itself shows. | engineering |
| Q3 | Which `set-presence` values should be allowed (V-6)? | (a) the `setPresence` documented pairs †: availability `Available\|Busy\|Away\|DoNotDisturb`, activity `Available\|InACall\|InAConferenceCall\|Away\|Presenting`; (b) leave them free-form | **(a)**, with corrected help text. Mark it "unverified live" per D24. | engineering |
| Q4 | When should the G-9 Premium hint appear (V-10)? | (a) on 403 only; (b) on 403 and 404 with conditional wording; (c) drop it | **(b)**. Which status Graph returns for a Premium plan is undocumented †, so keeping both statuses with "If this is…" wording avoids false certainty. | engineering |
| Q5 | Setup wizard on non-TTY stdin (V-15)? | (a) refuse at the start with exit 2; (b) accept piped answers through one line reader, and exit 2 on EOF | **(b)**. It makes the wizard scriptable and testable without a pty. Device-code login also works without a TTY. | engineering |
| Q6 | `planner --percent` range (V-8)? | (a) keep `0\|50\|100`; (b) any integer 0–100 (Graph stores it †) | **(a)**. It's the current contract and matches the three Planner UI states. Revisit in a feature spec. | engineering |
| Q7 | Is `auth status`'s `Logged in as …` (`auth.ts:57`) a status line or data? | (a) data, stays on stdout; (b) status, moves to `info()` | **(a)**. It's the command's answer. A `-o json` form belongs in an agent-readiness spec, not this one. | engineering |

## Acceptance criteria

Run in-process via `test/helpers/cli.ts`, using mocked routes and the exit-handler capture:

1. No `throw new Error(` remains in `src/commands/**` or `src/utils.ts`. A grep-based test in `test/` enforces this, with `docs-gen.ts` exempt.
2. Each of these exits **2**, with nothing on stdout and no recorded Graph call:
   - `user set-ooo --status bogus`
   - `user set-ooo --status disabled --audience nobody`
   - `user set-timezone Not/AZone`
   - `user set-language xx_bad`
   - `user set-presence --availability Offline --activity Away`
   - `user set-presence … --duration 30`
   - `user upload-photo /nonexistent.jpg`
   - `user upload-photo x.webp` (existing file)
   - `planner update-task T --percent 30`
   - `planner update-task T --percent 50abc`
   - `planner update-plan-details P --categories '{bad'`
   - `planner update-plan-details P --categories '{"foo":1}'`
   - `mail send --to bad …`
   - `cal create … --attendees bad`
   - `cal free-busy --emails bad …`
   - `cal free-busy … --interval abc`
   - `wf email-to-task M --due 2026-02-30`
   - `wf focus-time --work-start abc`
   - `todo update L T --status done`
   - `drive cp ./missing /x`
   - `printf '2\n' | auth setup` (EOF before the client ID)
3. `--profile p user set-presence …` (client ID only in the profile) sends `sessionId` equal to the profile's client ID. With no client ID anywhere, it exits **3** with code `no_client_id`.
4. `user set-timezone UTC -q`, `set-language en-US -q` and `set-presence … -q` print nothing on stdout or stderr. Without `-q`, the confirmation goes to stderr.
5. `mail folder Nope` and `wf email-to-task` with zero lists exit **5**.
6. `--account bob auth accounts` marks bob "(active)". With no saved selection, it marks `accounts[0]`.
7. A 404 from `/planner/tasks/X` carries a `fix` that starts with "If this is a Premium plan".
8. `drive share … -o yaml` and `wf email-to-task … -o yaml` print valid YAML of the response.
9. If Q1 is (b): with no client ID, `--dry-run mail list` exits 0 and prints the request preview.
10. All existing tests and `npm run build` pass. `user.ts:212-214` is gone.

## Suggested task breakdown

| Task | Findings | Files | Depends on |
|---|---|---|---|
| VT1 — utils errors and prompter | V-11 (utils part), V-13 (utils part), V-16 | `src/utils.ts`, `test/utils.test.ts` | — |
| VT2 — user.ts consistency | V-1, V-2, V-4, V-5, V-6, V-7 | `src/commands/user.ts`, `src/auth.ts` (`clientId` getter), `src/graph-client.ts` (accessor), `test/commands/misc.test.ts` | — |
| VT3 — planner | V-8, V-9, V-10 | `src/commands/planner.ts`, `test/commands/planner-todo.test.ts` | — |
| VT4 — email callers | V-11 (callers, teams `usageEmail` removal), V-12 | `src/commands/teams.ts`, `src/commands/calendar.ts`, `test/commands/calendar.test.ts` | VT1 |
| VT5 — auth accounts | V-14 | `src/commands/auth.ts`, `src/auth.ts`, `test/auth-features.test.ts` | VT2 (both touch `src/auth.ts`) |
| VT6 — setup wizard | V-15 | `src/commands/setup.ts`, `test/commands/setup.test.ts` | VT1 |
| VT7 — remaining errors, choices, numbers, -o | V-13 (caller), V-18, V-19, V-20, V-21 | `src/commands/{drive,mail,workflow,todo,calendar,excel}.ts`, matching tests | VT1, VT4 (both touch `calendar.ts`) |
| VT8 — dry-run without a client ID (if Q1 = b) | V-3 | `src/index.ts`, `src/auth.ts`, `src/graph-client.ts`, `test/index.test.ts` | VT2, VT5 |
| VT9 — guard test plus snapshot refresh | AC 1, AC 10 | `test/` (new grep test), `test/__snapshots__` | VT1–VT8 |
