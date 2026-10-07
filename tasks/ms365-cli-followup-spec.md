# ms365-cli — Follow-up Spec (deferred opportunities)

Status: DRAFT (not adversarially reviewed; open questions in §8). Date: 2026-10-06.
Source: all 25 rows of `tasks/opportunities.md`, each re-checked against branch `build/ms365-cli-hardening` at `bb72388` (static read; file:line anchors are current). Consumable by `/build-loop tasks/ms365-cli-followup-spec.md` once §8 is resolved.
Builds on `tasks/ms365-cli-hardening-spec.md` (same priority scale, confidence tags and conventions: exits through `handleErrors`, stdout carries data only, global flags before the subcommand, tests via the built binary + fake Graph in `test/e2e/run-cli.ts`).

Priority: **P0** data loss / security / wrong-target or duplicate writes. **P1** documented feature broken. **P2** wrong edge behavior / UX / scripting contract. **P3** cleanup, docs.
Confidence: `[verified]` reproduced; `[code]` read from source, unambiguous; `[graph-doc]` depends on Graph API semantics — confirm against Microsoft docs before fixing; `[INFERENCE]` derived, not observed.

## 0. Baseline and mapping

| Check | Result |
|---|---|
| `npm test` | 39 files, 579 tests pass; `tsc --noEmit` 0 errors |
| `npm audit --omit=dev` | 0 vulnerabilities |
| Opportunities still holding | 23 of 25 as written; #20 and #21 partly fixed already but the remaining problem is wider than the row |

| Opportunity # | Spec item |
|---|---|
| 21, 7 | DRY-1, DRY-2 |
| 18 | OUT-1 |
| 8, 9, 10 | CAL-3, CAL-4, CHOICE-1 |
| 11 | TZ-1 |
| 20 | CHOICE-1 |
| 17 | USER-4 |
| 4 | MAIL-9 |
| 5 | MAIL-10 |
| 6, 13, 16 | PAGE-2 |
| 14 | DRIVE-3 |
| 12 | SP-3 |
| 15 | CLEAN-1 |
| 19 | META-2 |
| 22 | SCHEMA-2 |
| 24, 25 | REMOTE-4 |
| 3, 23 | REMOTE-5 |
| 1 | REMOTE-2 (later) |
| 2 | REMOTE-3 (later) |

## 1. Dry-run and output contract

### DRY-1 (P2) Success messages printed under `--dry-run` `[code]`
- Under `--dry-run`, `GraphClient` returns `{ data: { dryRun, method, url, body? }, status: 0 }` (`src/graph-client.ts:151-156,177-179`). No central check exists; each command must test `status !== 0` itself. Only `teams.ts:15` `post()` and `user.ts:172,180,207,234,256,267` do.
- 28 unguarded success `info()` lines claim work that did not happen, e.g. `ms365 --dry-run mail send …` prints `Email sent successfully.` (`mail.ts:300`) [INFERENCE: derived from code, not run]:
  - `calendar.ts:191,227,261,278` (delete, accept, decline, tentative)
  - `contacts.ts:136` (delete)
  - `drive.ts:118` (delete)
  - `mail.ts:300,310,324,401,456,511,560,576,589,601,623,635,647,662,728,871,896` (send, delete, move, draft, send-draft, delete-attachment, reply, forward, mark-read, mark-unread, flag, unflag, complete, categorize, shared-send, batch-read, batch-move)
  - `planner.ts:185,194,203` (delete-task, delete-plan, delete-bucket)
  - `todo.ts:131` (delete)
  - `workflow.ts:428` (file-announce; prints `Announced "undefined"` [INFERENCE])
- Related: `drive.ts:42` prints `Uploading …` before the preview; `workflow.ts:196` meeting-prep prints `No upcoming meetings in the next 24 hours.` in dry-run because the preview has no data.
- **Fix:** one helper (e.g. `done(res, message)` in `src/utils.ts` or `src/graph-helpers.ts`) that calls `info()` only when the response was real (`status !== 0`). Apply it to all 28 lines and move `teams.ts` `post()` and the guarded `user.ts` lines onto it. Batch commands (`mail.ts:871,896`): under dry-run print nothing or `Would mark N message(s)…` — pick one rule and use it everywhere (§8 Q1). `drive upload`: skip the `Uploading` line under dry-run. Meeting-prep: under dry-run, skip the empty-state line.
- **Accept:** one table-driven `--dry-run` e2e test covering at least `mail send`, `mail reply`, `mail flag`, `mail batch-move`, `calendar accept`, `drive upload`, `workflow file-announce`, `workflow meeting-prep`: stderr has the `[dry-run]` preview and none of the success/empty-state text; exit 0. Delete commands are refused by `confirm()` without a TTY, so they are covered by the shared helper rather than e2e. No source-text grep tests. README `252-253,847-850` updated.

### DRY-2 (P3) `workflow email-to-task --dry-run` without `--list` fails `[code]`
- `workflow.ts:266-273` looks up the default To Do list; under dry-run the preview has no `value`, so it throws `No default To Do list found; pass --list <listId>.` (exit 1, no POST preview) [INFERENCE].
- With `--list L1` the title becomes `[Email] (no subject)` because the message is the dry-run object.
- **Fix:** under dry-run, use the placeholder `<default-list>` as the list ID (same pattern as planner's `'<dry-run>'` ETag, `planner.ts:27`), and `<subject>` for the title.
- **Accept:** `ms365 --dry-run workflow email-to-task M1 -o json` exits 0, stderr shows `POST …/me/todo/lists/%3Cdefault-list%3E/tasks` (or the unencoded form the URL builder produces), and no request is made.

### OUT-1 (P3) Teams send commands drop the created message `[code]`
- `teams.ts:13-16` `post()` discards `res.data`; `teams send-channel-message` (85), `teams reply-channel-message` (99), `chat send` (146), `chat reply` (155) print nothing on stdout, so callers cannot get the new message ID.
- Precedent: `workflow file-announce` prints the posted message with `-o` (`workflow.ts:428-429`, CHANGELOG WF-3).
- **Fix:** `.addOption(outputOption())` on the four commands; print `formatOutput(res.data, opts.output)` on stdout; confirmation stays on stderr (via DRY-1 helper).
- **Accept:** e2e with a fake 201 `{id:'m1',…}`: `chat send C1 --body hi -o json` stdout parses and has `id: 'm1'`; dry-run stdout is the dry-run object. Update `schema` ENDPOINTS for `chat.send`/`chat.reply` (response shape), README `554-555,568-569`.

## 2. Calendar, To Do, time zones

### CAL-3 (P3) `calendar free-busy --interval` unvalidated `[code]`
- `calendar.ts:289,300`: `parseInt(opts.interval)`; `abc` → `NaN` → `null` in the body, `15x` → `15`.
- **Fix:** integer option parser (CLEAN-1's exported `integer(min, max)`), range 5-1440 `[graph-doc]` (confirm bounds for `availabilityViewInterval` in the getSchedule docs).
- **Accept:** `--interval abc`, `--interval 15x`, `--interval 0` exit 1 before any request; `--interval 15` sends `15`.

### CAL-4 (P3) Inconsistent time zone defaults `[code]`
- `calendar create --timezone` default `UTC` (`calendar.ts:104`), `update` falls back to `'UTC'` (`158-159`, help `147`), `free-busy` default `UTC` (`288`); `decline` uses `localTimeZone()` (`251`). README `451-452` documents the split.
- **Fix (pending §8 Q2):** default every calendar `--timezone` to `localTimeZone()` (`src/graph-helpers.ts`), which already falls back to UTC when the runtime cannot name the zone.
- **Accept:** with `TZ=America/New_York`, `--dry-run cal create … ` body has `timeZone: 'America/New_York'`; with `TZ=JST-9` it is `UTC`; explicit `--timezone` wins. README/CHANGELOG updated (interface change).

### TZ-1 (P3) To Do and mail-flag dates written as UTC midnight `[code]`
- `todo.ts:84,109`, `workflow.ts:289`: `dueDateTime = { dateTime: '<date>T00:00:00', timeZone: 'UTC' }`; `mail.ts:617-618` (flag start/due) does the same. Clients east/west of UTC may show the previous/next day. Readers already compensate (`workflow.ts:69-79` `dueDay()`); planner writes local midnight (`planner.ts:48-52`).
- **Fix (pending §8 Q3):** send `timeZone: localTimeZone()` with `T00:00:00` for all four writers; keep the reader tolerant of old UTC-midnight values.
- **Accept:** with `TZ=Asia/Tokyo`, `todo create L1 --title t --due 2026-10-07` body has `{dateTime:'2026-10-07T00:00:00', timeZone:'Asia/Tokyo'}`; `workflow digest` still buckets an old `T00:00:00 UTC` task on its intended day.

## 3. Input validation

### CHOICE-1 (P3) Fixed-value options accept free text `[code]`
- Already fixed: `user set-ooo --status/--audience`, `set-presence`, `todo`, `workflow --importance`, `planner --percent`, `drive --conflict`, `-o`.
- Still free text (values from the code/help):
  - `calendar create-calendar --color` (`calendar.ts:331`) and `update-calendar --color` (`349`, no list in help): auto, lightBlue, lightGreen, lightOrange, lightGray, lightYellow, lightTeal, lightPink, lightBrown, lightRed, maxColor
  - `drive share --type` (`drive.ts:162`): view, edit, embed; `--scope` (`163`): anonymous, organization, users
  - `mail create-rule --importance` (`mail.ts:758`, used `792`): low, normal, high
  - `excel create-chart --series-by` (`excel.ts:55`): auto, columns, rows
  - `excel create-chart --type` (`excel.ts:53`): Graph chart-type enum `[graph-doc]` (full list not in code; §8 Q4)
  - `search query --entity` (`search.ts:51`): comma list, validated at runtime by `entityTypes()` (`25-36`) but invisible to the command tree, so `ms365 serve` validation and generated skills don't see it.
- Why it matters: `command-tree.ts:115` copies `argChoices`; remote validation rejects bad values (`command-tree.ts:219-222`); skills render `one of: …` (`generate-skills.ts:33`).
- **Fix:** `new Option(...).choices([...])` for each fixed list (shared `CALENDAR_COLORS` for both calendar commands). For `--entity`, add a way for the command tree to carry a comma-list choice set (e.g. an `OptionNode.listChoices` read from a custom property on the Option) and validate it in `validateInvocation`; keep `entityTypes()` bucket rules.
- **Accept:** each listed option rejects a bad value with exit 1 and `Allowed choices` before any request, locally and via `ms365-remote` against `ms365 serve` (400 `invalid_args`); generated skills list the values. README `404,446,490-491,649,665-666`.

### USER-4 (P3) `user upload-photo` accepts any file as JPEG `[code]`
- `user.ts:223-225`: private map for jpg/jpeg/png/gif; any other extension (e.g. `.pdf`) uploads as `image/jpeg`. Description says "(JPEG, PNG, GIF — max 4MB)" (`user.ts:215`).
- **Fix:** use `detectMimeType` (`src/mime.ts`); accept only `image/jpeg`, `image/png`, `image/gif` (the documented set; §8 Q5) and exit 1 otherwise, before reading the file.
- **Accept:** `.pdf` and `.svg` → exit 1, no request; `.PNG` → PUT with `image/png`.

## 4. Mail

### MAIL-9 (P3) `mail attachments` downloads every attachment's content to list it `[code]`
- `mail.ts:462` lists `/me/messages/{id}/attachments` with no `$select`, so Graph returns `contentBytes` for each attachment; `mail read` already selects (`mail.ts:148`).
- **Fix:** `$select=id,name,contentType,size,isInline,lastModifiedDateTime`. `[graph-doc]` confirm `$select` on the attachment collection (it is a polymorphic type; `contentBytes` is on `fileAttachment`).
- **Accept:** outbound URL carries the `$select`; output columns unchanged except `contentBytes` gone; README/schema checked.

### MAIL-10 (P3) Folder options take IDs only `[code]`
- `mail folders --parent` (`mail.ts:328,332-333`), `create-folder --parent` (`355,359-360`), `create-rule --move-to` (`754,780`) take raw IDs, unlike `--folder` elsewhere, which goes through `resolveFolder` (`mail.ts:38-73`).
- Trap: `resolveFolder` returns well-known names unchanged (`mail.ts:40`). Fine in a URL path, wrong for `moveToFolder`, which needs a real ID.
- **Fix:** resolve all three through `resolveFolder`; for `--move-to`, convert a well-known name with `GET /me/mailFolders/{wellKnown}?$select=id`. Dry-run keeps the name (`mail.ts:48`).
- **Accept:** `create-rule --move-to Archive` body has `moveToFolder: '<real id>'` from the fake GET; `folders --parent Inbox` requests `/me/mailFolders/inbox/childFolders`; ambiguous names error as `--folder` does. README `399,404`.

## 5. Paging, files, SharePoint

### PAGE-2 (P3) Collection commands that fetch one page only `[graph-doc]`
- Single request today:
  - `planner plans` (`planner.ts:72-78`), `planner my-tasks` (`84-90`), `planner tasks/buckets/bucket-tasks` via `get()` (`61-69,92,95,97`)
  - `excel worksheets` (`excel.ts:33-39`)
  - `drive permissions` (`drive.ts:191-197`)
  - `groups owners` (`groups.ts:79-86`; `groups members` already pages, `69-76`)
  - `teams list/members/channels` (`teams.ts:22-59`)
  - `calendar rooms/room-lists` (beta, `calendar.ts:309-325`)
- `mail rules` sends `$top` only when passed (`mail.ts:732-735`); Graph docs don't list `$top` for `messageRules`.
- **Fix:** per endpoint, check the Graph docs for `$top` and `@odata.nextLink`. Where `nextLink` is possible, use `listRequest` with `--all`/`--page-limit`/`--page-delay` (following `nextLink` is safe even if Graph never sends one). Add `--top` only where `$top` is documented; otherwise expose a `paginationOptions` variant without `--top`. For `mail rules`: drop `--top` if undocumented.
- **Accept:** each command with `--all` follows a two-page fake response and prints one combined JSON array; commands without documented `$top` reject `--top` as unknown option (exit 1). Record the doc link per endpoint in the commit/CHANGELOG. README `548-551,732`.

### DRIVE-3 (P3) `drive download` takes a redundant parent ID `[code]`
- `drive.ts:94,97`: `download <driveId> <itemId> <childId>` → `/items/{itemId}/children/{childId}/content`; `/drives/{d}/items/{childId}/content` is enough.
- **Fix (pending §8 Q6, breaking):** `download <driveId> <itemId>` → `/drives/{d}/items/{i}/content`; remove the 3-argument form (clean cutover).
- **Accept:** 2-arg form downloads byte-identical via fake Graph; 3-arg form exits 1 (`too many arguments`). README `482,942`, CHANGELOG "Breaking", capabilities unchanged (`drive.read`).

### SP-3 (P3) `sp item` calls an undocumented endpoint `[graph-doc]`
- `sharepoint.ts:124-131`: `GET /sites/{id}/items/{id}`; not in Graph v1.0 docs (same class as the removed `sp items`, SP-1).
- **Fix (pending §8 Q7):** remove the command; point users at `sp list-items`/`drive` equivalents.
- **Accept:** `sp item` exits 1 (unknown command), like the existing `sp items is gone` test; remove from `src/serve/capabilities.ts:101`, README `590`, CHANGELOG `16`; counts 175 → 174, sharepoint 11 → 10 (README `61,1027-1045`).

## 6. Hygiene and metadata

### CLEAN-1 (P3) Duplicate option parsers and JSON field helpers `[code]`
- Integer parsers: `list.ts:14-22` `integer(min)` (private), `search.ts:38-43` `size()`, inline `--sequence` parser (`mail.ts:760-763`), `serve.ts:19-29` `port()`/`seconds()`.
- Field narrowing: `drive.ts:54-56` `field()`, `graph-client.ts:516-520` `stringField()`, inline at `planner.ts:29`; similar `values()`/`collectionItems` in `workflow.ts:85`, `graph-client.ts:511-513`, `list.ts:80`, `sharepoint.ts:148`.
- **Fix:** export `integer(min, max?)` from `src/list.ts` (or a new `src/options.ts`) and use it for all five; one `field()`/`stringField()`/`items()` set in `src/graph-helpers.ts`, replacing the local copies. Behaviour-preserving.
- **Accept:** existing tests pass unchanged; error messages for bad integers identical across commands (one parametrized test over `--top`, `search --size`, `serve --port`, `create-rule --sequence`, `free-busy --interval`).

### META-2 (P3) Version hard-coded in four places `[code]`
- `package.json:3` is the source; copies at `src/index.ts:38`, `src/remote.ts:49`, `src/commands/generate-skills.ts:70,105` (SKILL.md front-matter).
- **Fix:** `src/version.ts` exporting `VERSION` from `../package.json` (`resolveJsonModule` already on; tsup inlines it [INFERENCE]); use it in all four places.
- **Accept:** temporarily bump `package.json` version in a test fixture build or assert `ms365 --version`, `ms365-remote --version` and a generated SKILL.md all equal `package.json`'s version.

### SCHEMA-2 (P3) `ms365 schema` covers 35 of 175 commands `[code]`
- `src/commands/schema.ts:28-378` has 35 hand-written entries; the command tree has 175 group leaves (+ schema, generate-skills, serve). Only guard: `generate-skills.test.ts:54-58` checks hints point at existing keys.
- **Fix (pending §8 Q8):** generate the catalog from the command tree (method/path come from a per-command annotation on the Commander command, e.g. `.endpoint('GET', '/me/messages')` stored as metadata), keep hand-written parameter/response detail where it exists; `schema --list` then covers every leaf.
- **Accept:** a test that every leaf has a schema entry and every entry is a leaf; `ms365 schema mail.send` output unchanged for the existing 35.

## 7. Remote access

### REMOTE-4 (P3) `ms365 serve` operability `[code]`
- `-q` hides lifecycle lines: `server.ts:410,418` (listening), `429` (received signal), `439` (stopped) go through `info()`, which `index.ts:86` silences. The TCP warning (`412-416`) uses `console.error` and stays.
- Limits are constants: `serve.ts:63-65` (`maxStdoutBytes` 50 MiB, `maxFileBytes`, `concurrency: 4`); `MAX_FILE_BYTES` (`server.ts:17`) also sizes the request body (`19`), output collection (`164-166`) and the 413 message (`88`).
- **Fix:** lifecycle lines always to stderr (service log, not a confirmation); add `--concurrency <n>` (1-64), `--max-stdout-mb <n>`, `--max-file-mb <n>` with CLEAN-1's integer parser; thread the file cap into `server.ts` instead of the module constant.
- **Accept:** `ms365 -q serve --socket …` still prints `listening on`; `--concurrency 1` serializes two slow fake requests; `--max-file-mb 1` makes a 2 MB upload return 413 with the 1 MB figure. README `904,959`.

### REMOTE-5 (P3) Transfers buffered in memory, capped at 25 MB `[code]`
- Upload: `readBody` concatenates the whole body (`server.ts:75-86`), parsed with `new Request(…).formData()` (`114`).
- Download: `collectOutputs` reads each file and base64-encodes it into the JSON reply (`server.ts:161-173,205`); `remote.ts:213` decodes in memory. A 25 MB file is ~34 MB of JSON. README `957` documents `files: [{name, contentType, size, base64}]`.
- **Fix (pending §8 Q9, wire-protocol change):** downloads: `/v1/execute` returns `files: [{name, contentType, size, url}]` where `url` is a one-time `GET /v1/files/{request_id}/{name}` streamed from the child's temp dir, deleted after first read or a TTL; uploads: stream the multipart body to a temp file with a size cap instead of buffering. Keep capability checks on the file endpoint (same caller identity as the execute request).
- **Accept:** a 100 MB fake download round-trips byte-identical through `ms365-remote` with server RSS growth well under the file size; the file URL is single-use and refused for a different caller; upload over the cap → 413 without buffering the body. README/CHANGELOG (protocol version note).

### REMOTE-2 later: Tailscale Service (`svc:ms365`) — opportunity #1
Unchanged from hardening spec §7. Blocked on a live test that `Tailscale-App-Capabilities` reaches `svc:` grants. Not built in this run.

### REMOTE-3 later: approval queue and per-agent rate limits — opportunity #2
Unchanged from hardening spec §7 (`src/serve/capabilities.ts:13` grants are all-or-nothing; only global `executor.ts:60` concurrency). Not built in this run.

## 8. Open questions

| ID | Question | Options | Recommendation |
|---|---|---|---|
| Q1 | DRY-1: what do write commands print under `--dry-run`? | nothing beyond the preview / `Would …` line | **Nothing**: the `[dry-run]` preview already says what would happen |
| Q2 | CAL-4: calendar `--timezone` default | local zone everywhere / UTC everywhere / keep split | **Local everywhere** (matches reads and decline) |
| Q3 | TZ-1: To Do / mail-flag date time zone | local zone / keep UTC | **Local zone** |
| Q4 | CHOICE-1: `excel create-chart --type` | full Graph enum as choices / leave free text | **Full enum** from Graph docs `[graph-doc]` |
| Q5 | USER-4: accepted photo types | jpeg/png/gif only / any `image/*` | **jpeg/png/gif** (documented set) |
| Q6 | DRIVE-3: change `drive download` arguments (breaking) | 2-arg cutover / keep 3-arg | **2-arg cutover**, CHANGELOG "Breaking" |
| Q7 | SP-3: `sp item` | remove / keep undocumented | **Remove** |
| Q8 | SCHEMA-2: catalog source | per-command annotation / generated from code paths / test-only guard | **Per-command annotation** |
| Q9 | REMOTE-5: transfer design | one-time file URLs + streamed upload / chunked binary in execute reply / keep 25 MB base64 | **One-time file URLs** |
| Q10 | REMOTE-2/3 in scope? | build now / keep later | **Keep later** (needs live Tailscale test) |

## 9. Sequencing

1. CLEAN-1 (shared `integer()` + field helpers) — CAL-3 and REMOTE-4 use it.
2. DRY-1 helper, then in parallel by file ownership: mail (DRY-1 mail lines, MAIL-9, MAIL-10, TZ-1 flag) · cal/wf/todo (DRY-1, DRY-2, CAL-3, CAL-4, TZ-1) · files (DRIVE-3, SP-3, PAGE-2 planner/excel/drive) · org (OUT-1, USER-4, PAGE-2 teams/groups).
3. CHOICE-1 after the slices (touches `command-tree.ts` and several command files).
4. META-2, SCHEMA-2.
5. REMOTE-4, then REMOTE-5.
6. README/CHANGELOG, full gate, built-binary smoke (hardening `tasks/build-loop.md` steps 1-4 plus: `--dry-run mail send` prints no success line; `chat send -o json` with fake Graph prints the message).

## 10. Non-goals
- New Graph features.
- A `--yes` flag for destructive operations (agent-safety stance unchanged).
- Live-tenant verification beyond reading Microsoft docs for `[graph-doc]` items.
