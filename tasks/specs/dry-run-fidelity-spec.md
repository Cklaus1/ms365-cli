# DRAFT — Dry-run fidelity: truthful, crash-free, offline, no secrets

**Status:** DRAFT — must pass its own Step 1 adversarial review before any build. Not scheduled.

## Summary

`--dry-run` is the main way an agent previews a write. Right now it is wrong in four ways. It prints success lines for actions that never happened ("Email sent successfully."). It crashes or exits 1 in multi-step commands because the placeholder response `{dryRun,method,url}` gets read as real data. It needs a client ID, so it does not work fully offline. It prints request bodies, including the `drive share --password-stdin` password, to stdout, mixed in with the command's JSON. This spec makes dry-run output honest and parseable, and makes it exit 0 for every valid invocation. Evidence comes from `test/__snapshots__/snapshot.test.ts.snap`, plus a sweep of all 191 leaf commands run under `--dry-run` with every optional value flag set, against commit 41c41f3.

## Findings

| ID | Sev | Location | Finding | Failure scenario | Fix |
|---|---|---|---|---|---|
| DR-1 | P2 | `utils.ts:24` `info()`; callers: mail.ts:477,487,500,672,751,798,824,837,849,873,885,897,912,978,1119; calendar.ts:251,293,324,334,352; teams.ts:112,136,224,257; todo.ts:138,232; contacts.ts:118; drive.ts:140; notes.ts:82; excel.ts:127; planner.ts:277,294,312; user.ts:152,215; workflow.ts:427 | Post-write confirmations print whether or not anything was sent. The sweep confirmed 26 of them under `--dry-run`, and the delete lines also print after a TTY "y". | An agent runs `--dry-run mail send`, sees "Email sent successfully.", and reports the mail as sent. | Add `success(msg)` in utils. It is a no-op under dry-run, and a single trailer replaces it (OQ2). Route every caller listed here through it. |
| DR-2 | P3 `[ADDED]` | `user.ts:138,238,250` | `set-presence`, `set-timezone` and `set-language` print their confirmation with `console.log`, so it goes to **stdout**, unlike every other command (stderr through `info`). It also prints under dry-run. | `ms365 user set-timezone X -o json \| jq` gets non-JSON. Dry-run says "Timezone set to X." | Use `success()` (DR-1). |
| DR-3 | P2 | `graph-client.ts:166-169` (request), `:478` (batch) | Dry-run prints `options.body` as-is. `drive share --password-stdin` (drive.ts:209-210) puts `"password":"s3cret"` on stdout. Confirmed by running it. | A password read from stdin specifically to keep it out of argv shows up in terminal scrollback, CI logs and agent transcripts. | Pass the JSON body through `redactForPreview()`, which replaces values of keys matching `/^(password\|.*secret\|.*token\|authorization\|cookie)$/i` (case-insensitive, recursive) with `"<redacted>"`. Apply the same to `Headers:` keys. If a body is not JSON, print it unchanged. |
| DR-4 | P3 `[ADDED]` | `graph-client.ts:168`; inline attachments from mail.ts `inlineAttachment` (send/reply/forward/draft) | `contentBytes` (base64 file contents) prints in full. A 3 MB inline attachment is a 4 MB stdout line. | The preview floods the terminal or agent context, and file contents leak into logs. | `redactForPreview` replaces each `contentBytes` with `"<base64, N bytes>"`. |
| DR-5 | — | `graph-client.ts:165-178` | Authorization header leak in dry-run. `[VERIFIED: false positive — the dry-run branch returns before getToken()/the Authorization header is built (graph-client.ts:178+); options.headers only carries Prefer/If-Match/Content-Type. DR-3's header-key redaction is kept as defence in depth.]` | — | — |
| DR-6 | P2 | `graph-client.ts:166-173,348-353,477-480`; `mcp-server.ts:235` | `[dry-run]` trace lines go to stdout ahead of the command's `-o` output. | `--dry-run cal create … -o json \| jq` fails to parse. Under `mcp serve --dry-run`, `parseJson(r.stdout)` returns undefined, so the tool result loses its structured content. | Send the trace lines to stderr (OQ1). The stdout placeholder gains a redacted `body` field, so the preview remains machine-readable. |
| DR-7 | P2 | `graph-client.ts:555-557` (`requestPagesStreaming`); crash at `workflow.ts:345` | Under dry-run the first page has no `value`, so the placeholder is yielded **as a collection item**. `workflow digest` then calls `parseGraphDateTime(undefined)`. | `--dry-run workflow digest` exits 1 with "Cannot read properties of undefined (reading 'dateTime')". | If `first.status === 0` (dry-run), yield `[]` and return. |
| DR-8 | P3 `[ADDED]` | same root, used through `formatter.ts:436-443` | Every `--all` list prints the placeholder as one row. | `--dry-run mail list --all -o table` shows a row `true GET https://…`. That is invented data. | Fixed by DR-7. |
| DR-9 | P2 `[ADDED]` | `workflow.ts:116-124,159` | `meeting-prep --event-id X` treats the placeholder as the event and crashes the same way. The snapshot does not cover this because it only passes mandatory options. | Exit 1 under dry-run. | Return early under dry-run (OQ6). |
| DR-10 | P2 | `workflow.ts:222-229` | `email-to-task` without `--list` gets no lists back from the placeholder and throws "No To Do lists found." | Exit 1 under dry-run, before the POST it was meant to preview. | Under dry-run, use the placeholder list ID `<first-list>` and still preview the POST. `resolve.ts:117` is the precedent. |
| DR-11 | P3 `[ADDED]` | `workflow.ts:243,401-416,427` | When a body is built from a preceding read, it fills in `undefined` or empty strings: "Date: undefined", `<b></b>`, and "Announced \"undefined\"". | The previewed body does not match what the real run would send, and gives no sign of it. | Under dry-run, values that come from a read use the marker `<from GET>`. The DR-1 trailer says "bodies use placeholders for values read from Graph". |
| DR-12 | P2 `[ADDED]` | `workflow.ts` standup (:68-101), focus-time (:452-547), meeting-prep (:147), digest | Read-aggregate workflows print a made-up report with exit 0: "Unread Emails: 0", "Available focus: 2h 58m", "No upcoming meetings". | The agent takes invented numbers as real. | Under dry-run, print the trace only. Text mode prints no report. `-o json` prints `{"dryRun":true}` (OQ6). |
| DR-13 | P2 | `mail.ts:759-768` | `download-attachment` reads `contentBytes` from the placeholder and throws "No content found". | Exit 1 under dry-run. | Under dry-run, print the placeholder and return, as `drive.ts:17-20` and `user.ts:186` already do. Write no file. |
| DR-14 | P3 | `mail.ts:39-52,527-529` (`resolveFolderName`) | The name lookup finds nothing in the placeholder, so `mail folder NAME` says "not found". | Exit 1 with a false "not found" (the snapshot shows this). | Under dry-run, a lookup with no listing returns the name unchanged, as `resolve.ts:117` does. |
| DR-15 | P2 | `user.ts:127-128,146-147` | `set-presence` and `clear-presence` read the client ID only from `process.env`. They ignore `--profile`'s `clientId` (index.ts:119) and fail under dry-run when the env var is unset. | Exit 1 in the snapshot. Real runs with a profile-only client ID also fail. | Expose `AuthManager.clientId` and read the ID from it (profile → env). Under dry-run with no client ID, use `<client-id>`. |
| DR-16 | P2 `[ADDED]` | `index.ts:116-129` → `auth.ts:259-266` | `getClient()` always constructs an `AuthManager`, and its constructor throws without a client ID. **Every** Graph command under `--dry-run` therefore exits 3 on a machine with no setup. Confirmed with `env -u MS365_CLI_CLIENT_ID`. | An agent can't preview anything in a fresh sandbox or CI. | Under dry-run, give `GraphClient` a lazy auth factory that is never built, because dry-run never calls `getToken()`. |
| DR-17 | P3 | `calendar.ts:220-223` | `cal update --start/--end` without `--timezone` previews the zone-read GET. **Decision: printing the GET is correct**, because the real run sends it. The PATCH body, though, falls back to the local zone because the placeholder carries no `originalStartTimeZone`. | The preview shows `America/New_York` while the real run would use the event's own zone. | Keep the GET. Mark the zone with `<from GET>` (DR-11). See OQ5. |
| DR-18 | P2 `[ADDED]` | `watch.ts:165-166` | `watch --reset` deletes the saved cursor file before any request, even under dry-run. | `--dry-run watch mail --reset` destroys local state that a preview should leave alone. | Skip `removeCursor` under dry-run. `saveCursor` is already skipped. |
| DR-19 | P2 `[ADDED]` | `utils.ts:35` `confirm()`; every delete | Under dry-run, deletes still demand a TTY confirmation and exit 7 without one. That covers 11 commands in the sweep. | An agent cannot preview a delete, the case where a preview matters most (findings-spec PM-7). | Under dry-run, skip the prompt and print `[dry-run] would ask: <message>` on stderr (OQ4). |
| DR-20 | P3 `[ADDED]` | planner.ts:29; mail.ts:80,178,1118; meetings.ts:17; calendar.ts:416; content.ts:282; drive.ts:17; user.ts:186; teams.ts:295 | Ten separate ways of sniffing `data.dryRun`. Each new multi-step command has to reinvent one, which is how DR-7 to DR-14 happened. | — | `GraphClient.isPreview(res)` (status 0), plus a lint-style test that bans `.dryRun` property reads outside graph-client.ts. |
| DR-21 | P3 `[ADDED]` | `test/snapshot.test.ts:92-97` | The harness records only lines that start with `Error:`, so commander's lowercase `error:` shows `"error": ""` for `watch RESOURCE`. It passes only mandatory options, which is how DR-9 was missed. It reads requests from stdout, which DR-6 changes. | Regressions in optional flags under dry-run go unseen. | Record stderr in full. Add a second pass that sets every optional value flag (as this sweep did) and asserts exit ∈ {0,2}. Read the trace from stderr. |

### Snapshot evidence: every non-zero exit or error under `--dry-run`

| Command (synthetic argv) | Exit | Classification |
|---|---|---|
| chat create --members x | 2 | Legit: invalid email in the synthetic args |
| excel add-rows / excel format-range / mail add-attachment / notes update | 2 | Legit: a required alternative flag (one of several) was not passed |
| mail create-rule --name x | 2 | Legit: the no-conditions safety guard |
| meetings all-transcripts | 2 | Legit: needs `--api-version beta` |
| planner bucket, bucket-tasks, delete-bucket, update-bucket | 2 | Legit: a bucket name needs `--plan`, and "BUCKETID" is not 28 characters long |
| watch RESOURCE | 2 | Legit: invalid choice. The empty error field is a harness gap (DR-21) |
| mail download-attachment | 1 | **Bug**: DR-13 |
| mail folder NAME | 1 | **Bug**: DR-14 |
| user set-presence, user clear-presence | 1 | **Bug**: DR-15 |
| workflow digest | 1 | **Bug**: DR-7 |
| workflow email-to-task | 1 | **Bug**: DR-10 |
| *(sweep only)* workflow meeting-prep --event-id x | 1 | **Bug**: DR-9 |
| *(sweep only)* 11 deletes without a TTY | 7 | **Bug by policy**: DR-19 |

The sweep also turned up two usage errors that exit 1 instead of 2: `email-to-task --due` with a bad date, and `planner update-task --percent 5`. They are out of theme and are noted for the exit-code backlog.

## Open questions

- **OQ1 — Which stream gets the dry-run trace? (needs-human)**
  - (a) Keep it on stdout.
  - (b) Send the `[dry-run]` lines to stderr. They are not affected by `--quiet`, because they are the point of a dry run. Stdout keeps the command's `-o` output, which is the placeholder `{dryRun,method,url,body}` with `body` redacted and parsed.
  - (c) Make stdout a single `{"dryRun":true,"requests":[…]}` document and suppress the command's own output.
  - **Recommendation: (b).** Stdout stays valid in every `-o` format, so jq, MCP `structuredContent` and `--fields` keep working. The preview stays machine-readable through the placeholder. (c) is more complete, but it changes the stdout shape for every command at once.
- **OQ2 — How to word success lines under dry-run? (engineering)**
  - (a) A wording per command, e.g. "(dry-run) Email not sent."
  - (b) `success()` stays silent under dry-run, and a `postAction` hook prints one trailer: `[dry-run] N request(s) previewed; nothing was sent.`
  - **Recommendation: (b).** It is one convention, and future commands can't forget it.
- **OQ3 — Redaction scope (engineering).** Should message and post bodies (`content`) be redacted too? **Recommendation: no.** Showing the body is the purpose of a preview, and the user wrote it. Redact the credential-like keys only (DR-3) and truncate `contentBytes` (DR-4). Add no `--show-secrets` escape hatch.
- **OQ4 — Should dry-run skip delete confirmations? (needs-human)** The `confirm()` doc says it "always prompts so agents can't bypass". **Recommendation: skip under dry-run.** Nothing is sent, so nothing is bypassed. The `would ask:` line keeps the prompt visible.
- **OQ5 — How should `cal update` get the zone? (engineering)**
  - (a) Print the GET and use the `<from GET>` marker.
  - (b) Hide reads that are only there to build the write.
  - (c) "Online dry-run": really send the GETs.
  - **Recommendation: (a).** (b) misrepresents what the real run sends. (c) breaks offline and needs auth (DR-16).
- **OQ6 — What should read-aggregate workflows print under dry-run? (engineering)**
  - (a) The trace only. `-o json` gives `{"dryRun":true}`.
  - (b) The report rendered with placeholders.
  - **Recommendation: (a).** A report built from no data is the untruthful output this spec removes.

## Acceptance criteria

1. Every leaf command in the dry-run snapshot, plus a second pass with every optional value flag set, exits 0 or 2 under `--dry-run`. Exit 1 or 3 fails the test, and so does exit 7 once OQ4 is accepted.
2. Under `--dry-run`, stderr contains none of the DR-1/DR-2 confirmation strings. It ends with exactly one trailer line `[dry-run] N request(s) previewed; nothing was sent.`, where N equals the number of trace request lines.
3. `echo s3cret | ms365 --dry-run drive share D I --password-stdin` contains no `s3cret` on stdout or stderr, and does contain `"password":"<redacted>"`. A unit test covers `redactForPreview` on nested keys, `Authorization`/`Cookie` headers and `contentBytes`.
4. Assuming OQ1 (b): `ms365 --dry-run cal create … -o json` stdout passes `JSON.parse`, has `dryRun:true`, and has a `body`. Every `[dry-run]` line is on stderr. An MCP tool call with `--dry-run` returns `structuredContent`.
5. With `MS365_CLI_CLIENT_ID` unset and an empty `HOME`, `ms365 --dry-run mail send …` exits 0 and makes no network call and no file write under the config directory.
6. `--dry-run workflow digest`, `standup` and `focus-time` print no report, exit 0, and in json mode print exactly `{"dryRun":true}`. `meeting-prep --event-id X` and `email-to-task M` (without `--list`) exit 0 and show the trace.
7. `--dry-run mail list --all -o json` prints `{"value":[]}`.
8. `--dry-run watch mail --reset` leaves an existing cursor file byte-for-byte unchanged.
9. `user set-presence` reads the client ID from the profile when `--profile` sets one. Under dry-run with no ID, the body shows `"sessionId":"<client-id>"`.
10. `grep -rn '\.dryRun\b' src --include=*.ts`, excluding graph-client.ts, returns nothing, and a test enforces it.

## Suggested task breakdown

| Task | Scope (files) | Covers | Depends on |
|---|---|---|---|
| DR-T1 | `graph-client.ts`: `isPreview()`, `redactForPreview()`, trace to stderr, enriched placeholder, request counter, `requestPagesStreaming` empty yield | DR-3, DR-4, DR-6, DR-7, DR-8 | OQ1 |
| DR-T2 | `index.ts`, `auth.ts`: lazy auth under dry-run; `AuthManager.clientId` getter | DR-16, half of DR-15 | — |
| DR-T3 | `utils.ts` (`success()`, dry-run flag, `confirm()` skip), `index.ts` postAction trailer | DR-1 (helper), DR-19 | DR-T1 (counter), OQ2, OQ4 |
| DR-T4 | Swap every DR-1/DR-2 caller to `success()`: mail, calendar, teams, todo, contacts, drive, notes, excel, planner, user, workflow | DR-1, DR-2 | DR-T3 |
| DR-T5 | `workflow.ts`: dry-run early returns, the `<from GET>` marker, the placeholder list ID | DR-9 to DR-12 | DR-T1, OQ6 |
| DR-T6 | `mail.ts` (download-attachment, resolveFolderName), `user.ts` presence, `watch.ts` reset, `calendar.ts` zone marker | DR-13, DR-14, DR-15, DR-17, DR-18 | DR-T1, DR-T2 |
| DR-T7 | Replace the ten ad-hoc `.dryRun` sniffs with `isPreview()`, plus the grep test | DR-20 | DR-T1, DR-T5, DR-T6 |
| DR-T8 | `test/snapshot.test.ts`: read stderr, add an optional-flags pass, assert AC1–AC2; regenerate the snapshot | DR-21, AC1–AC10 | DR-T1 to DR-T7 |
