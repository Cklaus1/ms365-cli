# DRAFT — Upload, Session and Send Robustness

**Status:** DRAFT — must pass its own Step 1 adversarial review before any build. Not scheduled.

**As of:** 2026-10-14 · **Scope:** `src/**` at `41c41f3` (branch `build/ms365-findings-spec`) · **Theme:** P3 opportunities from the last build run, plus related instances found nearby (`[ADDED]`). Format and severities follow `docs/ms365-findings-spec.md`. There are no legacy consumers (D27–D30), so no compat shims.

## Summary

Drive uploads go through one session uploader that retries and resumes (`GraphClient.uploadInSession`). Mail attachment uploads use a second, hand-rolled loop that has no retry, no resume, no timeout and no audit line. The idempotency ledger (D23) has two gaps. A crash leaves a `pending` claim that blocks the key for 24 h. A timeout whose outcome is unknown drops the claim, so the retry the key exists for sends the mail again. Smaller issues: `mail reply --html` is ignored, `shared-send` and the Teams sends have no `--idempotency-key`, one transient error ends `watch --follow`, and `audit.log` rotation can lose the rotated generation when several processes write at once.

## Findings

| ID | Sev | Location | Finding | Failure scenario | Fix |
|---|---|---|---|---|---|
| R-1 | P2 | `src/commands/mail.ts:72-114` | `uploadLargeAttachment` PUTs chunks with a bare `fetch`. It has no retry, no resume from `nextExpectedRanges`, no `AbortController` timeout, and no `audited()` wrapper. `GraphClient.uploadInSession` (`graph-client.ts:374-449`) already has all four. | A 429 or 503 on chunk 7 of a 40 MB attachment fails the whole `mail send --attach`. A stalled connection hangs forever because nothing times out. `audit.log` records the `createUploadSession` POST but none of the upload. | Move the chunk loop into one exported helper in `graph-client.ts`: `uploadToSession(uploadUrl, filePath, size, { chunkSize })`, plus `client.uploadSession(sessionEndpoint, sessionBody, filePath, size, { chunkSize, auditAs })`. Both drive uploads and mail use it. Mail passes `chunkSize = 3 MiB` (Outlook caps chunks at 4 MB †). Accept a 200/201 response with an empty body (Outlook returns 201 with no JSON † — `uploadInSession` currently calls `res.json()` unconditionally, `graph-client.ts:418`). If `querySessionOffset` fails, fall back to the last 202 `nextExpectedRanges` (GET on an Outlook uploadUrl is not documented †). On a final failure, DELETE the uploadUrl on a best-effort basis. Audit the upload as one PUT. |
| R-2 | P2 | `src/commands/mail.ts:777-797` | `mail reply` declares `--html` but never reads `opts.html`. The body always goes out as `comment`, which Graph treats as text †. | `mail reply X --html --body '<b>Approved</b>'` sends literal `<b>Approved</b>`. The agent sees "Reply sent." and has no sign anything went wrong. | With `--html`, POST `createReply`/`createReplyAll` (the draft path `sendViaDraft` already uses). GET the draft body, insert the HTML right after `<body…>`, PATCH the body back, then `/send`. Graph rejects `comment` and `message.body` sent together (400 †), and `message.body` alone drops the quoted thread †. Without `--html` the request stays as it is now. |
| R-3 | P3 | `src/commands/mail.ts:803-823` `[ADDED]` | `mail forward` has no `--html` at all, so `--comment` is always plain text. Teams sends do offer `--html`. | An agent that forwards with a formatted note can't send it as HTML. | Add `--html` and reuse the R-2 path (`createForward` → patch the body → send). |
| R-4 | P3 | `src/commands/mail.ts:957-980` | `mail shared-send` has no `--idempotency-key`. `send`, `send-draft`, `reply` and `forward` all have one (D23). | An agent that retries a timed-out shared-mailbox send cannot avoid sending a duplicate. | Add the `--idempotency-key` option and wrap the POST in `idempotentSend(client, key, 'mail shared-send', { userId, message }, …)`. |
| R-5 | P3 | `src/commands/teams.ts:96,116,212,228` `[ADDED]` | `teams send-channel-message`, `reply-channel-message`, `send` and `reply` are non-idempotent POSTs with no `--idempotency-key`. D23 covers only mail and events. | An agent retrying a timed-out `teams send` posts the message twice in the chat. | Reuse the local ledger (`withIdempotency`). `chatMessage` has no server-side dedupe id †, so the ledger is the only guard. Scope is OQ-3. |
| R-6 | P2 | `src/idempotency.ts:81-89, 108-112` | `withIdempotency` writes a `pending` claim before running `fn`. The claim is removed only in the `catch` (`:117`). SIGKILL, OOM, a `SIGINT` (Ctrl-C), or the MCP runner timeout killing the child all skip that `catch`. The claim then blocks the key for the full `TTL_MS` (24 h, `:28`). | An agent's `mail send --idempotency-key k` is killed mid-upload. Every retry with `k` fails `idempotency_in_progress` for 24 h. The suggested fix ("check whether it took effect before using a new key") gives no way to clear `k`. | Store `{ pid, host, startedAt }` on pending entries. In `checkKey`, a pending entry whose owner process is gone (same `host` and `process.kill(pid, 0)` throws `ESRCH`), or whose `startedAt` is older than `PENDING_STALE_MS` (1 h), is reported as `idempotency_outcome_unknown` instead of `in_progress`. Add `ms365 idempotency list` and `ms365 idempotency release <key>` (local file only, no Graph call). The new error's `fix` names the release command. Never auto-release (OQ-1). |
| R-7 | P2 | `src/idempotency.ts:113-119` with `src/graph-client.ts:194-209` `[ADDED]` | When `fn` throws, the claim is always dropped, including when the request may already have been processed. Examples: a POST timeout ("not retried; it may have been processed"), or a 5xx on the final `/send` / `/sendMail`. | `mail send --idempotency-key k` times out after Graph accepted it. The agent retries with `k` as D23 tells it to. The ledger is empty, so the mail goes out twice. This is the exact case the key exists for. | GraphClient tags errors from requests that may have reached the server: a non-idempotent request that timed out after it was sent, or a 5xx on POST (`CliError.maybeProcessed = true`). `withIdempotency` then keeps the entry as `status: 'unknown'` instead of deleting it. A retry fails with `idempotency_outcome_unknown` (exit 1), and the user clears it with `idempotency release` (R-6). Errors that are definitely not sent (`isPreSendError`, 4xx, usage) still drop the claim. |
| R-8 | P3 | `src/commands/mail.ts:177-181` `[ADDED]` | `sendViaDraft` creates a draft, then attaches, then sends. If an attachment upload fails, the draft is left in Drafts. Each retry adds another one. | Three failed `mail send --attach big.pdf` retries leave three orphan drafts with the full body, and the user may later send one by mistake. | Wrap the attach step. If it fails before `/send` is dispatched, DELETE the draft on a best-effort basis and say so in the error. Never delete after `/send` was dispatched (R-7 ambiguity). |
| R-9 | P3 | `src/graph-client.ts:9,193`; `src/commands/mail.ts:156-161,470,724-736` | The 30 s header timeout on large simple PUTs. | — | `[VERIFIED: already fixed — the simple PUT in uploadFile passes timeoutMs: 120_000 (graph-client.ts:363) and is capped at 4 MB; session chunks use 120 s (:401)]` `[ADDED]` Inline-attachment POSTs still use the 30 s default. These are `sendMail`, `reply`, `forward` and `add-attachment` with up to 3 MB, which is about 4 MB of base64 JSON. They are non-idempotent, so they are not retried. On a 1 Mbit/s uplink the body alone takes about 32 s, so `mail send --attach 2.9MB.pdf` fails with "may have been processed". Fix: in `request()`, when `options.timeoutMs` is unset, default it from the body size: `max(30 s, 30 s + bytes / 32 KiB·s⁻¹)`. |
| R-10 | P2 | `src/commands/watch.ts:181-197, 205-208` | Under `--follow`, any non-expiry error in a pass is rethrown (`:184`) and ends the loop. GETs are retried only `MAX_RETRIES = 3` times with roughly 7 s of backoff in total (`graph-client.ts:6`). | A watcher meant to run for days dies on the first minute-long network outage, a 503 burst, or a refresh failure after sleep or resume. The supervising agent stops getting events and the process exits. | With `--follow`, catch errors per pass, log a warning to stderr, and retry on the next interval with capped backoff (interval × 2ⁿ, at most 15 min). Exit only on `auth_required`, `delta_expired` (with `--strict`), usage errors, or N consecutive failures (`--max-failures`, default 10). Without `--follow`, behaviour stays as it is. |
| R-11 | P3 | `src/commands/watch.ts:101-104,193` | Cursor saved on crash. | — | `[VERIFIED: false positive — the cursor is written atomically (writeFileAtomic) only after a pass completes, so a crash never corrupts it or skips events]` A crash mid-pass re-emits that pass's changes on restart. That is at-least-once delivery, and neither `--help` nor `docs/` says so. Fix: document it, and say consumers should dedupe on `id` plus `lastModifiedDateTime`. |
| R-12 | P3 | `src/commands/watch.ts:203-211` `[ADDED]` | With `--follow`, the SIGINT handler only aborts the sleep. Ctrl-C during a long pass (many pages) does nothing until the pass ends, and a second Ctrl-C is swallowed too. | A user presses Ctrl-C on an initial `--include-existing` sync of a large drive and the process keeps running. | On the first SIGINT, stop after the current page without saving the cursor. On the second SIGINT, `process.exit(130)`. |
| R-13 | P3 | `src/audit.ts:51-54` | Rotation is stat, then `renameSync(file, file.1)`, with no lock. When two processes cross 5 MB together, A renames the log, appends one line to a fresh `audit.log`, and then B's rename moves that 1-line file over `audit.log.1`. Appends themselves are safe: `appendFileSync` opens with `O_APPEND` on every call. | During parallel agent runs, the whole 5 MB rotated generation is silently replaced by a single line. | Rotate inside `withFileLock(`${file}.rotate`)` and re-stat after taking the lock. Keep the lock-free fast path when the stat shows the log is under the limit. |

† depends on documented Graph behaviour and has not been run against a live tenant.

## Open questions

- **OQ-1 (needs-human): recovering from a stale or unknown claim (R-6, R-7).**
  - (a) Shorter pending TTL with auto-release. Simple, but it brings back duplicates.
  - (b) Owner-pid check that auto-releases dead claims. A dead owner says nothing about whether the send took effect.
  - (c) Pid check + staleness that report `outcome_unknown`, with an explicit `ms365 idempotency release <key>`.
  - **Recommend (c).** Auto-release defeats D23. The cost is a one-line manual step, and only after a crash.
- **OQ-2 (engineering): how `--html` reaches a reply (R-2).**
  - (a) `message.body` on `/reply`. One call, but it loses the quoted thread †.
  - (b) createReply → patch the body → send. Three calls, keeps the thread.
  - **Recommend (b).** It reuses `sendViaDraft`.
- **OQ-3 (needs-human): extend D23 to Teams sends (R-5)?**
  - (a) Mail and events only, as today.
  - (b) Add the key to the 4 Teams send commands.
  - **Recommend (b).** The ledger and the agent retry pattern are the same.
- **OQ-4 (engineering): the shared uploader API (R-1).** Expose a public `client.uploadSession(...)`, or keep `uploadLargeAttachment` in mail calling an exported free function? **Recommend a public `GraphClient` method.** It keeps read-only, dry-run and audit handling in one place.
- **OQ-5 (engineering): body-size timeout (R-9).**
  - (a) Scale the timeout in `request()`.
  - (b) Pass `timeoutMs` at the 4 call sites.
  - **Recommend (a).** It also covers future large POSTs.

## Acceptance criteria

1. **R-1:**
   - A mocked Outlook session that returns 503 on chunk 2, then 202 with `nextExpectedRanges`, completes, and every byte is sent exactly once after the resume point.
   - A 201 with an empty body succeeds.
   - A chunk that stalls past the timeout aborts and retries.
   - `audit.log` gets one PUT line for the attachment upload.
   - The drive upload tests in `test/upload.test.ts` still pass, unchanged.
2. **R-2/R-3:**
   - `mail reply --html` issues createReply → GET → PATCH (`contentType: 'HTML'`, user HTML before the quoted body) → send.
   - Without `--html`, it issues a single `/reply` with `comment`.
   - `forward --html` works the same way.
3. **R-4/R-5:** A second `shared-send` (and each Teams send) with the same key and request makes no Graph POST and prints the stored result. A different request with the same key fails with `idempotency_conflict`.
4. **R-6:**
   - A pending entry with a dead pid on the same host, or older than 1 h, makes `checkKey` throw `idempotency_outcome_unknown`. A live pid gives `idempotency_in_progress`.
   - `ms365 idempotency release k` removes the entry, and the next send proceeds.
   - `idempotency list` prints the entries as JSON.
5. **R-7:**
   - A POST to `/me/sendMail` that times out leaves an `unknown` entry, and the retry with the same key does not POST.
   - A pre-send connection error or a 400 still drops the claim.
6. **R-8:** A failed attachment upload in `sendViaDraft` triggers one DELETE of the draft. A failure at `/send` triggers no DELETE.
7. **R-9:** `request()` with a 4 MB body and no `timeoutMs` arms a timer of at least 150 s, checked with fake timers.
8. **R-10:**
   - With `--follow --max-passes 3`, a pass that throws 503 is logged and followed by another pass.
   - `auth_required` exits immediately.
   - 10 consecutive failures exit 1.
9. **R-12:** A first SIGINT during a pass exits 130 after the current page, and the cursor file is unchanged.
10. **R-13:** Two rotations interleaved under the test harness never replace a `.1` that is over 5 MB with a smaller file.
11. Full suite green. The `generate-skills`, `describe` and completion snapshots are updated for the new options and commands.

## Suggested task breakdown

| Task | Scope (files) | Covers | Depends on |
|---|---|---|---|
| RT-1 Shared session uploader | `src/graph-client.ts`, `test/upload.test.ts` | R-1 (core: empty-201, chunkSize, DELETE on failure) | — |
| RT-2 Mail uses the shared uploader | `src/commands/mail.ts`, `test/commands/mail*.test.ts` | R-1 | RT-1 |
| RT-3 Body-size timeout | `src/graph-client.ts`, `test/graph-client-core.test.ts` | R-9 | — |
| RT-4 `maybeProcessed` tagging | `src/errors.ts`, `src/graph-client.ts` | R-7 (signal) | — |
| RT-5 Ledger owner, unknown state, release | `src/idempotency.ts`, new `src/commands/idempotency.ts`, `src/index.ts`, `test/idempotency.test.ts` | R-6, R-7 | RT-4, OQ-1 |
| RT-6 Reply/forward `--html` | `src/commands/mail.ts`, tests | R-2, R-3 | OQ-2 |
| RT-7 Keys on shared-send and Teams | `src/commands/mail.ts`, `src/commands/teams.ts`, tests | R-4, R-5 | RT-5, OQ-3 |
| RT-8 Orphan-draft cleanup | `src/commands/mail.ts`, tests | R-8 | RT-2, RT-4 |
| RT-9 Resilient `watch --follow` + SIGINT + docs | `src/commands/watch.ts`, `test/commands/watch*.test.ts`, docs | R-10, R-11, R-12 | — |
| RT-10 Locked audit rotation | `src/audit.ts`, `test/audit.test.ts` | R-13 | — |
| RT-11 Snapshots and generated docs | `test/__snapshots__`, `docs/` | AC 11 | RT-5, RT-6, RT-7, RT-9 |
