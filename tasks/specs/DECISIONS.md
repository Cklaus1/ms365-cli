# Follow-up build: Step 2 decisions (all 5 specs)

Applies to tasks/specs/{dry-run-fidelity, query-pagination, robustness, validation-consistency, performance-dx}-spec.md.

**Context:** the user has no legacy consumers (D27–D30), so output and flag changes are reversible engineering calls, not product decisions. Every open question is resolved below or tagged needs-human. **⚠** marks a close call.

## Engineering decisions (canon)

| # | Source | Question → decision | Why |
|---|---|---|---|
| F1 | DR OQ1 (was needs-human), X-9 conflict | Dry-run trace lines → **stderr**, and `--quiet` does **not** suppress them. Stdout keeps the command's `-o` output (a placeholder with the body redacted). X-9's `info()` routing is rejected. | Stdout stays parseable (jq, MCP structuredContent, `--fields`). The trace is the point of a dry run. No legacy consumers. |
| F2 | DR OQ2 | `success()` helper is silent under dry-run; `handleErrors` prints one trailer: `[dry-run] N request(s) previewed; nothing was sent.` | One convention that future commands can't forget. |
| F3 | DR OQ3 | Redact only credential-like keys (DR-3); truncate `contentBytes` (DR-4); bodies stay visible; no `--show-secrets`. | Showing the body is the purpose of a preview. |
| F4 ⚠ | DR OQ4 (was needs-human) | Under `--dry-run`, `confirm()` resolves without prompting and prints a `would ask:` line. This is accepted **only together with DR-22**: offline/local commands (auth profile delete, logout, watch --reset) don't act under dry-run. | Nothing is sent, so nothing is bypassed. DR-22 closes the local-side-effect hole the reviewer found. |
| F5 | DR OQ5 | `cal update` prints the zone GET and uses a `<from GET>` marker. | Truthful about what the real run sends. |
| F6 | DR OQ6 | Read-aggregate workflows print the trace only; `-o json` gives `{"dryRun":true}`. | No invented reports. |
| F7 | X OQ-1 | Build X-1 (lazy MSAL) only, then re-measure (XT4). X-4/X-5 (warm server) are needs-human, see N1. OQ-2/OQ-3 are moot this run. | The gate needs a logged-in measurement. |
| F8 | X OQ-4, OQ-8 (were needs-human) | `drive copy <driveId> <itemId> --to-parent <id> [--name] [--wait]` (ID-based); keep both `cp` (upload) and `copy` (server-side), cross-referenced in help. | Consistent with `drive files`/`delete`. Path form can follow later. |
| F9 | X OQ-6 | No `id` column in tables; the table footer points to `--fields id,…`. X-13 uses `$expand`. | IDs truncate to 50 characters. |
| F10 | X OQ-7 | MCP subscriptions `--reset` on the first subscribe of each server session. | Avoids replaying days of changes. |
| F11 | Q OQ-1 | `mail-folder` resolution is exact, case-insensitive only (a `match: 'exact'` flag on TypeSpec). | A prefix match must never pick a `move` destination. |
| F12 | Q OQ-2 (was needs-human) | Path syntax `Inbox/Projects`, plus a breadth-first search fallback on a top-level miss (cap 500 folders; two matches → `ambiguous_name`). | Predictable for agents; reversible. |
| F13 | Q OQ-3 | Unpaged endpoints use `printList` anyway. | One envelope and one set of flags (D27). |
| F14 | Q OQ-4 | Pre-flight uses CNF with per-option scopes (any-of inside each clause). | Keeps false positives near zero. |
| F15 | Q OQ-5 (was needs-human) | No change: login keeps `User.Read.All`; only the catalog changes (Q-18). | Status quo, so no new exposure. |
| F16 | Q OQ-6 | Implement directory `$search` `\"` escaping as documented, marked †. | D24. |
| F17 | R OQ-1 (was needs-human) | Pending claims record pid, host and start time. A dead or stale owner reports `outcome_unknown` (exit 1). `ms365 idempotency list` and `ms365 idempotency release <key>` are added. Never auto-release. | Auto-release would defeat D23. The conservative option adds no risk. |
| F18 | R OQ-2 | `reply`/`forward --html` go through createReply/createForward → patch the body → send (reusing sendViaDraft). | Keeps the quoted thread. |
| F19 | R OQ-3 (was needs-human) | `--idempotency-key` also on the 4 Teams send commands; `idempotentSend` moves to `idempotency.ts` (in RT-5). | Same retry pattern as mail; the change is additive. |
| F20 | R OQ-4 | Public `GraphClient.uploadSession(...)` shared by drive and mail. | Read-only, dry-run and audit handling stay in one place. |
| F21 | R OQ-5 | `send()` scales the timeout with body size. | Also covers future large POSTs. |
| F22 ⚠ | R OQ-6 | A SIGINT that cuts a `watch` pass short exits **130**; D2 gets a footnote. SIGINT during the sleep still exits 0. | Shell convention; "interrupted" stays distinguishable from a generic failure. |
| F23 | V Q1 | `--dry-run` works without a client ID (lazy auth), owned by DR-T2. | Preview before setup. |
| F24 | V Q2 | `set-timezone` accepts an `Intl`-valid IANA name or a Windows-shaped name. | Catches typos without rejecting names Outlook shows. |
| F25 | V Q3 | `set-presence` validates the documented availability/activity pairs, marked †. | D24. |
| F26 | V Q4 | The Premium hint shows on 403 and 404 with "If this is a Premium plan…" wording. | Which status Graph returns is undocumented. |
| F27 | V Q5 | The setup wizard accepts piped answers through one line reader; EOF → exit 2. | Scriptable and testable. |
| F28 | V Q6 | `planner --percent` stays 0/50/100. | Current contract. |
| F29 | V Q7 | `auth status` "Logged in as…" stays on stdout (it's data). | It's the command's answer. |
| F30 | Cross-spec | One owner per overlap (from the Step 1 reviews): <br>• DR-16 ≡ V-3 ≡ X-3 → DR-T2 (XT1 lands first) <br>• DR-15 ≡ V-2 → DR-T2 passes the resolved client ID into GraphClient (the V-2 getter is rejected) <br>• DR-2 ≡ V-1 → DR-T4 <br>• DR-6 ≡ X-9 → DR-T1 <br>• DR-14 → Q-15/QP-8 (`resolveFolderName` is deleted) <br>• V-18 split per the V spec (mail:529 → Q-15; mail:768 → DR-13; workflow:229 → DR-10) <br>• V-22 → Q-9/QP-2 <br>• V-23 → VT2 <br>• mail.ts:80 sniff → RT-2 <br>• R-10 `--max-failures` is hidden from MCP (XT5) <br>• all snapshot/docs regeneration happens once, in the FIN task | Each change lands once. |

## needs-human (excluded from this run)

| # | Item | Why | Recommendation |
|---|---|---|---|
| N1 | X-4/X-5 warm in-process MCP server (XT10, XT11; X-2/XT4 gate) | Needs a logged-in measurement of per-call token cost (gate: more than 50 ms per call) | Run `ms365 auth login`, then the XT4 measurement; build only if the gate passes. |
| N2 | X-12 / X OQ-5 PowerShell `.ps1` shim completion | Needs a Windows machine | Test on Windows; option (a), document the limitation, if it fails. |
| N3 | Live-tenant validation of every † item across the 5 specs | No tenant (refresh token expired) | After `ms365 auth login`, run the snapshot commands for real against a test tenant. |
