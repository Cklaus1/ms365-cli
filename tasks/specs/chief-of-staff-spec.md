# DRAFT — Chief-of-staff view and aggregate activity reports

**Status:** DRAFT — must pass its own Step 1 adversarial review before any build is accepted. **Implementation exists but is uncommitted** (written ahead of this spec in error; it is held back until this spec is reviewed). Prefix: CS-
**As of:** 2026-10-15 · **Base:** `main` at `c8c9b53` + uncommitted `auth grant-scopes`, `reports`, `chief` work.
**Depends on:** tasks/specs/identity-profiles-spec.md (notice + opt-out, admin-only decline list).

## Summary

Two read-only features that give the CEO a picture of what is happening, without collecting anyone's messages:
1. **`ms365 chief brief`** — a briefing of the *signed-in person's own* work: today's meetings, high-importance and flagged mail, overdue tasks, recent Teams chats. It uses only that person's own delegated sign-in, so it shows nothing they couldn't already open.
2. **`ms365 reports …`** — Microsoft's aggregate usage reports: tenant-level daily counts (email, Teams, active users) and a department roll-up. Counts only: **no message content, no per-person rows, and no department group smaller than 5 people.**

3. **`ms365 audit-trail`** — oversight of the monitoring itself: who accessed a mailbox and how (admins, delegates, apps), which apps hold broad mail access, and which admin roles or consents changed. It reads Microsoft's audit log, so it shows *access events*, never the content that was accessed.

## Design rules (owner decisions and constraints)

- **No bulk collection of content.** ms365 gets no command that reads other people's mail, chats or calendars. Reports are Microsoft-computed counts. (Owner asked for bulk collection on 2026-10-15; declined as mass-surveillance tooling. The owner then asked for "activity collection" and "chief of staff view", which this spec interprets as the two features above.)
- **Small groups are never shown.** Department roll-ups pool groups under the minimum into one "Other" row and drop even that pool if it is under the minimum, so no row describes fewer than 5 people. The floor is not configurable below 5.
- **Decline list honored.** People who declined monitoring (identity spec ID-18/19) are excluded from roll-ups. Roll-up takes an exclusion set; wiring it to the SharePoint decline list is task CST-5 (blocked on IDT-12).
- **Untrusted text.** Mail and chat previews from other people are marked `untrusted: true` and truncated to 200 characters.

## Permissions

| Feature | Delegated scopes (added to `--org-mode` login and the setup wizard) | Also needed |
|---|---|---|
| `chief brief` | `Calendars.Read`, `Mail.Read`, `Tasks.Read`, `Chat.Read` (all already requested) | Nothing |
| `reports email|teams|active-users` | `Reports.Read.All` | Admin consent (use `ms365 auth grant-scopes --org-mode`), **and** the signed-in person must hold an admin role that can read reports (Reports Reader, Global Reader, or Global Administrator) |
| `reports by-department` | `Reports.Read.All`, `User.Read.All` | Same as above; the tenant setting "Display concealed user, group, and site names in all reports" must be **off**, otherwise users can't be matched to departments (the command stops with a clear message) |
| `audit-trail mailbox`, `audit-trail changes` | `AuditLogsQuery.Read.All` (Graph audit log query API) and `AuditLog.Read.All` (directory audit), both added to the ms365 app's scope list (CS-15) † | Admin consent, **and** a Purview audit role for the signed-in person (View-Only Audit Logs or Audit Reader). `MailItemsAccessed` (message reads) exists only with Audit (Premium), included in E5. Without it, only basic events appear. |
| `audit-trail apps` | `Application.Read.All` | Admin consent; reads service principals and their app-role assignments (which apps hold `Mail.Read`-type roles) |

No application permissions and no app-only sign-in are involved.

## Findings

| ID | Sev | Finding | Status | Fix |
|---|---|---|---|---|
| CS-1 | P2 | There is no single view of "what needs my attention" across calendar, mail, tasks and Teams. | Implemented (uncommitted): `src/commands/chief.ts` | `chief brief` runs 6 reads in parallel; one source failing (e.g. no Teams consent) shows "unavailable" for that section only. `--days 1-14`, `--timezone`, `-o text|json|…`. An "attention" list summarises: next meeting, overlapping meetings, unread high-importance mail, flagged mail, overdue tasks. |
| CS-2 | P2 | Tenant-level activity trends aren't available from the CLI. | Implemented (uncommitted): `src/commands/reports.ts` | `reports email|teams|active-users --period D7|D30|D90|D180` call Graph's `getEmailActivityCounts`, `getTeamsUserActivityCounts`, `getOffice365ActiveUserCounts` with `$format=application/json`. |
| CS-3 | P2 | Department-level patterns need per-user data internally, which must never leak. | Implemented (uncommitted) | `reports by-department` reads the per-user detail report and the user directory, sums by department in memory, and prints only department rows. Output never contains a UPN (test asserts this). `--min-group-size` accepts 5 or more only. |
| CS-4 | P3 | JSON field names differ between CSV and JSON report formats, and the live response shape is unverified. | † unverified live | Keys are normalised (lower-case, no spaces) so `Send Count` and `sendCount` both work. |
| CS-5 | P2 | The decline list can't yet exclude people from roll-ups. | Open | `rollUp()` already accepts an exclusion set; load it from the admin-only decline list once IDT-12 lands, and fail closed if it can't be read. |
| CS-6 | P3 | `chief brief` Teams section uses `$expand=lastMessagePreview` with `$orderby` on `/me/chats`. | † unverified live | If Graph rejects it, the section reports "unavailable" and the rest of the briefing is unaffected. |
| CS-7 | P3 | The briefing shows only the signed-in person's data; there is no way to produce one for someone else. | By design | A briefing for another person would need that person's own consent grant (identity spec ID-13) and is out of scope here. |
| CS-8 | P3 | Dry-run must not invent a briefing. | Implemented | Under `--dry-run` the requests are traced and nothing is printed. |
| CS-9 | P2 | Nobody can see who accessed a mailbox through ms365. | Open | `ms365 audit-trail mailbox <upn> [--since <t>] [--operations …] [--by <upn>]`: creates an audit log query (`POST /security/auditLog/queries` filtered to the mailbox and the Exchange workload), polls until it succeeds, and lists the records: time, operation, actor (user or app), logon type (owner / delegate / admin), client IP, count of items. **Never** prints subjects, bodies or recipients. † |
| CS-10 | P2 | Broad app access to mail is invisible. | Open | `ms365 audit-trail apps`: lists service principals holding mail-read application roles (`Mail.Read`, `Mail.ReadBasic.All`, `Mail.ReadWrite`, `Chat.Read.All`, `ChannelMessage.Read.All` …) with who consented and when, plus any Exchange RBAC scoping the tool can see, flagging apps with tenant-wide scope. † |
| CS-11 | P3 | Role and consent changes aren't easy to review. | Open | `ms365 audit-trail changes --since <t>`: directory audit events for admin role assignments, app consents and permission grants (`GET /auditLogs/directoryAudits`, filtered by activity), newest first. |
| CS-12 | P3 | Audit coverage depends on licensing and retention, and a partial log can look like "nobody accessed it". | Open | Each command prints a coverage line to stderr: the earliest date the log can answer for (retention), and whether message-read events (`MailItemsAccessed`) are available; when they aren't, it says "message reads are not logged on this licence" instead of returning an empty list as if it were clean. |
| CS-13 | P3 | Audit queries are asynchronous and can be large. | Open | The command polls with backoff (default timeout 5 minutes, `--timeout`), honors the API's `recordCountLimit` / `isRecordCountLimitExceeded` flags by saying the result was truncated and how to narrow it, streams pages through `printList`, and supports `--dry-run` (traces the query creation, no polling). |
| CS-14 | P3 | The oversight data itself is sensitive. | Open | Audit-trail output goes through the normal audit log (an admin running a mailbox query is recorded), is excluded from MCP tools by default (admin action, like `auth grant-scopes`), and `--dry-run`/`--read-only` behave as elsewhere. Creating a query is a POST, so `--read-only` blocks it with a clear message. |
| CS-15 | P2 | The audit search API can't be reached through the Azure CLI, so `audit-trail` must run on ms365's own app registration. | Verified 2026-10-15 | The `az` Graph token (app "Microsoft Azure CLI") carries `AuditLog.Read.All` (directory audits work) but **not** `AuditLogsQuery.Read.All`, and that can't be added to Microsoft's own app. A mailbox-access query through `az rest` therefore returns 403 (`User:… dont have any permissions`) even after Audit Reader applies. Fix: (1) add `AuditLogsQuery.Read.All` and `AuditLog.Read.All` to `ORG_SCOPES` in `src/auth.ts`, the setup wizard and the catalog; (2) run `ms365 --org-mode auth grant-scopes` (az supplies the admin sign-in for that step only); (3) `audit-trail` calls the API with the person's own ms365 sign-in. |
| CS-16 | P3 | Two separate prerequisites look like one. | Verified 2026-10-15 | The person running `audit-trail` needs **both** (a) the Purview **Audit Reader** role group (View-Only Audit Logs; add via Purview → Settings → Roles and scopes → Role groups) and (b) the Graph `AuditLogsQuery.Read.All` consent for the ms365 app. Missing (a) gives 403 `dont have any permissions`; missing (b) gives a Graph insufficient-privileges error. The command maps each to a distinct hint (exit 4). Purview role changes can take up to an hour or more to apply. As of this date both cklaus@fusen.world and CK-Admin@fusen.world were added to Audit Reader. |

## Open questions

| # | Question | Options | Recommendation | Tag |
|---|---|---|---|---|
| OQ-1 | Smallest group shown in roll-ups | 5 (current), 10 | Keep 5 as the floor; the owner can raise it with `--min-group-size` | decided |
| OQ-2 | Which reports to add next | Microsoft 365 Copilot usage, SharePoint/OneDrive activity, Teams meeting counts, mailbox storage | Add as separate `reports` subcommands after live verification of the current ones | engineering |
| OQ-3 | Should `chief brief` include company activity trends (a one-line summary from the reports)? | (a) no; (b) opt-in `--with-activity` for people who hold a reports role | **(b)**, later | engineering |
| OQ-4 | Should the briefing be pushed (daily Teams message or email to the person) instead of pulled? | (a) pull only; (b) a scheduled `chief brief` plus `chat send`/`mail send` to oneself | **(b)** as a follow-up, using `--idempotency-key` per day | engineering |
| OQ-5 | Should each employee get a periodic "who accessed your mailbox" notice? | (a) no, admins only; (b) monthly Teams message per person | **RESOLVED by the owner (2026-10-15): (a) No.** Access information stays with admins; no per-employee notices. CST-10 is dropped. | decided |
| OQ-6 | Who may run `audit-trail` | (a) any admin with a Purview audit role; (b) a second approver for mailbox queries | **RESOLVED by the owner (2026-10-15): (a) admins.** Any admin holding a Purview audit role; every query is recorded in the audit log. Enforced by Graph itself (the API refuses callers without the role), not by a CLI check. | decided |

## Acceptance criteria

1. `chief brief` with one failing source still prints the other sections, and the failing one says "unavailable: <reason>". Exit code 0.
2. `chief brief --dry-run` prints nothing to stdout, traces every request, exits 0.
3. Mail and chat previews carry `untrusted: true` and never exceed 200 characters.
4. No `reports` output contains a UPN or any per-person value (test over `by-department` with fixtures including small groups).
5. No roll-up row describes fewer than the minimum group size; a pool under the minimum is omitted.
6. `--min-group-size` below 5 exits 2 with no request.
7. Concealed names produce a clear error naming the admin setting (exit 1), not an empty report.
8. A 403 from Graph maps to exit 4; the hint says the signed-in person needs a reports-reading admin role plus admin consent.
9. Full suite, typecheck, `docs:check`, smoke and the dry-run snapshot stay green (the new commands appear in the snapshot).
10. `audit-trail mailbox` output never includes a message subject, body or recipient (test over fixtures that contain them).
11. Without `MailItemsAccessed` coverage, the command says so on stderr and does not present an empty result as "no access".
12. An audit query that exceeds the record limit reports truncation and exits 0; one that times out exits 1 with a "narrow the window" hint; a 403 exits 4 and names the audit role and consent needed.
13. `audit-trail` is hidden from MCP tools, honors `--dry-run`, and is blocked by `--read-only` (exit 4).

## Suggested task breakdown

| Task | Scope | Covers | State |
|---|---|---|---|
| CST-1 | `src/commands/reports.ts`, tests, catalog, scopes, wizard group | CS-2, CS-3, CS-4 | written; tests passing |
| CST-2 | `src/commands/chief.ts`, tests (to write), catalog | CS-1, CS-6, CS-8 | written; **tests not written yet** |
| CST-3 | Live check against the tenant after `grant-scopes --org-mode` adds `Reports.Read.All`; confirm JSON shapes | CS-4, CS-6 | not started |
| CST-4 | Regenerate README and the dry-run snapshot | AC 9 | not started |
| CST-5 | Decline-list exclusion in roll-ups | CS-5 | blocked on identity spec IDT-12 |
| CST-6 | Daily push of the briefing | OQ-4 | later |
| CST-7 | `src/auth.ts` + `src/commands/setup.ts` + `src/catalog.ts` (add `AuditLogsQuery.Read.All`, `AuditLog.Read.All`; CS-15); distinct 403 hints for the two prerequisites (CS-16); `src/commands/audit-trail.ts`: query create/poll/records, `mailbox` and `changes`, coverage line, truncation handling; tests with mocked async flow | CS-9, CS-11, CS-12, CS-13, CS-14, CS-15, CS-16 | — |
| CST-8 | `audit-trail apps` (service principals and app-role assignments) | CS-10 | CST-7 |
| CST-9 | `ms365 --org-mode auth grant-scopes`, then a live mailbox-access query as cklaus and as CK-Admin; confirm API shapes and licence coverage | † items, CS-15, CS-16 | CST-7, Audit Reader applied |
| CST-10 | ~~Monthly per-employee access notice~~ — dropped (OQ-5 resolved: no) | — | — |
