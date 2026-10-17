# Flight plan: admin oversight

**Mission:** fix the shipped `reports` bugs, remove `chief brief`, and add read-only `audit-trail` commands (mailbox, changes, apps, access, offline import) that show access metadata and grants, never content.

**Decisions:**
- **E1:** no new report types this run (after AOT-7).
- **E2:** reports reader is format-agnostic (CSV on v1.0, JSON when returned), never beta.
- **E3:** `audit-trail apps` shows facts only, no risk score.
- **E4 ⚠ close call:** `Directory.Read.All` (broad, read-only, org-mode) to list delegated grants, because Learn lists it as least privileged; AOT-7 probes the narrower `DelegatedPermissionGrant.Read.All`.
- **E5:** wave 1 AOT-1 ∥ AOT-2, then serial AOT-3 → 5 → 6 → 4.

**needs-human:** AOT-7 (live tenant: grant new scopes, run real queries, confirm † items).

**Step 1 revisions:** 9 revised (incl. `Directory.Read.All` replaces `DelegatedPermissionGrant.Read.All`; `OFFLINE_COMMANDS` can't see `audit-trail import` → AO-34 fix; shared mailboxes look disabled → `disabled-or-shared`; streaming CSV for 141 MB exports), 2 added (AO-33, AO-34).

**Critical path:** AOT-3 — async query lifecycle tests + allow-list property test across all output formats.

**Shape:** 6 tasks, longest chain 5, 1 parallel pair. Test `npm test`+typecheck (~7 s); smoke `scripts/smoke.sh`; reviewer fable; budget unbounded. Baseline 1664/1664, smoke 14/0/0.

**First:** AOT-1 ∥ AOT-2, then AOT-3.
