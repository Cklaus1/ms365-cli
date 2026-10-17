# Flight plan: follow-up specs (DR, Q, R, V, X)

**Mission:** make dry-run truthful, offline and leak-free; finish pagination and folder resolution; harden uploads, idempotency and watch; make validation and exit codes consistent; cut startup with lazy MSAL. The five specs are built in one run with each overlap owned once.

**Decisions** (tasks/specs/DECISIONS.md):
- **F1** the dry-run trace goes to stderr and `--quiet` doesn't suppress it; stdout stays parseable. X-9 is rejected.
- **F2** `success()` is silent under dry-run, plus one trailer line. **F3** redact credential keys only.
- **F4 ⚠** under dry-run, `confirm()` doesn't prompt and prints `would ask:`; accepted only together with DR-22 (local commands don't act).
- **F5** `cal update` uses a `<from GET>` marker. **F6** read-aggregate workflows print the trace only.
- **F7** X-1 lazy MSAL only; the warm MCP server is N1.
- **F8** `drive copy` is ID-based; `cp` and `copy` coexist. **F9** no `id` table column; `$expand`. **F10** MCP subscribe → `--reset`.
- **F11** mail-folder names match exactly. **F12** `Inbox/Projects` path syntax, plus a BFS fallback capped at 500.
- **F13** unpaged endpoints use `printList` too. **F14** pre-flight uses CNF per-option scopes. **F15** login keeps `User.Read.All`. **F16** directory `$search` escaping †.
- **F17** idempotency claims record the owner, report `outcome_unknown`, and get `idempotency list`/`release`; never auto-release.
- **F18** `--html` replies go through createReply. **F19** idempotency keys on Teams sends. **F20** public `uploadSession()`. **F21** timeout scales with body size.
- **F22 ⚠** SIGINT mid-pass in `watch` → exit 130 (D2 footnote).
- **F23** dry-run works without a client ID (DR-T2). **F24** `set-timezone` accepts IANA or Windows-shaped names. **F25** presence pairs †. **F26** Premium hint wording. **F27** the wizard reads piped stdin. **F28** `--percent` stays 0/50/100. **F29** `auth status` stays on stdout.
- **F30** each cross-spec overlap gets a single owner; the snapshot and docs regenerate once (FIN).

**needs-human (excluded):** N1 warm MCP server (XT10/XT11, gated on a logged-in measurement); N2 PowerShell `.ps1` shim (needs Windows); N3 live-tenant validation of † items.

**Step 1 revisions:** about 45 markers across the 5 specs.
- Added: DR-22, Q-20, R-14, R-15, V-22, V-23, X-15.
- Confirmed false positives: DR-5, Q-5, Q-7.
- Already fixed: V-17.

**Critical path:**
- **DR-T1 + DR-T2 (preview path, lazy auth):** a no-client-ID dry-run matrix (exit 0, trace on stderr, stdout parseable, no secrets).
- **RT-1:** byte-exact resume for drive and mail.
- **RT-5:** an 8-process claim stress test plus pid-death detection.

**Shape:**
- 47 spec tasks (after 5 folds), 4 waves, 6 + 2 + 9 + 2 tracks; the longest chain is W2 core (7 tasks).
- Test: `npm test` + typecheck. Run: `node dist/index.js`. Smoke: `scripts/smoke.sh` (9 → 12 checks). Suite budget 4.5 s, run budget unbounded, reviewer fable.
- Baseline: 854/854 tests, smoke 9/0/0.

**First:** W1 runs in parallel: XT1, VT1, RT-1→3→4, QP-7, RT-10 + XT8, VT3.
