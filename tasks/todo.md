# todo — derived view of commit history on `build/ms365-cli-hardening`

## Tier 0 — core (main)
- [x] CORE-2 typecheck gate (cc6c29b)
- [x] CORE-1 exit path + e2e harness (critical path) (119c0ac)
- [x] CORE-3 + CORE-6 preAction hook / program.opts() (7ab70e4; + positional options, spec [REVISED])
- [x] CORE-7 dotenv quiet (7ab70e4)
- [x] CORE-5 confirm → CancelledError (core part) (af26654)

## Tier 1 — parallel
- [x] AUTH: AUTH-1/2/3, msal bump (Q8), keytar removal, GRAPH-7 auth side, AUTH-4, auth status exit (CORE-5), setup re-prompt (MISC-1) (b138d27, merged)
- [x] Node 24 LTS target (user decision; c57b227)
- [x] GRAPH: GRAPH-1, GRAPH-5/7, GRAPH-4, GRAPH-2, GRAPH-3, GRAPH-6, CORE-4 dry-run/streamPages, FMT-1, PAGE-1 helpers, quoting helpers (9415b9d, merged; 273 tests)

## Tier 2 — parallel command slices
- [x] MAIL: MAIL-1..8, CORE-4/5 mail parts (4372d01)
- [x] CALWF: WF-1..3, CAL-1/2, TODO-1, MISC-1 workflow parts (b8523ee)
- [x] FILES: DRIVE-1/2, EXCEL-1, SP-1/2, PLANNER-1/2, ONENOTE-1, SEARCH-1 (ab36c39; + 336eda7 base64 branch removed)
- [x] ORG: TEAMS-1/2, USER-1..3, GROUPS-1 (9341206)
- [x] META: SCHEMA-1, SKILLS-1 (f1d68f0; + 1913d82 src/commands/registry.ts)

## Tier 3 — remote
- [x] REMOTE-1a/b/c `ms365 serve` (32a4068; 495 tests; smoke step 4 passes)
- [x] REMOTE-1d `ms365-remote` + REMOTE-0 ssh mode (32a4068)

## Tier 4
- [x] §6 docs/hygiene, README, CHANGELOG, schema ENDPOINTS (f2cb9b9; README Remote Access section added after REMOTE merge); npm audit 0
- [x] Implementation review loop: 4 reviewers, 15 findings + 1 follow-up (parseDate), all reproduced fail-first, all fixed (afd2b72, 903fafc, merges fix/cmds, fix/remote 0295592, fix/core 2df98e3); 579 tests; smoke 1-4 + dot-segment + impossible-date pass
- [x] End-of-run report

## Follow-up spec (`tasks/ms365-cli-followup-spec.md`, branch `build/ms365-cli-followup`)
- [x] Q1-Q10 decided (recommendations), adversarial spec review + [REVISED] notes
- [x] T0: CLEAN-1, META-2, DRY-1 helper (6d7da8c); CAL-4/TZ-1 zone helpers (75872a5)
- [x] Slices: files 4402c86, mail 09f40ee, org 4d803e8, calwf 2e5a58d; SCHEMA-2 af81759; docs 1280769; REMOTE-4/5 42bc846
- [x] Implementation review: 2 reviewers, 3 findings (eventTime date-only/non-ISO 92193a0; DRIVE-3 test migration in fu/remote; PAGE-2 Breaking entry 6638d1f) + 1 smoke finding (serve repeated signal)
- [x] README/CHANGELOG, full gate, built-binary smoke 1-4 + follow-up checks
