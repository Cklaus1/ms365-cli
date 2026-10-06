# todo — derived view of commit history on `build/ms365-cli-hardening`

## Tier 0 — core (main)
- [ ] CORE-2 typecheck gate
- [ ] CORE-1 exit path + e2e harness (critical path)
- [ ] CORE-3 + CORE-6 preAction hook / program.opts()
- [ ] CORE-7 dotenv quiet
- [ ] CORE-5 confirm → CancelledError (core part)

## Tier 1 — parallel
- [ ] AUTH: AUTH-1/2/3, msal bump (Q8), keytar removal, GRAPH-7 auth side, AUTH-4, auth status exit (CORE-5), setup re-prompt (MISC-1)
- [ ] GRAPH: GRAPH-1, GRAPH-5/7, GRAPH-4, GRAPH-2, GRAPH-3, GRAPH-6, CORE-4 dry-run/streamPages, FMT-1, PAGE-1 helpers, quoting helpers, command test harness

## Tier 2 — parallel command slices
- [ ] MAIL: MAIL-1..8, CORE-4/5 mail parts
- [ ] CALWF: WF-1..3, CAL-1/2, TODO-1, MISC-1 workflow parts
- [ ] FILES: DRIVE-1/2, EXCEL-1, SP-1/2, PLANNER-1/2, ONENOTE-1, SEARCH-1
- [ ] ORG: TEAMS-1/2, USER-1..3, GROUPS-1
- [ ] META: SCHEMA-1, SKILLS-1 (command-tree introspection)

## Tier 3 — remote
- [ ] REMOTE-1a/b/c `ms365 serve`
- [ ] REMOTE-1d `ms365-remote` + REMOTE-0 ssh mode

## Tier 4
- [ ] §6 docs/hygiene, README, CHANGELOG, npm audit 0
- [ ] Implementation review loop
- [ ] End-of-run report
