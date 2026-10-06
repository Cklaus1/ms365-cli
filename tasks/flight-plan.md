# Flight plan — ms365-cli hardening

**Mission:** fix every reviewed defect in ms365-cli (exit codes, token cache, Graph client, ~60 command bugs) and ship `ms365 serve` + `ms365-remote` for tailnet agents, verified by a typecheck+test gate and a built-binary smoke.

**Decisions (user: Q1–Q9; engineering: rest)**
- Q1 remote model → `ms365 serve` + SSH interim · Q9 API → command-level `/v1/execute` allowlist
- Q2 setup → Graph Command Line Tools client ID default (ID confirmed on Microsoft Learn; retirement risk noted)
- Q3 config → keep cwd `.env`, documented · Q5 cache → file-only 0600, keytar removed · Q8 msal-node major bump with AUTH-2
- Q4 upload conflicts → `--conflict`, upload defaults `rename`
- ⚠ close call: Q11 `drive cp` defaults `replace` → explicit destination path; keeps cat→edit→cp round trip
- Q6 `@odata` stripping → formatter only · Q7 POST/PATCH retry → 429 only
- Q10 caller identity → `Tailscale-App-Capabilities`; X-Forwarded-For only audited
- Q12 `cal rooms` → beta `findRooms` (no admin-consent `Place.Read.All`)
- Q13 distribution lists → client-side `groupTypes` filter (no advanced query)
- ⚠ close call: Q14 serve transport → Unix socket 0600 default (loopback TCP lets any local process forge caps); 127.0.0.1 `--port` fallback with warning
- CORE-2 → `handleErrors<A extends any[]>`: 756 errors → 4 real ones; typed options interfaces per file as slices touch them

**Needs-human:** none. Out of scope (spec "later"): REMOTE-2, REMOTE-3 → opportunities.md.

**[REVISED] highlights (39 markers):** new CORE-6 (argv flag spoofing), CORE-7 (dotenv stdout banner, P1), GRAPH-7 (401 never refreshes), MISC-1; REMOTE: `-o json` only when declared, explicit child env, strict header parsing, Unix socket; GRAPH-3 photo split; TEAMS-1 narrowed + TEAMS-2 fix; CAL-1 fix changed; DRIVE-1 evidence corrected; AUTH-3 ships with AUTH-2 (one re-login).

**Critical path:** CORE-1 (built-binary exit-code e2e), AUTH-1/2 (cache file perms + concurrency proof), GRAPH-4 (formatter byte-identity invariant).

**Shape:** ~24 nodes in 5 tiers; longest chain CORE-2→CORE-1→CORE-3/6→CORE-7→GRAPH→FILES→REMOTE→DOCS→review (9); parallel tracks: 2 (tier 1), 5 (tier 2). TEST_CMD `npm test` (+tsc from CORE-2); RUN_CMD `node dist/index.js`; smoke = temp-dir dry-run JSON / not-logged-in exit 1 / schema without client ID / serve+remote. RUN_BUDGET unbounded. Baseline: 95/95 pass, tsc 756 errors, exit 143.

**First 3 tasks:** CORE-2 typecheck gate → CORE-1 exit path + e2e harness → CORE-3/CORE-6 preAction + program.opts().
