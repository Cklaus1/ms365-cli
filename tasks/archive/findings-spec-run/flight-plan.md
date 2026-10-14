# Flight plan: ms365 findings spec

**Mission:** fix every engineering finding, bring all dependencies to their latest stable versions, and ship the agent-native features (describe, MCP, NDJSON, watch, meetings) on the `build/ms365-findings-spec` branch, until the full suite, typecheck and the 9-check offline smoke suite are all green.

**Engineering decisions (canon):**
- **D1** exit → `process.exit(code)` after stdout/stderr drain, replacing SIGTERM. `process.exit` was shown to work.
- **D2** exit codes 0/1/2/3/4/5/6/7. No contract exists yet (everything exits 143).
- **D3** `--error-format json` is opt-in. Additive.
- **D4** responses use a caller-declared `responseType` (json/text/binary/stream). Sniffing MIME types always misses some.
- **D5** retries: 429 for any method; 5xx and timeouts for GET/PUT/DELETE/PATCH; POST only before the request is sent. Prevents duplicate sends.
- **D6** keep `@odata.type`/`etag`/`count`. Additive, and unblocks M-14 and M-16.
- **D7** ⚠ close call. Own `ICachePlugin` with a lockfile; keytar becomes optional and is used only to migrate. msal-node-extensions still uses keytar and adds a native runtime.
- **D8 / D25** ⚠ close call. Latest stable for everything: msal-node 7.1, commander 15, dotenv 18, TypeScript 7 (fall back to 5.x/6.x), vitest 5, tsup 9. `@types/node` stays on `^22`, and `engines` becomes `>=22.12` (Node 20 is end-of-life).
- **D9** time zone comes from `MS365_CLI_TIMEZONE`, then the local `Intl` zone. One source, so times don't shift twice.
- **D10** `chat reply` becomes a quoted reply via `messageReference`. Keeps the capability. †
- **D11** `create-rule` needs a condition or `--all-messages`; `--disabled` replaces `--enabled`. Explicit beats a prompt.
- **D12** never overwrite downloads without `--force`; file names use `basename`. Fails safe.
- **D13** wizard option [1] prints "not yet available" until N1 ships.
- **D14** wizard scopes are saved and used at run time.
- **D15** `/content` uses manual redirects and re-attaches the token only for Microsoft hosts. Ready for the 2027-04-01 tempauth retirement.
- **D16** ⚠ close call. MCP is hand-rolled stdio JSON-RPC that spawns the CLI per call. No new dependency.
- **D17** downloads stream.
- **D18** `-o ndjson`, `--fields` and `--cursor` are added without changing the envelope (the envelope is N5).
- **D19** `$batch` helper sends chunks of 20, with per-item retry on 429.
- **D20** `--wait` is a no-op (false positive).
- **D21** `--account` and `--profile`.
- **D22** small built-in natural-time parser. No dependency.
- **D23** `--idempotency-key`: `transactionId` for events, a 24 h ledger for mail.
- **D24** † items are built to Graph's documented shape and marked "unverified live".

**needs-human (excluded, 7 items):**
- N1 publish the default app (PM-1)
- N2 app-only sign-in (PM-2, app-only part)
- N3 policy file and plan/apply (PM-7)
- N4 npm name (PM-12)
- N5 JSON envelope
- N6 encryption at rest
- N7 live-tenant validation

**Step 1 revisions:** 38 markers.
- Raised: C-3 P1→P0.
- Lowered: C-6 P1→P2, C-12 P2→P3, G-3 P2→P3.
- Corrected: C-1's premise, C-16's fix, M-27.
- Added: C-20, M-30, M-31, and G-11 (your request).

**Critical path and extra gates:**
- **T02 exit contract:** spawn the built CLI and check the exit-code matrix.
- **T03 GraphClient core:** byte-exact round trips for 8 MIME types, and POST never retried after a timeout.
- **T08 auth cache:** round-trip persistence, plus 8 parallel writers leaving a valid cache.

**Shape:**
- 37 tasks across 5 tiers. Longest chain is 15 (T00→…→T44).
- Tier 2 has 10 tracks, 8 of them disjoint and run in parallel.
- Test: `npm test` (+ `tsc` from T20). Run: `node dist/index.js`. Smoke: `scripts/smoke.sh`.
- Suite budget ~2 s, run budget unbounded, reviewer fable.
- Baseline: 95/95 tests passing; smoke 1 pass, 6 fail, 2 skip.

**First 3 tasks:**
1. T00: dev toolchain upgrade.
2. T01: runtime dependency upgrade (msal 7, commander 15, dotenv 18).
3. T02: exit and error contract **[CP]**.
