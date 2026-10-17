# Flight plan: agent-host integration (AH-/AHT-)

**Mission:** make ms365 a safe Microsoft 365 transport for Chief and MCP hosts — profile scopes as a fail-closed ceiling, broker-supplied tokens, in-band completeness/refs, host-held watch checkpoints, machine-readable auth/health, then plan/apply, honest sends, calendar safety, MCP gaps and delegate reads.

**Plan options** (gates are cheap: ≈ 8 s per change, smoke ≈ 6 s; the cost is agent time):
- ▶ **Fastest @90 % (active):** 7–11 h, ~70 % inside; ~10 % a real problem slips (found by Chief's wrapper or the live run, ≈ 1–2 h to fix). *Done by tonight; out of 10 runs like this about 1 lets a small mistake through, fixed in about an hour later.*
- **Balanced:** 8–12 h, ~5 % slip. *One more hour; extra check on the two most important parts; about 1 in 20 lets a mistake through.*
- **Most reliable:** 9–14 h, ~2 % slip. *Done tomorrow morning; checks everything after every step; about 1 in 50.*
- ⚠ production-level: touches auth and credentials (token hand-off, profile ceilings, cache refusal) → Most reliable recommended (+2–3 h). Running Fastest @90 % (default option, not chosen by the user). Want a different one?

**Step 2/3 engineering decisions:**
- OQ-1 token hand-off → env `MS365_CLI_TOKEN` now, `MS365_CLI_TOKEN_COMMAND` helper in AHT-10 → one-shot env suits Chief G2.
- OQ-2 external mode → env/helper turns it on, `type: host` makes it mandatory → no cache fallback.
- OQ-3 ceiling scope → profiles only → single-user runs unchanged.
- OQ-4 `_ref` → behind `--refs`/`MS365_CLI_REFS=1`; OQ-5 NDJSON → `_meta` trailer; OQ-6 immutable ids → always for mail/calendar.
- OQ-7 plan → canonical JSON, sha256 without `digest`, caller path 0600, 15 min default TTL, `--plan-ttl` ≤ 24 h.
- OQ-9 reconcile → `internetMessageId`; OQ-10 lost response → exit 1 `outcome_unknown`; OQ-11 checkpoint → versioned base64 JSON.
- OQ-12 device-code relay → detached child + 0600 handle (AHT-10); OQ-13 403 → `AUTH_REQUIRED`/`access_denied`.
- OQ-15 delegate → per-command `--mailbox`; OQ-16 Chief scopes → canon from ADR 0013 D3; OQ-17 events window → roll at < ½ left with `gap window_rolled`; OQ-18 budget → env vars; OQ-19 mismatch → exit 2.
- E1 ceiling/external-token/childEnv as pure modules (`src/ceiling.ts`, `src/external-token.ts`, `src/child-env.ts`) → R-33 testability.
- E2 ⚠ close call: catalog `uses` is CNF (`string[][]`, each inner list any-of) → `user list` alternatives and `workflow digest`'s conjunction both expressible; every one of the 15 multi-scope entries decided explicitly.
- E3 cursor binding compares the decoded path case-insensitively and the version segment exactly → Graph nextLink path shapes are † (live).
- E4 no parallel dispatch → every task pair shares `catalog.ts`/README/snapshot.
- E5 fetch-stub `--import` preload (`test/fixtures/fetch-stub.mjs`) shared by spawned tests and the AHT-6a smoke checks → one offline Graph double.
- E6 identity for `_ref`/plans from the decoded token via one `GraphClient.identity()` (GET `/me` only for opaque tokens).
- E7 `health` is a top-level command in `src/commands/health.ts`; E8 plan/apply in `src/plan.ts` + `src/commands/apply.ts`.

**needs-human (excluded):** OQ-8 → AHT-9 non-TTY apply of delete/cancel (rec. b); OQ-14 → AHT-12 subscription helpers (rec. a); AHT-6b live read evidence (Chief MSC-H3); AHT-14 live write evidence. Pruned: 3 nodes + 1 sub-part.

**Step 1 revisions:** R-1…R-39 (39 markers): incl. ceiling before dry-run/offline returns (R-3), `--cursor` scope hole (R-4), `.env` token exclusion (R-10), host profile refuses opaque tokens (R-11), error JSON = last stderr line (R-15), host-bounded retries (R-16), checkpoint `@-`/`new` (R-19), coverage folder/window (R-20), serial G2 order (R-35).

**Critical path:** AHT-1 ceiling + AHT-2 external token. Gate: catalog-wide property test (refuse iff a `uses` clause unmet; monotone in the ceiling) + zero-fetch refusals end-to-end (PREFLIGHT=0, opaque, `--dry-run`) + no MSAL load / no config-dir writes in external mode.

**Shape:** 12 tasks, 2 tiers (G2: 6, G3: 6), longest chain 12 (serial), 0 parallel tracks. Test `npm test`+tsc (6.5 s, 1784/1784), smoke `scripts/smoke.sh` (15/0/0), reviewer = self (adversarial), budget unbounded.

**First 3:** AHT-1 ceiling → AHT-2 external token → AHT-3 refs/completeness.
