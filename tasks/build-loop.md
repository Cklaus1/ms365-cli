# Build Loop: agent-host integration (AH- / AHT-)

**Spec:** tasks/specs/agent-host-integration-spec.md (Step 1 + Step 2 done 2026-10-17, f1cc32c; Step 3 decisions E1–E8 below) · **Branch:** `build/agent-host-integration` (off `origin/main` 2be3fce)
**Previous harnesses:** tasks/archive/ (admin-oversight-run archived by this run).

## Resolved placeholders

| Placeholder | Value |
|---|---|
| `<TEST_CMD>` | `npm test` + `npm run typecheck` (+ `npm run docs:check` for tasks that add options/commands) |
| `<RUN_CMD>` / `<SMOKE_SCENARIO>` | `node dist/index.js` / `scripts/smoke.sh` (15 checks today) + AHT-6a host-mode checks, all offline: `type: host` profile ceiling refusal under `--dry-run` → exit 4, last stderr line `{"error":{"code":"scope_ceiling"}}`; `describe --json` lists `--refs`, `watch --checkpoint`/`--select`, `health`, `auth status -o`; `auth status -o json` with an env JWT → `source:"external"`, no MSAL load; `watch mail --checkpoint new` against a fetch-stub preload → last line `type:"checkpoint"`, no `watch/` dir. Signals = exit code + parsed JSON field. |
| `<FULL_SUITE_BUDGET>` | measured 6.5 s (`npm test`, 1784 tests) + 0.24 s tsc + 0.23 s docs:check → full gate after every task |
| `<RUN_BUDGET>` | unbounded |
| `<RELIABILITY>` | Fastest at 90 % (default option, not chosen by the user) |
| `<REVIEW_MODEL>` | invoking model (self-review, adversarial) — Step 1 was done by the spec reviewer |

**Baseline (2be3fce + spec commits, Node 24.21.0):** `npm test` 58 files / 1784 tests passed, 6.0 s (6.5 s wall); `tsc --noEmit` 0 errors (0.24 s); `docs:check` clean (0.23 s); `scripts/smoke.sh` pass=15 fail=0 skip=0 (6.1 s incl. `tsup` build). Known-failure list: empty.

## Dev-loop cost baseline
One change lands for ≈ 8 s of local CPU (suite + tsc + docs:check); smoke ≈ 6 s. No GPU, device or CI runner. The only scarce resources are (a) the live tenant + Chief's Entra app + the principal's time (needs-human, excluded: AHT-6b/AHT-14) and (b) this single agent's serial attention (R-35: G2 tasks share `index.ts`, `profiles.ts`, `commands/auth.ts`, `formatter.ts`, `watch.ts`). Spawned-entry tests (≈ 0.6 s each through tsx) are capped at 12 new for the AH subset.

## Plan options
| Option | Wall time | P(inside range) | P(real problem slips) | Slip cost |
|---|---|---|---|---|
| **Fastest @90 % (active)** | 7–11 h [INFERENCE] | ~70 % | ~10 % | found by Chief's MSC wrapper or the AHT-6b live run; ≈ 1–2 h to fix |
| Balanced | 8–12 h [INFERENCE] | ~70 % | ~5 % | as above, rarer |
| Most reliable | 9–14 h [INFERENCE] | ~65 % | ~2 % | as above, rare |
- Fastest: suite+tsc every task, docs:check on option changes, smoke once per tier, one implementation review at the end.
- Balanced: Fastest + an extra adversarial review after AHT-1/AHT-2 (critical path).
- Most reliable: smoke + docs:check every task, adversarial review per tier, fail-first re-proved for every AC.

## Task DAG

```mermaid
graph LR
  subgraph G2[Tier G2 — Gate-2 subset, serial R-35]
    AHT1[AHT-1 ceiling: host type, uses, preAction, cursor bind, MCP, describe ★] --> AHT2[AHT-2 external token ★]
    AHT2 --> AHT3[AHT-3 immutable ids, --refs _ref, complete/_meta]
    AHT3 --> AHT4[AHT-4 watch --checkpoint, gap/backfill, window roll, --select]
    AHT4 --> AHT5[AHT-5 auth status -o json, health, retry budget, retryAfterMs]
    AHT5 --> AHT6a[AHT-6a offline conformance pack + smoke]
  end
  subgraph G3[Tier G3 — Gate 3 / Operator seat, serial]
    AHT7[AHT-7 plan/apply core, If-Match, --bcc, --body-file, sha256 attachments] --> AHT8[AHT-8 honest send + reconcile]
    AHT7 --> AHT9[AHT-9 calendar plans: series, instances, cancel, notifies — TTY part only]
    AHT7 --> AHT10[AHT-10 MCP gap/resume/body-file, token helper, login --start/--wait]
    AHT11[AHT-11 --mailbox delegate reads]
  end
  AHT6a --> AHT7
  AHT3 --> AHT7
  AHT2 --> AHT10
  AHT4 --> AHT10
  AHT1 --> AHT11
  AHT4 --> AHT11
  AHT9 --> AHT11
  AHT10 --> AHT11
  AHT8 --> AHT13[AHT-13 docs: agent-host guide, mapping tables, skills, snapshot]
  AHT11 --> AHT13
  AHT6a -.needs-human MSC-H3.-> AHT6b[AHT-6b live read evidence]
  AHT9 -.needs-human OQ-8.-> AHT9b[AHT-9 non-TTY apply delete/cancel]
  AHT4 -.needs-human OQ-14.-> AHT12[AHT-12 subscription helpers]
  AHT9 -.needs-human tenant.-> AHT14[AHT-14 live write evidence]
```
★ = critical path. Execution order (topological, spec priority G2 first): AHT-1 → 2 → 3 → 4 → 5 → 6a → 7 → 8 → 9 → 10 → 11 → 13. Self-check: 12 buildable nodes, every non-needs-human AHT item exactly once (AHT-6 split a/b, AHT-9 split TTY/non-TTY), acyclic (the order above is a valid topological sort).

## Track / contention graph
All tracks share implicitly: `src/catalog.ts`, `README.md` (generated), `test/__snapshots__/snapshot.test.ts.snap`, `tasks/*`. Per-task files: AHT-1 `index.ts, profiles.ts, catalog.ts, ceiling.ts(new), formatter.ts, mcp-server.ts, describe.ts`; AHT-2 `index.ts, profiles.ts, auth.ts, commands/auth.ts, errors.ts, graph-client.ts, external-token(new), child-env(new), grant-scopes.ts, keychain.ts, setup.ts`; AHT-3 `graph-client.ts, formatter.ts, mail.ts, calendar.ts, watch.ts, workflow.ts, refs(new)`; AHT-4 `watch.ts`; AHT-5 `commands/auth.ts, health(new), graph-client.ts, errors.ts`; AHT-7 `mail.ts, calendar.ts, plan(new), apply(new)`; AHT-8 `mail.ts, idempotency.ts`; AHT-9 `calendar.ts`; AHT-10 `mcp-server.ts, auth.ts, commands/auth.ts`; AHT-11 `mail.ts, calendar.ts, watch.ts`. No pair of ready tasks is disjoint once `catalog.ts`/README/snapshot are counted → **parallel dispatch loop disabled** (E4); one agent, serial.

## Loops
- **Outer:** take the next node in the order above whose dependencies are DONE and that is not needs-human; run the inner loop; mark DONE in `todo.md`; re-evaluate.
- **Inner (one task):** read the touched code → fail-first tests for each behaviour change (run only the task's test files) → implement → task tests green → orchestrator gate in a separate step: `npm test` + `npm run typecheck` (+ `docs:check`, `npm run docs` to regenerate when options changed) → self-review checklist: diff confined to the task's files; no TODO/dead code; no secret/real address/tenant id/token in any file (role refs only; test JWTs are synthetic); reuse existing helpers (`scopeSatisfied`, `tokenInfo`, `writeFileAtomic`, `printList`, `runCommand`, ledger) before new code; no speculative abstraction; matches the touched file's pattern; diff size proportional → commit with explicit paths (`git commit -m … -- <paths>`) → `chown -R cklaus:cklaus` the repo.
- **Meta:** on any correction/failure append mistake → rule to `tasks/lessons.md`; on divergence update the spec (`[REVISED: build …]`) and the divergence log below. A rule firing twice becomes a mechanism (script/guard test).
- **Full-suite regression:** after every task (the gate costs ≈ 8 s), compared with the 1784/1784 baseline. A new failure blocks unless it is on the known-failure list (empty) via the skill's tally + 3× re-run-through-gate procedure.
- **Smoke/integration:** once per tier (after AHT-6a and after AHT-13): `scripts/smoke.sh` (builds `dist/`, runs in a mktemp dir, 30 s per check), then `git status --short --ignored` must show nothing new beyond `dist/`, `node_modules/` and the owner's untracked files. MCP long-lived server behaviour (AHT-10) is covered in-process by `test/mcp.test.ts` with the injectable `Runner`.
- **Implementation review:** when every buildable node is DONE, one read-only adversarial reviewer against the finished code and ACs 1–32; each finding reproduced (fail-first test or observed output) before acting; dispositions recorded in the end-of-run report.
- **Dev-loop cost:** per tier, split agent time into implementing / gate-waiting / rework; the largest non-useful share gets an action or an `opportunities.md` entry with numbers.
- **Re-plan:** a task whose diff grows past ~2× its expected size, or whose 2nd attempt fails, is re-scoped in the spec before attempt 3.
- **Contract loop:** any change to the cross-spec contract (spec line 286) is sent to `agent://BuildMsc` with the exact change and recorded as `[REVISED: build …]`; changes received from BuildMsc are applied in the next task boundary.

## Gates
- **Critical path: AHT-1 (ceiling) and AHT-2 (external token source).** Everything Chief relies on sits on them (host profile, `scope_ceiling`, `token_*` codes; AHT-5 health `uses`, AHT-7 apply re-check, AHT-11 `.Shared` selector). Extra proof: (a) a property test over **every** catalog path × a family of ceilings: the check refuses iff some `uses` clause has no scope satisfied by the ceiling, and adding scopes to a ceiling never turns an allowed command into a refused one (monotonicity); (b) end-to-end through the real entry: refusal makes zero fetch calls with `MS365_CLI_PREFLIGHT=0`, an opaque token and `--dry-run`; (c) AHT-2: external mode never loads MSAL and leaves the config dir unchanged (file-list diff), and every refusal code makes zero fetch calls. A BLOCKED AHT-1 or AHT-2 halts the run.
- **Dry-run snapshot:** changes only in AHT-3 (immutable-id header) and wherever a task adds a request-shape change it names; never `-u` blindly — diff read first.
- **Dependency hygiene:** no new npm dependency (JWT decode, sha256, canonical JSON all from `node:crypto`/in-repo).
- **Gate cost budget:** suite ≤ 15 s (spec), smoke ≤ 60 s. Above 2× → review before re-running.
- **Statistical gates:** none in this spec (no sampling gate).
- **Skip unchanged inputs:** smoke skipped when the `dist/` sha256 equals the last green run's.
- **Poison-task ceiling:** attempts 1–2 diagnose/re-plan; attempt 3 fails → BLOCKED, dependents blocked, logged.
- **Resumability:** commits are the source of truth; `todo.md` is derived from `git log --grep '^AHT-'`.

## Stop condition
```
DONE = AHT-1..5, 6a, 7, 8, 9 (TTY part), 10, 11, 13 DONE or BLOCKED-and-logged
       AND npm test + typecheck: no new failures vs the 1784 baseline; docs:check clean
       AND every AH-* item addressed or tagged needs-human (AHT-6b, AHT-9 non-TTY/OQ-8, AHT-12/OQ-14, AHT-14)
       AND scripts/smoke.sh passes with 0 FAIL / 0 SKIP incl. the AHT-6a host-mode checks
       AND the implementation review ran, every disposition recorded
```

## Known-failure list
(empty)

## Divergence log
