# todo: ms365 findings spec build

Derived view of commit history (`git log --grep '^T[0-9]'`). The commits are the source of truth.


## Tier 0
- [x] T00: Dev toolchain upgrade: typescript, vitest, tsup, tsx, @types/node@22; `engines >=22.12`; CI matrix 22/24; README requirements and badge
- [x] T01: Runtime dependency upgrade: dotenv 18, commander 15, @azure/msal-node 7.1.0
- [x] T02: **[CP]** Exit and error contract: `ExitCode` and `CliError`; `handleErrors` → `process.exit` after drain; `--error-format json`; map Graph status codes to exit codes; `confirm` refusal → 7; wrap `auth` commands; test hook so tests don't exit
- [x] T03: **[CP]** GraphClient core: method-aware `fetchWithRetry`; `responseType` json/text/binary/stream; pagination through the shared path with headers and 401 refresh; keep `@odata.type`/`etag`/`count`; return metadata; manual redirect for `/content`; `requestRaw`
- [x] T04: Upload: session URL fix; chunk retry and resume via `nextExpectedRanges`; timeouts; dry-run prints the session URL
- [x] T05: Startup: argv flags parsed by commander (including `=` forms); validate `api-version`; skip auth for offline commands; lazy-load command modules; `schema [endpoint]`
- [x] T06: Shared helpers: time-zone resolver (D9), natural-time parser (D22), `encodePath`, `odataQuote`, `kqlQuote`, `safeOutputPath` with `--force` (D12), and the rule that status lines use `info()`
- [x] T07: Command test harness: register a module against a fake `GraphClient` that records requests; run one command and get `{calls, stdout, stderr, exitCode}`

## Tier 1
- [ ] T08: **[CP]** Auth cache rework: `ICachePlugin` with lockfile and unique temp files; persist after a silent acquire when `fromCache===false`; selected account stored in a file only; full logout cleanup; keytar → optionalDependency used once for migration, with the probe memoised; `forceRefresh`; account selected after login; `null` login result; `Place.Read.All`, meeting scopes and Mail scopes added to scope lists

## Tier 2
- [x] T10: mail.ts fixes
- [ ] T11: calendar.ts fixes
- [x] T12: workflow.ts fixes
- [x] T13: excel.ts fixes
- [x] T14: teams.ts fixes
- [x] T15: drive.ts fixes and streamed downloads
- [x] T16: todo.ts and planner.ts fixes
- [x] T17: user.ts fixes
- [ ] T18: Setup wizard, `auth` command, generate-skills
- [x] T19: sharepoint, groups, contacts, onenote, search fixes
- [ ] T20: CI typecheck (`tsc --noEmit`) and a dry-run endpoint snapshot for every command

## Tier 3
- [ ] T21: Streaming `--all`; `-o ndjson`; `--fields`; `--cursor`
- [ ] T22: `$batch` helper; `batch-read` and `batch-move` use it

## Tier 4
- [ ] T30: `describe --json` catalogue (args, options, destructive, idempotent, scopes)
- [ ] T31: Generate the README command table and SKILL.md from the catalogue; CI drift check
- [ ] T32: `mcp serve` (stdio JSON-RPC that spawns the CLI per call)
- [ ] T33: `meetings` group; `cal create --teams`; `cal find-time`
- [ ] T34: `watch` (delta with a saved cursor, NDJSON output)
- [ ] T35: `mail send/draft/reply --attach`
- [ ] T36: `--body text\
- [ ] T37: Audit log for non-GET calls, `audit tail`, `--actor`
- [ ] T38: `--idempotency-key` (D23)
- [ ] T39: `auth scopes`; pre-flight `missing_scope` error
- [ ] T40: `--account` and `--profile` (D21)
- [ ] T41: `auth login --browser` (PKCE loopback)
- [ ] T42: Name resolution: `resolve` command; names accepted for todo lists, team/channel and planner plan
- [ ] T43: Coverage gaps: excel tables and rows, todo checklist, teams chat create and @mentions, Outlook notes, multi-entity search
- [ ] T44: Shell completion and default table columns

## needs-human (excluded)
- [ ] N1
- [ ] N2
- [ ] N3
- [ ] N4
- [ ] N5
- [ ] N6
- [ ] N7
