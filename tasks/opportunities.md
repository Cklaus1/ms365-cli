# Opportunities (deferred work)

Things discovered mid-build that have no assigned task. Each entry gives a proposed severity, and none of them are acted on during this run.

- **[P3] `--dry-run mail send` prints "Email sent successfully."** Found in the Step 0 smoke run. The dry-run message misleads both agents and people. Proposal: print "(dry-run) not sent" instead.
- **[P3] Warm in-process MCP server** (alternative to D16). Spawning the CLI per call costs process start-up each time. Proposal: an embedded mode where `handleErrors` doesn't exit.
- **[P3] `drive copy --wait`** (D20). There's no copy command. Add `drive copy`, which returns a monitor URL, and poll it when `--wait` is given.
- **[P3] MCP notifications on delta changes** (§6 event-driven, "later").
- **[P3] Lazy-load command modules and MSAL** (P-7, measured during T05). `--version` takes 70 ms; the potential saving is about 30 ms. It needs stub command registration so `--help` still works.
- **[P2] KQL quoting gaps outside mail.** `contacts list --search`, `search people` and `groups list --search` (`groups.ts:28`) build `"…"` without `kqlQuote`, so a term containing `"` breaks the query. (Found by T19.)
- **[P3] Teams commands with no `--all`.** `teams members`, `channels`, `list` and `chat get` don't paginate. (T14)
- **[P2] More user.ts status lines on stdout.** `set-presence`, `set-timezone` and `set-language` print to stdout (C-19 pattern). The presence commands throw a plain `Error` for a missing client ID. `set-ooo --status`/`--audience` and `set-timezone` values aren't validated. The `upload-photo` `status >= 400` check never runs. (T17)
- **[P3] Planner usage errors exit 1 instead of 2.** `planner update-task --percent` validation and an unguarded `JSON.parse` in `update-plan-details --categories` both exit 1. The G-9 Premium hint also shows on a genuine 404. (T16)
- **[P2] Dry-run leaks the share password.** `GraphClient` dry-run prints request bodies to stdout, including `drive share --password` values. Redact known secret fields in dry-run output. (T15)
- **[P3] Mail folder names in move/triage.** `move --folder` and `triage --folder` don't resolve folder names. Attachment upload-session chunks have no retry or resume, unlike `uploadFile`. (T10)
- **[P2] Calendar time-zone defaults.** `calendar.ts` `--timezone` defaults to `'UTC'`, which T11 is fixing (M-9). Logged by T12; no action needed if T11 covers it.
- **[P3] free-busy --emails isn't validated**, and dry-run `cal update` prints the zone-read GET. (T11)
- **[P3] Setup wizard with piped stdin.** With non-TTY stdin, the second prompt never gets its input and the wizard exits 0 silently (this predates the run). `setup.ts` `prompt()` duplicates the readline code in `confirm()`. (T18)
