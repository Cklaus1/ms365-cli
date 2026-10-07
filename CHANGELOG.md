# Changelog

All notable changes to this project are documented here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). IDs in brackets refer to items in `tasks/ms365-cli-hardening-spec.md`; IDs marked `fu:` refer to `tasks/ms365-cli-followup-spec.md`.

## [Unreleased]

### Breaking changes

- **Node.js 24 or later is required** (`engines.node >=24`; was >=18). [§6 Node 24, Q8]
- **OS keychain storage removed; everyone must run `ms365 auth login` once.** `keytar` is gone and tokens live only in `~/.config/ms365-cli/token-cache.json` (0600, directory 0700). Tokens that were stored only in the keychain are not migrated, and the new `MailboxSettings.ReadWrite` scope invalidates existing consent. Users of the old file fallback keep the same path. [AUTH-1, AUTH-2, AUTH-3]
- **Global flags must precede the subcommand.** `ms365 --dry-run mail send …` works; `ms365 mail send … --dry-run` now fails with `unknown option` (exit 1). An argument value such as `--body --dry-run` can no longer switch a mode. [CORE-6]
- **Exit codes are now meaningful:** 0 on success, 1 on any failure (previously every run exited 143 under plain `node`). Declined or non-interactive delete confirmations, partially failed `mail batch-read`/`batch-move`, and `auth status` without a login now exit 1. [CORE-1, CORE-5]
- **Dry-run preview lines go to stderr**, always (even with `--quiet`); stdout carries only data, and commands with `-o` print a `{ dryRun, method, url, body }` object. [CORE-4]
- **Downloads without `--out` stream raw bytes to stdout** (`drive download`, `drive download-shared`, like `drive cat`); the old default file names `download_<id>`/`shared_download` are gone. `onenote page` prints raw HTML instead of a `{rawResponse}` JSON wrapper. An existing output file is never overwritten without the new `--force`. [GRAPH-2, DRIVE-2, MAIL-1, ONENOTE-1, USER-3]
- **`drive upload` no longer replaces existing files by default:** new `--conflict rename|replace|fail`, default `rename` (`drive cp` defaults to `replace`). [GRAPH-3]
- **`sp items` removed** — Graph cannot enumerate `/sites/{id}/items`. Use `sp item <siteId> <itemId>`, `sp drives` or `sp lists`. [SP-1]
- **`mail triage` now reads the inbox** (was all folders); `mail list` still lists all folders and says so; use `--folder inbox` for the inbox. [MAIL-6]
- **`mail folder`/`--folder` name resolution is exact:** names are matched as typed (nested folders included); an unknown or ambiguous name is an error instead of a silently wrong target. [MAIL-3]
- **`user set-ooo --audience` is no longer defaulted to `all`;** it is left unchanged when omitted. `--start`/`--end` without `Z`/offset are local time, and are rejected unless `--status scheduled`. [USER-1]
- **`user set-presence` sets your preferred presence** via `setUserPreferredPresence` (was the app-session `setPresence`). `--availability` is validated, `--activity` is optional and must match (`--activity Focusing` is rejected). `clear-presence` uses `clearUserPreferredPresence`. [USER-2]
- **`workflow` JSON shapes changed:** `unreadEmails` is a number (the inbox `unreadItemCount`, was a string such as `"1+"`); `digest` reports `week: {start, end}` as local dates, `overdueTasks[].due` as `YYYY-MM-DD`, and `tasksDueThisWeek` no longer includes overdue tasks; `meeting-prep -o json` prints `null` when there is no meeting; `file-announce` prints the posted message. [WF-1, WF-2, WF-3]
- **`schema --list` prints JSON** (`{endpoint, method, path, description}` array) instead of a grouped text listing. [SCHEMA-1]
- `search query --entity` rejects types that Graph cannot combine (e.g. `driveItem,message`). [SEARCH-1]
- `planner --due` dates mean that day in the local time zone (were midnight UTC). [PLANNER-2]
- **`drive download` takes `<driveId> <itemId>`** and fetches `/drives/{driveId}/items/{itemId}/content`; the old 3-argument form exits 1 (`too many arguments`). [fu:DRIVE-3]
- **Calendar `--timezone` defaults to the local time zone instead of UTC** (`cal create`, `free-busy`, `update`'s fallback, `decline --propose-*`). Wall-clock times in a zone Graph does not document (e.g. `Europe/Amsterdam`, or a POSIX `TZ` like `JST-9`) are converted to the UTC instant; `decline` no longer sends the bare zone name or a UTC-labelled wall-clock time. Event times must be ISO 8601: a date alone is local midnight, and non-ISO text (`Oct 7 9:00`) or impossible days/times (`2026-02-30`, `25:00`) exit 1 before any request. [fu:CAL-4]
- **Drained list commands print a JSON array instead of `{ "value": [...] }`** with `-o json`: `planner plans`, `my-tasks`, `tasks`, `buckets`, `bucket-tasks`, `excel worksheets`, `drive permissions`, `teams list`, `teams channels`. Scripts reading `.value` must read the array directly. [fu:PAGE-2]
- **`ms365 serve` file protocol**: output files come back as `files:[{name, contentType, size, url}]` with a one-time `GET /v1/files/{token}` URL instead of `base64`. Input files are uploaded with a raw-body `PUT /v1/uploads` and referenced as `{"$upload": "<token>"}`; multipart and `{"$file"}` are removed, and `/v1/execute` accepts JSON only (1 MB). Update `ms365-remote` and `ms365 serve` together. [fu:REMOTE-5]

### Fixed

- `ms365 -q serve` no longer hides its listening, signal and stopped lines (they always go to stderr). A SIGTERM sent right after the listening line now stops the server gracefully instead of killing it, and a repeated SIGTERM/SIGINT during shutdown (GNU `timeout` signals both the process and its group) no longer kills it before running requests finish and the socket is removed. [fu:REMOTE-4]
- The process exits once, after stdout/stderr are flushed; large `-o json` output is no longer truncated and broken pipes exit cleanly. [CORE-1]
- The codebase type-checks; `npm test` runs `tsc --noEmit` before the tests. [CORE-2]
- `schema`, `generate-skills` and `auth` commands no longer need a client ID or token cache. [CORE-3]
- dotenv no longer prints a banner to stdout. [CORE-7]
- Headless Linux logins are persisted instead of crashing on the keychain. [AUTH-1]
- 401 responses refresh the token and retry once with the new token, including on later pages. [GRAPH-5, GRAPH-7]
- Binary downloads (attachments, files, photos) are saved byte-for-byte. [GRAPH-2]
- Upload sessions resume after a failed chunk and are cancelled when they give up; upload progress respects `--quiet`. [GRAPH-3, MISC-1]
- `@odata.etag`, `@odata.count`, `@odata.deltaLink` and `@odata.type` are kept for the commands that need them and stripped only on output. [GRAPH-4]
- Paging (`--all`) uses the same request core as the first page, keeping headers, the retry policy and error messages. `--page-limit`/`--page-delay` work on every list command. [GRAPH-5, PAGE-1]
- `-o yaml` quotes values that YAML would misread, including YAML 1.1 forms such as date-only strings, `0b…`, digits with `_`, `=` and `<<`; table cells escape line breaks; CSV keeps negative numbers. [FMT-1]
- Downloads with `--out` and no `--force` work on filesystems without hard links (vfat/exFAT, some FUSE mounts); they still never overwrite an existing file. [GRAPH-2]
- `MS365_CLI_LOG_LEVEL` / `MS365_CLI_LOG_FILE` set in `.env` take effect.
- With `MS365_CLI_TOKEN` set, Graph commands no longer need a writable `~/.config/ms365-cli`.
- `mail create-rule` sends the required `sequence`, needs at least one action, and `--enabled`/`--no-enabled` work; `update-rule --enabled yes` enables. `--from` matches exact sender addresses. [MAIL-2]
- `mail flag --due` sends the required start date (new `--start`). [MAIL-4]
- `mail triage` no longer fails with `InefficientFilter`; `--include-read` keeps your `--filter`. [MAIL-5]
- `mail reply --html` sends an HTML body; `mail read -o json|yaml` includes attachments; triage CSV/YAML/text show full IDs. [MAIL-7]
- Quotes in `--search` values are escaped (`mail list`, `contacts list`, `groups list`, `search people`). [MAIL-8, GROUPS-1]
- Calendar and workflow times are read in your local time zone; day and week boundaries are local, including To Do due dates in `workflow digest`. A `TZ` the runtime cannot name falls back to UTC. [WF-1, WF-3]
- `workflow meeting-prep` skips all-day and in-progress events; `focus-time` handles multi-day events and validates `--work-start`/`--work-end` (`HH:MM`); `digest` scans every To Do list; `email-to-task` uses the default list and honours `-o`; `file-announce` posts escaped HTML. [WF-3, MISC-1]
- `cal rooms`/`room-lists` call the beta endpoints they need. [CAL-1]
- `cal decline` proposes new times in `--timezone` (default local) instead of UTC and rejects `--no-notify` with a proposal or a lone `--propose-end`; `cal update --timezone` alone is rejected; `cal view` gains `--top`. [CAL-2]
- `todo --importance`/`--status` are validated; update commands with no field fail with "Nothing to update" instead of sending an empty PATCH (todo, contacts, calendar, planner). [TODO-1]
- Drive paths with `#`, `?`, `%` or spaces reach the right file (`drive ls/cat/cp`, `sp site-by-path`). [DRIVE-1]
- `excel sort-range` calls `sort/apply`; `format-range` PATCHes font and fill separately and `--font-bold` is optional. [EXCEL-1]
- `sp delta` prints `nextLink`/`deltaLink` and resumes with `--delta-link` (both the `sites/delta` and `sites/delta()` link forms). [SP-2]
- Dates that do not exist (`2026-02-30`) or are incomplete (`2026-12`) are rejected instead of rolling over to another day (`todo`, `mail flag`, `planner`, `user set-ooo`, `workflow email-to-task`).
- Planner ETag lookups go through the Graph client (API version, retry, read-only); `--dry-run` updates/deletes work without login. [PLANNER-1]
- `planner create-plan` uses `container` instead of the deprecated `owner`. [PLANNER-2]
- `onenote page` prints raw HTML; `onenote pages` pages. [ONENOTE-1]
- `search query --top` is bounded to 1–1000. [SEARCH-1]
- `chat reply` uses `replyWithQuote` (chats have no `/replies`). [TEAMS-2]
- `user download-photo` uses the real image extension and does not overwrite; `upload-photo` prints no success in dry-run. [USER-3]
- `groups distribution-lists` excludes Microsoft 365 groups. [GROUPS-1]
- `ms365 auth setup` re-prompts on unknown menu input and exits 1 when input ends. [AUTH-4, MISC-1]
- `--dry-run` write commands no longer print success lines (`Email sent successfully.`, `Contact deleted.`, …): mail, calendar accept/decline/tentative/delete, todo delete, contacts delete, `drive delete`, `planner delete-task/delete-plan/delete-bucket` and `workflow file-announce`. `mail batch-read`/`batch-move` print no summary, `drive upload` no `Uploading …` line, and `workflow meeting-prep` no "No upcoming meetings". [fu:DRY-1]
- `workflow email-to-task --dry-run` without `--list` previews the POST to `/me/todo/lists/<default-list>/tasks` with the title `[Email] <subject>` instead of failing. [fu:DRY-2]
- `cal free-busy --interval` must be an integer from 5 to 1440 (default 30, sent as a number); `abc` or `15x` used to send `null` or `15`. [fu:CAL-3]
- `user upload-photo` no longer uploads every file as `image/jpeg`: the type comes from the extension, only `.jpg`/`.jpeg`/`.png`/`.gif` are accepted, and other files exit 1 before the file is read. [fu:USER-4]

### Added

- Built-in client ID (Microsoft Graph Command Line Tools, `14d82eec-204b-4c2f-b7e8-296a70dab67e`): `ms365 auth login` works without an app registration. `ms365 auth setup` offers quick setup or a custom app; tenant-block errors point to the custom-app path. [AUTH-4, Q2]
- `MailboxSettings.ReadWrite` scope for mail rules, mailbox settings and auto-replies; "Permissions changed … run: ms365 auth login" when consent is outdated. [AUTH-3]
- `--read-only` / `MS365_CLI_READ_ONLY` now allow read-only POST actions (`/search/query`, free/busy, `findMeetingTimes`). [GRAPH-6]
- `--force` on every file-writing command (`drive download`, `download-shared`, `cat`, `mail download-attachment`, `onenote page`, `user download-photo`); `--out` on `onenote page`; `user download-photo --out` is optional (extension from the content type). `mail download-attachment --out` accepts a directory. [GRAPH-2, MAIL-1, ONENOTE-1, USER-3]
- `--conflict rename|replace|fail` on `drive upload` and `drive cp`. [GRAPH-3]
- `--top`/`--all`/`--page-limit`/`--page-delay` on every list command. [PAGE-1]
- `mail create-rule --sequence` and `--no-enabled`; `mail flag --start`; `mail triage --all`. [MAIL-2, MAIL-4, MAIL-5]
- `cal decline --timezone`. [CAL-2]
- `sp delta --delta-link`. [SP-2]
- `schema [endpoint]` with `-o`; unknown endpoints suggest close matches. [SCHEMA-1]
- `generate-skills` derives its skills from the registered commands, covering every group (including `auth`, `user`, `groups`, `chat`). [SKILLS-1]
- `workflow file-announce -o`. [WF-3]
- `ms365 serve` and `ms365-remote`: agents on other tailnet machines share one Microsoft 365 login. `ms365 serve --app-cap <name>` listens on a 0600 Unix socket (loopback TCP fallback with `--port`) behind `tailscale serve --accept-app-caps`. Each request is checked against the `Tailscale-App-Capabilities` grant (per-command capabilities such as `mail.read`, `mail.send`), runs the CLI in an isolated child process, and is written to a JSON-lines audit log. `ms365-remote` (`MS365_REMOTE_URL`) takes the same commands, uploads local input files and saves downloads locally; `MS365_REMOTE_SSH=<host>` runs commands over Tailscale SSH instead. See README "Remote Access (Tailscale)". [REMOTE-0, REMOTE-1]
- `teams send-channel-message`, `teams reply-channel-message`, `chat send` and `chat reply` take `-o` and print the created message on stdout, so callers get its ID; the confirmation stays on stderr. `schema chat.send`/`chat.reply` describe this response. [fu:OUT-1]
- `teams members` and `groups owners` take `--top`/`--all`/`--page-limit`/`--page-delay` (team-list-members and group-list-owners document `$top` and `@odata.nextLink`). [fu:PAGE-2]
- `mail folders --parent`, `mail create-folder --parent` and `mail create-rule --move-to` accept folder names (well-known, display name, nested) as well as IDs; well-known `--move-to` names are converted to the real folder ID. [fu:MAIL-10]
- `ms365 serve --concurrency <n>` (1-64, default 4), `--max-stdout-mb <n>` (default 50) and `--max-file-mb <n>` (default 25). [fu:REMOTE-4]

### Changed

- `ms365 serve` and `ms365-remote` stream file transfers to and from disk. The default file cap stays 25 MB, but `--max-file-mb` can raise it without memory growing with file size (a 100 MB round trip grew server memory by a few MB). [fu:REMOTE-5]
- `planner plans`, `my-tasks`, `tasks`, `buckets`, `bucket-tasks` (planneruser-list-plans, planneruser-list-tasks, plannerplan-list-tasks, plannerplan-list-buckets, plannerbucket-list-tasks), `excel worksheets` (workbook-list-worksheets), `drive permissions` (driveitem-list-permissions), `teams list` (user-list-joinedteams) and `teams channels` (channel-list) follow every `@odata.nextLink` and print one combined result. [fu:PAGE-2]
- Options with a documented value set reject other values before any request (`Allowed choices`): `drive share --type view|edit|embed` / `--scope anonymous|organization|users`, `excel create-chart --type` (the 74 ExcelApi 1.1 `Excel.ChartType` values) / `--series-by auto|columns|rows`, `cal create-calendar`/`update-calendar --color` (`calendarColor` values), `mail create-rule --importance low|normal|high`, and each item of `search query --entity` (also checked by `ms365 serve` and `ms365-remote`, and listed in generated skills). [fu:CHOICE-1]
- `mail flag` start/due and To Do due dates (`todo create/update --due`, `workflow email-to-task --due`) are midnight in the local time zone when Graph documents it (else UTC); the default `mail flag` start is the local date, not the UTC date. [fu:TZ-1]
- `mail attachments` selects `id,name,contentType,size,isInline,lastModifiedDateTime` and no longer downloads `contentBytes`. [fu:MAIL-9]
- `sp item` help explains that it addresses one item through the site's non-enumerable `items` relationship. [fu:SP-3]
- `ms365 schema` is generated from per-command Graph annotations and covers all 168 Graph-calling commands (was 35); the 35 hand-written entries are unchanged. `schema --list` prints one row per Graph request, and multi-request commands show `requests: [...]`. Generated SKILL.md files carry one `Endpoint details: ms365 schema <group>.<command>` line instead of one hint per endpoint. [fu:SCHEMA-2]
- One shared integer option parser; the error wording changed for `mail create-rule --sequence` (`Expected an integer >= 0.`), `serve --port` (`Expected an integer from 0 to 65535.`) and `serve --timeout` (`Expected an integer >= 1.`). Accepted ranges are unchanged. [fu:CLEAN-1]
- `ms365 --version`, `ms365-remote --version` and the generated SKILL.md front matter read the version from `package.json` (one `src/version.ts`) instead of hard-coded copies. [fu:META-2]

### Security

- `mail download-attachment` reduces the sender-supplied file name to its base name, so an attachment cannot write outside the target directory (path traversal). [MAIL-1]
- Non-idempotent requests (`POST`/`PATCH`, and uploads with `--conflict rename`) are retried only on 429, never after a timeout or 5xx, so a send cannot be delivered twice; the error says the request may have been applied. `teams`/`chat` send commands no longer exit 1 after a successful post. [GRAPH-1, GRAPH-3, TEAMS-1]
- `sp delta --delta-link` accepts only `https://graph.microsoft.com/…/sites/delta` links, so a pasted URL cannot send your bearer token to another host. [SP-2]
- `.`/`..` path segments are rejected before any request: `sp site-by-path contoso.sharepoint.com ../../me/messages` used to read the mailbox, and under `ms365 serve` a `sharepoint.read` grant could reach any Graph resource.
- `ms365 serve`: a malformed request target no longer crashes the service; download output is capped at 25 MB while the command runs (the child is killed past it); a repeated `Tailscale-App-Capabilities` header is denied; audit lines are bounded. `ms365-remote` refuses server-chosen file names that start with `.` and names profile photos locally. [REMOTE-1]
- `mail create-rule` asks for confirmation before creating a rule with no conditions (a catch-all that could delete or forward every message); `mail delete-attachment` asks too. [MAIL-2, MAIL-7]
- Token cache files are 0600 in a 0700 directory, re-applied on every start, written atomically; `device-code.txt` is no longer written. [AUTH-2]
- `@azure/msal-node` 7 (clears the `uuid` advisory). [Q8]
