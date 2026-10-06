# opportunities — deferred work (no task, no failing test)

| # | Item | Proposed severity | Source |
|---|---|---|---|
| 1 | REMOTE-2: move `ms365 serve` behind a Tailscale Service (`svc:ms365`); first live-test that `Tailscale-App-Capabilities` arrives for `svc:` grants | P3 | spec §7 (later) |
| 2 | REMOTE-3: per-capability approval queue (mail.send, deletes) and per-agent rate limits | P3 | spec §7 (later) |
| 3 | REMOTE file transfer streaming beyond the 25 MB per-request cap | P3 | spec REMOTE-1c |
| 4 | `mail attachments` lists without `$select`, so Graph returns `contentBytes` for every attachment; add `$select=id,name,contentType,size,isInline,lastModifiedDateTime` | P3 | MailSlice |
| 5 | `create-rule --move-to`, `folders --parent`, `create-folder --parent` take IDs only; could use `resolveFolder` (well-known names need converting to real IDs for `moveToFolder`) | P3 | MailSlice |
| 6 | `messageRules` `$top` support undocumented; `mail rules` sends `$top` only when passed | P3 | MailSlice |
| 7 | `workflow email-to-task --dry-run` without `--list` fails ("No default To Do list found"); fall back to a placeholder list ID | P3 | CalWfSlice |
| 8 | `calendar free-busy --interval` not validated (NaN → `null` in body) | P3 | CalWfSlice |
| 9 | `calendar create/update --timezone` defaults to UTC while `decline` defaults to local; consider local everywhere | P3 | CalWfSlice |
| 10 | `calendar create-calendar/update-calendar --color` free-form; use `.choices()` | P3 | CalWfSlice |
| 11 | To Do `dueDateTime` written as `<date>T00:00:00` UTC (todo create/update, email-to-task); other clients may show it shifted | P3 | CalWfSlice |
| 12 | `sp item` (`/sites/{id}/items/{id}`) is not a documented v1.0 endpoint either; consider removing with SP-1 | P3 | FilesSlice |
| 13 | Planner lists, excel worksheets, drive permissions stay single-request; confirm `$top`/`nextLink` support before adding `paginationOptions` | P3 | FilesSlice |
| 14 | `drive download <driveId> <itemId> <childId>` uses `/items/{itemId}/children/{childId}/content`; `/items/{childId}/content` would suffice (interface change) | P3 | FilesSlice |
| 15 | Export `integer()` option parser from `list.ts` (search.ts has a local 1-1000 parser); shared JSON `field()` narrowing helper (drive.ts, planner.ts, graph-client.ts each have one) | P3 | FilesSlice |
| 16 | `groups owners`, `teams list/members/channels` single request, no pagination | P3 | OrgSlice |
| 17 | `user upload-photo` has its own mime map; could use `src/mime.ts` and reject non-images | P3 | OrgSlice |
| 18 | Teams send commands print nothing on stdout, so callers can't get the created message id; add `-o` | P3 | OrgSlice |
| 19 | `.version('0.1.0')` hard-coded in `index.ts`; read from package.json | P3 | MetaSlice |
| 20 | Free-text option values (`calendar --color`, `drive share --type/--scope`, `user set-ooo --status`, `search --entity`) lack `.choices()`; REMOTE-1 validation and generated skills would then enforce/document them | P3 | MetaSlice |
| 21 | In `--dry-run`, `mail send/reply` and others still print success messages ("Email sent successfully."); only teams and `user set-presence` check `status !== 0` | P2 | DocsTrack |
| 22 | Only 35 of 175 commands have a `schema` ENDPOINTS entry; generate the catalog or annotate per command to stop drift | P3 | DocsTrack |
| 23 | `/v1/execute` buffers downloads as base64 in memory (25 MB → ~34 MB JSON); stream instead | P3 | RemoteTrack (same as #3) |
| 24 | `ms365 -q serve` also hides the "listening on" line (goes through `info()`) | P3 | RemoteTrack |
| 25 | `serve` concurrency (4), stdout cap (50 MB) and file cap (25 MB) are constants, not flags | P3 | RemoteTrack |
