# todo: agent-host integration build (derived from commits)

## Tier G2 (Gate-2 subset)
- [ ] AHT-1 ceiling: `type: host`, catalog `uses`, preAction check, cursor binding, MCP + describe surface [critical path]
- [ ] AHT-2 external token source, `.env` exclusion, JWT checks, auth refusals, `childEnv()` [critical path]
- [ ] AHT-3 immutable ids, `--refs` `_ref`, `complete`/`_meta`, `--cursor` help
- [ ] AHT-4 `watch --checkpoint`, gap/backfill, window roll, change-line fields, `--select`
- [ ] AHT-5 `auth status -o json`, `auth scopes` renames, `health`, retry budget env, `retryAfterMs`
- [ ] AHT-6a offline conformance pack (smoke host-mode checks, dist sha256)

## Tier G3
- [ ] AHT-7 plan/apply core, If-Match, `--bcc`, `--body-file`, sha256 attachments
- [ ] AHT-8 honest send, correlation, `outcome_unknown`, `mail reconcile`
- [ ] AHT-9 calendar plans: series/type checks, `instances`, `cancel`, `notifies`/`scope` (TTY part)
- [ ] AHT-10 MCP buffer gap / resume / body via file; `MS365_CLI_TOKEN_COMMAND`; `auth login --start/--wait`
- [ ] AHT-11 `--mailbox` delegate reads; fold `mail shared-list`/`shared-get`
- [ ] AHT-13 docs: agent-host guide, mapping tables, skills, snapshot

## needs-human (excluded)
- [ ] AHT-6b live read evidence (Chief MSC-H3)
- [ ] AHT-9 non-TTY apply of delete/cancel (OQ-8)
- [ ] AHT-12 subscription helpers (OQ-14)
- [ ] AHT-14 live write evidence (Gate 3)
