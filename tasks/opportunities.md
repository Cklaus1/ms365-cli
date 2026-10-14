# Opportunities (deferred work)

Things discovered mid-build that have no assigned task. Each entry gives a proposed severity, and none of them are acted on during this run.

- **[P3] `--dry-run mail send` prints "Email sent successfully."** Found in the Step 0 smoke run. The dry-run message misleads both agents and people. Proposal: print "(dry-run) not sent" instead.
- **[P3] Warm in-process MCP server** (alternative to D16). Spawning the CLI per call costs process start-up each time. Proposal: an embedded mode where `handleErrors` doesn't exit.
- **[P3] `drive copy --wait`** (D20). There's no copy command. Add `drive copy`, which returns a monitor URL, and poll it when `--wait` is given.
- **[P3] MCP notifications on delta changes** (§6 event-driven, "later").
