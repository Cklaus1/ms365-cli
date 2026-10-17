# Lessons

Mistake → rule. Added only after a correction or failure.

1. **Concatenated `cat -n` output renumbers lines.** Line numbers I read for `graph-client.ts` and `errors.ts` were offset by the length of the files printed before them, so spec rows initially had wrong line numbers. **Rule:** read files one at a time with `cat -n` or Read before citing `file:line`.
2. **Commander `hook('preAction', fn)` passes `(hookedCommand, actionCommand)`.** When the hook is on `program`, the first argument is always the program, so a chain built from it is just `['ms365']`. The original "skip auth for `auth`" check silently never matched. **Rule:** derive the command chain from the **second** argument, and test hook-gated behaviour end to end.
3. **An empty-string env var is not "unset".** `process.env.X ?? default` keeps `''`. **Rule:** use `||` for string env fallbacks, and test with the variable set to `''`.
4. **A fail-first test for path traversal really does traverse.** T10's pre-fix run of `download-attachment ../../evil.sh` wrote `/evil.sh` outside the sandbox, and the cleanup was blocked by the rm safety check. **Rule:** run traversal tests from a temp directory nested at least as deep as the traversal (for example `tmp/a/b/c` for `../../`), and assert inside that root.
5. **Parallel agents running the full suite see each other's half-written edits**, which shows up as transient failures (T02 and T05 tests "failed" mid-edit). **Rule:** agents run only their own test files, and the orchestrator runs the full suite at commit time.
6. **A hung native thread blocks `process.exit()`.** keytar/libsecret on WSLg (DISPLAY set, no D-Bus session) hangs in D-Bus autolaunch. A JS-level timeout returns, but the process can't exit, which is why the original code used SIGTERM. **Rule:** never load a native module that may hang in the CLI process; isolate it in a child process with `spawnSync({timeout, killSignal: 'SIGKILL'})`.
7. **Chained `&&`/`;` scripts committed despite failing tests.** I ran the T08 commit in the same command as the suite, without making the commit depend on the suite's exit status. **Rule:** run the gate first, then commit in a separate step only after reading a green result.
8. **A fail-first run executes new tests against old code, and old code doesn't honor new isolation hooks.** I stashed `src/auth.ts` to prove the T08 tests failed. The old `auth.ts` ignored `MS365_CLI_CONFIG_DIR`, so the logout test deleted the developer's real `~/.config/ms365-cli/token-cache.json` and `selected-account.json`. They were already expired, but that's luck, not design. **Rule:** isolate the environment at the test-runner level (vitest `setupFiles` sets `HOME` and the config dir to a temp sandbox) so that *any* revision of the code under test is sandboxed. Never rely on a variable the new code introduces.
9. **Stashing for a fail-first check is a git state change.** A FIN agent ran `git stash push src/errors.ts` to prove a test failed, then popped it straight back. No harm was done, but in a shared tree with parallel agents a stash can capture another track's uncommitted work. **Rule:** prove fail-first by reverting the change in a scratch copy or by writing the test before the code. Never stash in a shared working tree.
10. **Static grep guards find real bugs in their first run.** QP-6's lint test found a `contacts folders` command that didn't page and an unquoted `users('${user}')` filter. Neither was in any spec. **Rule:** when a spec fixes a class of bug, land a guard test for the class, not just fixes for the listed instances.

## 11. Timing bounds in tests must hold under full-suite load
**Mistake:** the hung-keychain test asserted < 2.5 s; it passed alone 5/5 but failed once in the parallel full suite.
**Rule:** a wall-clock bound only has to separate "killed at the timeout" from "waited for the hang" — set it near the midpoint (stub hangs 30 s → bound 8 s), never just above the timeout.

## 12. Fixtures written from the spec share the spec's blind spots — smoke against one real record
**Mistake:** every audit-trail fixture put the calling app in `AppId`; real Exchange records put the resource (Microsoft Graph) there and the caller in `ClientAppId`. 1,781 tests passed while `--app` matched nothing on a real export.
**Rule:** when a feature parses an external record format, run it once against a real sample (or a fixture copied field-for-field from one, content scrubbed) before calling the task done.

## 13. A request-shape change must re-run the smoke check that asserts the old shape
**Mistake:** AHT-8 moved `mail send` to the draft path; the suite was updated, but `scripts/smoke.sh` still asserted `POST /me/sendMail` and stayed red for five tasks because smoke ran only once per tier.
**Rule:** when a task changes the requests a command sends, grep `scripts/smoke.sh` for that command and run smoke in that task's gate, whatever the tier cadence says.

## 14. "Nothing you call changed" must be checked field by field against what the peer parses
**Mistake:** I told BuildMsc AHT-11 changed nothing it calls; `describe --json` flat `scopes` for `mail list|get` gained `Mail.Read.Shared`, and Chief's scope check (which read the flat list) would have refused every session.
**Rule:** before announcing a cross-spec change as no-op, list every output field the peer reads (its parser, not the spec prose) and diff those fields on the built CLI.

## 15. A signed file's self-description is not a capability
**Mistake:** `apply` took the scope ceiling and the delete/cancel confirmation from the plan's `command` label, but the digest is computed over the file, so whoever writes the file chooses the label.
**Rule:** derive every safety decision from the bytes that will be acted on (the requests), and treat labels as claims to be checked against them.

## 16. A behaviour change rewrites its own help text, and help text is contract bytes
**Mistake:** I told ChiefFollowups AHT-9 would not change any fixture bytes; documenting the new non-TTY rule in `apply`'s description changed `describe.json` (the contract check caught it at once).
**Rule:** before promising a peer "no contract change", run `npm run contract:check` on the finished change; description strings are in `describe --json`.

## 17. Refusal tests cannot fail first when the old code refuses everything — mutate the new guards instead
**Mistake (near miss):** relaxing "never without a TTY" to "only when …", the mislabel / no-identity / dry-run refusal tests passed on the old code, so a red-then-green run proved nothing about them.
**Rule:** when a change relaxes a blanket refusal, prove each new guard by disabling it (one at a time) and watching its test fail; record that mutation run as the fail-first evidence.
