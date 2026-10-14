# Lessons

Mistake → rule. Added only after a correction or failure.

1. **Concatenated `cat -n` output renumbers lines.** Line numbers I read for `graph-client.ts` and `errors.ts` were offset by the length of the files printed before them, so spec rows initially had wrong line numbers. **Rule:** read files one at a time with `cat -n` or Read before citing `file:line`.
2. **Commander `hook('preAction', fn)` passes `(hookedCommand, actionCommand)`.** When the hook is on `program`, the first argument is always the program, so a chain built from it is just `['ms365']`. The original "skip auth for `auth`" check silently never matched. **Rule:** derive the command chain from the **second** argument, and test hook-gated behaviour end to end.
3. **An empty-string env var is not "unset".** `process.env.X ?? default` keeps `''`. **Rule:** use `||` for string env fallbacks, and test with the variable set to `''`.
4. **A fail-first test for path traversal really does traverse.** T10's pre-fix run of `download-attachment ../../evil.sh` wrote `/evil.sh` outside the sandbox, and the cleanup was blocked by the rm safety check. **Rule:** run traversal tests from a temp directory nested at least as deep as the traversal (for example `tmp/a/b/c` for `../../`), and assert inside that root.
5. **Parallel agents running the full suite see each other's half-written edits**, which shows up as transient failures (T02 and T05 tests "failed" mid-edit). **Rule:** agents run only their own test files, and the orchestrator runs the full suite at commit time.
