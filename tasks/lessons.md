# Lessons

Mistake → rule. Added only after a correction or failure.

1. **Concatenated `cat -n` output renumbers lines.** Line numbers I read for `graph-client.ts` and `errors.ts` were offset by the length of the files printed before them, so spec rows initially had wrong line numbers. **Rule:** read files one at a time with `cat -n` or Read before citing `file:line`.
2. **Commander `hook('preAction', fn)` passes `(hookedCommand, actionCommand)`.** When the hook is on `program`, the first argument is always the program, so a chain built from it is just `['ms365']`. The original "skip auth for `auth`" check silently never matched. **Rule:** derive the command chain from the **second** argument, and test hook-gated behaviour end to end.
3. **An empty-string env var is not "unset".** `process.env.X ?? default` keeps `''`. **Rule:** use `||` for string env fallbacks, and test with the variable set to `''`.
