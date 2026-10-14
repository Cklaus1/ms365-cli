# Lessons

Mistake → rule. Added only after a correction or failure.

1. **Concatenated `cat -n` output renumbers lines.** Line numbers I read for `graph-client.ts` and `errors.ts` were offset by the length of the files printed before them, so spec rows initially had wrong line numbers. **Rule:** read files one at a time with `cat -n` or Read before citing `file:line`.
