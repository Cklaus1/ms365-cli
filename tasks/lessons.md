# lessons — mistake → rule (corrections/failures only)


- Line-number edits after a whole-file write hit the wrong lines twice (index.ts, auth-cache test) → re-read the exact range before any numbered edit to a file just written.
- Removing argv scanning did not stop `--body --dry-run` flipping dry-run: Commander parses program options anywhere by default → for "value can't flip a mode" claims, test the parser, not just the code that reads it.
- Spec accept commands can name flags that don't exist (`schema -o json`) → run each accept command once by hand before encoding it in a test.
- Setting `"lib": ["ES2024"]` silently dropped the DOM lib and retyped `fetch().json()` as `unknown` (3 errors in other tracks' files) → when bumping the TS target, leave `lib` at its default unless removing DOM is the intended change.
