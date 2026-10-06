# lessons — mistake → rule (corrections/failures only)


- Line-number edits after a whole-file write hit the wrong lines twice (index.ts, auth-cache test) → re-read the exact range before any numbered edit to a file just written.
- Removing argv scanning did not stop `--body --dry-run` flipping dry-run: Commander parses program options anywhere by default → for "value can't flip a mode" claims, test the parser, not just the code that reads it.
- Spec accept commands can name flags that don't exist (`schema -o json`) → run each accept command once by hand before encoding it in a test.
- Setting `"lib": ["ES2024"]` silently dropped the DOM lib and retyped `fetch().json()` as `unknown` (3 errors in other tracks' files) → when bumping the TS target, leave `lib` at its default unless removing DOM is the intended change.
- Percent-encoding each path segment did not stop `..` traversal: WHATWG URL (fetch) resolves `.`/`..`/`%2e` segments after the dry-run preview is printed → guard at the single URL builder, and test the URL the transport actually sends, not the string the CLI built.
- Node joins repeated custom headers with ", ", so an `Array.isArray(header)` check was dead in the server and only its unit test exercised it → test header rules over a raw socket request; read `req.headersDistinct`.
- A subagent's relative-path edits landed in the main checkout instead of its worktree → give worktree agents absolute paths and check `git status` in main before merging.
- `import './env.js'` first does not order side effects once tsup splits shared chunks → make singletons read env lazily instead of relying on import order.
