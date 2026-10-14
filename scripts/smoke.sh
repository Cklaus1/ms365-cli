#!/usr/bin/env bash
# Offline smoke suite for ms365-cli (build-loop smoke/integration gate).
# Each check has a concrete signal (exit code + output match). Checks whose
# feature doesn't exist yet report SKIP; the final stop condition requires 0 SKIP.
# Usage: scripts/smoke.sh            (builds first)
#        SMOKE_NO_BUILD=1 scripts/smoke.sh
set -u
cd "$(dirname "$0")/.."
[ -z "${SMOKE_NO_BUILD:-}" ] && npm run -s build >/dev/null 2>&1
CLI="node dist/index.js"
TMP="$(mktemp -d)"; trap 'rm -rf "$TMP"' EXIT
pass=0; fail=0; skip=0

check() { # name expected_exit cmd... ; output in $TMP/out, $TMP/err
  local name="$1" want="$2"; shift 2
  timeout 30 "$@" >"$TMP/out" 2>"$TMP/err"; local got=$?
  if [ "$got" = 124 ]; then echo "FAIL  $name (hang/timeout)"; fail=$((fail+1)); return 1; fi
  if [ "$got" != "$want" ]; then echo "FAIL  $name (exit $got, want $want): $(head -c 200 "$TMP/err" "$TMP/out" | tr '\n' ' ')"; fail=$((fail+1)); return 1; fi
  return 0
}
ok()   { echo "PASS  $1"; pass=$((pass+1)); }
bad()  { echo "FAIL  $1"; fail=$((fail+1)); }
skipc(){ echo "SKIP  $1"; skip=$((skip+1)); }
has_cmd() { $CLI --help 2>/dev/null | grep -qE "^  $1( |\b)"; }

# 1. version
check "version" 0 $CLI --version && { grep -qx "0.1.0" "$TMP/out" && ok "version" || bad "version output"; }

# 2. schema --list works offline (no auth)
check "schema --list" 0 env MS365_CLI_CLIENT_ID=  $CLI schema --list && ok "schema --list"

# 3. dry-run send exits 0 and prints the request
check "dry-run mail send" 0 $CLI --dry-run mail send --to a@b.com --subject s --body b && {
  # F1: the dry-run trace goes to stderr; stdout carries only the command's output
  if grep -q "\[dry-run\]" "$TMP/out"; then bad "dry-run trace leaked to stdout";
  elif grep -q "\[dry-run\] POST https://graph.microsoft.com/v1.0/me/sendMail" "$TMP/err"; then ok "dry-run mail send";
  else bad "dry-run mail send trace"; fi; }

# 4. dry-run large upload targets the correct upload-session URL
head -c 12582912 /dev/zero > "$TMP/big.bin"
check "dry-run drive cp 12MB" 0 $CLI --dry-run drive cp "$TMP/big.bin" /smoke/big.bin && {
  if grep -q "content/createUploadSession" "$TMP/err"; then bad "upload session URL (has /content/)";
  elif grep -q ":/createUploadSession" "$TMP/err"; then ok "dry-run drive cp 12MB";
  else bad "upload session URL not shown"; fi; }

# 5. auth failure maps to exit code 3 with an actionable hint (no valid login available offline)
HOME_SAVE="$HOME"; export HOME="$TMP/home"; mkdir -p "$HOME"
check "auth failure exit 3" 3 env MS365_CLI_TOKEN=  MS365_CLI_CLIENT_ID=00000000-0000-0000-0000-000000000000 $CLI --read-only user me && {
  grep -qi "auth login" "$TMP/err" && ok "auth failure exit 3" || bad "auth failure hint"; }

# 6. JSON error format
check "json error" 3 env MS365_CLI_TOKEN= MS365_CLI_CLIENT_ID=00000000-0000-0000-0000-000000000000 $CLI --error-format json --read-only user me && {
  node -e 'const e=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8").trim().split("\n").pop()); if(e.error.exitCode!==3) process.exit(1)' "$TMP/err" && ok "json error" || bad "json error shape"; }
export HOME="$HOME_SAVE"

# 7. describe --json catalog
if has_cmd describe; then
  check "describe --json" 0 $CLI describe --json && {
    node -e 'const c=JSON.parse(require("fs").readFileSync(process.argv[1],"utf8")); if(!Array.isArray(c.commands)||c.commands.length<100) process.exit(1)' "$TMP/out" && ok "describe --json" || bad "describe --json shape"; }
else skipc "describe --json"; fi

# 8. MCP server answers tools/list
if has_cmd mcp; then
  printf '%s\n' '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"smoke","version":"0"}}}' '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' > "$TMP/mcp.in"
  timeout 20 $CLI mcp serve < "$TMP/mcp.in" > "$TMP/out" 2>"$TMP/err"
  node -e 'const l=require("fs").readFileSync(process.argv[1],"utf8").trim().split("\n").map(JSON.parse); const r=l.find(m=>m.id===2); if(!r||!r.result||r.result.tools.length<50) process.exit(1)' "$TMP/out" && ok "mcp tools/list" || bad "mcp tools/list"
else skipc "mcp tools/list"; fi

# 9. confirmation refused (non-TTY) exits 7
check "confirm refused exit 7" 7 env MS365_CLI_TOKEN=dummy $CLI todo delete L T < /dev/null && ok "confirm refused exit 7"

echo "smoke: pass=$pass fail=$fail skip=$skip"
[ "$fail" = 0 ]
